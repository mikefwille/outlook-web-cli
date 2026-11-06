package cli

import (
	"encoding/json"
	"fmt"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"regexp"
	"time"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/capture"
)

// The helper owns its browser lock and bounded recovery/cleanup lifecycle.
// Keep it outside MCP and don't automatically start it for ordinary reads.
type recoverResult struct {
	Version          int    `json:"m365_session"`
	Action           string `json:"action"`
	Service          string `json:"service"`
	OK               bool   `json:"ok"`
	Status           string `json:"status"`
	IdentityVerified bool   `json:"identity_verified"`
	Browser          string `json:"recovery_browser"`
	Reason           string `json:"reason"`
	Stage            string `json:"stage"`
	Secondary        struct {
		Attachments struct {
			Status string `json:"status"`
		} `json:"attachments"`
	} `json:"secondary_audiences"`
}

var safeRecoveryCode = regexp.MustCompile(`^[a-z0-9_]{1,80}$`)

type boundedHelperOutput struct {
	data     []byte
	overflow bool
}

func (b *boundedHelperOutput) Write(p []byte) (int, error) {
	n := len(p)
	if len(b.data)+n > 32768 {
		b.overflow = true
		return n, nil
	}
	b.data = append(b.data, p...)
	return n, nil
}

func parseRecoverResult(data []byte) (recoverResult, error) {
	return parseSessionHelperResult(data, "recover")
}

func parseSessionHelperResult(data []byte, action string) (recoverResult, error) {
	var r recoverResult
	if json.Unmarshal(data, &r) != nil || r.Version != 1 || r.Action != action || r.Service != "outlook" {
		return r, fmt.Errorf("session helper returned an invalid response")
	}
	if !r.OK {
		if !safeRecoveryCode.MatchString(r.Reason) {
			r.Reason = "unknown"
		}
		if !safeRecoveryCode.MatchString(r.Stage) {
			r.Stage = "unknown"
		}
		return r, fmt.Errorf("session recovery failed: %s (stage: %s)", r.Reason, r.Stage)
	}
	if r.Status != "service_read_ok" || !r.IdentityVerified || (action == "recover" && r.Browser != "started_and_stopped" && r.Browser != "reused_existing_owner") {
		return r, fmt.Errorf("session helper did not verify the recovered Outlook account")
	}
	switch r.Secondary.Attachments.Status {
	case "captured_from_cache", "renewed_in_browser", "not_captured", "not_renewed":
	default:
		r.Secondary.Attachments.Status = "unknown"
	}
	return r, nil
}

func findSessionHelper() (string, error) {
	if path, err := exec.LookPath("m365-session"); err == nil {
		return path, nil
	}
	executable, err := os.Executable()
	if err == nil {
		candidate := filepath.Join(filepath.Dir(executable), "m365-session")
		if info, err := os.Stat(candidate); err == nil && info.Mode().IsRegular() && info.Mode().Perm()&0111 != 0 {
			return candidate, nil
		}
	}
	return "", fmt.Errorf("m365-session helper not found beside the CLI or on PATH")
}

func runSessionRecovery(helper string) (recoverResult, error) {
	return runSessionHelper(helper, "recover", "")
}

func runSessionHelper(helper, action, exportID string) (recoverResult, error) {
	// The helper enforces its own deadline and stops only the browser owner it
	// started. Do not kill it mid-cleanup on a shorter outer timeout.
	args := []string{action, "outlook"}
	if exportID != "" {
		args = append(args, exportID)
	}
	command := exec.Command(helper, args...)
	var output boundedHelperOutput
	command.Stdout = &output
	command.Stderr = io.Discard
	runErr := command.Run()
	if output.overflow {
		return recoverResult{}, fmt.Errorf("session helper response exceeded the limit")
	}
	result, err := parseSessionHelperResult(output.data, action)
	if err != nil {
		return result, err
	}
	if runErr != nil {
		return result, fmt.Errorf("session helper failed after reporting success")
	}
	return result, nil
}

func newOutlookRecoverCmd(flags *rootFlags) *cobra.Command {
	var execute bool
	cmd := &cobra.Command{Use: "recover", Short: "Recover the saved Outlook web session using the local helper", Long: "Preview session recovery by default. With --execute, the local m365-session helper may start the saved browser profile, sign in through recognized prompts, verify an Outlook read, and stop an owner it started. Existing browser state may be changed. This command is not exposed through MCP.", Args: cobra.NoArgs,
		RunE: func(cmd *cobra.Command, _ []string) error {
			if !execute || flags.dryRun {
				return flags.printJSON(cmd, map[string]any{"operation": "auth " + cmd.Name(), "preview": true, "helper": "m365-session recover outlook", "may_start_browser": true, "state_change": false})
			}
			helper, err := findSessionHelper()
			if err != nil {
				return err
			}
			result, err := runSessionRecovery(helper)
			if err != nil {
				return err
			}
			session, _, err := currentSession()
			if err != nil {
				return fmt.Errorf("helper reported recovery, but the current session could not be loaded")
			}
			_, _, folderErr := capture.LoadRuntimeWeb("folders_list", session)
			_, _, mailErr := capture.PrepareRuntimeRead("mail_list", session)
			return flags.printJSON(cmd, map[string]any{"operation": "auth " + cmd.Name(), "state": "available", "recovery_browser": result.Browser, "expires_at": session.ExpiresAt.UTC().Format(time.RFC3339), "templates_compatible": folderErr == nil && mailErr == nil, "attachment_audience": result.Secondary.Attachments.Status, "jwt_signature_verified": false})
		},
	}
	cmd.Flags().BoolVar(&execute, "execute", false, "Run the explicit local session recovery")
	return cmd
}
