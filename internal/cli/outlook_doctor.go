package cli

import (
	"fmt"
	"time"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/capture"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

// pp:data-source live
func newOutlookDoctorCmd(flags *rootFlags) *cobra.Command {
	var live bool
	var failOn string
	cmd := &cobra.Command{Use: "doctor", Short: "Check the saved Outlook session and operation templates", Long: "Checks the account-bound session and basic operation templates locally. With --live, also performs a bounded read-only FindFolder request. It never refreshes or opens a browser.", Annotations: map[string]string{"mcp:read-only": "true"}, RunE: func(cmd *cobra.Command, args []string) error {
		if len(args) != 0 {
			return usageErr(fmt.Errorf("doctor takes no positional arguments"))
		}
		if failOn != "" && failOn != "warn" && failOn != "error" {
			return usageErr(fmt.Errorf("--fail-on must be warn or error"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "doctor", "live_check": live})
		}
		// The root gate deliberately defers doctor errors so a structured
		// platform report can be emitted. Preserve that contract before
		// inspecting any Outlook session or making a live request.
		if registeredPlatformSource != nil {
			if flags.platformSession == nil {
				return fmt.Errorf("verified client profile session is required")
			}
			report, err := platformDoctorV2Report(cmd.Context(), flags.platformSession)
			if err != nil {
				return err
			}
			if err = flags.printJSON(cmd, report); err != nil {
				return err
			}
			return flags.platformGateError
		}
		current, operation, err := currentSession()
		if err != nil {
			return fmt.Errorf("saved Outlook session unavailable; use the session helper to refresh: %w", err)
		}
		_, _, folderErr := capture.LoadRuntimeWeb("folders_list", current)
		_, _, mailErr := capture.PrepareRuntimeRead("mail_list", current)
		state := "ok"
		if folderErr != nil || mailErr != nil {
			state = "warning"
		}
		result := map[string]any{"source": "local", "state": state, "session": "available", "expires_at": current.ExpiresAt.UTC().Format(time.RFC3339), "folder_template_available": folderErr == nil, "mail_template_available": mailErr == nil, "live_checked": false, "attachment_audience_checked": false}
		if live {
			client, err := transport.New(current.AccountKey, current, nil)
			if err != nil {
				return err
			}
			if _, err = client.Do(cmd.Context(), current.AccountKey, operation); err != nil {
				return fmt.Errorf("Outlook live session check failed: %w", err)
			}
			result["source"] = "live"
			result["live_checked"] = true
			result["live_result"] = "success"
		}
		if err = flags.printJSON(cmd, result); err != nil {
			return err
		}
		if state == "warning" && failOn == "warn" {
			return fmt.Errorf("one or more operation templates are unavailable")
		}
		return nil
	}}
	cmd.Flags().BoolVar(&live, "live", false, "Make one read-only Outlook session check")
	cmd.Flags().StringVar(&failOn, "fail-on", "", "Exit nonzero on warn or error")
	return cmd
}
