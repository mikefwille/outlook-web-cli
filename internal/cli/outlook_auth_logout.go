package cli

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"syscall"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/capture"
)

var outlookExportID = regexp.MustCompile(`^[a-z0-9]+-[a-f0-9]{8}$`)

func outlookPointerPath() (string, error) {
	if os.Getenv("OUTLOOK_WEB_SESSION_FILE") != "" {
		return "", fmt.Errorf("unset OUTLOOK_WEB_SESSION_FILE before deactivating the shared session")
	}
	if path := os.Getenv("OUTLOOK_WEB_SESSION_POINTER"); path != "" {
		return filepath.Clean(path), nil
	}
	home, err := os.UserHomeDir()
	if err != nil {
		return "", err
	}
	return filepath.Join(home, ".local", "share", "outlook-web-dev", "private", "current-outlook.json"), nil
}
func safeOutlookPointer(path string) (bool, error) {
	if filepath.Base(path) != "current-outlook.json" {
		return false, fmt.Errorf("Outlook session pointer path is not recognized")
	}
	dir, err := os.Lstat(filepath.Dir(path))
	if err != nil || !dir.IsDir() || dir.Mode()&os.ModeSymlink != 0 || dir.Mode().Perm()&0077 != 0 {
		return false, fmt.Errorf("private Outlook session directory is unavailable or unsafe")
	}
	owner, ok := dir.Sys().(*syscall.Stat_t)
	if !ok || owner.Uid != uint32(os.Getuid()) {
		return false, fmt.Errorf("private Outlook session directory has the wrong owner")
	}
	info, err := os.Lstat(path)
	if os.IsNotExist(err) {
		return false, nil
	}
	if err != nil || !info.Mode().IsRegular() || info.Mode().Perm()&0077 != 0 {
		return false, fmt.Errorf("Outlook session pointer is unavailable or unsafe")
	}
	data, err := capture.ReadPrivateEvidence(path)
	if err != nil {
		return false, fmt.Errorf("Outlook session pointer is unavailable or unsafe")
	}
	var pointer struct {
		Version  int    `json:"version"`
		Service  string `json:"service"`
		ExportID string `json:"export_id"`
	}
	if json.Unmarshal(data, &pointer) != nil || pointer.Version != 1 || pointer.Service != "outlook" || !outlookExportID.MatchString(pointer.ExportID) {
		return false, fmt.Errorf("Outlook session pointer is not recognized")
	}
	after, err := os.Lstat(path)
	if err != nil || !os.SameFile(info, after) {
		return false, fmt.Errorf("Outlook session pointer changed while checking")
	}
	return true, nil
}
func newOutlookLogoutCmd(flags *rootFlags) *cobra.Command {
	var execute bool
	cmd := &cobra.Command{Use: "logout", Short: "Preview or deactivate the local Outlook session", Long: "Removes only the active Outlook session pointer after an explicit --execute. It does not revoke Microsoft tokens, delete saved exports, sign out the browser profile, or affect the Teams pointer.", RunE: func(cmd *cobra.Command, args []string) error {
		if len(args) != 0 {
			return usageErr(fmt.Errorf("auth logout takes no arguments"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "auth logout", "state_change": false})
		}
		path, err := outlookPointerPath()
		if err != nil {
			return err
		}
		if !execute {
			present, err := safeOutlookPointer(path)
			if err != nil {
				return err
			}
			return flags.printJSON(cmd, map[string]any{"preview": true, "operation": "auth logout", "session_present": present, "browser_profile_retained": true, "tokens_revoked": false})
		}
		if err := mutationInHarness(); err != nil {
			return err
		}
		if _, err := safeOutlookPointer(path); err != nil {
			return err
		}
		// The shared helper uses this same cooperative lock during pointer
		// publication. It prevents a simultaneous supported refresh.
		lock := filepath.Join(filepath.Dir(path), ".m365-session-outlook.lock")
		if err := os.Mkdir(lock, 0700); err != nil {
			return fmt.Errorf("Outlook session is locked or its private directory is unavailable")
		}
		defer os.Remove(lock)
		present, err := safeOutlookPointer(path)
		if err != nil {
			return err
		}
		if present {
			if err = os.Remove(path); err != nil {
				return fmt.Errorf("cannot deactivate Outlook session pointer")
			}
		}
		return flags.printJSON(cmd, map[string]any{"operation": "auth logout", "deactivated": present, "browser_profile_retained": true, "tokens_revoked": false})
	}}
	cmd.Flags().BoolVar(&execute, "execute", false, "Deactivate the local Outlook session pointer once")
	return cmd
}
