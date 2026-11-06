package cli

import (
	"fmt"
	"path/filepath"
	"time"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/capture"
)

func newOutlookImportCmd(flags *rootFlags) *cobra.Command {
	var execute bool
	cmd := &cobra.Command{Use: "import <export-id>", Short: "Preview or activate a saved Outlook session export", Long: "Imports an existing owner-only Outlook FindFolder export produced by the local m365-session helper. Preview validates its local shape, expiry and compatibility with the current private operation bundle. With --execute, the helper verifies an authenticated read and expected identity before atomically activating it. This is not a general cookie importer or multi-account switcher.", RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "auth import"})
		}
		id, err := oneID(args)
		if err != nil {
			return err
		}
		if !outlookExportID.MatchString(id) {
			return usageErr(fmt.Errorf("export ID is not in the saved session format"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "auth import"})
		}
		pointer, err := outlookPointerPath()
		if err != nil {
			return err
		}
		if _, err = safeOutlookPointer(pointer); err != nil {
			return err
		}
		session, _, err := capture.LoadFindFolder(filepath.Join(filepath.Dir(pointer), "session-outlook-"+id+".json"))
		if err != nil {
			return fmt.Errorf("saved Outlook export is unavailable, expired, or invalid")
		}
		_, _, folderErr := capture.LoadRuntimeWeb("folders_list", session)
		_, _, mailErr := capture.PrepareRuntimeRead("mail_list", session)
		compatible := folderErr == nil && mailErr == nil
		if !execute {
			return flags.printJSON(cmd, map[string]any{"operation": "auth import", "preview": true, "session_available": true, "expires_at": session.ExpiresAt.UTC().Format(time.RFC3339), "templates_compatible": compatible, "network_checked": false})
		}
		if !compatible {
			return fmt.Errorf("saved export does not match the current private operation bundle; no changes made")
		}
		if err = mutationInHarness(); err != nil {
			return err
		}
		helper, err := findSessionHelper()
		if err != nil {
			return err
		}
		if _, err = runSessionHelper(helper, "activate", id); err != nil {
			return fmt.Errorf("session activation did not complete; inspect auth status before retrying: %w", err)
		}
		active, _, err := currentSession()
		if err != nil || active.AccountKey != session.AccountKey {
			return fmt.Errorf("session activation was reported but the active account could not be confirmed; inspect auth status")
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "operation": "auth import", "activated": true, "expires_at": active.ExpiresAt.UTC().Format(time.RFC3339), "templates_compatible": true, "jwt_signature_verified": false})
	}}
	cmd.Flags().BoolVar(&execute, "execute", false, "Validate and activate this saved session export")
	return cmd
}
