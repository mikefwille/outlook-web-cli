package cli

import (
	"encoding/json"
	"fmt"
	"strings"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/exchange"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/mailactions"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

// pp:data-source live
func newMailMoveCmd(flags *rootFlags, kind string) *cobra.Command {
	var destinationID, key string
	var execute bool
	cmd := &cobra.Command{Use: kind + " <id>", Short: "Preview or " + kind + " one message", Long: "Moves one message rather than its conversation. Preview resolves the destination and current version. Outlook returns a replacement item ID and change key. The version check is before the write and is not atomic.", Annotations: map[string]string{"pp:data-source": "live"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": strings.TrimPrefix(cmd.CommandPath(), cmd.Root().Name()+" ")})
		}
		id, err := oneID(args)
		if err != nil {
			return err
		}
		if kind == "move" && destinationID == "" {
			return usageErr(fmt.Errorf("--folder-id is required"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": strings.TrimPrefix(cmd.CommandPath(), cmd.Root().Name()+" "), "id": id, "folder_id": destinationID})
		}
		response, err := getMessage(cmd.Context(), id)
		if err != nil {
			return err
		}
		got, raw, status := exchange.ParseMessageGet(response.Body)
		if status != transport.EnvelopeOK || got != id {
			return &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
		}
		var item messageActionState
		if json.Unmarshal(raw, &item) != nil || item.ItemID.ID != id || item.ItemID.ChangeKey == "" || item.Parent.ID == "" {
			return fmt.Errorf("message location or version is unavailable; no changes made")
		}
		var destination exchange.Folder
		switch {
		case kind == "archive":
			destination, err = resolveDistinguishedFolder(cmd.Context(), "archive")
		case kind == "restore" && destinationID == "":
			destination, err = resolveDistinguishedFolder(cmd.Context(), "inbox")
		default:
			destination, err = folderLookup(cmd.Context(), destinationID)
		}
		if err != nil {
			return err
		}
		if destination.ID.ID == "" || destination.ID.ID == item.Parent.ID || destination.FolderClass != "IPF.Note" {
			return fmt.Errorf("destination must be a different mail folder")
		}
		if kind == "restore" {
			deleted, e := resolveDistinguishedFolder(cmd.Context(), "deleteditems")
			if e != nil {
				return e
			}
			if deleted.ID.ID != item.Parent.ID {
				return fmt.Errorf("restore requires a message in Deleted Items")
			}
		}
		preview := map[string]any{"preview": true, "operation": strings.TrimPrefix(cmd.CommandPath(), cmd.Root().Name()+" "), "scope": "single message", "id": id, "subject": item.Subject, "folder_id": item.Parent.ID, "destination_id": destination.ID.ID, "destination_name": destination.DisplayName, "change_key": item.ItemID.ChangeKey, "version_check": "before write; not atomic", "returns_replacement_id": true}
		if !execute {
			return flags.printJSON(cmd, preview)
		}
		if key == "" {
			return usageErr(fmt.Errorf("--change-key from the preview is required"))
		}
		if key != item.ItemID.ChangeKey {
			return &transport.Error{Kind: transport.KindConflict}
		}
		if err := mutationInHarness(); err != nil {
			return err
		}
		detail, _, err := outlookCapture("OUTLOOK_WEB_MAIL_MOVE", false)
		if err != nil {
			return err
		}
		current, _, err := currentSession()
		if err != nil {
			return err
		}
		session, operation, err := mailactions.PrepareMove(detail, current, id, destination.ID.ID)
		if err != nil {
			return err
		}
		moved, err := outlookOperation(cmd.Context(), session, operation)
		if err != nil {
			return err
		}
		result, state := mailactions.ParseMove(moved.Body)
		if state != transport.EnvelopeOK {
			return &transport.Error{Kind: transport.KindAmbiguous, Status: moved.Status}
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "accepted": true, "operation": strings.TrimPrefix(cmd.CommandPath(), cmd.Root().Name()+" "), "scope": "single message", "original_id": id, "destination_id": destination.ID.ID, "result": result})
	}}
	if kind != "archive" {
		cmd.Flags().StringVar(&destinationID, "folder-id", "", "Destination folder ID")
	}
	cmd.Flags().StringVar(&key, "change-key", "", "Expected version from preview")
	cmd.Flags().BoolVar(&execute, "execute", false, "Apply the previewed move once")
	return cmd
}
