package cli

import (
	"context"
	"encoding/json"
	"fmt"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/exchange"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/mailactions"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

// messageLocation independently reads the requested identity and its folder.
func messageLocation(ctx context.Context, id string) (messageActionState, error) {
	response, err := getMessage(ctx, id)
	if err != nil {
		return messageActionState{}, err
	}
	got, raw, state := exchange.ParseMessageGet(response.Body)
	var item messageActionState
	if state != transport.EnvelopeOK || got != id || json.Unmarshal(raw, &item) != nil || item.ItemID.ID != id || item.ItemID.ChangeKey == "" || item.Parent.ID == "" {
		return messageActionState{}, &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
	}
	return item, nil
}

func verifyCopyReadback(ctx context.Context, id, copyID, destination string, before messageActionState, read func(context.Context, string) (messageActionState, error)) (bool, bool) {
	original, originalErr := read(ctx, id)
	copied, copyErr := read(ctx, copyID)
	return originalErr == nil && original.Parent.ID == before.Parent.ID && original.Subject == before.Subject,
		copyErr == nil && copied.Parent.ID == destination && copied.Subject == before.Subject
}

// pp:data-source live
func newMailCopyCmd(flags *rootFlags) *cobra.Command {
	var destinationID, changeKey string
	var execute bool
	cmd := &cobra.Command{Use: "copy <id>", Short: "Preview or copy one message into a mail folder", Long: "Copies exactly one selected message, not its conversation. Preview reads its current version and the destination. --execute requires the previewed change key and attempts CopyItem once. After acceptance, it independently reads the original and returned copy. The version check is before the write and is not atomic; reconcile an incomplete or unknown outcome before retrying.", Annotations: map[string]string{"pp:data-source": "live", "pp:happy-args": "id=example-id;--folder-id=example-folder;--dry-run"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "mail copy"})
		}
		id, err := oneID(args)
		if err != nil {
			return err
		}
		if destinationID == "" {
			return usageErr(fmt.Errorf("--folder-id is required"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "mail copy", "id": id, "folder_id": destinationID})
		}
		ctx, cancel := boundCtx(cmd.Context(), flags)
		defer cancel()
		source, err := messageLocation(ctx, id)
		if err != nil {
			return err
		}
		folder, err := folderLookup(ctx, destinationID)
		if err != nil {
			return err
		}
		if folder.ID.ID != destinationID || folder.FolderClass != "IPF.Note" {
			return fmt.Errorf("destination must be an accessible mail folder")
		}
		preview := map[string]any{"preview": true, "operation": "mail copy", "scope": "single message", "id": id, "subject": source.Subject, "folder_id": source.Parent.ID, "destination_id": folder.ID.ID, "destination_name": folder.DisplayName, "change_key": source.ItemID.ChangeKey, "version_check": "before write; not atomic", "returns_new_copy_id": true}
		if !execute {
			return flags.printJSON(cmd, preview)
		}
		if changeKey == "" {
			return usageErr(fmt.Errorf("--change-key from the preview is required"))
		}
		if changeKey != source.ItemID.ChangeKey {
			return &transport.Error{Kind: transport.KindConflict}
		}
		if err := mutationInHarness(); err != nil {
			return err
		}
		detail, _, err := outlookCapture("OUTLOOK_WEB_MAIL_COPY", false)
		if err != nil {
			return err
		}
		current, _, err := currentSession()
		if err != nil {
			return err
		}
		session, op, err := mailactions.PrepareCopy(detail, current, id, destinationID)
		if err != nil {
			return err
		}
		response, err := outlookOperation(ctx, session, op)
		if err != nil {
			return err
		}
		created, state := mailactions.ParseCopy(response.Body)
		if state != transport.EnvelopeOK || created.ID == id {
			return &transport.Error{Kind: transport.KindAmbiguous, Status: response.Status}
		}
		originalOK, copyOK := verifyCopyReadback(ctx, id, created.ID, destinationID, source, messageLocation)
		result := map[string]any{"source": "live", "accepted": true, "operation": "mail copy", "scope": "single message", "original_id": id, "destination_id": destinationID, "copy": created, "original_readback_verified": originalOK, "copy_readback_verified": copyOK}
		if !originalOK || !copyOK {
			result["reconcile_before_retry"] = true
			if err := printJSONFilteredKeep(cmd.OutOrStdout(), result, flags, "accepted", "original_id", "copy", "original_readback_verified", "copy_readback_verified", "reconcile_before_retry"); err != nil {
				return err
			}
			return &transport.Error{Kind: transport.KindPartial, Status: response.Status}
		}
		return printJSONFilteredKeep(cmd.OutOrStdout(), result, flags, "accepted", "original_id", "copy", "original_readback_verified", "copy_readback_verified")
	}}
	cmd.Flags().StringVar(&destinationID, "folder-id", "", "Destination mail folder ID")
	cmd.Flags().StringVar(&changeKey, "change-key", "", "Expected source version from the preview")
	cmd.Flags().BoolVar(&execute, "execute", false, "Copy this message once and read back both items")
	return cmd
}
