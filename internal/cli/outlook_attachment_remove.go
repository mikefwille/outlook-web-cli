package cli

import (
	"fmt"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/attachments"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func newDraftRemoveAttachmentCmd(flags *rootFlags) *cobra.Command {
	var key string
	var execute bool
	cmd := &cobra.Command{Use: "remove-attachment <draft-id> <attachment-id>", Short: "Preview or remove one attachment from a draft", Long: "Reads the current draft and attachment metadata. Only the selected file attachment is removed; Outlook returns a new draft change key. Inline status is reported when available. The expected version is checked before the write and included in the captured request; atomic conflict behavior has not been established.", Annotations: map[string]string{"pp:data-source": "live"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "drafts remove-attachment"})
		}
		if len(args) != 2 || args[0] == "" || args[1] == "" {
			return usageErr(fmt.Errorf("provide a draft ID and attachment ID"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "drafts remove-attachment", "draft_id": args[0], "attachment_id": args[1]})
		}
		draft, err := readDraft(cmd.Context(), args[0])
		if err != nil {
			return err
		}
		list, err := readMessageAttachments(cmd.Context(), args[0])
		if err != nil {
			return err
		}
		if draft.ChangeKey != list.ChangeKey {
			return &transport.Error{Kind: transport.KindConflict}
		}
		var selected *attachments.Metadata
		for i := range list.Attachments {
			if list.Attachments[i].ID == args[1] {
				selected = &list.Attachments[i]
				break
			}
		}
		if selected == nil {
			return fmt.Errorf("attachment is not present on the current draft")
		}
		if selected.Type != "FileAttachment:#Exchange" || selected.IsInline != nil && *selected.IsInline {
			return fmt.Errorf("only file attachments not reported as inline can be removed")
		}
		inlineStatus := "unknown"
		if selected.IsInline != nil && !*selected.IsInline {
			inlineStatus = "not_inline"
		}
		preview := map[string]any{"preview": true, "operation": "drafts remove-attachment", "draft_id": args[0], "subject": draft.Subject, "change_key": draft.ChangeKey, "attachment": selected, "inline_status": inlineStatus, "version_check": "before write; atomic conflict behavior unverified", "sends_mail": false}
		if !execute {
			return flags.printJSON(cmd, preview)
		}
		if key == "" {
			return usageErr(fmt.Errorf("--change-key from the preview is required"))
		}
		if key != draft.ChangeKey {
			return &transport.Error{Kind: transport.KindConflict}
		}
		if err := mutationInHarness(); err != nil {
			return err
		}
		detail, _, err := outlookCapture("OUTLOOK_WEB_DRAFT_REMOVE_ATTACHMENT", false)
		if err != nil {
			return err
		}
		current, _, err := currentSession()
		if err != nil {
			return err
		}
		session, operation, err := attachments.PrepareRemove(detail, current, args[0], key, args[1])
		if err != nil {
			return err
		}
		response, err := outlookOperation(cmd.Context(), session, operation)
		if err != nil {
			return err
		}
		removed, state := attachments.ParseRemoved(response.Body)
		if state != transport.EnvelopeOK || removed.RootItemID != args[0] {
			return &transport.Error{Kind: transport.KindAmbiguous, Status: response.Status}
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "removed": true, "result": removed, "sends_mail": false})
	}}
	cmd.Flags().StringVar(&key, "change-key", "", "Expected version from the preview")
	cmd.Flags().BoolVar(&execute, "execute", false, "Remove this attachment once")
	return cmd
}
