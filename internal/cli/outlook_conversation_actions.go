package cli

import (
	"fmt"
	"time"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/conversation"
)

func newConversationActionCmd(flags *rootFlags, kind string) *cobra.Command {
	var folderID, destination, through, completed string
	var execute bool
	cmd := &cobra.Command{Use: kind + " <conversation-id>", Short: "Preview or " + kind + " a whole conversation in a folder", Long: "Acts on the whole conversation within the specified context folder. Preview is the default; --execute applies the action once. Read-state actions suppress read receipts. A write with an unknown outcome must be reconciled before retrying.", Annotations: map[string]string{"pp:data-source": "live"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "threads " + kind})
		}
		id, err := oneID(args)
		if err != nil {
			return err
		}
		if folderID == "" {
			return usageErr(fmt.Errorf("--folder-id is required"))
		}
		if kind != "move" {
			if _, err := time.Parse(time.RFC3339Nano, through); err != nil {
				return usageErr(fmt.Errorf("--through requires an RFC3339 timestamp"))
			}
		}
		if kind == "complete" {
			if _, err := time.Parse(time.RFC3339Nano, completed); err != nil {
				return usageErr(fmt.Errorf("--completed-at requires an RFC3339 timestamp"))
			}
		}
		if kind == "move" && (destination == "" || destination == folderID) {
			return usageErr(fmt.Errorf("--destination-folder-id must identify a different folder"))
		}
		p := conversation.Params{ConversationID: id, FolderID: folderID, DestinationID: destination, Through: through, CompletedAt: completed}
		preview := map[string]any{"preview": true, "operation": "threads " + kind, "conversation_id": id, "folder_id": folderID, "scope": "whole conversation in context folder"}
		if kind != "move" {
			preview["through"] = through
		}
		if kind == "mark-read" || kind == "mark-unread" {
			preview["suppress_read_receipts"] = true
		}
		if kind == "complete" {
			preview["completed_at"] = completed
		}
		if kind == "move" {
			preview["destination_folder_id"] = destination
		}
		if !execute || flags.dryRun {
			return flags.printJSON(cmd, preview)
		}
		if err := mutationInHarness(); err != nil {
			return err
		}
		prefix := map[string]string{"mark-read": "OUTLOOK_WEB_THREAD_MARK_READ", "mark-unread": "OUTLOOK_WEB_THREAD_MARK_UNREAD", "flag": "OUTLOOK_WEB_THREAD_FLAG", "complete": "OUTLOOK_WEB_THREAD_COMPLETE", "move": "OUTLOOK_WEB_THREAD_MOVE", "restore": "OUTLOOK_WEB_THREAD_RESTORE"}[kind]
		detail, _, err := outlookCapture(prefix, false)
		if err != nil {
			return err
		}
		current, _, err := currentSession()
		if err != nil {
			return err
		}
		session, op, err := conversation.Prepare(kind, detail, current, p)
		if err != nil {
			return err
		}
		if _, err = outlookOperation(cmd.Context(), session, op); err != nil {
			return err
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "accepted": true, "operation": "threads " + kind, "conversation_id": id, "folder_id": folderID})
	}}
	cmd.Flags().StringVar(&folderID, "folder-id", "", "Context folder containing the conversation")
	if kind != "move" {
		cmd.Flags().StringVar(&through, "through", "", "Observed conversation sync time as RFC3339")
	}
	if kind == "move" {
		cmd.Flags().StringVar(&destination, "destination-folder-id", "", "Destination folder ID")
	}
	if kind == "complete" {
		cmd.Flags().StringVar(&completed, "completed-at", "", "Completion time as RFC3339")
	}
	cmd.Flags().BoolVar(&execute, "execute", false, "Apply the previewed action once")
	return cmd
}
