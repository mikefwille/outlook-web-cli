package cli

import (
	"context"
	"fmt"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/attachments"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/exchange"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func init() {
	registerNovelCommand(func(root *cobra.Command, flags *rootFlags) {
		group := &cobra.Command{Use: "attachments", Short: "Inspect message attachments"}
		group.AddCommand(&cobra.Command{Use: "list <message-id>", Short: "List the attachments on a message or draft", Long: "Reads attachment metadata from the current message. Reported size is Outlook's value and is not a verified downloaded byte count.", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live"}, RunE: func(cmd *cobra.Command, args []string) error {
			if flags.dryRun && requiredInputWaived() {
				return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "attachments list"})
			}
			id, err := oneID(args)
			if err != nil {
				return err
			}
			if flags.dryRun {
				return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "attachments list", "message_id": id})
			}
			v, err := readMessageAttachments(cmd.Context(), id)
			if err != nil {
				return err
			}
			return flags.printJSON(cmd, map[string]any{"source": "live", "message": v, "size_note": "reported_size is provided by Outlook and is not a verified downloaded byte count"})
		}})
		addAttachmentDownloads(group, flags)
		root.AddCommand(group)
	})
}
func readMessageAttachments(ctx context.Context, id string) (attachments.MessageAttachments, error) {
	r, err := getMessage(ctx, id)
	if err != nil {
		return attachments.MessageAttachments{}, err
	}
	got, item, state := exchange.ParseMessageGet(r.Body)
	if state != transport.EnvelopeOK || got != id {
		return attachments.MessageAttachments{}, &transport.Error{Kind: transport.KindUpstream, Status: r.Status}
	}
	v, state := attachments.ParseMessage(item)
	if state != transport.EnvelopeOK || v.MessageID != id {
		return attachments.MessageAttachments{}, fmt.Errorf("complete attachment metadata is unavailable")
	}
	return v, nil
}
