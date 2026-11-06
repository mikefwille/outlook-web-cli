package cli

import (
	"encoding/json"
	"fmt"
	"time"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/exchange"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/mailactions"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

type messageActionState struct {
	ItemID struct {
		ID        string `json:"Id"`
		ChangeKey string `json:"ChangeKey"`
	} `json:"ItemId"`
	Parent struct {
		ID string `json:"Id"`
	} `json:"ParentFolderId"`
	Subject string `json:"Subject"`
	IsDraft *bool  `json:"IsDraft"`
	IsRead  *bool  `json:"IsRead"`
	Flag    *struct {
		Status string `json:"FlagStatus"`
	} `json:"Flag"`
}

// pp:data-source live
func newMailActionCmd(flags *rootFlags, kind string) *cobra.Command {
	var key, completed string
	var execute bool
	cmd := &cobra.Command{Use: kind + " <id>", Short: "Preview or " + kind + " one message", Long: "Acts on the selected message, not its whole conversation. Preview reads the current state and change key. The captured Outlook action uses AlwaysOverwrite, so the version check before execution is not atomic. Read and trash actions suppress read receipts; trash moves the item to Deleted Items.", Annotations: map[string]string{"pp:data-source": "live"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "mail " + kind})
		}
		id, err := oneID(args)
		if err != nil {
			return err
		}
		if kind == "complete" {
			if completed == "" && execute && !flags.dryRun {
				return usageErr(fmt.Errorf("--completed-at from the preview is required"))
			}
			if completed == "" {
				completed = time.Now().UTC().Format(time.RFC3339Nano)
			}
			if _, err := time.Parse(time.RFC3339Nano, completed); err != nil {
				return usageErr(fmt.Errorf("--completed-at must be an RFC3339 timestamp"))
			}
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "mail " + kind, "id": id})
		}
		response, err := getMessage(cmd.Context(), id)
		if err != nil {
			return err
		}
		got, raw, state := exchange.ParseMessageGet(response.Body)
		if state != transport.EnvelopeOK || got != id {
			return &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
		}
		var item messageActionState
		if json.Unmarshal(raw, &item) != nil || item.ItemID.ID != id || item.ItemID.ChangeKey == "" || item.IsDraft == nil || *item.IsDraft {
			return fmt.Errorf("item is not a verified non-draft message; no changes made")
		}
		if (kind == "mark-read" || kind == "mark-unread") && item.IsRead == nil {
			return fmt.Errorf("message read state is unavailable; no changes made")
		}
		if (kind == "flag" || kind == "unflag" || kind == "complete") && (item.Flag == nil || item.Flag.Status == "") {
			return fmt.Errorf("message flag state is unavailable; no changes made")
		}
		preview := map[string]any{"preview": true, "operation": "mail " + kind, "scope": "single message", "id": id, "subject": item.Subject, "folder_id": item.Parent.ID, "change_key": item.ItemID.ChangeKey, "version_check": "before write; not atomic"}
		if item.IsRead != nil {
			preview["is_read"] = *item.IsRead
		}
		if item.Flag != nil {
			preview["flag_status"] = item.Flag.Status
		}
		if kind == "mark-read" || kind == "mark-unread" || kind == "trash" {
			preview["suppress_read_receipts"] = true
		}
		if kind == "trash" {
			preview["destination"] = "Deleted Items"
		}
		if kind == "complete" {
			preview["completed_at"] = completed
			preview["wire_shape"] = "composed from observed single-message flag and conversation completion; item-level live verification pending"
		}
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
		prefix := map[string]string{"mark-read": "OUTLOOK_WEB_MAIL_MARK_READ", "mark-unread": "OUTLOOK_WEB_MAIL_MARK_UNREAD", "flag": "OUTLOOK_WEB_MAIL_FLAG", "unflag": "OUTLOOK_WEB_MAIL_UNFLAG", "complete": "OUTLOOK_WEB_MAIL_FLAG", "trash": "OUTLOOK_WEB_MAIL_TRASH"}[kind]
		detail, _, err := outlookCapture(prefix, false)
		if err != nil {
			return err
		}
		current, _, err := currentSession()
		if err != nil {
			return err
		}
		var s transport.Session
		var op transport.Operation
		if kind == "complete" {
			s, op, err = mailactions.PrepareComplete(detail, current, id, completed)
		} else {
			s, op, err = mailactions.Prepare(kind, detail, current, id)
		}
		if err != nil {
			return err
		}
		result, err := outlookOperation(cmd.Context(), s, op)
		if err != nil {
			return err
		}
		changed, status := mailactions.ParseResult(kind, result.Body)
		if status != transport.EnvelopeOK || kind != "trash" && changed.ID != id {
			return &transport.Error{Kind: transport.KindAmbiguous, Status: result.Status}
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "accepted": true, "operation": "mail " + kind, "scope": "single message", "result": changed})
	}}
	cmd.Flags().StringVar(&key, "change-key", "", "Expected version from the preview")
	if kind == "complete" {
		cmd.Flags().StringVar(&completed, "completed-at", "", "Completion timestamp from the preview (RFC3339)")
	}
	cmd.Flags().BoolVar(&execute, "execute", false, "Apply the previewed action once")
	return cmd
}
