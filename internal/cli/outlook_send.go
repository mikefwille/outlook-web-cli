package cli

import (
	"context"
	"encoding/json"
	"fmt"
	"os"
	"time"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/capture"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/drafts"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/exchange"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func newDraftSendCmd(flags *rootFlags) *cobra.Command {
	var execute bool
	var expected string
	cmd := &cobra.Command{Use: "send <id>", Short: "Preview or explicitly send a saved draft",
		Long: "Read and preview the current body, To/Cc/Bcc recipients, and saved attachment metadata. Attachments can be inspected with attachments get or download. Supply --execute and the preview's --change-key to submit it. The server acceptance is not proof of delivery.",
		RunE: func(cmd *cobra.Command, args []string) error {
			if flags.dryRun && requiredInputWaived() {
				return json.NewEncoder(cmd.OutOrStdout()).Encode(map[string]any{"operation": "send_draft", "preview": true, "sends_mail": true})
			}
			if len(args) != 1 || args[0] == "" {
				return usageErr(fmt.Errorf("supply exactly one draft ID"))
			}
			if execute && !flags.dryRun && expected == "" {
				return usageErr(fmt.Errorf("--change-key from the preview is required with --execute"))
			}
			if execute && !flags.dryRun {
				if err := mutationInHarness(); err != nil {
					return err
				}
			}
			draft, err := readDraft(cmd.Context(), args[0])
			if err != nil {
				return err
			}
			compose, err := readDraftCompose(cmd.Context(), args[0])
			if err != nil {
				return err
			}
			if compose.ID != args[0] || compose.ChangeKey != draft.ChangeKey || compose.Subject != draft.Subject {
				return &transport.Error{Kind: transport.KindConflict}
			}
			metadata, err := readMessageAttachments(cmd.Context(), args[0])
			if err != nil {
				return err
			}
			if metadata.ChangeKey != compose.ChangeKey || draft.HasAttachments == nil || metadata.HasAttachments != *draft.HasAttachments {
				return &transport.Error{Kind: transport.KindConflict}
			}
			attachmentStatus := "unknown"
			if draft.HasAttachments != nil {
				if *draft.HasAttachments {
					attachmentStatus = "present"
				} else {
					attachmentStatus = "none"
				}
			}
			if len(metadata.Attachments) > 0 {
				attachmentStatus = "present"
			}
			if !execute || flags.dryRun {
				return json.NewEncoder(cmd.OutOrStdout()).Encode(map[string]any{"operation": "send_draft", "preview": true, "draft": compose, "attachment_status": attachmentStatus, "attachments": metadata.Attachments, "attachment_size_note": "reported_size is Outlook's value, not a verified byte count", "sends_mail": true})
			}
			if expected != compose.ChangeKey {
				return &transport.Error{Kind: transport.KindConflict}
			}
			if attachmentStatus == "unknown" {
				return fmt.Errorf("cannot send until attachment status is known")
			}
			if err := submitDraft(cmd.Context(), compose); err != nil {
				return err
			}
			return json.NewEncoder(cmd.OutOrStdout()).Encode(map[string]any{"source": "live", "accepted": true, "delivery": "unknown", "sends_mail": true})
		}}
	cmd.Flags().StringVar(&expected, "change-key", "", "Version returned by the send preview")
	cmd.Flags().BoolVar(&execute, "execute", false, "Submit the previewed draft for delivery")
	return cmd
}

func readDraftCompose(ctx context.Context, id string) (drafts.ComposeSnapshot, error) {
	detail, request := os.Getenv("OUTLOOK_WEB_DRAFT_COMPOSE_DETAILS"), os.Getenv("OUTLOOK_WEB_DRAFT_COMPOSE_REQUEST")
	current, _, err := currentSession()
	if err != nil {
		return drafts.ComposeSnapshot{}, err
	}
	var session transport.Session
	var op transport.Operation
	if detail == "" && request == "" {
		session, op, err = capture.PrepareRuntimeDraft("draft_compose", current, id, "", "", "", drafts.Recipients{}, nil)
	} else if detail != "" && request != "" {
		session, op, err = capture.PrepareDraftOperation(detail, request, "GetItem", current, id, "", "", "")
	} else {
		return drafts.ComposeSnapshot{}, fmt.Errorf("compose template configuration is incomplete")
	}
	if err != nil {
		return drafts.ComposeSnapshot{}, err
	}
	op.Validate = exchange.ValidateMessageGet
	client, err := transport.New(session.AccountKey, session, nil)
	if err != nil {
		return drafts.ComposeSnapshot{}, err
	}
	ctx, cancel := context.WithTimeout(ctx, 30*time.Second)
	defer cancel()
	response, err := client.Do(ctx, session.AccountKey, op)
	if err != nil {
		return drafts.ComposeSnapshot{}, err
	}
	compose, state := drafts.ParseCompose(response.Body)
	if state != transport.EnvelopeOK {
		return drafts.ComposeSnapshot{}, fmt.Errorf("draft recipients or complete body could not be verified")
	}
	return compose, nil
}

func submitDraft(ctx context.Context, compose drafts.ComposeSnapshot) error {
	detail, request := os.Getenv("OUTLOOK_WEB_DRAFT_SEND_DETAILS"), os.Getenv("OUTLOOK_WEB_DRAFT_SEND_REQUEST")
	current, _, err := currentSession()
	if err != nil {
		return err
	}
	var session transport.Session
	var op transport.Operation
	if detail == "" && request == "" {
		session, op, err = capture.PrepareRuntimeDraft("draft_send", current, compose.ID, compose.ChangeKey, compose.Subject, "", drafts.Recipients{}, nil)
	} else if detail != "" && request != "" {
		session, op, err = capture.PrepareDraftSend(detail, request, current, compose.ID, compose.ChangeKey, compose.Subject)
	} else {
		return fmt.Errorf("send template configuration is incomplete")
	}
	if err != nil {
		return err
	}
	client, err := transport.New(session.AccountKey, session, nil)
	if err != nil {
		return err
	}
	ctx, cancel := context.WithTimeout(ctx, 30*time.Second)
	defer cancel()
	_, err = client.Do(ctx, session.AccountKey, op)
	return err
}
