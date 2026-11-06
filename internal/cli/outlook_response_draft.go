package cli

import (
	"context"
	"encoding/json"
	"fmt"
	"os"
	"strings"
	"time"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/attachments"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/capture"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/drafts"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/exchange"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

type responseSource struct {
	drafts.ReplySource
	Subject        string `json:"Subject"`
	ItemClass      string `json:"ItemClass"`
	IsDraft        *bool  `json:"IsDraft"`
	HasAttachments *bool  `json:"HasAttachments"`
	ItemID         struct {
		ID        string `json:"Id"`
		ChangeKey string `json:"ChangeKey"`
	} `json:"ItemId"`
}

func newResponseDraftCmd(flags *rootFlags, kind string) *cobra.Command {
	var text, textFile, subject string
	var to, cc, bcc []string
	var execute bool
	cmd := &cobra.Command{Use: kind + " <id>", Short: "Preview or create a " + kind + " draft", Annotations: map[string]string{"mcp:read-flags": "text-file"},
		Long: "Creates a server draft only. Replies derive recipients from the source when none are supplied; explicit To/Cc/Bcc override that selection. Forwarded file attachments are checked against the original by SHA-256 when available. Review the saved draft before sending.",
		RunE: func(cmd *cobra.Command, args []string) error {
			if flags.dryRun && requiredInputWaived() {
				return json.NewEncoder(cmd.OutOrStdout()).Encode(map[string]any{"operation": kind + "_draft", "preview": true, "sends_mail": false})
			}
			if len(args) != 1 || args[0] == "" {
				return usageErr(fmt.Errorf("supply exactly one source message ID"))
			}
			content, err := draftTextInput(cmd, text, textFile)
			if err != nil {
				return err
			}
			recipients, err := drafts.ParseRecipients(to, cc, bcc)
			if err != nil {
				return err
			}
			automatic := kind != "forward" && len(recipients.To)+len(recipients.Cc)+len(recipients.Bcc) == 0
			if kind == "forward" && len(recipients.To)+len(recipients.Cc)+len(recipients.Bcc) == 0 {
				return usageErr(fmt.Errorf("supply at least one explicit recipient"))
			}
			if execute && !flags.dryRun {
				if err := mutationInHarness(); err != nil {
					return err
				}
			}
			source, err := readResponseSourceWithReplyTo(cmd.Context(), args[0], automatic)
			if err != nil {
				return err
			}
			if automatic {
				var aliases []string
				if kind == "reply-all" {
					aliases, err = readAccountAliases(cmd.Context())
					if err != nil {
						return err
					}
				}
				recipients, err = drafts.ResolveReply(source.ReplySource, kind == "reply-all", aliases)
				if err != nil {
					return err
				}
			}
			var original attachments.MessageAttachments
			var originalFingerprints map[string]int
			if kind == "forward" {
				original, err = readMessageAttachments(cmd.Context(), source.ItemID.ID)
				if err != nil {
					return err
				}
				if original.ChangeKey != source.ItemID.ChangeKey {
					return &transport.Error{Kind: transport.KindConflict}
				}
				if execute && !flags.dryRun && len(original.Attachments) > 0 {
					originalFingerprints, err = attachmentFingerprints(cmd.Context(), source.ItemID.ID, original.Attachments)
					if err != nil {
						return err
					}
				}
			}
			if !cmd.Flags().Changed("subject") {
				prefix := "Re: "
				if kind == "forward" {
					prefix = "Fwd: "
				}
				if strings.HasPrefix(strings.ToLower(source.Subject), strings.ToLower(prefix)) {
					subject = source.Subject
				} else {
					subject = prefix + source.Subject
				}
			}
			if !execute || flags.dryRun {
				mode := "explicit"
				if automatic {
					mode = "derived_from_source"
				}
				result := map[string]any{"operation": kind + "_draft", "preview": true, "source_id": args[0], "source_subject": source.Subject, "subject": subject, "text": content, "recipients": recipients, "recipient_mode": mode, "sends_mail": false}
				if kind == "forward" {
					result["source_attachments"] = original.Attachments
					result["attachment_review"] = "Saved file attachments are checked by SHA-256 after draft creation; at most 20 files and 32 MiB total."
				}
				return json.NewEncoder(cmd.OutOrStdout()).Encode(result)
			}
			created, err := createResponseDraft(cmd.Context(), kind, source, subject, content, recipients)
			if err != nil {
				return err
			}
			result := map[string]any{"source": "live", "operation": kind + "_draft", "created": created, "sends_mail": false}
			if kind == "forward" && len(original.Attachments) > 0 {
				result["attachment_verification"] = "incomplete_review_required"
				saved, e := readMessageAttachments(cmd.Context(), created.ID)
				if e == nil {
					check, e := attachmentFingerprints(cmd.Context(), created.ID, saved.Attachments)
					if e == nil && fingerprintsEqual(originalFingerprints, check) {
						result["attachment_verification"] = "sha256_match"
					}
				}
			}
			return json.NewEncoder(cmd.OutOrStdout()).Encode(result)
		}}
	cmd.Flags().StringVar(&subject, "subject", "", "Override the response subject")
	cmd.Flags().StringVar(&text, "text", "", "Plain-text response to prepend")
	cmd.Flags().StringVar(&textFile, "text-file", "", "Read UTF-8 response from a local file, or - for stdin")
	cmd.Flags().StringArrayVar(&to, "to", nil, "To recipient; repeat for more than one")
	cmd.Flags().StringArrayVar(&cc, "cc", nil, "Cc recipient; repeat for more than one")
	cmd.Flags().StringArrayVar(&bcc, "bcc", nil, "Bcc recipient; repeat for more than one")
	cmd.Flags().BoolVar(&execute, "execute", false, "Save the response as a draft")
	return cmd
}

func readResponseSource(ctx context.Context, id string) (responseSource, error) {
	return readResponseSourceWithReplyTo(ctx, id, false)
}
func readResponseSourceWithReplyTo(ctx context.Context, id string, includeReplyTo bool) (responseSource, error) {
	response, err := getMessageWithReplyTo(ctx, id, includeReplyTo)
	if err != nil {
		return responseSource{}, err
	}
	got, raw, state := exchange.ParseMessageGet(response.Body)
	var source responseSource
	if state != transport.EnvelopeOK || got != id || json.Unmarshal(raw, &source) != nil || source.ItemID.ID != id || source.ItemID.ChangeKey == "" || source.IsDraft == nil || *source.IsDraft || !strings.HasPrefix(strings.ToLower(source.ItemClass), "ipm.note") {
		return responseSource{}, fmt.Errorf("source is not a verified ordinary email message")
	}
	return source, nil
}

func createResponseDraft(ctx context.Context, kind string, source responseSource, subject, text string, recipients drafts.Recipients) (drafts.Created, error) {
	prefix := map[string]string{"reply": "OUTLOOK_WEB_REPLY_CREATE", "reply-all": "OUTLOOK_WEB_REPLY_ALL_CREATE", "forward": "OUTLOOK_WEB_FORWARD_CREATE"}[kind]
	if prefix == "" {
		return drafts.Created{}, &transport.Error{Kind: transport.KindInput}
	}
	detail, request := os.Getenv(prefix+"_DETAILS"), os.Getenv(prefix+"_REQUEST")
	current, _, err := currentSession()
	if err != nil {
		return drafts.Created{}, err
	}
	var session transport.Session
	var op transport.Operation
	if detail == "" && request == "" {
		name := kind + "_create"
		if kind == "reply-all" {
			name = "reply_all_create"
		}
		session, op, err = capture.PrepareRuntimeDraft(name, current, source.ItemID.ID, source.ItemID.ChangeKey, subject, text, recipients, nil)
	} else if detail != "" && request != "" {
		session, op, err = capture.PrepareResponseDraft(detail, request, current, kind, source.ItemID.ID, source.ItemID.ChangeKey, subject, text, recipients)
	} else {
		return drafts.Created{}, fmt.Errorf("response template configuration is incomplete")
	}
	if err != nil {
		return drafts.Created{}, err
	}
	client, err := transport.New(session.AccountKey, session, nil)
	if err != nil {
		return drafts.Created{}, err
	}
	ctx, cancel := context.WithTimeout(ctx, 30*time.Second)
	defer cancel()
	response, err := client.Do(ctx, session.AccountKey, op)
	if err != nil {
		return drafts.Created{}, err
	}
	created, state := drafts.ParseCreate(response.Body)
	if state != transport.EnvelopeOK {
		return drafts.Created{}, &transport.Error{Kind: transport.KindAmbiguous, Status: response.Status}
	}
	return created, nil
}
