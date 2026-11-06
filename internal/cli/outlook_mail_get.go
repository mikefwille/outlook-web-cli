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

func newMailGetCmd(flags *rootFlags) *cobra.Command {
	return &cobra.Command{Use: "get <id>", Short: "Read a message by ID",
		Long:        "Read the observed message shape without intentionally changing its read state. Body size and recipient bounds are preserved, including server truncation metadata.",
		Annotations: map[string]string{"mcp:read-only": "true", "pp:method": "POST"},
		RunE: func(cmd *cobra.Command, args []string) error {
			if flags.dryRun && requiredInputWaived() {
				return json.NewEncoder(cmd.OutOrStdout()).Encode(map[string]any{"operation": "get_message", "preview": true})
			}
			if len(args) != 1 || args[0] == "" {
				return usageErr(fmt.Errorf("supply exactly one message ID"))
			}
			if flags.dryRun {
				return json.NewEncoder(cmd.OutOrStdout()).Encode(map[string]any{"operation": "get_message", "preview": true, "id": args[0]})
			}
			response, err := getMessage(cmd.Context(), args[0])
			if err != nil {
				return err
			}
			id, item, state := exchange.ParseMessageGet(response.Body)
			if state != transport.EnvelopeOK || id != args[0] {
				return &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
			}
			return printJSONFilteredKeep(cmd.OutOrStdout(), map[string]any{"source": "live", "item": item}, flags, "source", "item")
		}}
}

func getMessage(ctx context.Context, id string) (transport.Response, error) {
	return getMessageWithReplyTo(ctx, id, false)
}

func getMessageWithReplyTo(ctx context.Context, id string, includeReplyTo bool) (transport.Response, error) {
	return getMessageWithFields(ctx, id, includeReplyTo, false)
}

func getMessageWithFields(ctx context.Context, id string, includeReplyTo, includeCategories bool) (transport.Response, error) {
	detail, request := os.Getenv("OUTLOOK_WEB_MAIL_GET_DETAILS"), os.Getenv("OUTLOOK_WEB_MAIL_GET_REQUEST")
	if detail == "" {
		detail = os.Getenv("OUTLOOK_WEB_DRAFT_GET_DETAILS")
	}
	if request == "" {
		request = os.Getenv("OUTLOOK_WEB_DRAFT_GET_REQUEST")
	}
	current, _, err := currentSession()
	if err != nil {
		return transport.Response{}, err
	}
	var session transport.Session
	var op transport.Operation
	if detail == "" && request == "" {
		session, op, err = capture.PrepareRuntimeDraft("mail_get", current, id, "", "", "", drafts.Recipients{}, nil)
	} else if detail != "" && request != "" {
		session, op, err = capture.PrepareDraftOperation(detail, request, "GetItem", current, id, "", "", "")
	} else {
		return transport.Response{}, fmt.Errorf("message template configuration is incomplete")
	}
	if err != nil {
		return transport.Response{}, err
	}
	if includeReplyTo {
		op, err = exchange.IncludeReplyTo(op)
		if err != nil {
			return transport.Response{}, err
		}
	}
	if includeCategories {
		op, err = exchange.IncludeCategories(op)
		if err != nil {
			return transport.Response{}, err
		}
	}
	op.Validate = exchange.ValidateMessageGet
	client, err := transport.New(session.AccountKey, session, nil)
	if err != nil {
		return transport.Response{}, err
	}
	ctx, cancel := context.WithTimeout(ctx, 30*time.Second)
	defer cancel()
	return client.Do(ctx, session.AccountKey, op)
}
