// Provisional recipient-free SaveOnly draft creation, based on a captured UI
// fixture. No send or arbitrary RPC action is exposed by this command.
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
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func newDraftsCmd(flags *rootFlags) *cobra.Command {
	parent := &cobra.Command{Use: "drafts", Short: "Inspect and manage server drafts"}
	var subject, text, textFile, html, htmlFile string
	var to, cc, bcc []string
	var execute bool
	create := &cobra.Command{Use: "create", Short: "Preview or save a draft", Args: cobra.NoArgs,
		Long:        "Create a draft with SaveOnly. It previews by default; --execute saves the draft but never sends it. Recipients may be repeated; files may be added separately with drafts attach.",
		Annotations: map[string]string{"mcp:destructive": "false", "mcp:read-flags": "text-file,html-file"},
		RunE: func(cmd *cobra.Command, _ []string) error {
			if flags.dryRun && requiredInputWaived() {
				return json.NewEncoder(cmd.OutOrStdout()).Encode(map[string]any{"operation": "create_draft", "preview": true, "sends_mail": false})
			}
			if !cmd.Flags().Changed("subject") {
				return usageErr(fmt.Errorf("--subject is required"))
			}
			content, htmlContent, err := draftBodyInput(cmd, text, textFile, html, htmlFile)
			if err != nil {
				return err
			}
			recipients, err := drafts.ParseRecipients(to, cc, bcc)
			if err != nil {
				return err
			}
			if !execute || flags.dryRun {
				preview := map[string]any{"operation": "create_draft", "preview": true, "subject": subject, "text": content, "recipients": recipients, "sends_mail": false}
				if htmlContent != nil {
					preview["html"] = *htmlContent
				}
				return json.NewEncoder(cmd.OutOrStdout()).Encode(preview)
			}
			if err := mutationInHarness(); err != nil {
				return err
			}
			detail, request := os.Getenv("OUTLOOK_WEB_DRAFT_CREATE_DETAILS"), os.Getenv("OUTLOOK_WEB_DRAFT_CREATE_REQUEST")
			current, _, err := currentSession()
			if err != nil {
				return err
			}
			var session transport.Session
			var op transport.Operation
			if detail == "" && request == "" {
				session, op, err = capture.PrepareRuntimeDraft("draft_create", current, "", "", subject, content, recipients, nil)
			} else if detail != "" && request != "" {
				session, op, err = capture.PrepareDraftCreateWithRecipients(detail, request, current, subject, content, recipients)
			} else {
				return fmt.Errorf("draft template configuration is incomplete")
			}
			if err != nil {
				return err
			}
			if htmlContent != nil {
				op.Body, err = drafts.SetHTMLBody(op.Body, "CreateItem", *htmlContent)
				if err != nil {
					return err
				}
			}
			client, err := transport.New(session.AccountKey, session, nil)
			if err != nil {
				return err
			}
			ctx, cancel := context.WithTimeout(cmd.Context(), 30*time.Second)
			defer cancel()
			response, err := client.Do(ctx, session.AccountKey, op)
			if err != nil {
				return err
			}
			created, state := drafts.ParseCreate(response.Body)
			if state != transport.EnvelopeOK {
				return &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
			}
			return json.NewEncoder(cmd.OutOrStdout()).Encode(map[string]any{"source": "live", "created": created, "sends_mail": false})
		}}
	create.Flags().StringVar(&subject, "subject", "", "Draft subject")
	create.Flags().StringVar(&text, "text", "", "Plain-text body")
	create.Flags().StringVar(&textFile, "text-file", "", "Read UTF-8 body from a local file, or - for stdin")
	create.Flags().StringVar(&html, "html", "", "Rich HTML body; --text supplies an optional plain-text companion")
	create.Flags().StringVar(&htmlFile, "html-file", "", "Read rich HTML from a local file, or - for stdin")
	create.Flags().StringArrayVar(&to, "to", nil, "To recipient; repeat for more than one")
	create.Flags().StringArrayVar(&cc, "cc", nil, "Cc recipient; repeat for more than one")
	create.Flags().StringArrayVar(&bcc, "bcc", nil, "Bcc recipient; repeat for more than one")
	create.Flags().BoolVar(&execute, "execute", false, "Save the draft on the server")
	parent.AddCommand(create)
	get := &cobra.Command{Use: "get <id>", Short: "Read a draft by ID", Annotations: map[string]string{"mcp:read-only": "true"},
		RunE: func(cmd *cobra.Command, args []string) error {
			if flags.dryRun && requiredInputWaived() {
				return json.NewEncoder(cmd.OutOrStdout()).Encode(map[string]any{"operation": "get_draft", "preview": true})
			}
			if len(args) != 1 || args[0] == "" {
				return usageErr(fmt.Errorf("supply exactly one draft ID"))
			}
			if flags.dryRun {
				return json.NewEncoder(cmd.OutOrStdout()).Encode(map[string]any{"operation": "get_draft", "preview": true, "id": args[0]})
			}
			draft, err := readDraft(cmd.Context(), args[0])
			if err != nil {
				return err
			}
			return json.NewEncoder(cmd.OutOrStdout()).Encode(map[string]any{"source": "live", "draft": draft})
		}}
	parent.AddCommand(get)
	var updateSubject, updateText, updateTextFile, updateHTML, updateHTMLFile, updateKey string
	var updateTo, updateCc, updateBcc []string
	var clearTo, clearCc, clearBcc bool
	var updateExecute bool
	update := &cobra.Command{Use: "update <id>", Short: "Preview or update a draft", Annotations: map[string]string{"mcp:read-flags": "text-file,html-file"},
		Long: "Change only the supplied subject, plain-text or HTML body, or recipients; omitted fields remain untouched. The expected change key is checked before the update and the request uses NeverOverwrite. It previews unless --execute is supplied.",
		RunE: func(cmd *cobra.Command, args []string) error {
			if flags.dryRun && requiredInputWaived() {
				return json.NewEncoder(cmd.OutOrStdout()).Encode(map[string]any{"operation": "update_draft", "preview": true, "sends_mail": false})
			}
			if len(args) != 1 || args[0] == "" || updateKey == "" {
				return usageErr(fmt.Errorf("draft ID and --change-key are required"))
			}
			hasSubject := cmd.Flags().Changed("subject")
			hasText := cmd.Flags().Changed("text") || cmd.Flags().Changed("text-file") || cmd.Flags().Changed("html") || cmd.Flags().Changed("html-file")
			hasRecipients := cmd.Flags().Changed("to") || cmd.Flags().Changed("cc") || cmd.Flags().Changed("bcc") || clearTo || clearCc || clearBcc
			if !hasSubject && !hasText && !hasRecipients {
				return usageErr(fmt.Errorf("provide a subject, text or recipient change"))
			}
			content, htmlContent, err := draftBodyInput(cmd, updateText, updateTextFile, updateHTML, updateHTMLFile)
			if err != nil {
				return err
			}
			if clearTo && cmd.Flags().Changed("to") || clearCc && cmd.Flags().Changed("cc") || clearBcc && cmd.Flags().Changed("bcc") {
				return fmt.Errorf("cannot both replace and clear the same recipient list")
			}
			parsed, err := drafts.ParseRecipients(updateTo, updateCc, updateBcc)
			if err != nil {
				return err
			}
			changes := &drafts.RecipientChanges{}
			previewChanges := map[string]any{}
			if cmd.Flags().Changed("to") || clearTo {
				changes.To = &parsed.To
				previewChanges["to"] = parsed.To
			}
			if cmd.Flags().Changed("cc") || clearCc {
				changes.Cc = &parsed.Cc
				previewChanges["cc"] = parsed.Cc
			}
			if cmd.Flags().Changed("bcc") || clearBcc {
				changes.Bcc = &parsed.Bcc
				previewChanges["bcc"] = parsed.Bcc
			}
			if !updateExecute || flags.dryRun {
				preview := map[string]any{"operation": "update_draft", "preview": true, "id": args[0], "change_key": updateKey, "recipient_changes": previewChanges, "sends_mail": false}
				if hasSubject {
					preview["subject"] = updateSubject
				}
				if hasText {
					preview["text"] = content
					if htmlContent != nil {
						preview["html"] = *htmlContent
					}
				}
				return json.NewEncoder(cmd.OutOrStdout()).Encode(preview)
			}
			if err := mutationInHarness(); err != nil {
				return err
			}
			draft, err := readDraft(cmd.Context(), args[0])
			if err != nil {
				return err
			}
			if draft.ChangeKey != updateKey {
				return &transport.Error{Kind: transport.KindConflict}
			}
			response, err := executeDraftOperationSelectedHTML(cmd.Context(), "UpdateItem", args[0], updateKey, updateSubject, content, changes, hasSubject, hasText, htmlContent)
			if err != nil {
				return err
			}
			updated, state := drafts.ParseUpdate(response.Body)
			if state != transport.EnvelopeOK || updated.ID != args[0] {
				return &transport.Error{Kind: transport.KindAmbiguous, Status: response.Status}
			}
			return json.NewEncoder(cmd.OutOrStdout()).Encode(map[string]any{"source": "live", "updated": updated, "sends_mail": false})
		}}
	update.Flags().StringVar(&updateKey, "change-key", "", "Version returned by drafts get or create")
	update.Flags().StringVar(&updateSubject, "subject", "", "Replacement subject")
	update.Flags().StringVar(&updateText, "text", "", "Replacement plain-text body")
	update.Flags().StringVar(&updateTextFile, "text-file", "", "Read replacement UTF-8 body from a file, or - for stdin")
	update.Flags().StringVar(&updateHTML, "html", "", "Replacement rich HTML body")
	update.Flags().StringVar(&updateHTMLFile, "html-file", "", "Read replacement HTML from a file, or - for stdin")
	update.Flags().StringArrayVar(&updateTo, "to", nil, "Replace To recipients; repeat for more than one")
	update.Flags().StringArrayVar(&updateCc, "cc", nil, "Replace Cc recipients; repeat for more than one")
	update.Flags().StringArrayVar(&updateBcc, "bcc", nil, "Replace Bcc recipients; repeat for more than one")
	update.Flags().BoolVar(&clearTo, "clear-to", false, "Remove all To recipients")
	update.Flags().BoolVar(&clearCc, "clear-cc", false, "Remove all Cc recipients")
	update.Flags().BoolVar(&clearBcc, "clear-bcc", false, "Remove all Bcc recipients")
	update.Flags().BoolVar(&updateExecute, "execute", false, "Save the update")
	parent.AddCommand(update)
	var discardKey string
	var discardExecute bool
	discard := &cobra.Command{Use: "discard <id>", Short: "Preview or move a draft to Deleted Items",
		Long: "Verify the current item is a draft with the expected change key, then move it to Deleted Items without notifications. It previews unless --execute is supplied.",
		RunE: func(cmd *cobra.Command, args []string) error {
			if flags.dryRun && requiredInputWaived() {
				return json.NewEncoder(cmd.OutOrStdout()).Encode(map[string]any{"operation": "discard_draft", "preview": true, "sends_mail": false})
			}
			if len(args) != 1 || args[0] == "" || discardKey == "" {
				return usageErr(fmt.Errorf("draft ID and --change-key are required"))
			}
			if !discardExecute || flags.dryRun {
				return json.NewEncoder(cmd.OutOrStdout()).Encode(map[string]any{"operation": "discard_draft", "preview": true, "id": args[0], "change_key": discardKey, "destination": "deleted_items", "sends_mail": false})
			}
			if err := mutationInHarness(); err != nil {
				return err
			}
			draft, err := readDraft(cmd.Context(), args[0])
			if err != nil {
				return err
			}
			if draft.ChangeKey != discardKey {
				return &transport.Error{Kind: transport.KindConflict}
			}
			_, err = executeDraftOperation(cmd.Context(), "DeleteItem", args[0], discardKey, "", "")
			if err != nil {
				return err
			}
			return json.NewEncoder(cmd.OutOrStdout()).Encode(map[string]any{"source": "live", "discarded": true, "destination": "deleted_items", "sends_mail": false})
		}}
	discard.Flags().StringVar(&discardKey, "change-key", "", "Version returned by drafts get or create")
	discard.Flags().BoolVar(&discardExecute, "execute", false, "Move the draft to Deleted Items")
	parent.AddCommand(discard)
	parent.AddCommand(newDraftListCmd(flags))
	parent.AddCommand(newDraftSendCmd(flags))
	parent.AddCommand(newDraftAttachCmd(flags))
	parent.AddCommand(newDraftRemoveAttachmentCmd(flags))
	return parent
}

func readDraft(ctx context.Context, id string) (drafts.Draft, error) {
	response, err := executeDraftOperation(ctx, "GetItem", id, "", "", "")
	if err != nil {
		return drafts.Draft{}, err
	}
	draft, state := drafts.ParseGet(response.Body)
	if state != transport.EnvelopeOK || draft.ID != id {
		return drafts.Draft{}, &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
	}
	return draft, nil
}

func executeDraftOperation(ctx context.Context, action, id, key, subject, text string) (transport.Response, error) {
	return executeDraftOperationWithRecipients(ctx, action, id, key, subject, text, nil)
}

func executeDraftOperationWithRecipients(ctx context.Context, action, id, key, subject, text string, recipients *drafts.RecipientChanges) (transport.Response, error) {
	return executeDraftOperationSelected(ctx, action, id, key, subject, text, recipients, true, true)
}

func executeDraftOperationSelected(ctx context.Context, action, id, key, subject, text string, recipients *drafts.RecipientChanges, includeSubject, includeText bool) (transport.Response, error) {
	return executeDraftOperationSelectedHTML(ctx, action, id, key, subject, text, recipients, includeSubject, includeText, nil)
}

func executeDraftOperationSelectedHTML(ctx context.Context, action, id, key, subject, text string, recipients *drafts.RecipientChanges, includeSubject, includeText bool, htmlContent *string) (transport.Response, error) {
	prefix := map[string]string{"GetItem": "OUTLOOK_WEB_DRAFT_GET", "UpdateItem": "OUTLOOK_WEB_DRAFT_UPDATE", "DeleteItem": "OUTLOOK_WEB_DRAFT_DISCARD"}[action]
	if prefix == "" {
		return transport.Response{}, &transport.Error{Kind: transport.KindInput}
	}
	detail, request := os.Getenv(prefix+"_DETAILS"), os.Getenv(prefix+"_REQUEST")
	current, _, err := currentSession()
	if err != nil {
		return transport.Response{}, err
	}
	var session transport.Session
	var op transport.Operation
	if detail == "" && request == "" {
		name := map[string]string{"GetItem": "draft_get", "UpdateItem": "draft_update", "DeleteItem": "draft_discard"}[action]
		session, op, err = capture.PrepareRuntimeDraft(name, current, id, key, subject, text, drafts.Recipients{}, recipients)
	} else if detail != "" && (action == "DeleteItem" && request == "" || action != "DeleteItem" && request != "") {
		session, op, err = capture.PrepareDraftOperationWithRecipients(detail, request, action, current, id, key, subject, text, recipients)
	} else {
		return transport.Response{}, fmt.Errorf("draft template configuration is incomplete")
	}
	if err != nil {
		return transport.Response{}, err
	}
	if action == "UpdateItem" && (!includeSubject || !includeText) {
		if op.Headers.Get("X-Owa-Urlpostdata") != "" {
			return transport.Response{}, &transport.Error{Kind: transport.KindInput}
		}
		op.Body, err = drafts.SelectUpdateFields(op.Body, includeSubject, includeText)
		if err != nil {
			return transport.Response{}, err
		}
	}
	if htmlContent != nil {
		if action != "UpdateItem" || !includeText || op.Headers.Get("X-Owa-Urlpostdata") != "" {
			return transport.Response{}, &transport.Error{Kind: transport.KindInput}
		}
		op.Body, err = drafts.SetHTMLBody(op.Body, action, *htmlContent)
		if err != nil {
			return transport.Response{}, err
		}
	}
	client, err := transport.New(session.AccountKey, session, nil)
	if err != nil {
		return transport.Response{}, err
	}
	ctx, cancel := context.WithTimeout(ctx, 30*time.Second)
	defer cancel()
	return client.Do(ctx, session.AccountKey, op)
}

// newMailComposeCmd reuses the same draft constructor and semantics as
// drafts create; composing does not submit a message.
func newMailComposeCmd(flags *rootFlags) *cobra.Command {
	group := newDraftsCmd(flags)
	cmd := findSubcommand(group, "create")
	group.RemoveCommand(cmd)
	cmd.Use = "compose"
	cmd.Short = "Preview or save a new mail draft"
	return cmd
}
