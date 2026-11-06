package cli

import (
	"context"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"io"
	"mime"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/attachments"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func newDraftAttachCmd(flags *rootFlags) *cobra.Command {
	var file, key, expected, mimeType string
	var execute bool
	cmd := &cobra.Command{Use: "attach <id>", Short: "Preview or attach one local file to a draft", Annotations: map[string]string{"mcp:hidden": "true", "mcp:read-flags": "file"},
		Long: "Reads one local file of at most 4 MiB. Preview shows its size and SHA-256; execute requires that hash and the draft's current change key. This command does not send the draft.",
		RunE: func(cmd *cobra.Command, args []string) error {
			if flags.dryRun && requiredInputWaived() {
				return json.NewEncoder(cmd.OutOrStdout()).Encode(map[string]any{"operation": "attach_file", "preview": true, "sends_mail": false})
			}
			if len(args) != 1 || args[0] == "" || file == "" || key == "" {
				return usageErr(fmt.Errorf("draft ID, --file and --change-key are required"))
			}
			if execute && !flags.dryRun {
				if err := mutationInHarness(); err != nil {
					return err
				}
			}
			data, name, err := readAttachmentFile(file)
			if err != nil {
				return err
			}
			typ := mimeType
			if typ == "" {
				typ = mime.TypeByExtension(filepath.Ext(name))
				if typ == "" {
					typ = "application/octet-stream"
				}
			}
			mediaType, _, err := mime.ParseMediaType(typ)
			if err != nil || strings.ContainsAny(typ, "\r\n\x00") {
				return usageErr(fmt.Errorf("invalid MIME type"))
			}
			typ = mediaType
			sum := sha256.Sum256(data)
			hash := hex.EncodeToString(sum[:])
			if !execute || flags.dryRun {
				return json.NewEncoder(cmd.OutOrStdout()).Encode(map[string]any{"operation": "attach_file", "preview": true, "id": args[0], "change_key": key, "name": name, "content_type": typ, "size": len(data), "sha256": hash, "sends_mail": false})
			}
			if len(expected) != 64 || !strings.EqualFold(hash, expected) {
				return usageErr(fmt.Errorf("--sha256 must match the previewed file"))
			}
			draft, err := readDraft(cmd.Context(), args[0])
			if err != nil {
				return err
			}
			if draft.ChangeKey != key {
				return &transport.Error{Kind: transport.KindConflict}
			}
			detail, _, err := outlookCapture("OUTLOOK_WEB_DRAFT_ATTACH", false)
			if err != nil {
				return err
			}
			current, _, err := currentSession()
			if err != nil {
				return err
			}
			session, op, err := attachments.PrepareUpload(detail, current, args[0], key, name, typ, data)
			if err != nil {
				return err
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
			uploaded, state := attachments.ParseUploaded(response.Body)
			if state != transport.EnvelopeOK || uploaded.RootItemID != args[0] {
				return &transport.Error{Kind: transport.KindAmbiguous, Status: response.Status}
			}
			return json.NewEncoder(cmd.OutOrStdout()).Encode(map[string]any{"source": "live", "attached": uploaded, "name": name, "size": len(data), "sha256": hash, "sends_mail": false})
		}}
	cmd.Flags().StringVar(&file, "file", "", "Local file to attach")
	cmd.Flags().StringVar(&key, "change-key", "", "Current version returned by drafts get")
	cmd.Flags().StringVar(&expected, "sha256", "", "SHA-256 returned by the preview")
	cmd.Flags().StringVar(&mimeType, "content-type", "", "Override the inferred MIME type")
	cmd.Flags().BoolVar(&execute, "execute", false, "Upload this file to the draft")
	return cmd
}

func readAttachmentFile(path string) ([]byte, string, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, "", fmt.Errorf("cannot open attachment file")
	}
	defer f.Close()
	before, err := f.Stat()
	if err != nil || !before.Mode().IsRegular() || before.Size() > attachments.MaxUploadBytes {
		return nil, "", fmt.Errorf("attachment must be a regular file of at most 4 MiB")
	}
	data, err := io.ReadAll(io.LimitReader(f, attachments.MaxUploadBytes+1))
	if err != nil || len(data) > attachments.MaxUploadBytes {
		return nil, "", fmt.Errorf("cannot read attachment within size limit")
	}
	after, err := f.Stat()
	if err != nil || before.Size() != after.Size() || !before.ModTime().Equal(after.ModTime()) || int64(len(data)) != after.Size() {
		return nil, "", fmt.Errorf("attachment changed while reading")
	}
	name := filepath.Base(path)
	if name == "." || name == "" || strings.ContainsAny(name, "/\\\r\n\x00") {
		return nil, "", fmt.Errorf("invalid attachment name")
	}
	return data, name, nil
}
