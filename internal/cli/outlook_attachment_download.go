package cli

import (
	"context"
	"crypto/sha256"
	"encoding/base64"
	"encoding/hex"
	"fmt"
	"os"
	"path/filepath"
	"time"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/attachments"
)

func attachmentPrivatePaths() (string, string, error) {
	pointer := os.Getenv("OUTLOOK_WEB_SESSION_POINTER")
	if pointer == "" {
		h, err := os.UserHomeDir()
		if err != nil {
			return "", "", err
		}
		pointer = filepath.Join(h, ".local", "share", "outlook-web-dev", "private", "current-outlook.json")
	}
	dir := filepath.Dir(pointer)
	return filepath.Join(dir, "outlook-attachment-route-v1.json"), filepath.Join(dir, "outlook-attachment-current-msal.json"), nil
}
func readAttachmentBytes(ctx context.Context, messageID, attachmentID string, max int64) (attachments.Metadata, []byte, error) {
	m, err := readMessageAttachments(ctx, messageID)
	if err != nil {
		return attachments.Metadata{}, nil, err
	}
	var selected *attachments.Metadata
	for i := range m.Attachments {
		if m.Attachments[i].ID == attachmentID {
			selected = &m.Attachments[i]
			break
		}
	}
	if selected == nil {
		return attachments.Metadata{}, nil, fmt.Errorf("attachment is not present on this message")
	}
	current, _, err := currentSession()
	if err != nil {
		return attachments.Metadata{}, nil, err
	}
	route, credential, err := attachmentPrivatePaths()
	if err != nil {
		return attachments.Metadata{}, nil, err
	}
	s, err := attachments.LoadDownloadSession(route, credential, current)
	if err != nil {
		return attachments.Metadata{}, nil, err
	}
	ctx, cancel := context.WithTimeout(ctx, 30*time.Second)
	defer cancel()
	b, err := attachments.Download(ctx, current, s, attachmentID, nil, max)
	if err != nil {
		return attachments.Metadata{}, nil, err
	}
	return *selected, b, nil
}
func addAttachmentDownloads(group *cobra.Command, flags *rootFlags) {
	get := &cobra.Command{Use: "get <message-id> <attachment-id>", Short: "Read an attachment as base64, up to 1 MiB", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "attachments get"})
		}
		if len(args) != 2 || args[0] == "" || args[1] == "" {
			return usageErr(fmt.Errorf("supply a message ID and attachment ID"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "attachments get", "message_id": args[0], "attachment_id": args[1]})
		}
		meta, data, err := readAttachmentBytes(cmd.Context(), args[0], args[1], 1<<20)
		if err != nil {
			return err
		}
		sum := sha256.Sum256(data)
		return flags.printJSON(cmd, map[string]any{"source": "live", "attachment": meta, "size": len(data), "sha256": hex.EncodeToString(sum[:]), "content_base64": base64.StdEncoding.EncodeToString(data)})
	}}
	var output string
	download := &cobra.Command{Use: "download <message-id> <attachment-id>", Short: "Save an attachment to a new local file", Long: "Downloads up to 16 MiB to the specified new path; existing files are never overwritten.", Annotations: map[string]string{"mcp:hidden": "true", "mcp:write-flags": "output"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "attachments download"})
		}
		if len(args) != 2 || args[0] == "" || args[1] == "" || output == "" {
			return usageErr(fmt.Errorf("supply message ID, attachment ID, and --output"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "attachments download", "output": output})
		}
		meta, data, err := readAttachmentBytes(cmd.Context(), args[0], args[1], attachments.MaxDownloadBytes)
		if err != nil {
			return err
		}
		f, err := os.OpenFile(output, os.O_CREATE|os.O_EXCL|os.O_WRONLY, 0600)
		if err != nil {
			return fmt.Errorf("cannot create output file; it may already exist")
		}
		ok := false
		defer func() {
			if !ok {
				_ = os.Remove(output)
			}
		}()
		if _, err = f.Write(data); err != nil {
			f.Close()
			return fmt.Errorf("cannot write attachment output")
		}
		if err = f.Close(); err != nil {
			return fmt.Errorf("cannot finish attachment output")
		}
		ok = true
		sum := sha256.Sum256(data)
		return flags.printJSON(cmd, map[string]any{"source": "live", "attachment": meta, "output": output, "size": len(data), "sha256": hex.EncodeToString(sum[:])})
	}}
	download.Flags().StringVar(&output, "output", "", "New local file to create")
	group.AddCommand(get, download)
}
