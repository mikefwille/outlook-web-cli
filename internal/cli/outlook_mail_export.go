package cli

import (
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"os"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/exchange"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

// writeNewMailExport creates only the requested new file, with private mode.
func writeNewMailExport(path string, data []byte) error {
	f, err := os.OpenFile(path, os.O_WRONLY|os.O_CREATE|os.O_EXCL, 0600)
	if err != nil {
		return fmt.Errorf("cannot create export file; it may already exist")
	}
	complete := false
	defer func() {
		if !complete {
			_ = os.Remove(path)
		}
	}()
	if _, err = f.Write(data); err != nil {
		_ = f.Close()
		return fmt.Errorf("cannot write export file")
	}
	if err = f.Sync(); err != nil {
		_ = f.Close()
		return fmt.Errorf("cannot sync export file")
	}
	if err = f.Close(); err != nil {
		return fmt.Errorf("cannot close export file")
	}
	complete = true
	return nil
}

// pp:data-source live
func newMailExportCmd(flags *rootFlags) *cobra.Command {
	var output string
	var execute bool
	cmd := &cobra.Command{Use: "export <id>", Short: "Export the observed message as JSON", Long: "Reads one message and exports the observed bounded Outlook JSON shape. This is not an original MIME/EML export and bodies may be truncated. With --output the command previews the new file by default; --execute creates it once without overwriting an existing path. Without --output it returns JSON on stdout.", Annotations: map[string]string{"mcp:read-only": "true", "mcp:write-flags": "output", "pp:data-source": "live"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "mail export", "format": "outlook-json"})
		}
		id, err := oneID(args)
		if err != nil {
			return err
		}
		if execute && (output == "" || output == "-") {
			return usageErr(fmt.Errorf("--execute requires a new --output path"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "mail export", "id": id, "format": "outlook-json", "output": output})
		}
		response, err := getMessage(cmd.Context(), id)
		if err != nil {
			return err
		}
		got, item, state := exchange.ParseMessageGet(response.Body)
		if state != transport.EnvelopeOK || got != id {
			return &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
		}
		payload := map[string]any{"source": "live", "format": "outlook-json", "message": item, "body_note": "The observed message body can be truncated; this is not original MIME/EML."}
		if output == "" || output == "-" {
			return flags.printJSON(cmd, payload)
		}
		encoded, err := json.MarshalIndent(payload, "", "  ")
		if err != nil {
			return err
		}
		encoded = append(encoded, '\n')
		hash := sha256.Sum256(encoded)
		result := map[string]any{"operation": "mail export", "preview": !execute, "id": id, "format": "outlook-json", "output": output, "bytes": len(encoded), "sha256": hex.EncodeToString(hash[:]), "overwrites": false, "body_may_be_truncated": true}
		if !execute {
			return flags.printJSON(cmd, result)
		}
		if err := mutationInHarness(); err != nil {
			return err
		}
		if err := writeNewMailExport(output, encoded); err != nil {
			return err
		}
		result["source"] = "live"
		result["exported"] = true
		return flags.printJSON(cmd, result)
	}}
	cmd.Flags().StringVar(&output, "output", "", "New file path, or - for stdout")
	cmd.Flags().BoolVar(&execute, "execute", false, "Create the previewed output file")
	return cmd
}
