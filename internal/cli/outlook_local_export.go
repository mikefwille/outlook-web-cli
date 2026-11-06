package cli

import (
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"os"

	"github.com/spf13/cobra"
)

// pp:data-source local
func newLocalMailExportCmd(flags *rootFlags) *cobra.Command {
	var folderID, output string
	var ids []string
	var execute bool
	cmd := &cobra.Command{Use: "export", Short: "Export selected cached mail metadata", Long: "Exports 1–100 exact item IDs from one previously reconciled folder, including its snapshot time and source IDs. The output contains mail-list metadata only: it is not full messages, a thread bundle, an EML export, or a live server snapshot. With --output the default is a preview; --execute creates a private new file without overwriting. Without --output, JSON is returned on stdout.", Annotations: map[string]string{"mcp:read-only": "true", "mcp:write-flags": "output", "pp:data-source": "local"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "export", "source": "local"})
		}
		if len(args) != 0 || folderID == "" || len(ids) == 0 || len(ids) > 100 {
			return usageErr(fmt.Errorf("supply --folder-id and 1–100 --item-id values"))
		}
		if execute && (output == "" || output == "-") {
			return usageErr(fmt.Errorf("--execute requires a new --output path"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "export", "source": "local", "item_count": len(ids)})
		}
		store, err := openMailCache(false)
		if os.IsNotExist(err) {
			return fmt.Errorf("no local mail snapshot; run sync folder first")
		}
		if err != nil {
			return err
		}
		defer store.Close()
		snapshot, err := store.Export(cmd.Context(), folderID, ids)
		if err != nil {
			return err
		}
		payload := map[string]any{"source": "local", "format": "outlook-mail-metadata-json", "coverage": "explicit items from one published folder snapshot; mail-list metadata only", "snapshot": snapshot}
		if output == "" || output == "-" {
			return flags.printJSON(cmd, payload)
		}
		encoded, err := json.MarshalIndent(payload, "", "  ")
		if err != nil {
			return err
		}
		encoded = append(encoded, '\n')
		hash := sha256.Sum256(encoded)
		result := map[string]any{"operation": "export", "preview": !execute, "source": "local", "format": "outlook-mail-metadata-json", "output": output, "item_count": len(snapshot.Items), "synced_at": snapshot.SyncedAt, "bytes": len(encoded), "sha256": hex.EncodeToString(hash[:]), "overwrites": false}
		if !execute {
			return flags.printJSON(cmd, result)
		}
		if err := mutationInHarness(); err != nil {
			return err
		}
		if err := writeNewMailExport(output, encoded); err != nil {
			return err
		}
		result["exported"] = true
		return flags.printJSON(cmd, result)
	}}
	cmd.Flags().StringVar(&folderID, "folder-id", "", "ID of one previously reconciled folder")
	cmd.Flags().StringArrayVar(&ids, "item-id", nil, "Exact cached item ID; repeat up to 100 times")
	cmd.Flags().StringVar(&output, "output", "", "New private file path, or - for stdout")
	cmd.Flags().BoolVar(&execute, "execute", false, "Create the previewed output file")
	return cmd
}
