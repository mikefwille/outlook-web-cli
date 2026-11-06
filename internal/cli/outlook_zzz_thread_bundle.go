package cli

import (
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"os"
	"time"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/capture"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/exchange"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

// pp:data-source live
func init() {
	registerNovelCommand(func(root *cobra.Command, flags *rootFlags) {
		group := findSubcommand(root, "threads")
		if group != nil {
			group.AddCommand(newThreadBundleCmd(flags))
		}
	})
}
func newThreadBundleCmd(flags *rootFlags) *cobra.Command {
	var output string
	var limit int
	var execute bool
	cmd := &cobra.Command{Use: "bundle <conversation-id>", Short: "Export one bounded conversation view", Long: "Reads one observed GetConversationItems view for a specified conversation and returns bounded Outlook JSON with the source ID and read time. It does not fetch attachment blobs; bodies may be truncated and chronology or completeness is not guaranteed. File output previews first and --execute creates a private new file. Execution reads the conversation again, so the preview may differ.", Annotations: map[string]string{"mcp:read-only": "true", "mcp:write-flags": "output", "pp:data-source": "live", "pp:happy-args": "conversation-id=example-value"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "threads bundle"})
		}
		id, err := oneID(args)
		if err != nil {
			return err
		}
		if limit < 1 || limit > 100 {
			return usageErr(fmt.Errorf("--limit must be 1–100"))
		}
		if execute && (output == "" || output == "-") {
			return usageErr(fmt.Errorf("--execute requires a new --output path"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "threads bundle", "id": id, "limit": limit})
		}
		current, _, err := currentSession()
		if err != nil {
			return err
		}
		var session transport.Session
		var operation transport.Operation
		if file := os.Getenv("OUTLOOK_WEB_THREADS_GET_EVIDENCE"); file != "" {
			session, operation, err = capture.LoadObservedRead(file, "GetConversationItems")
			if err == nil {
				session, err = capture.RebindRead(session, current)
			}
		} else {
			session, operation, err = capture.PrepareRuntimeRead("threads_get", current)
		}
		if err != nil {
			return err
		}
		operation, err = exchange.ParameterizeRead(operation, "GetConversationItems", exchange.ReadParameters{ConversationID: id, Limit: &limit})
		if err != nil {
			return err
		}
		response, err := outlookOperation(cmd.Context(), session, operation)
		if err != nil {
			return err
		}
		if len(response.Body) > 4<<20 || !json.Valid(response.Body) {
			return fmt.Errorf("conversation response is invalid or exceeds the 4 MiB bundle limit")
		}
		payload := map[string]any{"source": "live", "format": "outlook-conversation-json", "conversation_id": id, "read_at": time.Now().UTC().Format(time.RFC3339), "limit": limit, "complete": false, "coverage": "one observed conversation view; bodies may be truncated; attachment blobs are not fetched", "response": json.RawMessage(response.Body)}
		if output == "" || output == "-" {
			return flags.printJSON(cmd, payload)
		}
		encoded, err := json.MarshalIndent(payload, "", "  ")
		if err != nil {
			return err
		}
		encoded = append(encoded, '\n')
		hash := sha256.Sum256(encoded)
		result := map[string]any{"preview": !execute, "operation": "threads bundle", "conversation_id": id, "output": output, "format": "outlook-conversation-json", "bytes": len(encoded), "sha256": hex.EncodeToString(hash[:]), "complete": false, "overwrites": false, "note": "execution reads the conversation again and can differ from this preview"}
		if !execute {
			return flags.printJSON(cmd, result)
		}
		if err = mutationInHarness(); err != nil {
			return err
		}
		if err = writeNewMailExport(output, encoded); err != nil {
			return err
		}
		result["exported"] = true
		return flags.printJSON(cmd, result)
	}}
	cmd.Flags().IntVar(&limit, "limit", 20, "Maximum requested conversation items (1–100)")
	cmd.Flags().StringVar(&output, "output", "", "New private file path, or - for stdout")
	cmd.Flags().BoolVar(&execute, "execute", false, "Create the previewed output file")
	return cmd
}
