package cli

import (
	"encoding/json"
	"fmt"
	"os"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/capture"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/exchange"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func init() {
	registerNovelCommand(func(root *cobra.Command, flags *rootFlags) {
		deleted := &cobra.Command{Use: "deleted", Short: "Inspect and restore soft-deleted mail"}
		var offset, limit int
		list := &cobra.Command{Use: "list", Short: "List a page of messages in Deleted Items", Long: "Resolves the current mailbox's Deleted Items folder and reads one bounded page using the observed mail-list projection. It does not cover the recoverable-items store.", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live"}, RunE: func(cmd *cobra.Command, args []string) error {
			if len(args) != 0 {
				return usageErr(fmt.Errorf("deleted list takes no arguments"))
			}
			if limit < 1 || limit > 1000 || offset < 0 || offset > 1000000 {
				return usageErr(fmt.Errorf("--limit must be 1–1000 and --offset must be nonnegative"))
			}
			if flags.dryRun {
				return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "deleted list", "limit": limit, "offset": offset})
			}
			folder, err := resolveDistinguishedFolder(cmd.Context(), "deleteditems")
			if err != nil {
				return err
			}
			current, _, err := currentSession()
			if err != nil {
				return err
			}
			var session transport.Session
			var operation transport.Operation
			if file := os.Getenv("OUTLOOK_WEB_MAIL_LIST_EVIDENCE"); file != "" {
				session, operation, err = capture.LoadObservedRead(file, "FindItem")
				if err == nil {
					session, err = capture.RebindRead(session, current)
				}
			} else {
				session, operation, err = capture.PrepareRuntimeRead("mail_list", current)
			}
			if err != nil {
				return err
			}
			operation, err = exchange.ParameterizeRead(operation, "FindItem", exchange.ReadParameters{FolderID: folder.ID.ID, Limit: &limit, Offset: &offset})
			if err != nil {
				return err
			}
			response, err := outlookOperation(cmd.Context(), session, operation)
			if err != nil {
				return err
			}
			return flags.printJSON(cmd, map[string]any{"source": "live", "folder_id": folder.ID.ID, "scope": "mail items in Deleted Items", "response": json.RawMessage(response.Body)})
		}}
		list.Flags().IntVar(&offset, "offset", 0, "Deleted Items page offset")
		list.Flags().IntVar(&limit, "limit", 100, "Maximum items in this page (1–1000)")
		deleted.AddCommand(list, newMailMoveCmd(flags, "restore"))
		root.AddCommand(deleted)
	})
}
