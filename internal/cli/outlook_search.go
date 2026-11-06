package cli

import (
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"time"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/search"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func newMailSearchCmd(flags *rootFlags) *cobra.Command {
	var folder string
	var limit, offset int
	cmd := &cobra.Command{Use: "search <query>", Short: "Search Outlook conversations live", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live"},
		Long: "Search server conversations in all mailbox folders or a selected folder. The observed search includes Deleted Items in either scope. Results include upstream pagination and partial-result metadata.",
		RunE: func(cmd *cobra.Command, args []string) error {
			if flags.dryRun && requiredInputWaived() {
				return json.NewEncoder(cmd.OutOrStdout()).Encode(map[string]any{"operation": "search_conversations", "preview": true})
			}
			if len(args) != 1 || args[0] == "" {
				return usageErr(fmt.Errorf("supply exactly one search query"))
			}
			if flags.dryRun {
				return json.NewEncoder(cmd.OutOrStdout()).Encode(map[string]any{"operation": "search_conversations", "preview": true, "query": args[0], "folder_id": folder, "includes_deleted_items": true})
			}
			prefix := "OUTLOOK_WEB_SEARCH_ALL"
			if folder != "" {
				prefix = "OUTLOOK_WEB_SEARCH_FOLDER"
			}
			detail, request := os.Getenv(prefix+"_DETAILS"), os.Getenv(prefix+"_REQUEST")
			if detail == "" && request == "" {
				name := "search_all"
				if folder != "" {
					name = "search_folder"
				}
				detail = "runtime:" + name
			} else if detail == "" || request == "" {
				return fmt.Errorf("search template configuration is incomplete")
			}
			current, _, err := currentSession()
			if err != nil {
				return err
			}
			session, op, err := search.Prepare(detail, request, current, search.Params{Query: args[0], FolderID: folder, Offset: offset, Limit: limit})
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
			var e *transport.Error
			if err != nil && !(errors.As(err, &e) && e.Kind == transport.KindPartial) {
				return err
			}
			if printErr := printJSONFilteredKeep(cmd.OutOrStdout(), map[string]any{"source": "live", "entity": "conversations", "includes_deleted_items": true, "partial": err != nil, "response": json.RawMessage(response.Body)}, flags, "source", "entity", "includes_deleted_items", "partial", "response"); printErr != nil {
				return printErr
			}
			return err
		}}
	cmd.Flags().StringVar(&folder, "folder-id", "", "Search this folder; omitted searches all folders")
	cmd.Flags().IntVar(&limit, "limit", 25, "Maximum conversations in this page (1–100)")
	cmd.Flags().IntVar(&offset, "offset", 0, "Result offset")
	return cmd
}
