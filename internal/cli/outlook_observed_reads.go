// Provisional captured-query reads. Each accepts only its own allowlisted
// observed RPC and requires operation-specific private evidence in the process
// environment; no arbitrary action, URL or file path is exposed as an MCP arg.
package cli

import (
	"context"
	"encoding/json"
	"fmt"
	"os"
	"time"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/capture"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/exchange"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

type observedRead struct{ name, action, env string }

func newObservedReadGroup(name string, flags *rootFlags, reads []observedRead) *cobra.Command {
	group := &cobra.Command{Use: name, Short: "Inspect an observed Outlook query"}
	for _, spec := range reads {
		spec := spec
		var folderID, calendarID, start, end string
		var limit, offset int
		use := spec.name
		argsCheck := cobra.NoArgs
		if spec.action == "GetConversationItems" {
			use += " <conversation-id>"
			argsCheck = cobra.MaximumNArgs(1)
		}
		cmd := &cobra.Command{Use: use, Short: "Run an observed " + spec.action + " query",
			Long: "Use the observed Outlook request shape. Omitted fields retain the captured query scope; supported flags replace only their corresponding observed fields.",
			Args: argsCheck, Annotations: map[string]string{"mcp:read-only": "true", "pp:method": "POST"},
			RunE: func(cmd *cobra.Command, args []string) error {
				if flags.dryRun {
					return json.NewEncoder(cmd.OutOrStdout()).Encode(map[string]any{"dry_run": true, "operation": spec.action, "source": "live", "state_change": false})
				}
				if spec.action == "GetConversationItems" && (len(args) != 1 || args[0] == "") {
					return usageErr(fmt.Errorf("supply exactly one conversation ID"))
				}
				file := os.Getenv(spec.env)
				current, _, err := currentSession()
				if err != nil {
					return err
				}
				var session transport.Session
				var operation transport.Operation
				if file == "" {
					name := map[string]string{"FindConversation": "threads_list", "GetConversationItems": "threads_get", "FindItem": "mail_list", "GetCalendarFoldersV2": "calendars_list", "GetCalendarView": "calendar_agenda", "GetAccountInformation": "account_info"}[spec.action]
					session, operation, err = capture.PrepareRuntimeRead(name, current)
				} else {
					session, operation, err = capture.LoadObservedRead(file, spec.action)
					if err == nil {
						session, err = capture.RebindRead(session, current)
					}
				}
				if err != nil {
					return err
				}
				resolvedFolderID := folderID
				if spec.action == "FindItem" && resolvedFolderID == "" {
					inbox, e := resolveDistinguishedFolder(cmd.Context(), "inbox")
					if e != nil {
						return e
					}
					resolvedFolderID = inbox.ID.ID
				}
				params := exchange.ReadParameters{FolderID: resolvedFolderID, CalendarID: calendarID, Start: start, End: end}
				if spec.action == "GetCalendarView" && start == "" && end == "" {
					zone, e := time.LoadLocation("America/Denver")
					if e != nil {
						return e
					}
					now := time.Now().In(zone)
					params.Start = now.Format(time.RFC3339)
					params.End = now.AddDate(0, 0, 7).Format(time.RFC3339)
				}
				if len(args) > 0 {
					params.ConversationID = args[0]
				}
				if cmd.Flags().Changed("limit") {
					params.Limit = &limit
				}
				if cmd.Flags().Changed("offset") {
					params.Offset = &offset
				}
				operation, err = exchange.ParameterizeRead(operation, spec.action, params)
				if err != nil {
					return err
				}
				client, err := transport.New(session.AccountKey, session, nil)
				if err != nil {
					return err
				}
				ctx, cancel := context.WithTimeout(cmd.Context(), 30*time.Second)
				defer cancel()
				response, err := client.Do(ctx, session.AccountKey, operation)
				if err != nil {
					return err
				}
				return printJSONFilteredKeep(cmd.OutOrStdout(), map[string]any{"source": "live", "scope": "captured-query-with-overrides", "response": json.RawMessage(response.Body)}, flags, "source", "scope", "response")
			}}
		switch spec.action {
		case "FindConversation":
			cmd.Flags().IntVar(&limit, "limit", 0, "Maximum conversations in this page (1–1000)")
			cmd.Flags().IntVar(&offset, "offset", 0, "Page offset within the captured folder")
		case "FindItem":
			cmd.Flags().StringVar(&folderID, "folder-id", "", "Folder ID; omitted resolves the current Inbox")
			cmd.Flags().IntVar(&limit, "limit", 0, "Maximum items in this page (1–1000)")
			cmd.Flags().IntVar(&offset, "offset", 0, "Page offset")
		case "GetConversationItems":
			cmd.Flags().IntVar(&limit, "limit", 0, "Maximum conversation items (1–1000)")
		case "GetCalendarView":
			cmd.Flags().StringVar(&calendarID, "calendar-id", "", "Calendar folder ID; omitted uses the captured calendar")
			cmd.Flags().StringVar(&start, "from", "", "Range start with RFC3339 offset")
			cmd.Flags().StringVar(&end, "to", "", "Range end with RFC3339 offset")
		}
		group.AddCommand(cmd)
	}
	return group
}
