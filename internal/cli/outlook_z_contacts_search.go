// Observed live People search-suggestion command.
package cli

import (
	"fmt"
	"strings"

	"github.com/spf13/cobra"
	people "github.com/mikefwille/outlook-web-cli/internal/outlook/contacts"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func init() {
	registerNovelCommand(func(root *cobra.Command, flags *rootFlags) {
		group := findSubcommand(root, "contacts")
		if group == nil {
			group = &cobra.Command{Use: "contacts", Short: "Work with personal contacts"}
			root.AddCommand(group)
		}
		addOutlookContactsSearch(group, flags)
	})
}

// pp:data-source live
func addOutlookContactsSearch(group *cobra.Command, flags *rootFlags) {
	search := &cobra.Command{Use: "search <query>", Short: "Search up to six People suggestions, including directory", Long: "Returns up to six People suggestions from the captured search box. Results can include personal contacts and directory profiles, and are not an exhaustive search.", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live", "pp:happy-args": "query=example"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "contacts search"})
		}
		if len(args) != 1 || strings.TrimSpace(args[0]) == "" {
			return usageErr(fmt.Errorf("supply one search query"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "contacts search", "query": args[0], "scope": "contacts and directory suggestions"})
		}
		detail, request, err := outlookCapture("OUTLOOK_WEB_CONTACTS_SEARCH", true)
		if err != nil {
			return err
		}
		current, _, err := currentSession()
		if err != nil {
			return err
		}
		session, op, err := people.PrepareSearch(detail, request, current, args[0])
		if err != nil {
			return err
		}
		response, err := outlookOperation(cmd.Context(), session, op)
		if err != nil {
			return err
		}
		items, state := people.ParseSearch(response.Body)
		if state != transport.EnvelopeOK {
			return &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "scope": "personal contacts and directory", "kind": "search suggestions", "complete": false, "max_suggestions": 6, "results": items})
	}}
	group.AddCommand(search)
}
