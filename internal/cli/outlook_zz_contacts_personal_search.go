// Local filtering of the observed personal contacts list.
package cli

import (
	"fmt"
	"strconv"
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
		addOutlookPersonalSearch(group, flags)
	})
}

// pp:data-source live
func addOutlookPersonalSearch(group *cobra.Command, flags *rootFlags) {
	var limit, offset int
	var cursor, snapshot string
	cmd := &cobra.Command{Use: "search-personal [query]", Short: "Search your personal contacts", Long: "Fetches the observed personal contacts list and applies a case-insensitive substring search to names, email addresses, phone numbers, companies, and job titles. Pagination re-fetches the page and checks a snapshot hash. Server-cap uncertainty is reported.", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live", "pp:happy-args": "query=example"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "contacts search-personal"})
		}
		if len(args) != 1 || strings.TrimSpace(args[0]) == "" || len(args[0]) > 512 {
			return usageErr(fmt.Errorf("supply one nonempty search query of at most 512 bytes"))
		}
		if limit < 1 || limit > 1000 || offset < 0 || offset > 1_000_000 {
			return usageErr(fmt.Errorf("--limit must be 1–1000 and --offset must be nonnegative"))
		}
		if offset > 0 && snapshot == "" {
			return usageErr(fmt.Errorf("--snapshot is required with a nonzero --offset"))
		}
		if snapshot != "" && len(snapshot) != 64 {
			return usageErr(fmt.Errorf("--snapshot must be the value returned by search-personal"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "contacts search-personal", "query": args[0], "limit": limit, "offset": offset})
		}
		detail, _, err := outlookCapture("OUTLOOK_WEB_CONTACTS_LIST", false)
		if err != nil {
			return err
		}
		current, _, err := currentSession()
		if err != nil {
			return err
		}
		session, op, err := people.PrepareList(detail, current, limit, cursor)
		if err != nil {
			return err
		}
		response, err := outlookOperation(cmd.Context(), session, op)
		if err != nil {
			return err
		}
		page, state := people.ParseList(response.Body)
		if state != transport.EnvelopeOK {
			return &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
		}
		serverLimit := 0
		if top := op.Query.Get("$top"); top != "" {
			serverLimit, err = strconv.Atoi(top)
			if err != nil {
				return &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
			}
		}
		window, err := people.PersonalSearchWindow(page, args[0], offset, limit, cursor, snapshot, serverLimit)
		if err != nil {
			return err
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "scope": "captured Your contacts filter", "matching": "case-insensitive substring of names, emails, phones, companies, and job titles", "results": window.Results, "next_cursor": window.NextCursor, "next_offset": window.NextOffset, "snapshot": window.Snapshot, "next_snapshot": window.NextSnapshot, "has_more": window.HasMore, "complete": window.Complete, "server_cap_reached": window.ServerCapReached, "server_limit": window.ServerLimit, "source_fetched_count": len(page.Results), "matches_in_server_page": window.FetchedCount, "pagination": "Each page re-fetches the server page; a changed match snapshot requires restarting."})
	}}
	cmd.Flags().IntVar(&limit, "limit", 100, "Maximum matching contacts returned (1–1000)")
	cmd.Flags().IntVar(&offset, "offset", 0, "Local offset into matching contacts on this server page")
	cmd.Flags().StringVar(&snapshot, "snapshot", "", "next_snapshot returned with a nonzero next_offset")
	cmd.Flags().StringVar(&cursor, "cursor", "", "Server-provided next_cursor from the previous page")
	group.AddCommand(cmd)
}
