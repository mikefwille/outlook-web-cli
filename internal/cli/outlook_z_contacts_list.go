// Observed live personal-contacts list command.
package cli

import (
	"fmt"
	"github.com/spf13/cobra"
	people "github.com/mikefwille/outlook-web-cli/internal/outlook/contacts"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"strconv"
)

func init() {
	registerNovelCommand(func(root *cobra.Command, flags *rootFlags) {
		group := findSubcommand(root, "contacts")
		if group == nil {
			group = &cobra.Command{Use: "contacts", Short: "Work with personal contacts"}
			root.AddCommand(group)
		}
		addOutlookContactsList(group, flags)
	})
}

// pp:data-source live
func addOutlookContactsList(group *cobra.Command, flags *rootFlags) {
	var pageSize, offset int
	var cursor, snapshot string
	list := &cobra.Command{Use: "list", Short: "List a page of personal contacts", Long: "Fetches the observed Your contacts server page and slices it locally. Continue with next_offset and next_snapshot, plus next_cursor when present. Each call re-fetches the server page and rejects a changed snapshot. A sync_cursor is not a list page.", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live"}, RunE: func(cmd *cobra.Command, args []string) error {
		if len(args) != 0 {
			return usageErr(fmt.Errorf("contacts list takes flags only"))
		}
		if pageSize < 1 || pageSize > 1000 {
			return usageErr(fmt.Errorf("--limit must be between 1 and 1000"))
		}
		if offset < 0 || offset > 1_000_000 {
			return usageErr(fmt.Errorf("--offset must be between 0 and 1000000"))
		}
		if offset > 0 && snapshot == "" {
			return usageErr(fmt.Errorf("--snapshot is required with a nonzero --offset"))
		}
		if snapshot != "" && len(snapshot) != 64 {
			return usageErr(fmt.Errorf("--snapshot must be the value returned by contacts list"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "contacts list", "limit": pageSize, "offset": offset, "has_cursor": cursor != "", "has_snapshot": snapshot != ""})
		}
		detail, _, err := outlookCapture("OUTLOOK_WEB_CONTACTS_LIST", false)
		if err != nil {
			return err
		}
		current, _, err := currentSession()
		if err != nil {
			return err
		}
		session, op, err := people.PrepareList(detail, current, pageSize, cursor)
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
		// A server continuation without $top has an unknown cap. Do not claim
		// it is exhausted merely because it has no further link.
		serverLimit := 0
		if top := op.Query.Get("$top"); top != "" {
			serverLimit, err = strconv.Atoi(top)
			if err != nil {
				return &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
			}
		}
		window, err := people.WindowList(page, offset, pageSize, cursor, snapshot, serverLimit)
		if err != nil {
			return err
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "scope": "captured Your contacts filter", "results": window.Results, "next_cursor": window.NextCursor, "next_offset": window.NextOffset, "snapshot": window.Snapshot, "next_snapshot": window.NextSnapshot, "sync_cursor": window.SyncCursor, "has_more": window.HasMore, "complete": window.Complete, "server_cap_reached": window.ServerCapReached, "server_limit": window.ServerLimit, "fetched_count": window.FetchedCount, "offset": window.Offset, "pagination": "Each page re-fetches the server page; a changed snapshot requires restarting."})
	}}
	list.Flags().IntVar(&pageSize, "limit", 100, "Maximum contacts returned in this page (1–1000)")
	list.Flags().IntVar(&offset, "offset", 0, "Local offset into the re-fetched server page")
	list.Flags().StringVar(&snapshot, "snapshot", "", "next_snapshot returned with a nonzero next_offset")
	list.Flags().StringVar(&cursor, "cursor", "", "Server-provided next_cursor from the previous page")
	group.AddCommand(list)
}
