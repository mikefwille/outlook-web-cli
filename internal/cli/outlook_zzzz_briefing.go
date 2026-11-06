package cli

import (
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"os"
	"strings"
	"time"

	"github.com/spf13/cobra"
	cal "github.com/mikefwille/outlook-web-cli/internal/outlook/calendar"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/capture"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/exchange"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

// pp:data-source live
func init() {
	registerNovelCommand(func(root *cobra.Command, flags *rootFlags) { root.AddCommand(newOutlookBriefingCmd(flags)) })
}
func newOutlookBriefingCmd(flags *rootFlags) *cobra.Command {
	var folderID, calendarID, from, to, output string
	var limit int
	var execute bool
	cmd := &cobra.Command{Use: "briefing", Short: "Summarize one cached mail folder and calendar view", Long: "Combines recent and flagged metadata from one explicitly reconciled folder with one bounded live calendar view. Folder freshness and missing dates/flags are reported; the calendar view is explicitly incomplete. No full message bodies or attachment blobs are fetched. With --output the default is a preview; --execute creates a private new file, rereading live data first.", Annotations: map[string]string{"mcp:read-only": "true", "mcp:write-flags": "output", "pp:data-source": "live"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "briefing"})
		}
		validID := func(x string) bool {
			return x != "" && len(x) <= 4096 && strings.TrimSpace(x) == x && !strings.ContainsAny(x, "\x00\r\n")
		}
		a, e1 := time.Parse(time.RFC3339, from)
		b, e2 := time.Parse(time.RFC3339, to)
		if len(args) != 0 || !validID(folderID) || !validID(calendarID) || e1 != nil || e2 != nil || !a.Before(b) || b.Sub(a) > 28*24*time.Hour || limit < 1 || limit > 50 {
			return usageErr(fmt.Errorf("supply folder and calendar IDs, an RFC3339 range of at most 28 days, and --limit 1–50"))
		}
		if execute && (output == "" || output == "-") {
			return usageErr(fmt.Errorf("--execute requires a new --output path"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "briefing", "folder_id": folderID, "calendar_id": calendarID, "from": from, "to": to, "limit": limit, "calendar_complete": false})
		}
		store, err := openMailCache(false)
		if os.IsNotExist(err) {
			return fmt.Errorf("no local mail snapshot; run sync folder first")
		}
		if err != nil {
			return err
		}
		folder, err := store.BriefFolder(cmd.Context(), folderID, limit)
		store.Close()
		if err != nil {
			return err
		}
		current, _, err := currentSession()
		if err != nil {
			return err
		}
		var session transport.Session
		var op transport.Operation
		if file := os.Getenv("OUTLOOK_WEB_CALENDAR_AGENDA_EVIDENCE"); file != "" {
			session, op, err = capture.LoadObservedRead(file, "GetCalendarView")
			if err == nil {
				session, err = capture.RebindRead(session, current)
			}
		} else {
			session, op, err = capture.PrepareRuntimeRead("calendar_agenda", current)
		}
		if err != nil {
			return err
		}
		op, err = exchange.ParameterizeRead(op, "GetCalendarView", exchange.ReadParameters{CalendarID: calendarID, Start: from, End: to})
		if err != nil {
			return err
		}
		response, err := outlookOperation(cmd.Context(), session, op)
		if err != nil {
			return err
		}
		events, err := cal.BriefAgendaView(response.Body, limit)
		if err != nil {
			return &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
		}
		payload := map[string]any{"format": "outlook-briefing-json", "mail": map[string]any{"source": "local", "coverage": "published selected folder mail-list metadata only", "snapshot": folder}, "calendar": map[string]any{"source": "live", "calendar_id": calendarID, "from": from, "to": to, "read_at": time.Now().UTC().Format(time.RFC3339), "coverage": "one server-provided agenda view; omissions are possible", "agenda": events}}
		if output == "" || output == "-" {
			return flags.printJSON(cmd, payload)
		}
		encoded, err := json.MarshalIndent(payload, "", "  ")
		if err != nil {
			return err
		}
		encoded = append(encoded, '\n')
		hash := sha256.Sum256(encoded)
		result := map[string]any{"preview": !execute, "operation": "briefing", "format": "outlook-briefing-json", "output": output, "bytes": len(encoded), "sha256": hex.EncodeToString(hash[:]), "overwrites": false, "note": "execution rereads data; results can differ from preview"}
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
	cmd.Flags().StringVar(&folderID, "folder-id", "", "Previously reconciled mail folder ID")
	cmd.Flags().StringVar(&calendarID, "calendar-id", "", "Calendar folder ID")
	cmd.Flags().StringVar(&from, "from", "", "Calendar range start (RFC3339)")
	cmd.Flags().StringVar(&to, "to", "", "Calendar range end (RFC3339)")
	cmd.Flags().IntVar(&limit, "limit", 10, "Maximum recent mail, flagged mail and events each (1–50)")
	cmd.Flags().StringVar(&output, "output", "", "New private file path, or - for stdout")
	cmd.Flags().BoolVar(&execute, "execute", false, "Create the previewed output file")
	return cmd
}
