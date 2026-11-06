package cli

import (
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

func init() {
	registerNovelCommand(func(root *cobra.Command, flags *rootFlags) {
		group := findSubcommand(root, "calendar")
		if group == nil {
			group = &cobra.Command{Use: "calendar", Short: "Work with calendar appointments"}
			root.AddCommand(group)
		}
		addOutlookCalendarSubjectSearch(group, flags)
		addOutlookCalendarFindTimes(group, flags)
	})
}

// pp:data-source live
func addOutlookCalendarSubjectSearch(group *cobra.Command, flags *rootFlags) {
	var calendarID, from, to string
	var limit int
	cmd := &cobra.Command{Use: "search <query>", Short: "Search subjects in a bounded calendar view", Long: "Reads one calendar's agenda for an explicit range of at most 28 days, then searches returned Subject fields locally using a case-insensitive substring. Missing subjects, server truncation and other fields can hide matches; results are not exhaustive.", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live", "pp:happy-args": "query=example"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "calendar search", "state_change": false})
		}
		if len(args) != 1 || strings.TrimSpace(args[0]) == "" || len(args[0]) > 256 {
			return usageErr(fmt.Errorf("supply one nonempty subject query of at most 256 bytes"))
		}
		if calendarID == "" || strings.TrimSpace(calendarID) != calendarID || len(calendarID) > 4096 || strings.ContainsAny(calendarID, "\r\n\x00") {
			return usageErr(fmt.Errorf("--calendar-id must identify one calendar"))
		}
		start, err1 := time.Parse(time.RFC3339, from)
		end, err2 := time.Parse(time.RFC3339, to)
		if err1 != nil || err2 != nil || !start.Before(end) || end.Sub(start) > 28*24*time.Hour {
			return usageErr(fmt.Errorf("--from and --to must be RFC3339 endpoints for a range of at most 28 days"))
		}
		if limit < 1 || limit > 200 {
			return usageErr(fmt.Errorf("--limit must be between 1 and 200"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "calendar search", "state_change": false, "scope": "one calendar and requested range", "matching": "case-insensitive substring of returned Subject", "complete": false})
		}
		current, _, err := currentSession()
		if err != nil {
			return err
		}
		var session transport.Session
		var operation transport.Operation
		if file := os.Getenv("OUTLOOK_WEB_CALENDAR_AGENDA_EVIDENCE"); file != "" {
			session, operation, err = capture.LoadObservedRead(file, "GetCalendarView")
			if err == nil {
				session, err = capture.RebindRead(session, current)
			}
		} else {
			session, operation, err = capture.PrepareRuntimeRead("calendar_agenda", current)
		}
		if err != nil {
			return err
		}
		operation, err = exchange.ParameterizeRead(operation, "GetCalendarView", exchange.ReadParameters{CalendarID: calendarID, Start: from, End: to})
		if err != nil {
			return err
		}
		response, err := outlookOperation(cmd.Context(), session, operation)
		if err != nil {
			return err
		}
		results, matches, missing, truncated, err := cal.SearchAgendaSubjects(response.Body, args[0], limit)
		if err != nil {
			return &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "method": "client-side agenda subject search", "scope": "one calendar and requested range", "calendar_id": calendarID, "from": from, "to": to, "matching": "case-insensitive substring of returned Subject", "complete": false, "completeness_reason": "server view completeness and missing subjects are not established", "matched_in_response": matches, "items_without_subject": missing, "limit": limit, "truncated": truncated, "results": results})
	}}
	cmd.Flags().StringVar(&calendarID, "calendar-id", "", "Calendar folder ID (required)")
	cmd.Flags().StringVar(&from, "from", "", "Inclusive range start with RFC3339 offset (required)")
	cmd.Flags().StringVar(&to, "to", "", "Exclusive range end with RFC3339 offset (required)")
	cmd.Flags().IntVar(&limit, "limit", 50, "Maximum returned matches (1–200)")
	group.AddCommand(cmd)
}

// pp:data-source live
func addOutlookCalendarFindTimes(group *cobra.Command, flags *rootFlags) {
	var from, to string
	var duration, limit int
	cmd := &cobra.Command{Use: "find-times", Short: "Suggest free times for the signed-in account", Long: "Reads the captured self GetSchedule view and suggests overlapping windows whose 15-minute slots are all code 0 (observed self free). Other codes are excluded. The explicit range is at most 28 days; suggestions are bounded and do not schedule an event or check other people.", Args: cobra.NoArgs, Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live"}, RunE: func(cmd *cobra.Command, _ []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "calendar find-times", "state_change": false})
		}
		_, _, slots, err := cal.AvailabilityRange(from, to)
		if err != nil {
			return usageErr(err)
		}
		if duration < 15 || duration%15 != 0 || duration/15 > slots {
			return usageErr(fmt.Errorf("--duration-minutes must be a positive multiple of 15 no longer than the requested range"))
		}
		if limit < 1 || limit > 200 {
			return usageErr(fmt.Errorf("--limit must be between 1 and 200"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "calendar find-times", "state_change": false, "scope": "signed-in account only", "from": from, "to": to, "duration_minutes": duration, "complete": false})
		}
		detail, request, err := outlookCapture("OUTLOOK_WEB_CALENDAR_AVAILABILITY", true)
		if err != nil {
			return err
		}
		current, _, err := currentSession()
		if err != nil {
			return err
		}
		session, operation, principal, err := cal.PrepareAvailability(detail, request, current, "", from, to)
		if err != nil {
			return err
		}
		response, err := outlookOperation(cmd.Context(), session, operation)
		if err != nil {
			return err
		}
		availability, state := cal.ParseAvailability(response.Body, principal, slots)
		if state != transport.EnvelopeOK {
			return &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
		}
		if availability.Status != "ok" || availability.AvailabilityView == nil {
			kind := transport.KindUpstream
			if availability.ErrorKind == "permission_denied" {
				kind = transport.KindPermission
			}
			return &transport.Error{Kind: kind, Status: response.Status}
		}
		results, truncated, unknown, err := cal.FindSelfTimes(from, to, *availability.AvailabilityView, duration, limit)
		if err != nil {
			return &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "method": "client-side self availability", "scope": "signed-in account only", "from": from, "to": to, "duration_minutes": duration, "interval_minutes": 15, "matching": "all slots have observed self-free code 0", "complete": false, "completeness_reason": "other availability codes and schedules are not interpreted", "unknown_codes_present": unknown, "overlapping_suggestions": true, "limit": limit, "truncated": truncated, "results": results})
	}}
	cmd.Flags().StringVar(&from, "from", "", "Inclusive range start with RFC3339 offset (required)")
	cmd.Flags().StringVar(&to, "to", "", "Exclusive range end with RFC3339 offset (required)")
	cmd.Flags().IntVar(&duration, "duration-minutes", 30, "Duration in minutes, a multiple of 15")
	cmd.Flags().IntVar(&limit, "limit", 50, "Maximum returned suggestions (1–200)")
	group.AddCommand(cmd)
}
