package cli

import (
	"context"
	"fmt"
	"time"

	"github.com/spf13/cobra"
	cal "github.com/mikefwille/outlook-web-cli/internal/outlook/calendar"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func calendarSeriesOccurrences(ctx context.Context, id, from, to string) (cal.Occurrences, error) {
	detail, _, err := outlookCapture("OUTLOOK_WEB_CALENDAR_OCCURRENCES", false)
	if err != nil {
		return cal.Occurrences{}, err
	}
	current, _, err := currentSession()
	if err != nil {
		return cal.Occurrences{}, err
	}
	session, op, err := cal.PrepareOccurrences(detail, current, id, from, to)
	if err != nil {
		return cal.Occurrences{}, err
	}
	response, err := outlookOperation(ctx, session, op)
	if err != nil {
		return cal.Occurrences{}, err
	}
	result, state := cal.ParseOccurrences(response.Body)
	if state != transport.EnvelopeOK {
		return cal.Occurrences{}, &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
	}
	return result, nil
}

func init() {
	registerNovelCommand(func(root *cobra.Command, flags *rootFlags) {
		group := findSubcommand(root, "calendar")
		if group == nil {
			group = &cobra.Command{Use: "calendar", Short: "Work with calendar appointments"}
			root.AddCommand(group)
		}
		addOutlookOccurrences(group, flags)
	})
}

// pp:data-source live
func addOutlookOccurrences(group *cobra.Command, flags *rootFlags) {
	var from, to string
	cmd := &cobra.Command{Use: "occurrences <series-id>", Short: "Expand a recurring series within a time range", Long: "Returns the series master, regular occurrences, exceptions and cancellations in the requested RFC3339 range using the captured Mountain-time calendar query.", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live", "pp:happy-args": "id=example-value"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "calendar occurrences"})
		}
		id, err := oneID(args)
		if err != nil {
			return err
		}
		a, e1 := time.Parse(time.RFC3339, from)
		b, e2 := time.Parse(time.RFC3339, to)
		if e1 != nil || e2 != nil || !a.Before(b) {
			return usageErr(fmt.Errorf("--from and --to need RFC3339 offsets and an end after the start"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "calendar occurrences", "id": id, "from": from, "to": to})
		}
		result, err := calendarSeriesOccurrences(cmd.Context(), id, from, to)
		if err != nil {
			return err
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "series_id": id, "from": from, "to": to, "result": result})
	}}
	cmd.Flags().StringVar(&from, "from", "", "Range start with RFC3339 offset")
	cmd.Flags().StringVar(&to, "to", "", "Range end with RFC3339 offset")
	group.AddCommand(cmd)
}
