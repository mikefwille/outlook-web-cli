package cli

import (
	"fmt"

	"github.com/spf13/cobra"
	cal "github.com/mikefwille/outlook-web-cli/internal/outlook/calendar"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func init() {
	registerNovelCommand(func(root *cobra.Command, flags *rootFlags) {
		group := findSubcommand(root, "calendar")
		if group == nil {
			group = &cobra.Command{Use: "calendar", Short: "Work with calendar appointments"}
			root.AddCommand(group)
		}
		addOutlookAvailability(group, flags)
	})
}

// pp:data-source live
func addOutlookAvailability(group *cobra.Command, flags *rootFlags) {
	var principal, from, to string
	cmd := &cobra.Command{Use: "availability", Short: "Read free/busy for one accessible principal", Long: "Queries Outlook's captured GetSchedule service in 15-minute intervals. Defaults to the signed-in account; --principal selects one address. The range is start-inclusive and end-exclusive, aligned to 15 minutes and at most 28 days. Codes 0 (free) and 2 (busy) were verified for the self query; other codes and other-principal permissions remain unverified.", Args: cobra.NoArgs, Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "calendar availability"})
		}
		_, _, slots, err := cal.AvailabilityRange(from, to)
		if err != nil {
			return usageErr(err)
		}
		if principal != "" && !cal.ValidAvailabilityPrincipal(principal) {
			return usageErr(fmt.Errorf("--principal must be one complete email address"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "calendar availability", "principal": func() string {
				if principal == "" {
					return "current account"
				}
				return principal
			}(), "from": from, "to": to, "interval_minutes": 15})
		}
		detail, request, err := outlookCapture("OUTLOOK_WEB_CALENDAR_AVAILABILITY", true)
		if err != nil {
			return err
		}
		current, _, err := currentSession()
		if err != nil {
			return err
		}
		session, op, resolved, err := cal.PrepareAvailability(detail, request, current, principal, from, to)
		if err != nil {
			return err
		}
		response, err := outlookOperation(cmd.Context(), session, op)
		if err != nil {
			return err
		}
		result, state := cal.ParseAvailability(response.Body, resolved, slots)
		if state != transport.EnvelopeOK {
			return &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "from": from, "to": to, "interval_minutes": 15, "availability": result, "observed_codes": map[string]string{"0": "free in the self fixture", "2": "busy in the self fixture"}})
	}}
	cmd.Flags().StringVar(&principal, "principal", "", "One accessible principal email; omitted uses the current account")
	cmd.Flags().StringVar(&from, "from", "", "Inclusive range start with RFC3339 offset")
	cmd.Flags().StringVar(&to, "to", "", "Exclusive range end with RFC3339 offset")
	group.AddCommand(cmd)
}
