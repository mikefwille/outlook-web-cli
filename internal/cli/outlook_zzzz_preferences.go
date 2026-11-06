package cli

import (
	"context"
	"fmt"
	"os"
	"strings"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/account"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func preferenceRequest(ctx context.Context, kind string, focused bool) (transport.Response, error) {
	name := map[string]string{"rules": "inbox_rules_list", "auto-replies": "preferences_auto_replies", "focused-update": "preference_focused_update"}[kind]
	if name == "" {
		return transport.Response{}, fmt.Errorf("unknown preference operation")
	}
	prefix := "OUTLOOK_WEB_" + strings.ToUpper(name)
	detail := os.Getenv(prefix + "_DETAILS")
	if os.Getenv(prefix+"_REQUEST") != "" {
		return transport.Response{}, fmt.Errorf("preference template uses a header body")
	}
	if detail == "" {
		detail = "runtime:" + name
	}
	current, _, err := currentSession()
	if err != nil {
		return transport.Response{}, err
	}
	s, op, err := account.PreparePreference(kind, detail, current, focused)
	if err != nil {
		return transport.Response{}, err
	}
	return outlookOperation(ctx, s, op)
}

// pp:data-source live
func init() {
	registerNovelCommand(func(root *cobra.Command, flags *rootFlags) {
		settings := findSubcommand(root, "settings")
		if settings == nil {
			settings = &cobra.Command{Use: "settings", Short: "Read and update observed mailbox settings"}
			root.AddCommand(settings)
		}
		settings.AddCommand(newFocusedInboxUpdate(flags))
		settings.AddCommand(&cobra.Command{Use: "auto-replies", Short: "Read the automatic-reply configuration", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live"}, RunE: func(cmd *cobra.Command, args []string) error {
			if len(args) != 0 {
				return usageErr(fmt.Errorf("settings auto-replies takes no arguments"))
			}
			if flags.dryRun {
				return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "settings auto-replies"})
			}
			response, err := preferenceRequest(cmd.Context(), "auto-replies", false)
			if err != nil {
				return err
			}
			value, state := account.ParseAutoReplies(response.Body)
			if state != transport.EnvelopeOK {
				return &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
			}
			return flags.printJSON(cmd, map[string]any{"source": "live", "auto_replies": value})
		}})
		rules := findSubcommand(root, "rules")
		if rules == nil {
			rules = &cobra.Command{Use: "rules", Short: "Read inbox rules"}
			root.AddCommand(rules)
		}
		rules.AddCommand(&cobra.Command{Use: "list", Short: "List the mailbox's inbox rules", Long: "Returns rule objects from the observed rule read without interpreting unobserved conditions or actions. Rule creation and renaming support only the observed narrow subject and category shape.", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live"}, RunE: func(cmd *cobra.Command, args []string) error {
			if len(args) != 0 {
				return usageErr(fmt.Errorf("rules list takes no arguments"))
			}
			if flags.dryRun {
				return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "rules list"})
			}
			response, err := preferenceRequest(cmd.Context(), "rules", false)
			if err != nil {
				return err
			}
			items, state := account.ParseRules(response.Body)
			if state != transport.EnvelopeOK {
				return &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
			}
			return flags.printJSON(cmd, map[string]any{"source": "live", "rules": items, "note": "Unknown rule fields are preserved; mutation support is limited to captured shapes."})
		}})
	})
}

func newFocusedInboxUpdate(flags *rootFlags) *cobra.Command {
	var focused, expected, execute bool
	cmd := &cobra.Command{Use: "update", Short: "Preview or change the focused-inbox setting", Long: "Updates only Focused Inbox. Preview reads the current value; execution requires --expect-focused-inbox and checks it before the write. This check is not atomic. Signature and automatic replies are not changed.", Annotations: map[string]string{"pp:data-source": "live", "pp:happy-args": "--focused-inbox=true"}, RunE: func(cmd *cobra.Command, args []string) error {
		if len(args) != 0 {
			return usageErr(fmt.Errorf("settings update takes flags only"))
		}
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "settings update"})
		}
		if !cmd.Flags().Changed("focused-inbox") {
			return usageErr(fmt.Errorf("--focused-inbox must be explicitly true or false"))
		}
		preview := map[string]any{"preview": true, "operation": "settings update", "focused_inbox": focused, "version_check": "current value before write; not atomic"}
		if flags.dryRun {
			return flags.printJSON(cmd, preview)
		}
		data, err := readOutlookConfiguration(cmd.Context())
		if err != nil {
			return err
		}
		current, state := account.ParseSettings(data)
		if state != transport.EnvelopeOK || current.FocusedInboxEnabled == nil {
			return &transport.Error{Kind: transport.KindUpstream}
		}
		preview["current_focused_inbox"] = *current.FocusedInboxEnabled
		if !execute {
			return flags.printJSON(cmd, preview)
		}
		if !cmd.Flags().Changed("expect-focused-inbox") {
			return usageErr(fmt.Errorf("--expect-focused-inbox is required with --execute"))
		}
		if *current.FocusedInboxEnabled != expected {
			return &transport.Error{Kind: transport.KindConflict}
		}
		if expected == focused {
			return usageErr(fmt.Errorf("focused inbox is already set to the requested value"))
		}
		if err = mutationInHarness(); err != nil {
			return err
		}
		if _, err = preferenceRequest(cmd.Context(), "focused-update", focused); err != nil {
			return err
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "accepted": true, "focused_inbox": focused, "next_step": "Read settings independently to verify the change."})
	}}
	cmd.Flags().BoolVar(&focused, "focused-inbox", false, "Desired focused-inbox state")
	cmd.Flags().BoolVar(&expected, "expect-focused-inbox", false, "Expected current state from a fresh preview")
	cmd.Flags().BoolVar(&execute, "execute", false, "Apply the previewed setting once")
	return cmd
}
