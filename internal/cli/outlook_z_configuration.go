package cli

import (
	"context"
	"fmt"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/account"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func readOutlookConfiguration(ctx context.Context) ([]byte, error) {
	detail, _, err := outlookCapture("OUTLOOK_WEB_ACCOUNT_ALIASES", false)
	if err != nil {
		return nil, err
	}
	current, _, err := currentSession()
	if err != nil {
		return nil, err
	}
	s, op, err := account.PrepareAliases(detail, current)
	if err != nil {
		return nil, err
	}
	response, err := outlookOperation(ctx, s, op)
	if err != nil {
		return nil, err
	}
	return response.Body, nil
}

// pp:data-source live
func init() {
	registerNovelCommand(func(root *cobra.Command, flags *rootFlags) {
		categories := &cobra.Command{Use: "categories", Short: "Read the mailbox's categories"}
		list := &cobra.Command{Use: "list", Short: "List the mailbox's master categories", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live"}, RunE: func(cmd *cobra.Command, args []string) error {
			if len(args) != 0 {
				return usageErr(fmt.Errorf("categories list takes no arguments"))
			}
			if flags.dryRun {
				return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "categories list"})
			}
			data, err := readOutlookConfiguration(cmd.Context())
			if err != nil {
				return err
			}
			items, state := account.ParseCategories(data)
			if state != transport.EnvelopeOK {
				return &transport.Error{Kind: transport.KindUpstream}
			}
			return flags.printJSON(cmd, map[string]any{"source": "live", "categories": items})
		}}
		get := &cobra.Command{Use: "get <id>", Short: "Read one category by ID", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live", "pp:happy-args": "id=example-value"}, RunE: func(cmd *cobra.Command, args []string) error {
			if flags.dryRun && requiredInputWaived() {
				return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "categories get"})
			}
			id, err := oneID(args)
			if err != nil {
				return err
			}
			if flags.dryRun {
				return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "categories get", "id": id})
			}
			data, err := readOutlookConfiguration(cmd.Context())
			if err != nil {
				return err
			}
			items, state := account.ParseCategories(data)
			if state != transport.EnvelopeOK {
				return &transport.Error{Kind: transport.KindUpstream}
			}
			for _, item := range items {
				if item.ID == id {
					return flags.printJSON(cmd, map[string]any{"source": "live", "category": item})
				}
			}
			return fmt.Errorf("category not found in the current mailbox")
		}}
		categories.AddCommand(list, get)
		root.AddCommand(categories)
		settings := &cobra.Command{Use: "settings", Short: "Read captured Outlook settings"}
		settings.AddCommand(&cobra.Command{Use: "get", Short: "Read signature and focused-inbox settings", Long: "Returns the observed signature and focused-inbox fields. Other settings, including automatic replies, are not included in this read.", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live"}, RunE: func(cmd *cobra.Command, args []string) error {
			if len(args) != 0 {
				return usageErr(fmt.Errorf("settings get takes no arguments"))
			}
			if flags.dryRun {
				return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "settings get"})
			}
			data, err := readOutlookConfiguration(cmd.Context())
			if err != nil {
				return err
			}
			value, state := account.ParseSettings(data)
			if state != transport.EnvelopeOK {
				return &transport.Error{Kind: transport.KindUpstream}
			}
			return flags.printJSON(cmd, map[string]any{"source": "live", "settings": value, "not_included": []string{"automatic_replies"}})
		}})
		root.AddCommand(settings)
	})
}
