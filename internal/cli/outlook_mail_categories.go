package cli

import (
	"encoding/json"
	"fmt"
	"strconv"
	"strings"
	"unicode/utf8"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/account"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/exchange"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/mailactions"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

type messageCategoryState struct {
	messageActionState
	Categories *[]string `json:"Categories"`
}

func newMailCategoriesCmd(flags *rootFlags) *cobra.Command {
	var action, name, key string
	var execute bool
	group := &cobra.Command{Use: "categories <id>", Short: "Read or change categories on one message", Long: "Read the message categories by default. Use --action add or remove with --category to preview an update, then provide the preview change key and --execute. Named get, add and remove subcommands are also available.", Annotations: map[string]string{"pp:data-source": "live"}}
	for _, kind := range []string{"get", "add", "remove"} {
		group.AddCommand(newMailCategoryCmd(flags, kind))
	}
	group.Flags().StringVar(&action, "action", "get", "Action: get, add, or remove")
	group.Flags().StringVar(&name, "category", "", "Existing master category name for add/remove")
	group.Flags().StringVar(&key, "change-key", "", "Expected message version from preview")
	group.Flags().BoolVar(&execute, "execute", false, "Apply a previewed add/remove once")
	group.RunE = func(cmd *cobra.Command, args []string) error {
		if action != "get" && action != "add" && action != "remove" {
			return usageErr(fmt.Errorf("--action must be get, add, or remove"))
		}
		if action == "get" && (name != "" || key != "" || execute) {
			return usageErr(fmt.Errorf("category, change-key and execute apply only to add/remove"))
		}
		child := findSubcommand(group, action)
		if child == nil {
			return fmt.Errorf("category action unavailable")
		}
		child.SetOut(cmd.OutOrStdout())
		child.SetErr(cmd.ErrOrStderr())
		child.SetContext(cmd.Context())
		if action != "get" {
			if err := child.Flags().Set("category", name); err != nil {
				return err
			}
			if err := child.Flags().Set("change-key", key); err != nil {
				return err
			}
			if err := child.Flags().Set("execute", strconv.FormatBool(execute)); err != nil {
				return err
			}
		}
		return child.RunE(child, args)
	}
	return group
}

// pp:data-source live
func newMailCategoryCmd(flags *rootFlags, kind string) *cobra.Command {
	var name, key string
	var execute bool
	cmd := &cobra.Command{Use: kind + " <id>", Short: kind + " categories on one message", Long: "Requests the message's current category array and preserves all other categories. Category changes use Outlook's observed AlwaysOverwrite action, so the preceding version check is not atomic.", Annotations: map[string]string{"pp:data-source": "live"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "mail categories " + kind})
		}
		id, err := oneID(args)
		if err != nil {
			return err
		}
		if kind != "get" && (name == "" || name != strings.TrimSpace(name) || len(name) > 255 || !utf8.ValidString(name) || strings.ContainsAny(name, "\x00\r\n")) {
			return usageErr(fmt.Errorf("--category must be a valid category name"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "mail categories " + kind, "id": id, "category": name})
		}
		response, err := getMessageWithFields(cmd.Context(), id, false, true)
		if err != nil {
			return err
		}
		got, raw, status := exchange.ParseMessageGet(response.Body)
		if status != transport.EnvelopeOK || got != id {
			return &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
		}
		var item messageCategoryState
		if json.Unmarshal(raw, &item) != nil || item.ItemID.ID != id || item.ItemID.ChangeKey == "" {
			return fmt.Errorf("message version is unavailable")
		}
		if item.Categories == nil {
			if kind == "get" {
				return flags.printJSON(cmd, map[string]any{"source": "live", "id": id, "change_key": item.ItemID.ChangeKey, "categories_complete": false})
			}
			return fmt.Errorf("message category array was not returned; no changes made")
		}
		current := append([]string{}, (*item.Categories)...)
		if len(current) > 128 {
			return fmt.Errorf("message has too many categories to update safely")
		}
		if kind == "get" {
			return flags.printJSON(cmd, map[string]any{"source": "live", "id": id, "change_key": item.ItemID.ChangeKey, "categories": current, "categories_complete": true})
		}
		matches := 0
		for _, existing := range current {
			if strings.EqualFold(existing, name) {
				matches++
			}
		}
		if matches > 1 {
			return fmt.Errorf("category name is ambiguous on this message")
		}
		if kind == "add" && matches == 0 {
			data, err := readOutlookConfiguration(cmd.Context())
			if err != nil {
				return err
			}
			masters, state := account.ParseCategories(data)
			if state != transport.EnvelopeOK {
				return &transport.Error{Kind: transport.KindUpstream}
			}
			found := ""
			for _, category := range masters {
				if strings.EqualFold(category.Name, name) {
					if found != "" {
						return fmt.Errorf("category name is ambiguous in the master list")
					}
					found = category.Name
				}
			}
			if found == "" {
				return fmt.Errorf("category not found in the master list")
			}
			name = found
		}
		next := make([]string, 0, len(current)+1)
		for _, existing := range current {
			if kind != "remove" || !strings.EqualFold(existing, name) {
				next = append(next, existing)
			}
		}
		if kind == "add" && matches == 0 {
			next = append(next, name)
		}
		changed := (kind == "add" && matches == 0) || (kind == "remove" && matches == 1)
		preview := map[string]any{"preview": true, "operation": "mail categories " + kind, "scope": "single message", "id": id, "change_key": item.ItemID.ChangeKey, "before": current, "after": next, "changed": changed, "version_check": "before write; not atomic"}
		if !execute {
			return flags.printJSON(cmd, preview)
		}
		if key == "" {
			return usageErr(fmt.Errorf("--change-key from the preview is required"))
		}
		if key != item.ItemID.ChangeKey {
			return &transport.Error{Kind: transport.KindConflict}
		}
		if !changed {
			return flags.printJSON(cmd, map[string]any{"source": "live", "accepted": false, "changed": false, "id": id})
		}
		if err := mutationInHarness(); err != nil {
			return err
		}
		detail, _, err := outlookCapture("OUTLOOK_WEB_MAIL_CATEGORIES", false)
		if err != nil {
			return err
		}
		session, _, err := currentSession()
		if err != nil {
			return err
		}
		bound, operation, err := mailactions.PrepareCategories(detail, session, id, next)
		if err != nil {
			return err
		}
		result, err := outlookOperation(cmd.Context(), bound, operation)
		if err != nil {
			return err
		}
		updated, state := mailactions.ParseCategoriesResult(result.Body)
		if state != transport.EnvelopeOK || updated.ID != id {
			return &transport.Error{Kind: transport.KindAmbiguous, Status: result.Status}
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "accepted": true, "changed": true, "operation": "mail categories " + kind, "scope": "single message", "result": updated, "categories": next})
	}}
	if kind == "get" {
		cmd.Annotations["mcp:read-only"] = "true"
	} else {
		cmd.Flags().StringVar(&name, "category", "", "Existing master category name")
		cmd.Flags().StringVar(&key, "change-key", "", "Expected version from preview")
		cmd.Flags().BoolVar(&execute, "execute", false, "Apply the previewed category change once")
	}
	return cmd
}
