package cli

import (
	"context"
	"encoding/json"
	"fmt"
	"os"
	"regexp"
	"strings"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/account"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/rules"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

var ruleVersionPattern = regexp.MustCompile(`^[a-f0-9]{64}$`)

func ruleCreatedResult(body []byte) (map[string]any, error) {
	var wrapper struct {
		Rule json.RawMessage `json:"InboxRule"`
	}
	if json.Unmarshal(body, &wrapper) != nil {
		return nil, fmt.Errorf("invalid created rule")
	}
	item, err := rules.Parse(wrapper.Rule)
	if err != nil {
		return nil, err
	}
	return map[string]any{
		"source": "live", "accepted": true, "identity_status": "list_read_required",
		"rule":      map[string]any{"name": item.Name, "enabled": item.Enabled, "subject_contains": item.Subject, "apply_category": item.Category},
		"next_step": "The accepted rule is enabled. Obtain its usable ID from a fresh rules list and act only if the match is unambiguous.",
	}, nil
}

func readRule(ctx context.Context, id string) (rules.Rule, error) {
	response, err := preferenceRequest(ctx, "rules", false)
	if err != nil {
		return rules.Rule{}, err
	}
	items, state := account.ParseRules(response.Body)
	if state != transport.EnvelopeOK {
		return rules.Rule{}, &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
	}
	var result *rules.Rule
	for _, raw := range items {
		item, err := rules.Parse(raw)
		if err != nil {
			return rules.Rule{}, fmt.Errorf("an inbox rule has an unsupported identity shape")
		}
		if item.ID == id {
			if result != nil {
				return rules.Rule{}, fmt.Errorf("inbox rule identity is not unique")
			}
			copy := item
			result = &copy
		}
	}
	if result == nil {
		return rules.Rule{}, fmt.Errorf("inbox rule not found")
	}
	return *result, nil
}

func ruleRequest(ctx context.Context, kind string, p rules.Params) (transport.Response, error) {
	name := map[string]string{"create": "rule_create", "update": "rule_update", "enable": "rule_enable", "disable": "rule_disable", "delete": "rule_delete"}[kind]
	if name == "" {
		return transport.Response{}, fmt.Errorf("unknown rule operation")
	}
	prefix := "OUTLOOK_WEB_" + strings.ToUpper(name)
	detail := os.Getenv(prefix + "_DETAILS")
	if os.Getenv(prefix+"_REQUEST") != "" {
		return transport.Response{}, fmt.Errorf("rule template uses a header body")
	}
	if detail == "" {
		detail = "runtime:" + name
	}
	current, _, err := currentSession()
	if err != nil {
		return transport.Response{}, err
	}
	session, op, err := rules.Prepare(kind, detail, current, p)
	if err != nil {
		return transport.Response{}, err
	}
	return outlookOperation(ctx, session, op)
}

// pp:data-source live
func init() {
	registerNovelCommand(func(root *cobra.Command, flags *rootFlags) {
		group := findSubcommand(root, "rules")
		if group == nil {
			group = &cobra.Command{Use: "rules", Short: "Work with inbox rules"}
			root.AddCommand(group)
		}
		group.AddCommand(&cobra.Command{Use: "get <id>", Short: "Read one inbox rule and its version fingerprint", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live", "pp:happy-args": "id=example-value"}, RunE: func(cmd *cobra.Command, args []string) error {
			if flags.dryRun && requiredInputWaived() {
				return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "rules get"})
			}
			id, err := oneID(args)
			if err != nil {
				return err
			}
			if flags.dryRun {
				return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "rules get", "id": id})
			}
			item, err := readRule(cmd.Context(), id)
			if err != nil {
				return err
			}
			return flags.printJSON(cmd, map[string]any{"source": "live", "rule": item})
		}})
		for _, kind := range []string{"create", "enable", "disable", "update", "delete"} {
			group.AddCommand(newRuleAction(flags, kind))
		}
		group.AddCommand(newRuleReorder(flags))
	})
}

func ruleInput(s string) bool {
	return len(s) > 0 && len(s) <= 255 && strings.TrimSpace(s) == s && !strings.ContainsAny(s, "\x00\r\n")
}

// composeRuleUpdate is called only after the existing complete-shape guard.
// SetInboxRule sends all three fields, so omitted inputs must retain the values
// from the freshly read rule instead of being cleared.
func composeRuleUpdate(current rules.Rule, p rules.Params, nameChanged, subjectChanged, categoryChanged bool) rules.Params {
	if !nameChanged {
		p.Name = current.Name
	}
	if !subjectChanged {
		p.Subject = current.Subject[0]
	}
	if !categoryChanged {
		p.Category = current.Category[0]
	}
	return p
}

func newRuleAction(flags *rootFlags, kind string) *cobra.Command {
	var name, subject, category, version string
	var execute, allowEnabled, restoreDisabled bool
	use := kind + " <id>"
	if kind == "create" {
		use = "create"
	}
	cmd := &cobra.Command{Use: use, Short: "Preview or " + kind + " an inbox rule", Long: "Create supports one subject-contains condition and one category action. The service creates the rule enabled: execution requires --allow-enabled, and no existing mail is run. Obtain its usable ID from a fresh rules list; the creation-response identity differs from the list identity. Existing-rule writes compare a fresh fingerprint before writing; the check is not atomic. Updates support only disabled rules with the complete observed simple shape. Subject or category edits may leave the rule enabled: execution requires --allow-enabled. --restore-disabled requests a checked disable if the readback shows it enabled.", Annotations: map[string]string{"pp:data-source": "live", "pp:happy-args": "id=example-value"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "rules " + kind})
		}
		id := ""
		if kind == "create" {
			if len(args) != 0 {
				return usageErr(fmt.Errorf("rules create takes flags only"))
			}
		} else {
			var err error
			id, err = oneID(args)
			if err != nil {
				return err
			}
			if len(id) > 4096 || strings.ContainsAny(id, "\x00\r\n") {
				return usageErr(fmt.Errorf("rule ID has an invalid format"))
			}
		}
		nameChanged := (kind == "create" || kind == "update") && cmd.Flags().Changed("name")
		subjectChanged := (kind == "create" || kind == "update") && cmd.Flags().Changed("subject-contains")
		categoryChanged := (kind == "create" || kind == "update") && cmd.Flags().Changed("category")
		if kind == "create" && !ruleInput(name) || kind == "update" && nameChanged && !ruleInput(name) {
			return usageErr(fmt.Errorf("--name must be 1–255 characters without leading or trailing whitespace"))
		}
		if kind == "create" && (!ruleInput(subject) || !ruleInput(category)) || kind == "update" && (subjectChanged && !ruleInput(subject) || categoryChanged && !ruleInput(category)) {
			return usageErr(fmt.Errorf("--subject-contains and --category must each be 1–255 characters"))
		}
		if kind == "update" && !nameChanged && !subjectChanged && !categoryChanged {
			return usageErr(fmt.Errorf("provide --name, --subject-contains, or --category"))
		}
		if kind == "update" && execute && !flags.dryRun && (subjectChanged || categoryChanged) && !allowEnabled {
			return usageErr(fmt.Errorf("--allow-enabled is required because a subject or category update may enable the rule"))
		}
		preview := map[string]any{"preview": true, "operation": "rules " + kind, "id": id}
		p := rules.Params{ID: id, Name: name, Subject: subject, Category: category}
		var before rules.Rule
		if kind == "create" {
			preview["name"] = name
			preview["subject_contains"] = subject
			preview["category"] = category
			preview["enabled_on_creation"] = true
			preview["run_existing_mail"] = false
			preview["requires_allow_enabled"] = true
		} else {
			preview["version_check"] = "before write; not atomic"
			if kind == "enable" {
				preview["enabled_after_execution"] = true
			}
			if kind == "update" {
				changes := map[string]any{}
				if nameChanged {
					preview["new_name"] = name
					changes["name"] = name
				}
				if subjectChanged {
					changes["subject_contains"] = subject
				}
				if categoryChanged {
					changes["category"] = category
				}
				preview["changes"] = changes
				preview["may_enable_rule"] = true
				preview["requires_allow_enabled"] = subjectChanged || categoryChanged
				preview["restore_disabled_if_enabled"] = restoreDisabled
			}
		}
		if flags.dryRun {
			return flags.printJSON(cmd, preview)
		}
		if kind == "create" {
			if !execute {
				return flags.printJSON(cmd, preview)
			}
			if !allowEnabled {
				return usageErr(fmt.Errorf("--allow-enabled is required: the new rule is enabled immediately"))
			}
		} else {
			item, err := readRule(cmd.Context(), id)
			if err != nil {
				return err
			}
			preview["current"] = item
			before = item
			if kind == "disable" && !item.Enabled {
				return fmt.Errorf("rule is already disabled")
			}
			if kind == "enable" && item.Enabled {
				return fmt.Errorf("rule is already enabled")
			}
			if kind == "update" && !item.SafeToRename() {
				return fmt.Errorf("only a disabled rule with the observed single-subject, single-category shape can be updated without dropping fields")
			}
			if !execute {
				return flags.printJSON(cmd, preview)
			}
			if !ruleVersionPattern.MatchString(version) {
				return usageErr(fmt.Errorf("--version must be the fingerprint from a fresh rule preview"))
			}
			if version != item.Version {
				return &transport.Error{Kind: transport.KindConflict}
			}
			p.IdentityDisplay = item.IdentityDisplay
			if kind == "update" {
				p = composeRuleUpdate(item, p, nameChanged, subjectChanged, categoryChanged)
			}
		}
		if err := mutationInHarness(); err != nil {
			return err
		}
		response, err := ruleRequest(cmd.Context(), kind, p)
		if err != nil {
			return err
		}
		if kind == "create" {
			result, err := ruleCreatedResult(response.Body)
			if err != nil {
				return &transport.Error{Kind: transport.KindAmbiguous, Status: response.Status}
			}
			return flags.printJSON(cmd, result)
		}
		if kind == "update" {
			return flags.printJSON(cmd, ruleUpdateOutcome(cmd.Context(), id, p, before, restoreDisabled, readRule, ruleRequest))
		}
		if kind == "enable" {
			observed, readErr := readRule(cmd.Context(), id)
			if readErr != nil || !observed.Enabled {
				return flags.printJSON(cmd, map[string]any{"source": "live", "accepted": true, "operation": "rules enable", "id": id, "readback_status": "unconfirmed", "next_step": "Inspect the rule without retrying the write."})
			}
			return flags.printJSON(cmd, map[string]any{"source": "live", "accepted": true, "operation": "rules enable", "id": id, "readback_status": "verified", "enabled": true})
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "accepted": true, "operation": "rules " + kind, "id": id, "next_step": "Read the rule list independently to verify the change."})
	}}
	if kind == "create" || kind == "update" {
		cmd.Flags().StringVar(&name, "name", "", "Rule display name")
	}
	if kind == "create" || kind == "update" {
		cmd.Flags().StringVar(&subject, "subject-contains", "", "Single subject text condition")
		cmd.Flags().StringVar(&category, "category", "", "Single category action")
	}
	if kind == "create" {
		cmd.Flags().BoolVar(&allowEnabled, "allow-enabled", false, "Acknowledge that the new rule is enabled immediately")
		cmd.Annotations["pp:happy-args"] = "--name=Example;--subject-contains=Example;--category=Example"
	}
	if kind == "update" {
		cmd.Flags().BoolVar(&allowEnabled, "allow-enabled", false, "Acknowledge that subject or category changes may enable the rule")
		cmd.Flags().BoolVar(&restoreDisabled, "restore-disabled", false, "Explicitly disable the rule again if verified enabled after the update")
	}
	if kind != "create" {
		cmd.Flags().StringVar(&version, "version", "", "Expected fingerprint from a fresh rule read")
	}
	cmd.Flags().BoolVar(&execute, "execute", false, "Apply the previewed action once")
	return cmd
}
