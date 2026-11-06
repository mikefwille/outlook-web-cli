package cli

import (
	"context"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"os"
	"strings"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/account"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/rules"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

type orderedRule struct {
	Rule     rules.Rule
	Priority int
}

func parseRuleOrder(raw []json.RawMessage) ([]orderedRule, string, error) {
	if len(raw) > 1000 {
		return nil, "", fmt.Errorf("rule ordering exceeds the 1000-rule safety bound")
	}
	items := make([]orderedRule, 0, len(raw))
	seen := map[string]bool{}
	hash := sha256.New()
	previous := 0
	for _, row := range raw {
		item, err := rules.Parse(row)
		var detail struct {
			Priority *int `json:"Priority"`
		}
		if err != nil || json.Unmarshal(row, &detail) != nil || detail.Priority == nil || *detail.Priority <= previous || seen[item.ID] {
			return nil, "", fmt.Errorf("rule list has incomplete or inconsistent ordering")
		}
		previous = *detail.Priority
		seen[item.ID] = true
		fmt.Fprintf(hash, "%d:%s:%s;", len(item.ID), item.ID, item.Version)
		items = append(items, orderedRule{Rule: item, Priority: *detail.Priority})
	}
	return items, hex.EncodeToString(hash.Sum(nil)), nil
}

func readRuleOrder(ctx context.Context) ([]orderedRule, string, error) {
	response, err := preferenceRequest(ctx, "rules", false)
	if err != nil {
		return nil, "", err
	}
	rows, status := account.ParseRules(response.Body)
	if status != transport.EnvelopeOK {
		return nil, "", &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
	}
	return parseRuleOrder(rows)
}

func requestedRuleOrder(current []orderedRule, ids []string) ([]rules.OrderMember, error) {
	if len(ids) < 2 || len(ids) > 1000 || len(ids) != len(current) {
		return nil, fmt.Errorf("provide every current rule ID exactly once, with at least two rules")
	}
	byID := make(map[string]orderedRule, len(current))
	for _, item := range current {
		byID[item.Rule.ID] = item
	}
	seen := map[string]bool{}
	out := make([]rules.OrderMember, 0, len(ids))
	changed := false
	for i, id := range ids {
		item, ok := byID[id]
		if !ok || seen[id] {
			return nil, fmt.Errorf("requested order must contain every current rule ID exactly once")
		}
		seen[id] = true
		if current[i].Rule.ID != id {
			changed = true
		}
		out = append(out, rules.OrderMember{ID: id, IdentityDisplay: item.Rule.IdentityDisplay, Enabled: item.Rule.Enabled})
	}
	if !changed {
		return nil, fmt.Errorf("requested rule order is already current")
	}
	return out, nil
}

func ruleContentWithoutPriority(raw json.RawMessage) (string, error) {
	var object map[string]json.RawMessage
	if json.Unmarshal(raw, &object) != nil || object == nil {
		return "", fmt.Errorf("invalid rule")
	}
	delete(object, "Priority")
	canonical, err := json.Marshal(object)
	return string(canonical), err
}

func orderMatches(before, after []orderedRule, desired []rules.OrderMember) bool {
	if len(before) != len(desired) || len(after) != len(desired) {
		return false
	}
	previous := make(map[string]string, len(before))
	for _, item := range before {
		content, err := ruleContentWithoutPriority(item.Rule.Raw)
		if err != nil {
			return false
		}
		previous[item.Rule.ID] = content
	}
	for i, item := range after {
		if item.Rule.ID != desired[i].ID || item.Rule.Enabled != desired[i].Enabled {
			return false
		}
		content, err := ruleContentWithoutPriority(item.Rule.Raw)
		if err != nil || previous[item.Rule.ID] != content {
			return false
		}
	}
	return true
}

func ruleOrderRequest(ctx context.Context, members []rules.OrderMember) (transport.Response, error) {
	prefix := "OUTLOOK_WEB_RULES_REORDER"
	detail := os.Getenv(prefix + "_DETAILS")
	if os.Getenv(prefix+"_REQUEST") != "" {
		return transport.Response{}, fmt.Errorf("rule ordering template uses a header body")
	}
	if detail == "" {
		detail = "runtime:rules_reorder"
	}
	current, _, err := currentSession()
	if err != nil {
		return transport.Response{}, err
	}
	s, op, err := rules.PrepareReorder(detail, current, members)
	if err != nil {
		return transport.Response{}, err
	}
	return outlookOperation(ctx, s, op)
}

// pp:data-source live
func newRuleReorder(flags *rootFlags) *cobra.Command {
	var ids []string
	var version string
	var execute, allowStateReapply bool
	cmd := &cobra.Command{Use: "reorder", Short: "Preview or reorder all inbox rules", Long: "Supply every current rule ID once in the desired order using repeated --id flags. Outlook's observed ordering request also reapplies each rule's enabled state. Execution requires --allow-state-reapply and the snapshot version from preview. The version check is not atomic; concurrent state changes could be overwritten. This does not run rules on existing messages. Behavior in mailboxes with sweep rules has not been established.", Annotations: map[string]string{"pp:data-source": "live", "pp:happy-args": "--id=example-value"}, RunE: func(cmd *cobra.Command, args []string) error {
		if len(args) != 0 {
			return usageErr(fmt.Errorf("rules reorder takes flags only"))
		}
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "rules reorder"})
		}
		if len(ids) < 2 || len(ids) > 1000 {
			return usageErr(fmt.Errorf("supply at least two --id values in the desired order"))
		}
		seen := map[string]bool{}
		for _, id := range ids {
			if len(id) > 4096 || strings.TrimSpace(id) != id || id == "" || strings.ContainsAny(id, "\x00\r\n") || seen[id] {
				return usageErr(fmt.Errorf("--id values must be distinct valid rule IDs"))
			}
			seen[id] = true
		}
		preview := map[string]any{"preview": true, "operation": "rules reorder", "requested_ids": ids, "reapplies_enabled_states": true, "version_check": "before write; not atomic", "run_existing_mail": false}
		if flags.dryRun {
			return flags.printJSON(cmd, preview)
		}
		if execute && !allowStateReapply {
			return usageErr(fmt.Errorf("--allow-state-reapply is required because ordering reapplies enabled states"))
		}
		if execute && !ruleVersionPattern.MatchString(version) {
			return usageErr(fmt.Errorf("--version must be the snapshot fingerprint from a fresh preview"))
		}
		current, snapshot, err := readRuleOrder(cmd.Context())
		if err != nil {
			return err
		}
		members, err := requestedRuleOrder(current, ids)
		if err != nil {
			return usageErr(err)
		}
		before := make([]map[string]any, 0, len(current))
		after := make([]map[string]any, 0, len(members))
		for _, item := range current {
			before = append(before, map[string]any{"id": item.Rule.ID, "name": item.Rule.Name, "enabled": item.Rule.Enabled, "priority": item.Priority})
		}
		for _, item := range members {
			after = append(after, map[string]any{"id": item.ID, "enabled": item.Enabled})
		}
		preview["current_order"], preview["requested_order"], preview["version"] = before, after, snapshot
		if !execute {
			return flags.printJSON(cmd, preview)
		}
		if version != snapshot {
			return &transport.Error{Kind: transport.KindConflict}
		}
		if err = mutationInHarness(); err != nil {
			return err
		}
		if _, err = ruleOrderRequest(cmd.Context(), members); err != nil {
			return err
		}
		observed, _, err := readRuleOrder(cmd.Context())
		result := map[string]any{"source": "live", "accepted": true, "operation": "rules reorder", "readback_status": "unconfirmed"}
		if err == nil && orderMatches(current, observed, members) {
			result["readback_status"] = "verified"
		} else {
			result["next_step"] = "Inspect the rule order and enabled states without retrying the write."
		}
		return flags.printJSON(cmd, result)
	}}
	cmd.Flags().StringArrayVar(&ids, "id", nil, "All rule IDs in the desired order; repeat this flag")
	cmd.Flags().StringVar(&version, "version", "", "Expected snapshot fingerprint from a fresh reorder preview")
	cmd.Flags().BoolVar(&allowStateReapply, "allow-state-reapply", false, "Acknowledge reapplication of all enabled states")
	cmd.Flags().BoolVar(&execute, "execute", false, "Apply the previewed ordering once")
	return cmd
}
