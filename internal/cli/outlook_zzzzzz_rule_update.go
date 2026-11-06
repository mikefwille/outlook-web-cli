package cli

import (
	"context"
	"encoding/json"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/rules"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

type ruleReader func(context.Context, string) (rules.Rule, error)
type ruleWriter func(context.Context, string, rules.Params) (transport.Response, error)

func comparableRuleUpdateFields(raw json.RawMessage) (string, bool) {
	var object map[string]json.RawMessage
	if json.Unmarshal(raw, &object) != nil || object == nil {
		return "", false
	}
	for _, key := range []string{"Name", "SubjectContainsWords", "ApplyCategory", "Enabled", "Description", "DescriptionTimeFormat", "DescriptionTimeZone", "Identity"} {
		delete(object, key)
	}
	encoded, err := json.Marshal(object)
	return string(encoded), err == nil
}

func ruleMatchesUpdate(item rules.Rule, id string, desired rules.Params, before rules.Rule) bool {
	if item.ID != id || item.Name != desired.Name || len(item.Subject) != 1 || item.Subject[0] != desired.Subject || len(item.Category) != 1 || item.Category[0] != desired.Category {
		return false
	}
	oldFields, ok := comparableRuleUpdateFields(before.Raw)
	if !ok {
		return false
	}
	newFields, ok := comparableRuleUpdateFields(item.Raw)
	return ok && oldFields == newFields
}

// The update has already been accepted when this runs. Every failure reports
// that acceptance explicitly and never repeats the update or a restore write.
func ruleUpdateOutcome(ctx context.Context, id string, desired rules.Params, before rules.Rule, restore bool, read ruleReader, write ruleWriter) map[string]any {
	out := map[string]any{"source": "live", "accepted": true, "operation": "rules update", "id": id, "restore_requested": restore}
	observed, err := read(ctx, id)
	if err != nil {
		out["readback_status"], out["restore_status"], out["next_step"] = "unavailable", "not_attempted", "Inspect the rule and its enabled state without retrying the update."
		return out
	}
	out["observed_enabled"] = observed.Enabled
	out["enabled_after_update"] = observed.Enabled
	if !ruleMatchesUpdate(observed, id, desired, before) {
		out["readback_status"], out["restore_status"], out["next_step"] = "mismatch", "not_attempted", "Inspect the rule contents and enabled state without retrying the update."
		return out
	}
	out["readback_status"] = "verified"
	if !observed.Enabled {
		out["restore_status"], out["final_enabled"] = "not_needed", false
		return out
	}
	if !restore {
		out["restore_status"], out["final_enabled"], out["next_step"] = "not_requested", true, "The rule is enabled. Disable it explicitly if desired."
		return out
	}
	latest, err := read(ctx, id)
	if err != nil || latest.Version != observed.Version || !latest.Enabled || !ruleMatchesUpdate(latest, id, desired, before) {
		out["restore_status"], out["next_step"] = "not_attempted_state_changed", "Inspect the rule before taking another action; its state changed after readback."
		return out
	}
	_, err = write(ctx, "disable", rules.Params{ID: id, IdentityDisplay: latest.IdentityDisplay})
	if err != nil {
		out["restore_status"], out["next_step"] = "unconfirmed", "The disable outcome is uncertain. Inspect the rule without retrying."
		return out
	}
	restored, err := read(ctx, id)
	if err != nil || restored.Enabled || !ruleMatchesUpdate(restored, id, desired, before) {
		out["restore_status"], out["next_step"] = "unconfirmed", "The disable was accepted but its final state is unconfirmed; inspect without retrying."
		return out
	}
	out["restore_status"], out["observed_enabled"], out["final_enabled"] = "verified", false, false
	return out
}
