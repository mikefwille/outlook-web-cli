package cli

import (
	"bytes"
	"encoding/json"
	"strings"
	"testing"
)

func TestRuleOrderPreservesEachEnabledState(t *testing.T) {
	raw := []json.RawMessage{
		json.RawMessage(`{"__type":"InboxRule:#Exchange","Identity":{"DisplayName":"first","RawIdentity":"id-a"},"Name":"A","Enabled":true,"Priority":1,"Unknown":"preserve"}`),
		json.RawMessage(`{"__type":"InboxRule:#Exchange","Identity":{"DisplayName":"second","RawIdentity":"id-b"},"Name":"B","Enabled":false,"Priority":2,"Unknown":"preserve"}`),
	}
	current, version, err := parseRuleOrder(raw)
	if err != nil || len(version) != 64 {
		t.Fatal("valid rule list rejected")
	}
	desired, err := requestedRuleOrder(current, []string{"id-b", "id-a"})
	if err != nil || desired[0].ID != "id-b" || desired[0].Enabled || desired[1].ID != "id-a" || !desired[1].Enabled {
		t.Fatal("reorder did not preserve per-rule states")
	}
	for _, bad := range [][]string{{"id-a", "id-a"}, {"id-a", "other"}, {"id-a", "id-b"}, {"id-a"}} {
		if _, err = requestedRuleOrder(current, bad); err == nil {
			t.Fatal("invalid or unchanged order accepted")
		}
	}
	post := []json.RawMessage{
		json.RawMessage(`{"__type":"InboxRule:#Exchange","Identity":{"DisplayName":"second","RawIdentity":"id-b"},"Name":"B","Enabled":false,"Priority":1,"Unknown":"preserve"}`),
		json.RawMessage(`{"__type":"InboxRule:#Exchange","Identity":{"DisplayName":"first","RawIdentity":"id-a"},"Name":"A","Enabled":true,"Priority":2,"Unknown":"preserve"}`),
	}
	after, _, err := parseRuleOrder(post)
	if err != nil || !orderMatches(current, after, desired) {
		t.Fatal("valid reorder readback rejected")
	}
	post[0] = json.RawMessage(`{"__type":"InboxRule:#Exchange","Identity":{"DisplayName":"second","RawIdentity":"id-b"},"Name":"B","Enabled":false,"Priority":1,"Unknown":"changed"}`)
	after, _, err = parseRuleOrder(post)
	if err != nil || orderMatches(current, after, desired) {
		t.Fatal("changed unknown rule field was accepted")
	}
	raw[1] = json.RawMessage(`{"__type":"InboxRule:#Exchange","Identity":{"DisplayName":"second","RawIdentity":"id-b"},"Name":"B","Enabled":false,"Priority":1}`)
	if _, _, err = parseRuleOrder(raw); err == nil {
		t.Fatal("inconsistent priority ordering accepted")
	}
}

func TestRuleReorderPreviewDisclosesStateReapplication(t *testing.T) {
	cmd := newRuleReorder(&rootFlags{dryRun: true})
	var out bytes.Buffer
	cmd.SetOut(&out)
	cmd.SetErr(&bytes.Buffer{})
	cmd.SetArgs([]string{"--id", "id-b", "--id", "id-a"})
	if err := cmd.Execute(); err != nil {
		t.Fatal(err)
	}
	var value map[string]any
	if json.Unmarshal(out.Bytes(), &value) != nil || value["preview"] != true || value["reapplies_enabled_states"] != true || value["run_existing_mail"] != false {
		t.Fatal("unsafe reorder preview")
	}
}

func TestRuleReorderRequiresExplicitStateReapplication(t *testing.T) {
	cmd := newRuleReorder(&rootFlags{})
	cmd.SetOut(&bytes.Buffer{})
	cmd.SetErr(&bytes.Buffer{})
	cmd.SetArgs([]string{"--id", "id-a", "--id", "id-b", "--version", strings.Repeat("a", 64), "--execute"})
	if err := cmd.Execute(); err == nil || !strings.Contains(err.Error(), "--allow-state-reapply") {
		t.Fatal("reorder did not require state-reapplication acknowledgement")
	}
}
