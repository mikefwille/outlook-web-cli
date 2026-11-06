package cli

import (
	"bytes"
	"encoding/json"
	"strings"
	"testing"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/rules"
)

func TestRuleCreationRequiresExplicitEnabledAcknowledgement(t *testing.T) {
	args := []string{"--name", "Example", "--subject-contains", "Unique marker", "--category", "Blue"}
	cmd := newRuleAction(&rootFlags{dryRun: true}, "create")
	var out bytes.Buffer
	cmd.SetOut(&out)
	cmd.SetErr(&bytes.Buffer{})
	cmd.SetArgs(args)
	if err := cmd.Execute(); err != nil {
		t.Fatal(err)
	}
	var p map[string]any
	if json.Unmarshal(out.Bytes(), &p) != nil || p["enabled_on_creation"] != true || p["run_existing_mail"] != false {
		t.Fatal("preview hides enabled behavior")
	}
	cmd = newRuleAction(&rootFlags{}, "create")
	cmd.SetOut(&bytes.Buffer{})
	cmd.SetErr(&bytes.Buffer{})
	cmd.SetArgs(append(args, "--execute"))
	if cmd.Execute() == nil {
		t.Fatal("enabled creation lacked explicit acknowledgement")
	}
}

func TestRuleUpdatePreviewAcceptsObservedFields(t *testing.T) {
	for _, tc := range []struct {
		args []string
		want map[string]any
		ok   bool
	}{
		{[]string{"rule-id", "--name", "New name"}, map[string]any{"name": "New name"}, true},
		{[]string{"rule-id", "--subject-contains", "new marker"}, map[string]any{"subject_contains": "new marker"}, true},
		{[]string{"rule-id", "--category", "Green"}, map[string]any{"category": "Green"}, true},
		{[]string{"rule-id", "--subject-contains", "new marker", "--category", "Green"}, map[string]any{"subject_contains": "new marker", "category": "Green"}, true},
		{[]string{"rule-id"}, nil, false},
		{[]string{"rule-id", "--subject-contains", ""}, nil, false},
	} {
		cmd := newRuleAction(&rootFlags{dryRun: true}, "update")
		var out bytes.Buffer
		cmd.SetOut(&out)
		cmd.SetErr(&bytes.Buffer{})
		cmd.SetArgs(tc.args)
		err := cmd.Execute()
		if (err == nil) != tc.ok {
			t.Fatalf("unexpected preview result for %v: %v", tc.args, err)
		}
		if !tc.ok {
			continue
		}
		var result struct {
			Changes   map[string]any `json:"changes"`
			Preview   bool           `json:"preview"`
			MayEnable bool           `json:"may_enable_rule"`
		}
		if json.Unmarshal(out.Bytes(), &result) != nil || !result.Preview || !result.MayEnable || len(result.Changes) != len(tc.want) {
			t.Fatal("invalid changes preview")
		}
		for key, expected := range tc.want {
			if result.Changes[key] != expected {
				t.Fatal("preview omitted or changed requested field")
			}
		}
	}
}

func TestRuleFieldUpdateRequiresEnabledAcknowledgement(t *testing.T) {
	for _, field := range []string{"--subject-contains", "--category"} {
		cmd := newRuleAction(&rootFlags{}, "update")
		cmd.SetOut(&bytes.Buffer{})
		cmd.SetErr(&bytes.Buffer{})
		cmd.SetArgs([]string{"rule-id", field, "replacement", "--execute"})
		if err := cmd.Execute(); err == nil || !strings.Contains(err.Error(), "--allow-enabled") {
			t.Fatal("field update did not require enabled-state acknowledgement")
		}
	}
}

func TestRuleUpdateComposesOmittedFieldsFromFreshRead(t *testing.T) {
	current := rules.Rule{Name: "Original name", Subject: []string{"original marker"}, Category: []string{"Blue"}}
	for _, tc := range []struct {
		input                               rules.Params
		name, subject, category             bool
		wantName, wantSubject, wantCategory string
	}{
		{rules.Params{Name: "Renamed"}, true, false, false, "Renamed", "original marker", "Blue"},
		{rules.Params{Subject: "new marker"}, false, true, false, "Original name", "new marker", "Blue"},
		{rules.Params{Category: "Green"}, false, false, true, "Original name", "original marker", "Green"},
		{rules.Params{Name: "Renamed", Subject: "new marker", Category: "Green"}, true, true, true, "Renamed", "new marker", "Green"},
	} {
		got := composeRuleUpdate(current, tc.input, tc.name, tc.subject, tc.category)
		if got.Name != tc.wantName || got.Subject != tc.wantSubject || got.Category != tc.wantCategory {
			t.Fatal("partial update did not preserve unspecified fields")
		}
	}
}

func TestCreatedRuleDoesNotExposeProvisionalIdentityAsUsableID(t *testing.T) {
	body := []byte(`{"InboxRule":{"__type":"InboxRule:#Exchange","Identity":{"DisplayName":"transient identity","RawIdentity":"transient identity"},"Name":"Example","Enabled":true,"SubjectContainsWords":["marker"],"ApplyCategory":["Blue"]}}`)
	result, err := ruleCreatedResult(body)
	if err != nil {
		t.Fatal(err)
	}
	encoded, err := json.Marshal(result)
	if err != nil {
		t.Fatal(err)
	}
	var output struct {
		IdentityStatus string         `json:"identity_status"`
		Rule           map[string]any `json:"rule"`
	}
	if json.Unmarshal(encoded, &output) != nil {
		t.Fatal("invalid created rule output")
	}
	if output.IdentityStatus != "list_read_required" {
		t.Fatal("create output did not require a fresh list identity")
	}
	if _, ok := output.Rule["id"]; ok {
		t.Fatal("create-only identity was exposed as a stable rule ID")
	}
	if bytes.Contains(encoded, []byte("transient identity")) {
		t.Fatal("create-only identity leaked through the result")
	}
}
