package webcapture_test

import (
	"encoding/json"
	people "github.com/mikefwille/outlook-web-cli/internal/outlook/contacts"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"strings"
	"testing"
)

func TestObservedPeopleSearchSuggestions(t *testing.T) {
	template := map[string]any{"Cvid": "fixture-correlation", "LogicalId": "fixture-logical", "Scenario": map[string]any{"Name": "owa.react.people"}, "EntityRequests": []any{map[string]any{"EntityType": "People", "Query": map[string]any{"QueryString": "fixture-name"}, "Size": 6, "Provenances": []any{"Mailbox", "Directory"}}}}
	d, r := captureFile(t, "POST", "/search/api/v1/suggestions?domain=People&n=1", "", template, false)
	_, op, err := people.PrepareSearch(d, r, accountSession(), "New query")
	if err != nil {
		t.Fatal(err)
	}
	if op.Class != transport.Read || strings.Contains(string(op.Body), "fixture-") || !strings.Contains(string(op.Body), "New query") {
		t.Fatal("search retained fixture input")
	}
	values, state := people.ParseSearch([]byte(`{"Groups":[{"Type":"People","Suggestions":[{"Id":"synthetic-id","DisplayName":"Example Name","EmailAddresses":["a@example.invalid"],"PeopleSubtype":"OrganizationUser","FeatureData":{"discard":99}}]}]}`))
	if state != transport.EnvelopeOK || len(values) != 1 || values[0].DisplayName != "Example Name" || values[0].PeopleSubtype != "OrganizationUser" {
		t.Fatal("suggestion response not parsed")
	}
	if encoded, _ := json.Marshal(values); strings.Contains(string(encoded), "FeatureData") {
		t.Fatal("irrelevant ranking payload retained")
	}
	for _, v := range []string{`{}`, `{"Groups":null}`, `{"Groups":[{}]}`, `{"Groups":[{"Suggestions":[{}]}]}`} {
		if _, state := people.ParseSearch([]byte(v)); state == transport.EnvelopeOK {
			t.Fatal("invalid suggestions accepted")
		}
	}
}
