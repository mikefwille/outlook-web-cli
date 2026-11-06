package webcapture_test

import (
	"encoding/json"
	"net/url"
	"strings"
	"testing"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/contactlists"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func contactListQuery(kind string) url.Values {
	q := url.Values{"n": {"1"}, "cri": {"synthetic-correlation"}, "cv": {"synthetic-cv"}}
	switch kind {
	case "list":
		q.Set("$top", "500")
		q.Set("$select", contactlists.Select)
		q.Set("$orderBy", "DisplayName Asc")
	case "create", "update":
		q.Set("$select", contactlists.Select)
	case "add-member", "remove-member":
		q.Set("$select", contactlists.MembersSelect)
		q.Set("$expand", "members")
	}
	return q
}

func TestContactListObservedOperationsReplaceFixtureData(t *testing.T) {
	member := map[string]any{"DisplayName": "Fixture Person", "Key": "fixture@example.invalid", "RoutingType": "SMTP", "Type": "OneOff"}
	for _, tc := range []struct {
		action, method, path string
		body                 any
		class                transport.Class
	}{
		{"list", "GET", contactlists.Path, nil, transport.Read},
		{"create", "POST", contactlists.Path, map[string]any{"DisplayName": "fixture list", "Members": []any{member}}, transport.Write},
		{"update", "PATCH", contactlists.Path + "/fixture-id", map[string]any{"DisplayName": "fixture list", "Notes": "fixture notes"}, transport.Write},
		{"add-member", "POST", contactlists.Path + "/fixture-id/AddMembers", map[string]any{"Members": []any{member}}, transport.Write},
		{"remove-member", "POST", contactlists.Path + "/fixture-id/DeleteMembers", map[string]any{"Members": []any{map[string]any{"Key": "fixture@example.invalid", "RoutingType": "SMTP", "Type": "OneOff"}}}, transport.Write},
		{"delete", "DELETE", contactlists.Path + "/fixture-id", nil, transport.Write},
	} {
		t.Run(tc.action, func(t *testing.T) {
			d, r := captureFile(t, tc.method, tc.path+"?"+contactListQuery(tc.action).Encode(), "", tc.body, false)
			p := contactlists.Params{ID: "new-id", Name: "New list", Notes: "new notes", Email: "new@example.invalid", MemberName: "New member"}
			if tc.action == "list" || tc.action == "create" {
				p.ID = ""
			}
			s, op, err := contactlists.Prepare(tc.action, d, r, accountSession(), p)
			if err != nil {
				t.Fatal(err)
			}
			if op.Method != tc.method || op.Class != tc.class || strings.Contains(op.Path, "fixture-id") || strings.Contains(string(op.Body), "fixture") || op.Query.Get("cri") == "synthetic-correlation" || s.Headers.Get("Cookie") != "fresh=synthetic" {
				t.Fatal("fixture data or captured identity survived operation preparation")
			}
			foreign := accountSession()
			foreign.AccountKey = "foreign"
			if _, _, err = contactlists.Prepare(tc.action, d, r, foreign, p); err == nil {
				t.Fatal("foreign account accepted")
			}
			if tc.action != "list" && tc.action != "create" && !strings.Contains(op.Path, "new-id") {
				t.Fatal("requested ID missing")
			}
		})
	}
}

func TestContactListRejectsUnsafeShapesAndChecksMembership(t *testing.T) {
	member := map[string]any{"DisplayName": "Fixture", "Key": "fixture@example.invalid", "RoutingType": "SMTP", "Type": "OneOff"}
	q := contactListQuery("create")
	q.Set("unexpected", "value")
	d, r := captureFile(t, "POST", contactlists.Path+"?"+q.Encode(), "", map[string]any{"DisplayName": "Fixture list", "Members": []any{member}}, false)
	p := contactlists.Params{Name: "Other list", Email: "one@example.invalid", MemberName: "One"}
	if _, _, err := contactlists.Prepare("create", d, r, accountSession(), p); err == nil {
		t.Fatal("unexpected query accepted")
	}
	q = contactListQuery("create")
	member["extra"] = "discard me"
	d, r = captureFile(t, "POST", contactlists.Path+"?"+q.Encode(), "", map[string]any{"DisplayName": "Fixture list", "Members": []any{member}}, false)
	if _, _, err := contactlists.Prepare("create", d, r, accountSession(), p); err == nil {
		t.Fatal("unexpected fixture field accepted")
	}
	item, state := contactlists.ParseItem([]byte(`{"Id":"id-1","@odata.etag":"v1","DisplayName":"A","Notes":"keep","Members":[{"Key":"one@example.invalid","Type":"OneOff","RoutingType":"SMTP"}]}`))
	if state != transport.EnvelopeOK || contactlists.CheckMember(item, "ONE@example.invalid", false) != nil || contactlists.CheckMember(item, "one@example.invalid", true) == nil {
		t.Fatal("membership matching failed")
	}
	item.Members = append(item.Members, json.RawMessage(`{"Type":"OneOff"}`))
	if contactlists.CheckMember(item, "two@example.invalid", true) == nil {
		t.Fatal("incomplete membership accepted")
	}
}

func TestContactListPageCompletenessAndMalformedResponses(t *testing.T) {
	for _, tc := range []struct {
		body         string
		ok, complete bool
	}{
		{`{"value":[]}`, true, true},
		{`{"value":[],"@odata.nextLink":"opaque"}`, true, false},
		{`{"value":null}`, false, false},
		{`{"value":[{"Id":"id"}]}`, false, false},
		{`{"error":{"message":"failure"}}`, false, false},
	} {
		page, state := contactlists.ParsePage([]byte(tc.body))
		if (state == transport.EnvelopeOK) != tc.ok || tc.ok && page.Complete != tc.complete {
			t.Fatal("invalid completeness result")
		}
	}
}
