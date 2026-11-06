package webcapture_test

import (
	"encoding/json"
	"testing"

	people "github.com/mikefwille/outlook-web-cli/internal/outlook/contacts"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func TestContactObservedInsertAndClear(t *testing.T) {
	template := map[string]any{"names": []any{map[string]any{"id": "fixture-name", "action": "PATCH", "body": map[string]any{"first": "fixture", "source": map[string]any{"type": "UserProvided"}}}}, "emails": []any{}, "phones": []any{}, "positions": []any{}, "notes": []any{}, "addresses": []any{}, "anniversaries": []any{}, "relationships": []any{}, "tags": []any{}, "webAccounts": []any{}, "websites": []any{}}
	d, r := captureFile(t, "POST", "/PeopleGraphVx/v1.0/contacts('fixture_contact')/UpdateCollections?n=1", "", template, false)
	session := accountSession()
	base := `{"id":"contact","changeKey":"v1","names":[{"id":"name-id","first":"Ada","source":{"type":"UserProvided"}}],"personSources":[{"type":"UserProvided","isEditable":true,"isDeletable":true}]}`
	empty, state := people.ParseContact([]byte(base), false)
	if state != transport.EnvelopeOK {
		t.Fatal("contact rejected")
	}
	params := people.Params{ID: "contact", Current: &empty, Email: strptr("a@example.invalid"), Phone: strptr("555"), Company: strptr("Company"), Notes: strptr("Note")}
	_, op, err := people.Prepare("update", d, r, session, params)
	if err != nil {
		t.Fatal(err)
	}
	root := unwrapBody(t, op)
	for _, k := range []string{"emails", "phones", "positions", "notes"} {
		v := root[k].([]any)
		if len(v) != 1 {
			t.Fatalf("missing %s", k)
		}
		entry := v[0].(map[string]any)
		if entry["action"] != "POST" || entry["id"] != nil || entry["body"] == nil {
			t.Fatalf("wrong %s insert shape", k)
		}
	}
	if body := root["positions"].([]any)[0].(map[string]any)["body"].(map[string]any); body["isCurrent"] != false || body["detail"].(map[string]any)["company"].(map[string]any)["displayName"] != "Company" {
		t.Fatal("wrong company insert")
	}
	full := `{"id":"contact","changeKey":"v2","names":[{"id":"name-id","first":"Ada","source":{"type":"UserProvided"}}],"emails":[{"id":"email-id","address":"a@example.invalid","type":"Work","source":{"type":"UserProvided"}}],"phones":[{"id":"phone-id","number":"555","type":"Business","source":{"type":"UserProvided"}}],"positions":[{"id":"position-id","detail":{"jobTitle":"","company":{"department":"","displayName":"Company","officeLocation":"","pronunciation":"","address":{"type":"Unknown"}}},"isCurrent":false,"categories":[],"manager":null,"colleagues":[],"source":{"type":"UserProvided"}}],"notes":[{"id":"note-id","detail":{"contentType":"text","content":"Note"},"source":{"type":"UserProvided"}}],"personSources":[{"type":"UserProvided","isEditable":true,"isDeletable":true}]}`
	contact, state := people.ParseContact([]byte(full), false)
	if state != transport.EnvelopeOK {
		t.Fatal("full contact rejected")
	}
	params = people.Params{ID: "contact", Current: &contact, ClearEmail: true, ClearPhone: true, ClearCompany: true, ClearNotes: true}
	_, op, err = people.Prepare("update", d, r, session, params)
	if err != nil {
		t.Fatal(err)
	}
	root = unwrapBody(t, op)
	for k, id := range map[string]string{"emails": "email-id", "phones": "phone-id", "positions": "position-id", "notes": "note-id"} {
		entry := root[k].([]any)[0].(map[string]any)
		if entry["action"] != "DELETE" || entry["id"] != id || entry["body"] != nil {
			t.Fatalf("wrong %s clear shape", k)
		}
	}
	if _, _, err = people.Prepare("update", d, r, session, people.Params{ID: "contact", Current: &empty, ClearEmail: true}); err == nil {
		t.Fatal("clear of absent email accepted")
	}
	if _, _, err = people.Prepare("update", d, r, session, people.Params{ID: "contact", Current: &contact, Email: strptr("new@example.invalid"), ClearEmail: true}); err == nil {
		t.Fatal("simultaneous set and clear accepted")
	}
	var raw map[string]any
	if json.Unmarshal([]byte(full), &raw) != nil {
		t.Fatal("invalid fixture")
	}
	raw["positions"].([]any)[0].(map[string]any)["detail"].(map[string]any)["jobTitle"] = "Engineer"
	b, _ := json.Marshal(raw)
	protected, _ := people.ParseContact(b, false)
	if _, _, err = people.Prepare("update", d, r, session, people.Params{ID: "contact", Current: &protected, ClearCompany: true}); err == nil {
		t.Fatal("clear company would erase job title")
	}
}
