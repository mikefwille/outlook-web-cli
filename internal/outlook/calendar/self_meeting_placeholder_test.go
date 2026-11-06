package calendar

import (
	"encoding/base64"
	"encoding/json"
	"net/http"
	"testing"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func TestSelfMeetingPlaceholderRequiresRuntimeAndRebinds(t *testing.T) {
	claims, _ := json.Marshal(map[string]any{"upn": "self@example.invalid", "oid": "self-object", "name": "Self"})
	session := transport.Session{Headers: http.Header{"Authorization": []string{"Bearer e30." + base64.RawURLEncoding.EncodeToString(claims) + ".signature"}}}
	item := func(email, object, name string) map[string]any {
		return map[string]any{"RequiredAttendees": []any{map[string]any{"IsCoauthor": false, "Mailbox": map[string]any{"EmailAddress": email, "AADObjectId": object, "Name": name, "MailboxType": "InternalMailbox", "RoutingType": "SMTP"}}}, "OptionalAttendees": []any{}, "Resources": []any{}}
	}
	placeholder := func() map[string]any { return item("template@example.invalid", "template-object", "template self") }
	if bindSelfMeetingAttendee(placeholder(), session, false) == nil {
		t.Fatal("placeholder accepted from an ordinary capture")
	}
	bound := placeholder()
	if err := bindSelfMeetingAttendee(bound, session, true); err != nil {
		t.Fatalf("runtime placeholder rejected: %v", err)
	}
	mailbox := obj(obj(bound["RequiredAttendees"].([]any)[0])["Mailbox"])
	if mailbox["EmailAddress"] != "self@example.invalid" || mailbox["AADObjectId"] != "self-object" || mailbox["Name"] != "Self" {
		t.Fatal("runtime placeholder was not rebound to the current account")
	}
	if bindSelfMeetingAttendee(item("other@example.invalid", "template-object", "template self"), session, true) == nil {
		t.Fatal("foreign attendee accepted as a runtime placeholder")
	}
	extra := placeholder()
	extra["OptionalAttendees"] = []any{map[string]any{"Mailbox": "other"}}
	if bindSelfMeetingAttendee(extra, session, true) == nil {
		t.Fatal("additional attendee accepted")
	}
}
