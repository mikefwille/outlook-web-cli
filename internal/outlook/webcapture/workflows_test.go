package webcapture_test

import (
	"context"
	"crypto/sha256"
	"encoding/base64"
	"encoding/hex"
	"encoding/json"
	"io"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	cal "github.com/mikefwille/outlook-web-cli/internal/outlook/calendar"
	people "github.com/mikefwille/outlook-web-cli/internal/outlook/contacts"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func shaAccount(tid, oid string) string {
	h := sha256.Sum256([]byte(tid + "\x00" + oid))
	return hex.EncodeToString(h[:])
}
func strptr(s string) *string { return &s }
func fakeToken(aud string) string {
	p, _ := json.Marshal(map[string]any{"tid": "tenant-test", "oid": "object-test", "aud": aud, "exp": time.Now().Add(time.Hour).Unix()})
	return "Bearer e30." + base64.RawURLEncoding.EncodeToString(p) + ".signature"
}
func accountSession() transport.Session {
	// The account partition key is normally produced by the browser-session adapter.
	// The same deterministic hash is computed here from entirely synthetic claims.
	h := shaAccount("tenant-test", "object-test")
	return transport.Session{AccountKey: h, Origin: "https://outlook.cloud.microsoft", ExpiresAt: time.Now().Add(time.Hour), Headers: http.Header{"Authorization": []string{fakeToken("aud-test")}, "Cookie": []string{"fresh=synthetic"}, "X-Anchormailbox": []string{"synthetic-mailbox"}}}
}
func captureFile(t *testing.T, method, path, action string, body any, inHeader bool) (string, string) {
	t.Helper()
	dir := t.TempDir()
	detail := filepath.Join(dir, "detail.txt")
	req := filepath.Join(dir, "request.json")
	b, _ := json.Marshal(body)
	status := "200"
	if method == "DELETE" {
		status = "204"
	}
	headers := "- :method:" + method + "\n- authorization:" + fakeToken("aud-test") + "\n- cookie:old=synthetic\n- x-anchormailbox:synthetic-mailbox\n"
	if action != "" {
		headers += "- action:" + action + "\n"
	}
	if inHeader {
		headers += "- x-owa-urlpostdata:" + url.QueryEscape(string(b)) + "\n"
	}
	txt := "## Request https://outlook.cloud.microsoft" + path + "\nStatus: " + status + "\n### Request Headers\n" + headers + "### Response Headers\n"
	if err := os.WriteFile(detail, []byte(txt), 0600); err != nil {
		t.Fatal(err)
	}
	if body == nil || inHeader {
		return detail, ""
	}
	if err := os.WriteFile(req, b, 0600); err != nil {
		t.Fatal(err)
	}
	return detail, req
}
func baseCalendar(typ string, body map[string]any) map[string]any {
	return map[string]any{"__type": typ, "Header": map[string]any{"TimeZoneContext": map[string]any{"TimeZoneDefinition": map[string]any{"Id": "Mountain Standard Time"}}}, "Body": body}
}
func captureCalendar(t *testing.T, action string, body any, header bool) (string, string) {
	return captureFile(t, "POST", "/owa/service.svc?action="+action+"&app=Calendar&n=1", action, body, header)
}
func unwrapBody(t *testing.T, op transport.Operation) map[string]any {
	t.Helper()
	b := op.Body
	if v := op.Headers.Get("X-Owa-Urlpostdata"); v != "" {
		s, e := url.QueryUnescape(v)
		if e != nil {
			t.Fatal(e)
		}
		b = []byte(s)
	}
	var d map[string]any
	if json.Unmarshal(b, &d) != nil {
		t.Fatal("invalid operation JSON")
	}
	return d
}
func TestCalendarRequestsReplaceFixtureIdentifiers(t *testing.T) {
	get := baseCalendar("GetCalendarEventJsonRequest:#Exchange", map[string]any{"EventIds": []any{map[string]any{"Id": "fixture-event"}}})
	d, r := captureCalendar(t, "GetCalendarEvent", get, false)
	s, op, err := cal.Prepare("GetCalendarEvent", d, r, accountSession(), cal.Params{ID: "requested-event"})
	if err != nil {
		t.Fatal(err)
	}
	if op.Class != transport.Read || strings.Contains(string(op.Body), "fixture-event") || !strings.Contains(string(op.Body), "requested-event") || s.Headers.Get("Cookie") != "fresh=synthetic" {
		t.Fatal("get not rebound or fixture ID retained")
	}
	create := baseCalendar("CreateItemJsonRequest:#Exchange", map[string]any{"SendMeetingInvitations": "SendToNone", "SavedItemFolderId": map[string]any{"BaseFolderId": map[string]any{"Id": "fixture-folder"}}, "Items": []any{map[string]any{"__type": "CalendarItem:#Exchange", "IsAllDayEvent": false, "ItemId": map[string]any{"Id": "fixture-event"}, "id": "fixture-event", "ClientSeriesId": "fixture-event", "ParentFolderId": map[string]any{"Id": "fixture-folder"}, "Resources": []any{}, "Attachments": []any{}, "Body": map[string]any{"BodyType": "HTML", "Value": ""}, "Locations": []any{}, "ExtendedProperty": []any{}}}})
	d, r = captureCalendar(t, "CreateCalendarEvent", create, false)
	_, op, err = cal.Prepare("CreateCalendarEvent", d, r, accountSession(), cal.Params{CalendarID: "requested-folder", Subject: "Synthetic appointment", Start: "2026-11-09T09:00:00-07:00", End: "2026-11-09T10:00:00-07:00"})
	if err != nil {
		t.Fatal(err)
	}
	if op.Class != transport.Write || strings.Contains(string(op.Body), "fixture-") || !strings.Contains(string(op.Body), "requested-folder") {
		t.Fatal("create retained fixture data")
	}
	bodyText, firstLocation := "Hello <reader>\nsecond", "Synthetic room"
	_, withFields, err := cal.Prepare("CreateCalendarEvent", d, r, accountSession(), cal.Params{CalendarID: "requested-folder", Subject: "Synthetic appointment", Start: "2026-11-09T09:00:00-07:00", End: "2026-11-09T10:00:00-07:00", BodyText: &bodyText, AddLocation: &firstLocation})
	if err != nil {
		t.Fatal(err)
	}
	created := unwrapBody(t, withFields)["Body"].(map[string]any)["Items"].([]any)[0].(map[string]any)
	if created["Body"].(map[string]any)["Value"] != "<div>Hello &lt;reader&gt;<br>second</div>" {
		t.Fatal("body text not encoded as HTML")
	}
	if len(created["Locations"].([]any)) != 1 || len(created["ExtendedProperty"].([]any)) != 1 {
		t.Fatal("missing initial location or source map")
	}
	body := create["Body"].(map[string]any)
	body["SendMeetingInvitations"] = "SendToAllAndSaveCopy"
	d, r = captureCalendar(t, "CreateCalendarEvent", create, false)
	if _, _, err = cal.Prepare("CreateCalendarEvent", d, r, accountSession(), cal.Params{CalendarID: "id", Subject: "x", Start: "2026-11-09T09:00:00-07:00", End: "2026-11-09T10:00:00-07:00"}); err == nil {
		t.Fatal("meeting invitation accepted")
	}
	update := baseCalendar("UpdateCalendarEventJsonRequest:#Exchange", map[string]any{"EventId": map[string]any{"Id": "fixture-event"}, "ItemChange": map[string]any{"ItemId": map[string]any{"Id": "fixture-event"}, "Updates": []any{map[string]any{"__type": "SetItemField:#Exchange", "Path": map[string]any{"PropertyName": "IsDefaultPreferences"}}, map[string]any{"__type": "SetItemField:#Exchange", "Path": map[string]any{"FieldURI": "Subject"}, "Item": map[string]any{"__type": "CalendarItem:#Exchange", "Subject": "fixture-subject"}}}}, "EventScope": 0, "ShouldSendUpdateToAttendees": false})
	d, r = captureCalendar(t, "UpdateCalendarEvent", update, true)
	_, op, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), cal.Params{ID: "new-event", Subject: "new subject"})
	if err != nil {
		t.Fatal(err)
	}
	parsed := unwrapBody(t, op)
	updates := parsed["Body"].(map[string]any)["ItemChange"].(map[string]any)["Updates"].([]any)
	if len(updates) != 1 {
		t.Fatal("unrelated calendar property change retained")
	}
	data, _ := json.Marshal(parsed)
	if strings.Contains(string(data), "fixture-") || !strings.Contains(string(data), "new-event") {
		t.Fatal("update retained fixture data")
	}
	del := baseCalendar("DeleteCalendarEventJsonRequest:#Exchange", map[string]any{"EventId": map[string]any{"Id": "fixture-event"}, "EventScope": 0})
	d, r = captureCalendar(t, "DeleteCalendarEvent", del, true)
	_, op, err = cal.Prepare("DeleteCalendarEvent", d, r, accountSession(), cal.Params{ID: "new-event"})
	if err != nil {
		t.Fatal(err)
	}
	data, _ = json.Marshal(unwrapBody(t, op))
	if strings.Contains(string(data), "fixture-event") {
		t.Fatal("delete retained fixture ID")
	}
}
func TestPersonalEventGuardAndResponseErrors(t *testing.T) {
	event := `{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Success","ResponseCode":"NoError","Items":[{"__type":"CalendarItem:#Exchange","ItemId":{"Id":"event","ChangeKey":"v1"},"IsMeeting":false,"IsRecurring":false,"IsOnlineMeeting":false,"IsOrganizer":true,"MeetingRequestWasSent":false,"CalendarItemType":"Single","RequiredAttendees":[],"OptionalAttendees":[],"Resources":[]}] }]}}}`
	e, s := cal.ParseEvent([]byte(event))
	if s != transport.EnvelopeOK || !e.SafePersonal() {
		t.Fatal("personal appointment rejected")
	}
	for _, old := range []string{`"IsMeeting":false`, `"IsRecurring":false`, `"IsOnlineMeeting":false`, `"IsOrganizer":true`, `"MeetingRequestWasSent":false`, `"CalendarItemType":"Single"`} {
		changed := strings.Replace(event, old, `"other":true`, 1)
		e, s = cal.ParseEvent([]byte(changed))
		if s != transport.EnvelopeOK || e.SafePersonal() {
			t.Fatal("incomplete event accepted for mutation")
		}
	}
	_, s = cal.ParseEvent([]byte(`{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Error","ResponseCode":"ErrorIrresolvableConflict"}]}}}`))
	if s != transport.EnvelopeConflict {
		t.Fatal("conflict not surfaced")
	}
}
func TestContactsRewriteAndValidate(t *testing.T) {
	session := accountSession()
	get := map[string]any{"ContactIds": []any{"fixture_contact"}}
	d, r := captureFile(t, "POST", "/PeopleGraphVx/v1.0/peopleLookup?n=1", "", get, false)
	_, op, err := people.Prepare("get", d, r, session, people.Params{ID: "new_contact"})
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(string(op.Body), "fixture_contact") || !strings.Contains(string(op.Body), "new_contact") {
		t.Fatal("lookup ID not rewritten")
	}
	contactJSON := `{"id":"new_contact","changeKey":"v1","names":[{"id":"name-7","first":"Ada","last":"Lovelace","middle":"M","source":{"type":"UserProvided"}}],"emails":[{"id":"email-1","address":"before@example.invalid","type":"Work","source":{"type":"UserProvided"}}],"phones":[{"id":"phone-1","number":"111","type":"Business","source":{"type":"UserProvided"}}],"positions":[{"id":"position-1","detail":{"jobTitle":"Engineer","company":{"department":"Research","displayName":"Old","officeLocation":"Remote","pronunciation":"","address":{"type":"Unknown"}}},"isCurrent":true,"categories":[],"source":{"type":"UserProvided"}}],"notes":[{"id":"note-1","detail":{"contentType":"text","content":"old notes"},"source":{"type":"UserProvided"}}],"personSources":[{"type":"UserProvided","isEditable":true,"isDeletable":true}]}`
	c, state := people.ParseContact([]byte(contactJSON), false)
	if state != transport.EnvelopeOK || !c.SafeEditable() || !c.SafeDeletable() {
		t.Fatal("personal contact rejected")
	}
	update := map[string]any{"names": []any{map[string]any{"id": "fixture-name", "action": "PATCH", "body": map[string]any{"first": "old", "last": "old", "source": map[string]any{"type": "UserProvided"}}}}, "emails": []any{}, "phones": []any{}}
	d, r = captureFile(t, "POST", "/PeopleGraphVx/v1.0/contacts('fixture_contact')/UpdateCollections?n=1", "", update, false)
	_, op, err = people.Prepare("update", d, r, session, people.Params{ID: "new_contact", First: "Ada", Last: "Byron", Current: &c, NameChange: true, Email: strptr("after@example.invalid"), Phone: strptr("222"), Company: strptr("New company"), Notes: strptr("new notes")})
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(op.Path, "fixture") || strings.Contains(string(op.Body), "fixture") || !strings.Contains(string(op.Body), "name-7") || !strings.Contains(string(op.Body), "Byron") {
		t.Fatal("contact update did not rebind values")
	}
	for _, v := range []string{"email-1", "after@example.invalid", "phone-1", "222", "position-1", "New company", "Research", "note-1", "new notes"} {
		if !strings.Contains(string(op.Body), v) {
			t.Fatalf("missing updated or preserved value %s", v)
		}
	}
	if strings.Contains(string(op.Body), "before@example.invalid") || strings.Contains(string(op.Body), "old notes") {
		t.Fatal("stale updated field retained")
	}
	create := map[string]any{"names": []any{map[string]any{"first": "fixture-name", "last": "", "source": map[string]any{"type": "UserProvided"}}}, "emails": []any{}, "phones": []any{}, "positions": []any{}, "notes": []any{}}
	cd, cr := captureFile(t, "POST", "/PeopleGraphVx/v1.0/contacts?n=1", "", create, false)
	_, co, err := people.Prepare("create", cd, cr, session, people.Params{First: "Ada", Email: strptr("new@example.invalid"), Phone: strptr("333"), Company: strptr("Company"), Notes: strptr("Notes")})
	if err != nil {
		t.Fatal(err)
	}
	for _, v := range []string{"new@example.invalid", "333", "Company", "Notes", "UserProvided"} {
		if !strings.Contains(string(co.Body), v) {
			t.Fatalf("missing create field %s", v)
		}
	}
	if strings.Contains(string(co.Body), "fixture-name") {
		t.Fatal("create retained fixture name")
	}
	created := unwrapBody(t, co)
	if e := created["emails"].([]any)[0].(map[string]any); e["type"] != "Work" || e["address"] != "new@example.invalid" {
		t.Fatal("wrong work email shape")
	}
	if p := created["phones"].([]any)[0].(map[string]any); p["type"] != "Business" || p["number"] != "333" {
		t.Fatal("wrong business phone shape")
	}
	if n := created["notes"].([]any)[0].(map[string]any)["detail"].(map[string]any); n["contentType"] != "text" || n["content"] != "Notes" {
		t.Fatal("wrong note shape")
	}
	if company := created["positions"].([]any)[0].(map[string]any)["detail"].(map[string]any)["company"].(map[string]any); company["displayName"] != "Company" {
		t.Fatal("wrong company shape")
	}

	// Updating just the email must not overwrite the name or other collections.
	_, emailOnly, err := people.Prepare("update", d, r, session, people.Params{ID: "new_contact", Current: &c, Email: strptr("only@example.invalid")})
	if err != nil {
		t.Fatal(err)
	}
	em := unwrapBody(t, emailOnly)
	for _, k := range []string{"names", "phones", "positions", "notes"} {
		if em[k] == nil {
			continue
		}
		if a, ok := em[k].([]any); !ok || len(a) != 0 {
			t.Fatalf("unexpected %s mutation", k)
		}
	}
	noFields, _ := people.ParseContact([]byte(`{"id":"new_contact","changeKey":"v1","names":[{"id":"name-1","first":"Ada","source":{"type":"UserProvided"}}],"personSources":[{"type":"UserProvided","isEditable":true}]}`), false)
	_, added, err := people.Prepare("update", d, r, session, people.Params{ID: "new_contact", Current: &noFields, Email: strptr("new@example.invalid")})
	if err != nil {
		t.Fatal(err)
	}
	entry := unwrapBody(t, added)["emails"].([]any)[0].(map[string]any)
	if entry["action"] != "POST" || entry["id"] != nil {
		t.Fatal("missing field did not use observed POST")
	}

	if _, _, err = people.Prepare("update", d, r, session, people.Params{ID: "new_contact')/anything", First: "X", Current: &c}); err == nil {
		t.Fatal("unsafe contact ID accepted")
	}
	d, r = captureFile(t, "DELETE", "/PeopleGraphVx/v1.0/contacts('fixture_contact')?n=1", "", nil, false)
	_, op, err = people.Prepare("delete", d, r, session, people.Params{ID: "new_contact"})
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(op.Path, "fixture") || op.Method != "DELETE" || op.Validate(nil) != transport.EnvelopeOK {
		t.Fatal("delete malformed")
	}
}

type roundTripFunc func(*http.Request) (*http.Response, error)

func (f roundTripFunc) RoundTrip(r *http.Request) (*http.Response, error) { return f(r) }
func TestCapturedOperationUsesFreshSessionAndPropagatesErrors(t *testing.T) {
	d, r := captureFile(t, "POST", "/PeopleGraphVx/v1.0/peopleLookup?n=1", "", map[string]any{"ContactIds": []any{"fixture"}}, false)
	s, op, err := people.Prepare("get", d, r, accountSession(), people.Params{ID: "synthetic"})
	if err != nil {
		t.Fatal(err)
	}
	client, err := transport.New(s.AccountKey, s, roundTripFunc(func(req *http.Request) (*http.Response, error) {
		if req.Header.Get("Cookie") != "fresh=synthetic" {
			t.Fatal("stale cookie used")
		}
		return &http.Response{StatusCode: 401, Header: http.Header{}, Body: io.NopCloser(strings.NewReader(""))}, nil
	}))
	if err != nil {
		t.Fatal(err)
	}
	_, err = client.Do(context.Background(), s.AccountKey, op)
	if e, ok := err.(*transport.Error); !ok || e.Kind != transport.KindAuth {
		t.Fatalf("expected auth error, got %v", err)
	}
	bad := accountSession()
	bad.Headers.Set("Authorization", fakeToken("different-audience"))
	if _, _, err = people.Prepare("get", d, r, bad, people.Params{ID: "synthetic"}); err == nil {
		t.Fatal("different audience accepted")
	}
}

func TestCalendarBodyAndAddedLocation(t *testing.T) {
	getResponse := `{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Success","ResponseCode":"NoError","Items":[{"__type":"CalendarItem:#Exchange","ItemId":{"Id":"event","ChangeKey":"v1"},"IsMeeting":false,"IsRecurring":false,"IsOnlineMeeting":false,"IsOrganizer":true,"MeetingRequestWasSent":false,"CalendarItemType":"Single","RequiredAttendees":[],"OptionalAttendees":[],"Resources":[],"Locations":[{"DisplayName":"Existing room","Id":"Existing room","PostalAddress":{"Type":"Home","LocationSource":"None"}}],"ExtendedProperty":[{"ExtendedFieldURI":{"PropertyName":"LocationSelectionSourceMap","DistinguishedPropertySetId":"PublicStrings"},"Value":"{\"Existing room\":\"ExplicitTyping\"}"}]}]}]}}}`
	event, state := cal.ParseEvent([]byte(getResponse))
	if state != transport.EnvelopeOK || !event.SafePersonal() {
		t.Fatal("synthetic event rejected")
	}
	bodyField := map[string]any{"__type": "SetItemField:#Exchange", "Path": map[string]any{"FieldURI": "Body"}, "Item": map[string]any{"__type": "CalendarItem:#Exchange", "Body": map[string]any{"BodyType": "HTML", "Value": "fixture-body"}}}
	prefField := map[string]any{"__type": "SetItemField:#Exchange", "Path": map[string]any{"PropertyName": "IsDefaultPreferences"}}
	sourceField := map[string]any{"__type": "SetItemField:#Exchange", "Path": map[string]any{"PropertyName": "LocationSelectionSourceMap", "PropertySetId": "00020329-0000-0000-C000-000000000046"}, "Item": map[string]any{"__type": "Message:#Exchange"}}
	locField := map[string]any{"__type": "SetItemField:#Exchange", "Path": map[string]any{"FieldURI": "Locations"}, "Item": map[string]any{"__type": "CalendarItem:#Exchange", "Locations": []any{"fixture-location"}}}
	template := baseCalendar("UpdateCalendarEventJsonRequest:#Exchange", map[string]any{"EventId": map[string]any{"Id": "fixture-event"}, "EventScope": 0, "ShouldSendUpdateToAttendees": false, "ItemChange": map[string]any{"ItemId": map[string]any{"Id": "fixture-event"}, "Updates": []any{bodyField, prefField, sourceField, locField}}})
	d, r := captureCalendar(t, "UpdateCalendarEvent", template, false)
	body, location := "Body <tag>\nsecond line", "New room"
	_, op, err := cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), cal.Params{ID: "event", Subject: "new subject", BodyText: &body, AddLocation: &location, Current: &event})
	if err != nil {
		t.Fatal(err)
	}
	decoded := unwrapBody(t, op)
	updates := decoded["Body"].(map[string]any)["ItemChange"].(map[string]any)["Updates"].([]any)
	if len(updates) != 4 {
		t.Fatalf("wrong selected fields: %d", len(updates))
	}
	b, _ := json.Marshal(decoded)
	encoded := string(b)
	for _, x := range []string{"Existing room", "New room", "second line", "LocationSelectionSourceMap"} {
		if !strings.Contains(encoded, x) {
			t.Fatalf("missing %s", x)
		}
	}
	if strings.Contains(encoded, "fixture-") || strings.Contains(encoded, "IsDefaultPreferences") || strings.Contains(encoded, "<tag>") {
		t.Fatal("fixture, unwanted property, or raw HTML was retained")
	}
	duplicate := "Existing room"
	if _, _, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), cal.Params{ID: "event", AddLocation: &duplicate, Current: &event}); err == nil {
		t.Fatal("duplicate location accepted")
	}
	event.Raw = []byte(strings.Replace(string(event.Raw), `"PropertyName":"LocationSelectionSourceMap"`, `"PropertyName":"Different"`, 1))
	if _, _, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), cal.Params{ID: "event", AddLocation: &location, Current: &event}); err == nil {
		t.Fatal("missing source metadata accepted")
	}
}

func TestCalendarCapturedTimeUpdate(t *testing.T) {
	response := `{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Success","ResponseCode":"NoError","Items":[{"__type":"CalendarItem:#Exchange","ItemId":{"Id":"event","ChangeKey":"v1"},"IsMeeting":false,"IsRecurring":false,"IsOnlineMeeting":false,"IsOrganizer":true,"IsAllDayEvent":false,"MeetingRequestWasSent":false,"CalendarItemType":"Single","StartTimeZoneId":"Mountain Standard Time","EndTimeZoneId":"Mountain Standard Time","RequiredAttendees":[],"OptionalAttendees":[],"Resources":[]}]}]}}}`
	event, state := cal.ParseEvent([]byte(response))
	if state != transport.EnvelopeOK || !event.SafeTimedMountain() {
		t.Fatal("synthetic timed event rejected")
	}
	field := func(name string, value any) map[string]any {
		return map[string]any{"__type": "SetItemField:#Exchange", "Path": map[string]any{"__type": "PropertyUri:#Exchange", "FieldURI": name}, "Item": map[string]any{"__type": "CalendarItem:#Exchange", name: value}}
	}
	pref := map[string]any{"__type": "SetItemField:#Exchange", "Path": map[string]any{"PropertyName": "IsDefaultPreferences"}}
	template := baseCalendar("UpdateCalendarEventJsonRequest:#Exchange", map[string]any{"EventId": map[string]any{"Id": "fixture-event"}, "EventScope": 0, "ShouldSendUpdateToAttendees": false, "ItemChange": map[string]any{"ItemId": map[string]any{"Id": "fixture-event"}, "Updates": []any{field("End", "fixture-end"), field("EndTimeZoneId", "Mountain Standard Time"), pref, field("OnlineMeetingProvider", nil), field("Start", "fixture-start"), field("StartTimeZoneId", "Mountain Standard Time")}}})
	d, r := captureCalendar(t, "UpdateCalendarEvent", template, false)
	p := cal.Params{ID: "event", Current: &event, Start: "2026-11-09T16:00:00Z", End: "2026-11-09T16:30:00Z"}
	_, op, err := cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), p)
	if err != nil {
		t.Fatal(err)
	}
	root := unwrapBody(t, op)
	updates := root["Body"].(map[string]any)["ItemChange"].(map[string]any)["Updates"].([]any)
	if len(updates) != 4 {
		t.Fatalf("expected four observed time fields, got %d", len(updates))
	}
	for i, want := range []struct{ name, value string }{{"End", "2026-11-09T09:30:00.000"}, {"EndTimeZoneId", "Mountain Standard Time"}, {"Start", "2026-11-09T09:00:00.000"}, {"StartTimeZoneId", "Mountain Standard Time"}} {
		f := updates[i].(map[string]any)
		if f["Path"].(map[string]any)["FieldURI"] != want.name || f["Item"].(map[string]any)[want.name] != want.value {
			t.Fatalf("wrong time field %s", want.name)
		}
	}
	b, _ := json.Marshal(root)
	if strings.Contains(string(b), "fixture") || strings.Contains(string(b), "OnlineMeetingProvider") || strings.Contains(string(b), "IsDefaultPreferences") {
		t.Fatal("unrelated template values retained")
	}
	p.End = ""
	if _, _, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), p); err == nil {
		t.Fatal("partial time change accepted")
	}
	p.End = "2026-11-09T16:30:00Z"
	badEvent := event
	badEvent.Raw = []byte(strings.Replace(string(event.Raw), "Mountain Standard Time", "Pacific Standard Time", 1))
	p.Current = &badEvent
	if _, _, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), p); err == nil {
		t.Fatal("unsupported time zone accepted")
	}
	p.Current = &event
	p.Start = "2026-11-01T01:30:00-06:00"
	p.End = "2026-11-01T02:30:00-07:00"
	if _, _, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), p); err == nil {
		t.Fatal("ambiguous DST time accepted")
	}
}

func TestCalendarLocationReplacementAndRemoval(t *testing.T) {
	response := `{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Success","ResponseCode":"NoError","Items":[{"__type":"CalendarItem:#Exchange","ItemId":{"Id":"event","ChangeKey":"v1"},"IsMeeting":false,"IsRecurring":false,"IsOnlineMeeting":false,"IsOrganizer":true,"MeetingRequestWasSent":false,"CalendarItemType":"Single","RequiredAttendees":[],"OptionalAttendees":[],"Resources":[],"Locations":[{"DisplayName":"Old room","Id":"Old room","PostalAddress":{"Type":"Home"}},{"DisplayName":"Keep room","Id":"Keep room","PostalAddress":{"Type":"Business"}}],"ExtendedProperty":[{"ExtendedFieldURI":{"PropertyName":"LocationSelectionSourceMap","DistinguishedPropertySetId":"PublicStrings"},"Value":"{\"Old room\":\"ExplicitTyping\",\"Keep room\":\"OtherSource\"}"}]}]}]}}}`
	event, state := cal.ParseEvent([]byte(response))
	if state != transport.EnvelopeOK || !event.SafePersonal() {
		t.Fatal("synthetic event rejected")
	}
	bodyField := map[string]any{"__type": "SetItemField:#Exchange", "Path": map[string]any{"FieldURI": "Body"}, "Item": map[string]any{"__type": "CalendarItem:#Exchange", "Body": map[string]any{"BodyType": "HTML", "Value": "fixture"}}}
	sourceField := map[string]any{"__type": "SetItemField:#Exchange", "Path": map[string]any{"PropertyName": "LocationSelectionSourceMap", "PropertySetId": "00020329-0000-0000-C000-000000000046"}, "Item": map[string]any{"__type": "Message:#Exchange"}}
	locField := map[string]any{"__type": "SetItemField:#Exchange", "Path": map[string]any{"FieldURI": "Locations"}, "Item": map[string]any{"__type": "CalendarItem:#Exchange", "Locations": []any{"fixture"}}}
	template := baseCalendar("UpdateCalendarEventJsonRequest:#Exchange", map[string]any{"EventId": map[string]any{"Id": "fixture"}, "EventScope": 0, "ShouldSendUpdateToAttendees": false, "ItemChange": map[string]any{"ItemId": map[string]any{"Id": "fixture"}, "Updates": []any{bodyField, sourceField, locField}}})
	d, r := captureCalendar(t, "UpdateCalendarEvent", template, false)
	old, newName := "Old room", "New room"
	_, op, err := cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), cal.Params{ID: "event", Current: &event, ReplaceLocation: &old, WithLocation: &newName})
	if err != nil {
		t.Fatal(err)
	}
	root := unwrapBody(t, op)
	updates := root["Body"].(map[string]any)["ItemChange"].(map[string]any)["Updates"].([]any)
	if len(updates) != 2 {
		t.Fatal("expected source map and locations only")
	}
	fields := updates[1].(map[string]any)["Item"].(map[string]any)["Locations"].([]any)
	if len(fields) != 2 || fields[0].(map[string]any)["DisplayName"] != "New room" || fields[1].(map[string]any)["DisplayName"] != "Keep room" {
		t.Fatal("replacement changed unrelated locations")
	}
	property := updates[0].(map[string]any)["Item"].(map[string]any)["ExtendedProperty"].([]any)[0].(map[string]any)
	var sources map[string]string
	if json.Unmarshal([]byte(property["Value"].(string)), &sources) != nil || sources["New room"] != "ExplicitTyping" || sources["Keep room"] != "OtherSource" || sources["Old room"] != "" {
		t.Fatal("replacement source map incorrect")
	}
	_, op, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), cal.Params{ID: "event", Current: &event, RemoveLocation: &old})
	if err != nil {
		t.Fatal(err)
	}
	root = unwrapBody(t, op)
	updates = root["Body"].(map[string]any)["ItemChange"].(map[string]any)["Updates"].([]any)
	fields = updates[1].(map[string]any)["Item"].(map[string]any)["Locations"].([]any)
	if len(fields) != 1 || fields[0].(map[string]any)["DisplayName"] != "Keep room" {
		t.Fatal("removal changed unrelated location")
	}
	keep := "Keep room"
	if _, _, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), cal.Params{ID: "event", Current: &event, ReplaceLocation: &old, WithLocation: &keep}); err == nil {
		t.Fatal("duplicate replacement accepted")
	}
	if _, _, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), cal.Params{ID: "event", Current: &event, RemoveLocation: &newName}); err == nil {
		t.Fatal("missing target accepted")
	}
	if _, _, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), cal.Params{ID: "event", Current: &event, AddLocation: &newName, RemoveLocation: &old}); err == nil {
		t.Fatal("multiple location operations accepted")
	}
	var item map[string]any
	if json.Unmarshal(event.Raw, &item) != nil {
		t.Fatal("invalid event")
	}
	item["Locations"] = item["Locations"].([]any)[:1]
	item["ExtendedProperty"].([]any)[0].(map[string]any)["Value"] = `{"Old room":"ExplicitTyping"}`
	last := event
	last.Raw, _ = json.Marshal(item)
	_, op, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), cal.Params{ID: "event", Current: &last, RemoveLocation: &old})
	if err != nil {
		t.Fatal(err)
	}
	root = unwrapBody(t, op)
	updates = root["Body"].(map[string]any)["ItemChange"].(map[string]any)["Updates"].([]any)
	fields = updates[1].(map[string]any)["Item"].(map[string]any)["Locations"].([]any)
	if len(fields) != 0 {
		t.Fatal("last location was not removed")
	}
	property = updates[0].(map[string]any)["Item"].(map[string]any)["ExtendedProperty"].([]any)[0].(map[string]any)
	if property["Value"] != "{}" {
		t.Fatal("empty source map not retained")
	}
}
