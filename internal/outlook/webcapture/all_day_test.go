package webcapture_test

import (
	"encoding/json"
	"strings"
	"testing"

	cal "github.com/mikefwille/outlook-web-cli/internal/outlook/calendar"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func TestAllDayInclusiveDateRange(t *testing.T) {
	for _, tc := range []struct {
		start, end, wireStart, wireEnd string
		valid                          bool
	}{
		{"2026-11-06", "2026-11-06", "2026-11-06T00:00:00.000", "2026-11-07T00:00:00.000", true},
		{"2026-11-06", "2026-11-08", "2026-11-06T00:00:00.000", "2026-11-09T00:00:00.000", true},
		{"2024-02-28", "2024-02-29", "2024-02-28T00:00:00.000", "2024-03-01T00:00:00.000", true},
		{"2026-11-08", "2026-11-06", "", "", false},
		{"2026-02-29", "2026-03-01", "", "", false},
		{"2026-1-01", "2026-01-01", "", "", false},
		{"0000-01-01", "0000-01-02", "", "", false},
		{"9999-12-31", "9999-12-31", "", "", false},
	} {
		a, b, err := cal.AllDayRange(tc.start, tc.end)
		if (err == nil) != tc.valid || tc.valid && (a != tc.wireStart || b != tc.wireEnd) {
			t.Fatalf("wrong all-day range for %q through %q: %q %q %v", tc.start, tc.end, a, b, err)
		}
	}
}

func allDayFixture(start, end string) string {
	return `{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Success","ResponseCode":"NoError","Items":[{"__type":"CalendarItem:#Exchange","ItemId":{"Id":"event","ChangeKey":"v1"},"Start":"` + start + `","End":"` + end + `","StartTimeZoneId":"Mountain Standard Time","EndTimeZoneId":"Mountain Standard Time","IsAllDayEvent":true,"IsMeeting":false,"IsRecurring":false,"IsOnlineMeeting":false,"IsOrganizer":true,"MeetingRequestWasSent":false,"CalendarItemType":"Single","RequiredAttendees":[],"OptionalAttendees":[],"Resources":[]}]}]}}}`
}

func TestAllDayGetAndCapturedCreateUpdate(t *testing.T) {
	data := allDayFixture("2026-03-07T00:00:00-07:00", "2026-03-09T00:00:00-06:00")
	e, state := cal.ParseEvent([]byte(data))
	if state != transport.EnvelopeOK || !e.SafeAllDayMountain() || e.StartDate != "2026-03-07" || e.EndDate != "2026-03-08" {
		t.Fatal("all-day get did not preserve inclusive dates across DST")
	}
	for _, altered := range []string{strings.Replace(data, `"IsAllDayEvent":true`, `"IsAllDayEvent":false`, 1), strings.Replace(data, `"IsMeeting":false`, `"IsMeeting":true`, 1), strings.Replace(data, `2026-03-09T00:00:00-06:00`, `2026-03-09T00:00:00-07:00`, 1), strings.Replace(data, `"EndTimeZoneId":"Mountain Standard Time"`, `"EndTimeZoneId":"Eastern Standard Time"`, 1)} {
		x, _ := cal.ParseEvent([]byte(altered))
		if x.SafeAllDayMountain() || x.StartDate != "" || x.EndDate != "" {
			t.Fatal("unsupported all-day event normalized")
		}
	}
	createItem := map[string]any{"__type": "CalendarItem:#Exchange", "IsAllDayEvent": true, "StartTimeZoneId": "Mountain Standard Time", "EndTimeZoneId": "Mountain Standard Time", "ItemId": map[string]any{"Id": "fixture-event"}, "id": "fixture-event", "ClientSeriesId": "fixture-event", "ParentFolderId": map[string]any{"Id": "fixture-folder"}, "Resources": []any{}, "Attachments": []any{}, "Body": map[string]any{"BodyType": "HTML", "Value": ""}, "Locations": []any{}, "ExtendedProperty": []any{}}
	create := baseCalendar("CreateItemJsonRequest:#Exchange", map[string]any{"SendMeetingInvitations": "SendToNone", "SavedItemFolderId": map[string]any{"BaseFolderId": map[string]any{"Id": "fixture-folder"}}, "Items": []any{createItem}})
	d, r := captureCalendar(t, "CreateCalendarEvent", create, false)
	body, location := "Body <safe>\nnext", "Synthetic room"
	p := cal.Params{CalendarID: "folder", Subject: "All day", StartDate: "2026-11-06", EndDate: "2026-11-07", BodyText: &body, AddLocation: &location}
	_, op, err := cal.Prepare("CreateCalendarEvent", d, r, accountSession(), p)
	if err != nil {
		t.Fatal(err)
	}
	created := unwrapBody(t, op)["Body"].(map[string]any)["Items"].([]any)[0].(map[string]any)
	if op.Class != transport.Write || created["Start"] != "2026-11-06T00:00:00.000" || created["End"] != "2026-11-08T00:00:00.000" || strings.Contains(string(op.Body), "fixture-") {
		t.Fatal("all-day create did not rebind dates and identifiers")
	}
	if created["Body"].(map[string]any)["Value"] != "<div>Body &lt;safe&gt;<br>next</div>" || len(created["Locations"].([]any)) != 1 || len(created["ExtendedProperty"].([]any)) != 1 {
		t.Fatal("all-day create lost body or location")
	}
	p.EndDate = ""
	if _, _, err = cal.Prepare("CreateCalendarEvent", d, r, accountSession(), p); err == nil {
		t.Fatal("missing end date accepted")
	}
	p.EndDate, p.Start = "2026-11-07", "2026-11-06T09:00:00-07:00"
	if _, _, err = cal.Prepare("CreateCalendarEvent", d, r, accountSession(), p); err == nil {
		t.Fatal("mixed all-day and timed create accepted")
	}

	fields := []any{}
	for _, v := range []struct{ name, value string }{{"End", "fixture-end"}, {"EndTimeZoneId", "Mountain Standard Time"}, {"Start", "fixture-start"}} {
		fields = append(fields, map[string]any{"__type": "SetItemField:#Exchange", "Path": map[string]any{"__type": "PropertyUri:#Exchange", "FieldURI": v.name}, "Item": map[string]any{"__type": "CalendarItem:#Exchange", v.name: v.value}})
	}
	update := baseCalendar("UpdateCalendarEventJsonRequest:#Exchange", map[string]any{"EventId": map[string]any{"Id": "fixture-event"}, "ItemChange": map[string]any{"ItemId": map[string]any{"Id": "fixture-event"}, "Updates": fields}, "EventScope": 0, "ShouldSendUpdateToAttendees": false})
	d, r = captureCalendar(t, "UpdateCalendarEvent", update, false)
	p = cal.Params{ID: e.ID, Current: &e, StartDate: "2026-11-06", EndDate: "2026-11-08"}
	_, op, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), p)
	if err != nil {
		t.Fatal(err)
	}
	x := unwrapBody(t, op)["Body"].(map[string]any)["ItemChange"].(map[string]any)["Updates"].([]any)
	if len(x) != 3 {
		t.Fatal("wrong all-day field count")
	}
	for i, want := range []struct{ name, value string }{{"End", "2026-11-09T00:00:00.000"}, {"EndTimeZoneId", "Mountain Standard Time"}, {"Start", "2026-11-06T00:00:00.000"}} {
		field := x[i].(map[string]any)
		if field["Path"].(map[string]any)["FieldURI"] != want.name || field["Item"].(map[string]any)[want.name] != want.value {
			t.Fatal("wrong all-day update field")
		}
	}
	encoded, _ := json.Marshal(unwrapBody(t, op))
	if strings.Contains(string(encoded), "fixture-") {
		t.Fatal("update retained fixture values")
	}
	p.Subject = "mixed"
	if _, _, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), p); err == nil {
		t.Fatal("mixed date and subject update accepted")
	}
}
