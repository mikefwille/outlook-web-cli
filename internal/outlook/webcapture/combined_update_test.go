package webcapture_test

import (
	"encoding/json"
	"strings"
	"testing"

	cal "github.com/mikefwille/outlook-web-cli/internal/outlook/calendar"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func TestCombinedTimedUpdatePreservesDuplicateLocations(t *testing.T) {
	raw := `{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Success","ResponseCode":"NoError","Items":[{"__type":"CalendarItem:#Exchange","ItemId":{"Id":"event","ChangeKey":"v1"},"Start":"2026-11-06T04:00:00-07:00","End":"2026-11-06T04:30:00-07:00","StartTimeZoneId":"Mountain Standard Time","EndTimeZoneId":"Mountain Standard Time","IsAllDayEvent":false,"IsMeeting":false,"IsRecurring":false,"IsOnlineMeeting":false,"IsOrganizer":true,"MeetingRequestWasSent":false,"CalendarItemType":"Single","RequiredAttendees":[],"OptionalAttendees":[],"Resources":[],"Locations":[{"DisplayName":"Duplicate room","Marker":"first"},{"DisplayName":"Duplicate room","Marker":"second"}],"ExtendedProperty":[]}]}]}}}`
	e, state := cal.ParseEvent([]byte(raw))
	if state != transport.EnvelopeOK || !e.SafeTimedMountain() {
		t.Fatal("timed fixture rejected")
	}
	field := func(name string, value any) map[string]any {
		return map[string]any{"__type": "SetItemField:#Exchange", "Path": map[string]any{"__type": "PropertyUri:#Exchange", "FieldURI": name}, "Item": map[string]any{"__type": "CalendarItem:#Exchange", name: value}}
	}
	updates := []any{field("Body", map[string]any{"BodyType": "HTML", "Value": "fixture-body"}), field("End", "fixture-end"), field("EndTimeZoneId", "Mountain Standard Time"), field("Locations", []any{map[string]any{"DisplayName": "fixture-location"}}), field("Start", "fixture-start"), field("StartTimeZoneId", "Mountain Standard Time")}
	req := baseCalendar("UpdateCalendarEventJsonRequest:#Exchange", map[string]any{"EventId": map[string]any{"Id": "fixture-event"}, "ItemChange": map[string]any{"ItemId": map[string]any{"Id": "fixture-event"}, "Updates": updates}, "EventScope": 0, "ShouldSendUpdateToAttendees": false})
	d, r := captureCalendar(t, "UpdateCalendarEvent", req, false)
	body, location := "Synthetic combined body", "New room"
	p := cal.Params{ID: e.ID, Current: &e, Start: "2026-11-06T07:00:00-07:00", End: "2026-11-06T07:30:00-07:00", BodyText: &body, AddLocation: &location}
	_, op, err := cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), p)
	if err != nil {
		t.Fatal(err)
	}
	fields := unwrapBody(t, op)["Body"].(map[string]any)["ItemChange"].(map[string]any)["Updates"].([]any)
	if len(fields) != 6 {
		t.Fatal("combined update field count incorrect")
	}
	for i, want := range []string{"Body", "End", "EndTimeZoneId", "Locations", "Start", "StartTimeZoneId"} {
		if fields[i].(map[string]any)["Path"].(map[string]any)["FieldURI"] != want {
			t.Fatal("combined update field order incorrect")
		}
	}
	locations := fields[3].(map[string]any)["Item"].(map[string]any)["Locations"].([]any)
	if len(locations) != 3 || locations[0].(map[string]any)["Marker"] != "first" || locations[1].(map[string]any)["Marker"] != "second" || locations[2].(map[string]any)["DisplayName"] != location {
		t.Fatal("existing duplicate locations not preserved")
	}
	encoded, _ := json.Marshal(unwrapBody(t, op))
	if strings.Contains(string(encoded), "fixture-") {
		t.Fatal("combined update retained fixture data")
	}
	p.AddLocation = nil
	_, op, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), p)
	if err != nil {
		t.Fatal(err)
	}
	fields = unwrapBody(t, op)["Body"].(map[string]any)["ItemChange"].(map[string]any)["Updates"].([]any)
	for _, x := range fields {
		if x.(map[string]any)["Path"].(map[string]any)["FieldURI"] == "Locations" {
			t.Fatal("unrequested duplicate locations changed")
		}
	}
	target := "Duplicate room"
	p.RemoveLocation = &target
	if _, _, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), p); err == nil {
		t.Fatal("ambiguous duplicate removal accepted")
	}
	withSource := strings.Replace(raw, `"ExtendedProperty":[]`, `"ExtendedProperty":[{"ExtendedFieldURI":{"PropertyName":"LocationSelectionSourceMap"},"Value":"{}"}]`, 1)
	protected, _ := cal.ParseEvent([]byte(withSource))
	p.Current = &protected
	p.RemoveLocation = nil
	p.AddLocation = &location
	if _, _, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), p); err == nil {
		t.Fatal("combined edit would discard location source metadata")
	}
}
