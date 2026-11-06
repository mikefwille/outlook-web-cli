package webcapture_test

import (
	"strings"
	"testing"

	cal "github.com/mikefwille/outlook-web-cli/internal/outlook/calendar"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func TestWeeklySeriesEditRebuildsCapturedRecurrence(t *testing.T) {
	master := `{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Success","ResponseCode":"NoError","Items":[{"__type":"CalendarItem:#Exchange","ItemId":{"Id":"master","ChangeKey":"v1"},"Start":"2026-11-06T04:00:00-07:00","End":"2026-11-06T04:30:00-07:00","StartTimeZoneId":"Mountain Standard Time","EndTimeZoneId":"Mountain Standard Time","IsMeeting":false,"IsRecurring":true,"IsOnlineMeeting":false,"IsOrganizer":true,"IsAllDayEvent":false,"MeetingRequestWasSent":false,"CalendarItemType":"RecurringMaster","RequiredAttendees":[],"OptionalAttendees":[],"Resources":[],"Recurrence":{"RecurrencePattern":{"__type":"WeeklyRecurrence:#Exchange","Interval":1,"DaysOfWeek":"Friday","FirstDayOfWeek":"Sunday"},"RecurrenceRange":{"__type":"EndDateRecurrence:#Exchange","StartDate":"2026-11-06T00:00:00-07:00","EndDate":"2026-11-20T00:00:00-07:00"}}}]}]}}}`
	e, state := cal.ParseEvent([]byte(master))
	if state != transport.EnvelopeOK || !e.SafeWeeklySeries() {
		t.Fatal("synthetic weekly master rejected")
	}
	field := func(name string, value any) map[string]any {
		return map[string]any{"__type": "SetItemField:#Exchange", "Path": map[string]any{"__type": "PropertyUri:#Exchange", "FieldURI": name}, "Item": map[string]any{"__type": "CalendarItem:#Exchange", name: value}}
	}
	recurrence := map[string]any{"__type": "RecurrenceType:#Exchange", "RecurrencePattern": map[string]any{"__type": "WeeklyRecurrence:#Exchange", "__typename": "WeeklyRecurrence", "DaysOfWeek": "fixture-day", "FirstDayOfWeek": "Sunday", "Interval": 1}, "RecurrenceRange": map[string]any{"__type": "EndDateRecurrence:#Exchange", "__typename": "EndDateRecurrence", "StartDate": "2000-01-01", "EndDate": "2000-01-02"}}
	template := func(updates []any, header bool) (string, string) {
		req := baseCalendar("UpdateCalendarEventJsonRequest:#Exchange", map[string]any{"EventId": map[string]any{"Id": "fixture-event"}, "ItemChange": map[string]any{"ItemId": map[string]any{"Id": "fixture-event"}, "Updates": updates}, "EventScope": 0, "ShouldSendUpdateToAttendees": false})
		return captureCalendar(t, "UpdateCalendarEvent", req, header)
	}
	d, r := template([]any{field("Subject", "fixture-subject")}, true)
	_, op, err := cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), cal.Params{ID: e.ID, Current: &e, Series: true, Subject: "Renamed"})
	if err != nil || op.Class != transport.Write {
		t.Fatalf("series subject rejected: %v", err)
	}
	otherZone, _ := cal.ParseEvent([]byte(strings.ReplaceAll(master, "Mountain Standard Time", "Eastern Standard Time")))
	if _, _, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), cal.Params{ID: otherZone.ID, Current: &otherZone, Series: true, Subject: "Renamed"}); err == nil {
		t.Fatal("uncaptured series time zone accepted")
	}
	d, r = template([]any{field("End", "fixture-end"), field("EndTimeZoneId", "Mountain Standard Time"), field("Recurrence", recurrence), field("Start", "fixture-start"), field("StartTimeZoneId", "Mountain Standard Time")}, false)
	p := cal.Params{ID: e.ID, Current: &e, Series: true, ResetExceptions: true, Start: "2026-11-06T06:00:00-07:00", End: "2026-11-06T06:30:00-07:00"}
	_, op, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), p)
	if err != nil {
		t.Fatal(err)
	}
	updates := unwrapBody(t, op)["Body"].(map[string]any)["ItemChange"].(map[string]any)["Updates"].([]any)
	if len(updates) != 5 {
		t.Fatal("series time field count incorrect")
	}
	for i, want := range []string{"End", "EndTimeZoneId", "Recurrence", "Start", "StartTimeZoneId"} {
		if updates[i].(map[string]any)["Path"].(map[string]any)["FieldURI"] != want {
			t.Fatal("series time field order incorrect")
		}
	}
	rebuilt := updates[2].(map[string]any)["Item"].(map[string]any)["Recurrence"].(map[string]any)
	pattern, dates := rebuilt["RecurrencePattern"].(map[string]any), rebuilt["RecurrenceRange"].(map[string]any)
	if pattern["DaysOfWeek"] != "Friday" || dates["StartDate"] != "2026-11-06" || dates["EndDate"] != "2026-11-20" || strings.Contains(string(op.Body), "fixture-") {
		t.Fatal("series time retained template values")
	}
	p.ResetExceptions = false
	if _, _, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), p); err == nil {
		t.Fatal("series time change without reset acknowledgment accepted")
	}
	p.ResetExceptions, p.Start, p.End = true, "2026-11-13T06:00:00-07:00", "2026-11-13T06:30:00-07:00"
	if _, _, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), p); err == nil {
		t.Fatal("series time change moved the start date")
	}
	d, r = template([]any{field("Recurrence", recurrence)}, false)
	p = cal.Params{ID: e.ID, Current: &e, Series: true, ResetExceptions: true, WeeklyUntil: "2026-11-27"}
	_, op, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), p)
	if err != nil {
		t.Fatal(err)
	}
	updates = unwrapBody(t, op)["Body"].(map[string]any)["ItemChange"].(map[string]any)["Updates"].([]any)
	if len(updates) != 1 {
		t.Fatal("series recurrence field count incorrect")
	}
	rebuilt = updates[0].(map[string]any)["Item"].(map[string]any)["Recurrence"].(map[string]any)
	dates = rebuilt["RecurrenceRange"].(map[string]any)
	if dates["StartDate"] != "2026-11-06" || dates["EndDate"] != "2026-11-27" {
		t.Fatal("series end date incorrectly rebuilt")
	}
	p.WeeklyUntil = "2026-11-05"
	if _, _, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), p); err == nil {
		t.Fatal("end before first occurrence accepted")
	}
	p.WeeklyUntil = "2028-11-27"
	if _, _, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), p); err == nil {
		t.Fatal("unsupported long range accepted")
	}
}
