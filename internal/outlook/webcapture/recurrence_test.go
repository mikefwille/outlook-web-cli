package webcapture_test

import (
	"encoding/json"
	"strings"
	"testing"

	cal "github.com/mikefwille/outlook-web-cli/internal/outlook/calendar"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func TestWeeklySeriesCreateAndOccurrenceSubject(t *testing.T) {
	recurrence := map[string]any{"__type": "RecurrenceType:#Exchange", "RecurrencePattern": map[string]any{"__type": "WeeklyRecurrence:#Exchange", "Interval": 1, "DaysOfWeek": "Friday", "FirstDayOfWeek": "Sunday"}, "RecurrenceRange": map[string]any{"__type": "EndDateRecurrence:#Exchange", "StartDate": "2000-01-01", "EndDate": "2000-01-02"}}
	item := map[string]any{"__type": "CalendarItem:#Exchange", "IsAllDayEvent": false, "ItemId": map[string]any{"Id": "fixture-event"}, "id": "fixture-event", "ClientSeriesId": "fixture-event", "ParentFolderId": map[string]any{"Id": "fixture-folder"}, "Resources": []any{}, "Attachments": []any{}, "Body": map[string]any{"BodyType": "HTML", "Value": ""}, "Locations": []any{}, "ExtendedProperty": []any{}, "Recurrence": recurrence}
	create := baseCalendar("CreateItemJsonRequest:#Exchange", map[string]any{"SendMeetingInvitations": "SendToNone", "SavedItemFolderId": map[string]any{"BaseFolderId": map[string]any{"Id": "fixture-folder"}}, "Items": []any{item}})
	d, r := captureCalendar(t, "CreateCalendarEvent", create, false)
	p := cal.Params{CalendarID: "folder", Subject: "Weekly", Start: "2026-11-06T09:00:00-07:00", End: "2026-11-06T09:30:00-07:00", WeeklyUntil: "2026-12-04"}
	_, op, err := cal.Prepare("CreateCalendarEvent", d, r, accountSession(), p)
	if err != nil {
		t.Fatal(err)
	}
	if op.Class != transport.Write || strings.Contains(string(op.Body), "fixture-") {
		t.Fatal("fixture identifiers retained")
	}
	created := unwrapBody(t, op)["Body"].(map[string]any)["Items"].([]any)[0].(map[string]any)
	rec := created["Recurrence"].(map[string]any)
	pat, dates := rec["RecurrencePattern"].(map[string]any), rec["RecurrenceRange"].(map[string]any)
	if pat["DaysOfWeek"] != "Friday" || dates["StartDate"] != "2026-11-06" || dates["EndDate"] != p.WeeklyUntil {
		t.Fatal("incorrect recurrence")
	}
	p.WeeklyUntil = "2026-11-05"
	if _, _, err = cal.Prepare("CreateCalendarEvent", d, r, accountSession(), p); err == nil {
		t.Fatal("end date before start accepted")
	}
	p.WeeklyUntil = "2026-12-04"
	p.BodyText = strptr("uncaptured body")
	if _, _, err = cal.Prepare("CreateCalendarEvent", d, r, accountSession(), p); err == nil {
		t.Fatal("unsupported body on recurring create accepted")
	}
	p.BodyText = nil
	p.WeeklyUntil = ""
	if _, _, err = cal.Prepare("CreateCalendarEvent", d, r, accountSession(), p); err == nil {
		t.Fatal("recurring template without recurrence request accepted")
	}

	base := `{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Success","ResponseCode":"NoError","Items":[{"__type":"CalendarItem:#Exchange","ItemId":{"Id":"instance","ChangeKey":"v1"},"IsMeeting":false,"IsRecurring":true,"IsOnlineMeeting":false,"IsOrganizer":true,"IsAllDayEvent":false,"StartTimeZoneId":"Mountain Standard Time","EndTimeZoneId":"Mountain Standard Time","MeetingRequestWasSent":false,"CalendarItemType":"Occurrence","RequiredAttendees":[],"OptionalAttendees":[],"Resources":[],"SeriesMasterItemId":{"Id":"master"}}]}]}}}`
	e, state := cal.ParseEvent([]byte(base))
	if state != transport.EnvelopeOK || !e.SafeOccurrence() {
		t.Fatal("valid occurrence rejected")
	}
	for _, v := range []string{strings.Replace(base, `"Occurrence"`, `"Exception"`, 1), strings.Replace(base, `"Occurrence"`, `"RecurringMaster"`, 1), strings.Replace(base, `"IsMeeting":false`, `"IsMeeting":true`, 1), strings.Replace(base, `"SeriesMasterItemId":{"Id":"master"}`, `"SeriesMasterItemId":{}`, 1)} {
		x, _ := cal.ParseEvent([]byte(v))
		want := strings.Contains(v, `"CalendarItemType":"Exception"`)
		if x.SafeOccurrence() != want {
			t.Fatal("incorrect occurrence safety decision")
		}
	}
	for _, v := range []string{
		strings.Replace(base, `"IsAllDayEvent":false`, `"IsAllDayEvent":true`, 1),
		strings.Replace(base, `"IsAllDayEvent":false,`, ``, 1),
		strings.ReplaceAll(base, `Mountain Standard Time`, `Eastern Standard Time`),
		strings.Replace(base, `"EndTimeZoneId":"Mountain Standard Time"`, `"EndTimeZoneId":"Eastern Standard Time"`, 1),
		strings.Replace(base, `"StartTimeZoneId":"Mountain Standard Time",`, ``, 1),
	} {
		unsafe, state := cal.ParseEvent([]byte(v))
		if state != transport.EnvelopeOK || unsafe.SafeOccurrence() {
			t.Fatal("uncaptured all-day or time-zone occurrence accepted")
		}
	}
	update := baseCalendar("UpdateCalendarEventJsonRequest:#Exchange", map[string]any{"EventId": map[string]any{"Id": "fixture-event"}, "ItemChange": map[string]any{"ItemId": map[string]any{"Id": "fixture-event"}, "Updates": []any{map[string]any{"__type": "SetItemField:#Exchange", "Path": map[string]any{"FieldURI": "Subject"}, "Item": map[string]any{"__type": "CalendarItem:#Exchange", "Subject": "fixture-subject"}}}}, "EventScope": 0, "ShouldSendUpdateToAttendees": false})
	d, r = captureCalendar(t, "UpdateCalendarEvent", update, true)
	_, op, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), cal.Params{ID: e.ID, Current: &e, Occurrence: true, Subject: "New subject"})
	if err != nil {
		t.Fatal(err)
	}
	encoded, _ := json.Marshal(unwrapBody(t, op))
	if strings.Contains(string(encoded), "fixture-") || !strings.Contains(string(encoded), "New subject") {
		t.Fatal("occurrence update not rebound")
	}
	for _, v := range []string{
		strings.Replace(base, `"IsAllDayEvent":false`, `"IsAllDayEvent":true`, 1),
		strings.ReplaceAll(base, `Mountain Standard Time`, `Eastern Standard Time`),
	} {
		unsafe, _ := cal.ParseEvent([]byte(v))
		if _, _, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), cal.Params{ID: unsafe.ID, Current: &unsafe, Occurrence: true, Subject: "New subject"}); err == nil {
			t.Fatal("uncaptured occurrence update prepared")
		}
	}
	if _, _, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), cal.Params{ID: e.ID, Current: &e, Occurrence: true, Subject: "New", Start: "2026-11-06T09:00:00-07:00"}); err == nil {
		t.Fatal("unsupported occurrence time change accepted")
	}
}

func TestExpandOccurrencesCapturedShape(t *testing.T) {
	body := map[string]any{"__type": "ExpandCalendarEventRequest:#Exchange", "EventId": map[string]any{"Id": "fixture-event"}, "WindowStart": "old", "WindowEnd": "old", "ReturnMaster": true, "ReturnRegularOccurrences": true, "ReturnExceptions": true, "ReturnCancellations": true}
	d, _ := captureCalendar(t, "ExpandCalendarEvent", baseCalendar("ExpandCalendarEventJsonRequest:#Exchange", body), true)
	_, op, err := cal.PrepareOccurrences(d, accountSession(), "master", "2026-11-06T09:00:00-07:00", "2026-11-20T09:00:00-07:00")
	if err != nil {
		t.Fatal(err)
	}
	result := unwrapBody(t, op)["Body"].(map[string]any)
	if op.Class != transport.Read || result["EventId"].(map[string]any)["Id"] != "master" || result["WindowStart"] != "2026-11-06T09:00:00.000" || result["WindowEnd"] != "2026-11-20T09:00:00.000" {
		t.Fatal("wrong expansion request")
	}
	if _, _, err = cal.PrepareOccurrences(d, accountSession(), "master", "2026-11-20T09:00:00-07:00", "2026-11-06T09:00:00-07:00"); err == nil {
		t.Fatal("reversed range accepted")
	}
	response := `{"Body":{"ResponseClass":"Success","ResponseCode":"NoError","RecurrenceMaster":{"ItemId":{"Id":"master"}},"Occurrences":[{"Subject":"one"}],"RegularOccurences":[],"ExceptionalOccurences":[],"CancelledOccurrences":[]}}`
	v, state := cal.ParseOccurrences([]byte(response))
	if state != transport.EnvelopeOK || len(v.Occurrences) != 1 {
		t.Fatal("captured expansion rejected")
	}
	if _, state = cal.ParseOccurrences([]byte(strings.Replace(response, `"CancelledOccurrences":[]`, `"CancelledOccurrences":null`, 1))); state != transport.EnvelopeError {
		t.Fatal("missing cancellation list accepted")
	}
}

func TestWeeklySeriesDeletionRequiresCompleteAttendeeFreeExpansion(t *testing.T) {
	masterJSON := `{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Success","ResponseCode":"NoError","Items":[{"__type":"CalendarItem:#Exchange","ItemId":{"Id":"master","ChangeKey":"v1"},"IsMeeting":false,"IsRecurring":true,"IsOnlineMeeting":false,"IsOrganizer":true,"IsAllDayEvent":false,"MeetingRequestWasSent":false,"CalendarItemType":"RecurringMaster","RequiredAttendees":[],"OptionalAttendees":[],"Resources":[],"Recurrence":{"RecurrencePattern":{"__type":"WeeklyRecurrence:#Exchange","Interval":1,"DaysOfWeek":"Friday","FirstDayOfWeek":"Sunday"},"RecurrenceRange":{"__type":"EndDateRecurrence:#Exchange","StartDate":"2026-11-06T00:00:00-07:00","EndDate":"2026-11-20T00:00:00-07:00"}}}]}]}}}`
	e, state := cal.ParseEvent([]byte(masterJSON))
	if state != transport.EnvelopeOK || !e.SafeWeeklySeries() {
		t.Fatal("valid weekly master rejected")
	}
	from, to, expected, err := e.WeeklyDeleteWindow()
	if err != nil || expected != 3 || from != "2026-11-06T00:00:00-07:00" || to != "2026-11-21T00:00:00-07:00" {
		t.Fatal("wrong series deletion window")
	}
	item := func(id, kind string) json.RawMessage {
		return json.RawMessage(`{"ItemId":{"Id":"` + id + `"},"SeriesMasterItemId":{"Id":"master"},"IsMeeting":false,"IsRecurring":true,"IsOnlineMeeting":false,"IsOrganizer":true,"IsAllDayEvent":false,"StartTimeZoneId":"Mountain Standard Time","EndTimeZoneId":"Mountain Standard Time","MeetingRequestWasSent":false,"CalendarItemType":"` + kind + `","RequiredAttendees":[],"OptionalAttendees":[],"Resources":[]}`)
	}
	a, b, c := item("one", "Occurrence"), item("two", "Occurrence"), item("three", "Exception")
	x := cal.Occurrences{Master: json.RawMessage(`{"ItemId":{"Id":"master"}}`), Occurrences: []json.RawMessage{a, b, c}, Regular: []json.RawMessage{a, b}, Exceptions: []json.RawMessage{c}, Cancellations: []json.RawMessage{}}
	if !x.SafeWeeklyDeletion("master", expected) {
		t.Fatal("complete safe expansion rejected")
	}
	partial := x
	partial.Occurrences = []json.RawMessage{a, c}
	if partial.SafeWeeklyDeletion("master", expected) {
		t.Fatal("incomplete expansion accepted")
	}
	unsafe := x
	unsafe.Occurrences = []json.RawMessage{a, b, json.RawMessage(strings.Replace(string(c), `"IsMeeting":false`, `"IsMeeting":true`, 1))}
	if unsafe.SafeWeeklyDeletion("master", expected) {
		t.Fatal("meeting exception accepted")
	}
	wrongZone := json.RawMessage(strings.ReplaceAll(string(c), "Mountain Standard Time", "Eastern Standard Time"))
	unsafe = x
	unsafe.Occurrences = []json.RawMessage{a, b, wrongZone}
	unsafe.Exceptions = []json.RawMessage{wrongZone}
	if unsafe.SafeWeeklyDeletion("master", expected) {
		t.Fatal("exception in an uncaptured zone accepted")
	}
	cancelled := x
	cancelled.Cancellations = []json.RawMessage{json.RawMessage(`{}`)}
	if cancelled.SafeWeeklyDeletion("master", expected) {
		t.Fatal("unsupported cancellation accepted")
	}
	deleteRequest := baseCalendar("DeleteCalendarEventJsonRequest:#Exchange", map[string]any{"__type": "DeleteCalendarEventRequest:#Exchange", "EventId": map[string]any{"Id": "fixture-id"}, "EventScope": 0})
	d, r := captureCalendar(t, "DeleteCalendarEvent", deleteRequest, true)
	_, op, err := cal.Prepare("DeleteCalendarEvent", d, r, accountSession(), cal.Params{ID: e.ID, Current: &e, Series: true})
	if err != nil {
		t.Fatal(err)
	}
	payload, _ := json.Marshal(unwrapBody(t, op))
	if op.Class != transport.Write || strings.Contains(string(payload), "fixture-id") || !strings.Contains(string(payload), "master") {
		t.Fatal("series delete not rebound")
	}
	if _, _, err = cal.Prepare("DeleteCalendarEvent", d, r, accountSession(), cal.Params{ID: "different", Current: &e, Series: true}); err == nil {
		t.Fatal("mismatched series master accepted")
	}
}
