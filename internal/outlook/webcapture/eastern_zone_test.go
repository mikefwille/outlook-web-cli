package webcapture_test

import (
	"strings"
	"testing"

	cal "github.com/mikefwille/outlook-web-cli/internal/outlook/calendar"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func easternHeader(root map[string]any) {
	root["Header"].(map[string]any)["TimeZoneContext"].(map[string]any)["TimeZoneDefinition"].(map[string]any)["Id"] = "Eastern Standard Time"
}

func TestEasternZoneCreateUpdateAndGet(t *testing.T) {
	raw := strings.ReplaceAll(allDayFixture("2026-11-06T02:00:00-07:00", "2026-11-06T02:30:00-07:00"), "Mountain Standard Time", "Eastern Standard Time")
	raw = strings.Replace(raw, `"IsAllDayEvent":true`, `"IsAllDayEvent":false`, 1)
	e, state := cal.ParseEvent([]byte(raw))
	zone, safe := e.SafeTimedZone()
	if state != transport.EnvelopeOK || !safe || zone != "Eastern Standard Time" || e.StartTimeZone != zone || e.EventStart != "2026-11-06T04:00:00-05:00" || e.EventEnd != "2026-11-06T04:30:00-05:00" {
		t.Fatal("Eastern event-local get incorrect")
	}
	item := map[string]any{"__type": "CalendarItem:#Exchange", "IsAllDayEvent": false, "StartTimeZoneId": zone, "EndTimeZoneId": zone, "ItemId": map[string]any{"Id": "fixture-event"}, "id": "fixture-event", "ClientSeriesId": "fixture-event", "ParentFolderId": map[string]any{"Id": "fixture-folder"}, "Resources": []any{}, "Attachments": []any{}, "Body": map[string]any{"BodyType": "HTML", "Value": ""}, "Locations": []any{}, "ExtendedProperty": []any{}}
	create := baseCalendar("CreateItemJsonRequest:#Exchange", map[string]any{"SendMeetingInvitations": "SendToNone", "SavedItemFolderId": map[string]any{"BaseFolderId": map[string]any{"Id": "fixture-folder"}}, "Items": []any{item}})
	easternHeader(create)
	d, r := captureCalendar(t, "CreateCalendarEvent", create, false)
	body, location := "Body <safe>", "Synthetic room"
	p := cal.Params{CalendarID: "folder", Subject: "Eastern", Start: "2026-11-06T02:00:00-07:00", End: "2026-11-06T02:30:00-07:00", TimeZone: "America/New_York", BodyText: &body, AddLocation: &location}
	_, op, err := cal.Prepare("CreateCalendarEvent", d, r, accountSession(), p)
	if err != nil {
		t.Fatal(err)
	}
	x := unwrapBody(t, op)["Body"].(map[string]any)["Items"].([]any)[0].(map[string]any)
	if x["Start"] != "2026-11-06T04:00:00.000" || x["End"] != "2026-11-06T04:30:00.000" || x["StartTimeZoneId"] != zone || strings.Contains(string(op.Body), "fixture-") {
		t.Fatal("Eastern create did not preserve event zone")
	}
	if x["Body"].(map[string]any)["Value"] != "<div>Body &lt;safe&gt;</div>" || len(x["Locations"].([]any)) != 1 || len(x["ExtendedProperty"].([]any)) != 1 {
		t.Fatal("Eastern create lost body or location")
	}
	p.BodyText, p.AddLocation = nil, nil
	p.WeeklyUntil = "2026-12-04"
	if _, _, err = cal.Prepare("CreateCalendarEvent", d, r, accountSession(), p); err == nil {
		t.Fatal("uncaptured Eastern recurrence accepted")
	}
	fields := []any{}
	for _, v := range []struct{ name, value string }{{"End", "fixture"}, {"EndTimeZoneId", zone}, {"Start", "fixture"}, {"StartTimeZoneId", zone}} {
		fields = append(fields, map[string]any{"__type": "SetItemField:#Exchange", "Path": map[string]any{"__type": "PropertyUri:#Exchange", "FieldURI": v.name}, "Item": map[string]any{"__type": "CalendarItem:#Exchange", v.name: v.value}})
	}
	update := baseCalendar("UpdateCalendarEventJsonRequest:#Exchange", map[string]any{"EventId": map[string]any{"Id": "fixture-event"}, "ItemChange": map[string]any{"ItemId": map[string]any{"Id": "fixture-event"}, "Updates": fields}, "EventScope": 0, "ShouldSendUpdateToAttendees": false})
	easternHeader(update)
	d, r = captureCalendar(t, "UpdateCalendarEvent", update, false)
	p = cal.Params{ID: e.ID, Current: &e, Start: "2026-11-06T05:00:00-05:00", End: "2026-11-06T05:30:00-05:00"}
	_, op, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), p)
	if err != nil {
		t.Fatal(err)
	}
	y := unwrapBody(t, op)["Body"].(map[string]any)["ItemChange"].(map[string]any)["Updates"].([]any)
	if len(y) != 4 || y[0].(map[string]any)["Item"].(map[string]any)["End"] != "2026-11-06T05:30:00.000" || y[2].(map[string]any)["Item"].(map[string]any)["Start"] != "2026-11-06T05:00:00.000" {
		t.Fatal("Eastern update did not preserve event-local time")
	}
	p.TimeZone = "Mountain Standard Time"
	if _, _, err = cal.Prepare("UpdateCalendarEvent", d, r, accountSession(), p); err == nil {
		t.Fatal("uncaptured timezone change accepted")
	}
}
