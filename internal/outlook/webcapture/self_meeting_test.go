package webcapture_test

import (
	"encoding/base64"
	"encoding/json"
	"strings"
	"testing"
	"time"

	cal "github.com/mikefwille/outlook-web-cli/internal/outlook/calendar"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func syntheticSelfSession() transport.Session {
	s := accountSession()
	claims, _ := json.Marshal(map[string]any{"tid": "tenant-test", "oid": "object-test", "aud": "aud-test", "exp": time.Now().Add(time.Hour).Unix(), "upn": "self@example.invalid", "name": "Synthetic Self"})
	s.Headers.Set("Authorization", "Bearer e30."+base64.RawURLEncoding.EncodeToString(claims)+".signature")
	return s
}

func TestSelfOnlineEventIsStrictlyBoundToCurrentAccount(t *testing.T) {
	self := syntheticSelfSession()
	attendee := map[string]any{"IsCoauthor": false, "Mailbox": map[string]any{"EmailAddress": "self@example.invalid", "Name": "fixture-name", "MailboxType": "InternalMailbox", "RoutingType": "SMTP", "AADObjectId": "object-test"}}
	item := map[string]any{"__type": "CalendarItem:#Exchange", "IsAllDayEvent": false, "IsOnlineMeeting": true, "OnlineMeetingProvider": "Unknown", "OnlineMeeting": map[string]any{}, "IsOrganizer": true, "IsResponseRequested": true, "StartTimeZoneId": "Mountain Standard Time", "EndTimeZoneId": "Mountain Standard Time", "ItemId": map[string]any{"Id": "fixture-event"}, "id": "fixture-event", "ClientSeriesId": "fixture-event", "ParentFolderId": map[string]any{"Id": "fixture-folder"}, "RequiredAttendees": []any{attendee}, "Resources": []any{}, "Attachments": []any{}, "Body": map[string]any{"BodyType": "HTML", "Value": ""}, "Locations": []any{}, "ExtendedProperty": []any{}}
	create := baseCalendar("CreateItemJsonRequest:#Exchange", map[string]any{"SendMeetingInvitations": "SendToNone", "SavedItemFolderId": map[string]any{"BaseFolderId": map[string]any{"Id": "fixture-folder"}}, "Items": []any{item}})
	d, r := captureCalendar(t, "CreateCalendarEvent", create, false)
	p := cal.Params{CalendarID: "folder", Subject: "Self", Start: "2026-11-06T04:30:00-07:00", End: "2026-11-06T05:00:00-07:00", SelfMeeting: true}
	_, op, err := cal.Prepare("CreateCalendarEvent", d, r, self, p)
	if err != nil || op.Class != transport.Write {
		t.Fatalf("self event create rejected: %v", err)
	}
	if strings.Contains(string(op.Body), "fixture-") || !strings.Contains(string(op.Body), "Synthetic Self") {
		t.Fatal("fixture identity was not rebound")
	}
	item["RequiredAttendees"] = []any{attendee, attendee}
	d, r = captureCalendar(t, "CreateCalendarEvent", create, false)
	if _, _, err = cal.Prepare("CreateCalendarEvent", d, r, self, p); err == nil {
		t.Fatal("multiple recipients accepted")
	}
	item["RequiredAttendees"] = []any{attendee}
	attendee["Mailbox"].(map[string]any)["EmailAddress"] = "other@example.invalid"
	d, r = captureCalendar(t, "CreateCalendarEvent", create, false)
	if _, _, err = cal.Prepare("CreateCalendarEvent", d, r, self, p); err == nil {
		t.Fatal("different attendee accepted")
	}

	raw := `{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Success","ResponseCode":"NoError","Items":[{"__type":"CalendarItem:#Exchange","ItemId":{"Id":"event","ChangeKey":"v1"},"StartTimeZoneId":"Mountain Standard Time","EndTimeZoneId":"Mountain Standard Time","IsAllDayEvent":false,"IsMeeting":false,"IsRecurring":false,"IsOnlineMeeting":true,"IsOrganizer":true,"MeetingRequestWasSent":true,"CalendarItemType":"Single","ResponseType":"Organizer","Organizer":{"Mailbox":{"EmailAddress":"self@example.invalid"}},"RequiredAttendees":[{"Mailbox":{"EmailAddress":"self@example.invalid"}}],"OptionalAttendees":[],"Resources":[]}]}]}}}`
	e, state := cal.ParseEvent([]byte(raw))
	if state != transport.EnvelopeOK {
		t.Fatal("synthetic self event rejected")
	}
	update := baseCalendar("UpdateCalendarEventJsonRequest:#Exchange", map[string]any{"EventId": map[string]any{"Id": "fixture-event"}, "ItemChange": map[string]any{"ItemId": map[string]any{"Id": "fixture-event"}, "Updates": []any{map[string]any{"__type": "SetItemField:#Exchange", "Path": map[string]any{"FieldURI": "Subject"}, "Item": map[string]any{"__type": "CalendarItem:#Exchange", "Subject": "fixture-subject"}}}}, "EventScope": 0, "ShouldSendUpdateToAttendees": false})
	d, r = captureCalendar(t, "UpdateCalendarEvent", update, true)
	_, _, err = cal.Prepare("UpdateCalendarEvent", d, r, self, cal.Params{ID: e.ID, Current: &e, SelfMeeting: true, Subject: "Renamed"})
	if err != nil {
		t.Fatal(err)
	}
	foreign, _ := cal.ParseEvent([]byte(strings.Replace(raw, `"RequiredAttendees":[{"Mailbox":{"EmailAddress":"self@example.invalid"}}]`, `"RequiredAttendees":[{"Mailbox":{"EmailAddress":"other@example.invalid"}}]`, 1)))
	if _, _, err = cal.Prepare("UpdateCalendarEvent", d, r, self, cal.Params{ID: foreign.ID, Current: &foreign, SelfMeeting: true, Subject: "Renamed"}); err == nil {
		t.Fatal("non-self attendee update accepted")
	}
	del := baseCalendar("DeleteCalendarEventJsonRequest:#Exchange", map[string]any{"EventId": map[string]any{"Id": "fixture-event"}, "EventScope": 0})
	d, r = captureCalendar(t, "DeleteCalendarEvent", del, true)
	if _, _, err = cal.Prepare("DeleteCalendarEvent", d, r, self, cal.Params{ID: e.ID, Current: &e, SelfMeeting: true}); err != nil {
		t.Fatal(err)
	}
	if _, _, err = cal.Prepare("DeleteCalendarEvent", d, r, self, cal.Params{ID: foreign.ID, Current: &foreign, SelfMeeting: true}); err == nil {
		t.Fatal("non-self attendee delete accepted")
	}
}
