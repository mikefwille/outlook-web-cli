package calendar

import (
	"encoding/json"
	"fmt"
	"strings"
	"time"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/webcapture"
)

type Occurrences struct {
	Master        json.RawMessage   `json:"master"`
	Occurrences   []json.RawMessage `json:"occurrences"`
	Regular       []json.RawMessage `json:"regular"`
	Exceptions    []json.RawMessage `json:"exceptions"`
	Cancellations []json.RawMessage `json:"cancellations"`
}

func PrepareOccurrences(detail string, current transport.Session, id, from, to string) (transport.Session, transport.Operation, error) {
	fail := func() (transport.Session, transport.Operation, error) {
		return transport.Session{}, transport.Operation{}, bad()
	}
	if !validID(id) {
		return fail()
	}
	a, e1 := time.Parse(time.RFC3339, from)
	b, e2 := time.Parse(time.RFC3339, to)
	if e1 != nil || e2 != nil || !a.Before(b) {
		return transport.Session{}, transport.Operation{}, fmt.Errorf("occurrence range needs RFC3339 offsets and an end after the start")
	}
	start, e1 := localTime(from)
	end, e2 := localTime(to)
	if e1 != nil {
		return transport.Session{}, transport.Operation{}, e1
	}
	if e2 != nil {
		return transport.Session{}, transport.Operation{}, e2
	}
	d, err := webcapture.Load(detail, "", current)
	if err != nil {
		return transport.Session{}, transport.Operation{}, err
	}
	if d.Method != "POST" || d.URL.Path != "/owa/service.svc" || d.URL.Query().Get("action") != "ExpandCalendarEvent" || d.URL.Query().Get("app") != "Calendar" || d.Session.Headers.Get("Action") != "ExpandCalendarEvent" || !d.HeaderBody {
		return fail()
	}
	var root map[string]any
	if json.Unmarshal(d.Body, &root) != nil || root["__type"] != "ExpandCalendarEventJsonRequest:#Exchange" {
		return fail()
	}
	h := obj(root["Header"])
	tz := obj(obj(h["TimeZoneContext"])["TimeZoneDefinition"])
	body := obj(root["Body"])
	if tz["Id"] != "Mountain Standard Time" || body["__type"] != "ExpandCalendarEventRequest:#Exchange" {
		return fail()
	}
	for _, k := range []string{"ReturnMaster", "ReturnRegularOccurrences", "ReturnExceptions", "ReturnCancellations"} {
		if body[k] != true {
			return fail()
		}
	}
	body["EventId"] = objectID(id)
	body["WindowStart"] = start
	body["WindowEnd"] = end
	encoded, err := json.Marshal(root)
	if err != nil {
		return fail()
	}
	return d.Session, d.Operation("POST", d.URL.Path, encoded, transport.Read, validateOccurrences), nil
}
func ParseOccurrences(data []byte) (Occurrences, transport.EnvelopeResult) {
	var root struct {
		Body struct {
			ResponseClass, ResponseCode                                                 string
			RecurrenceMaster                                                            json.RawMessage
			Occurrences, RegularOccurences, ExceptionalOccurences, CancelledOccurrences json.RawMessage
		}
	}
	if json.Unmarshal(data, &root) != nil || root.Body.ResponseClass != "Success" || root.Body.ResponseCode != "NoError" {
		return Occurrences{}, transport.EnvelopeError
	}
	var master map[string]json.RawMessage
	if len(root.Body.RecurrenceMaster) == 0 || string(root.Body.RecurrenceMaster) == "null" || json.Unmarshal(root.Body.RecurrenceMaster, &master) != nil || master == nil {
		return Occurrences{}, transport.EnvelopeError
	}
	arrays := []json.RawMessage{root.Body.Occurrences, root.Body.RegularOccurences, root.Body.ExceptionalOccurences, root.Body.CancelledOccurrences}
	parsed := make([][]json.RawMessage, len(arrays))
	for i, v := range arrays {
		if len(v) == 0 || string(v) == "null" || json.Unmarshal(v, &parsed[i]) != nil {
			return Occurrences{}, transport.EnvelopeError
		}
		if parsed[i] == nil {
			parsed[i] = []json.RawMessage{}
		}
	}
	return Occurrences{Master: root.Body.RecurrenceMaster, Occurrences: parsed[0], Regular: parsed[1], Exceptions: parsed[2], Cancellations: parsed[3]}, transport.EnvelopeOK
}
func validateOccurrences(data []byte) transport.EnvelopeResult {
	_, s := ParseOccurrences(data)
	return s
}

// WeeklyDeleteWindow limits deletion checks to the observed weekly pattern,
// at most one year, with the first occurrence on the pattern's weekday.
func (e Event) WeeklyDeleteWindow() (string, string, int, error) {
	badWindow := func() (string, string, int, error) { return "", "", 0, fmt.Errorf("unsupported weekly range") }
	if !e.SafeWeeklySeries() {
		return badWindow()
	}
	var x struct {
		Recurrence struct {
			Pattern struct{ DaysOfWeek string }         `json:"RecurrencePattern"`
			Range   struct{ StartDate, EndDate string } `json:"RecurrenceRange"`
		} `json:"Recurrence"`
	}
	if json.Unmarshal(e.Raw, &x) != nil || len(x.Recurrence.Range.StartDate) < 10 || len(x.Recurrence.Range.EndDate) < 10 {
		return badWindow()
	}
	start, e1 := time.Parse("2006-01-02", x.Recurrence.Range.StartDate[:10])
	end, e2 := time.Parse("2006-01-02", x.Recurrence.Range.EndDate[:10])
	days := int(end.Sub(start).Hours() / 24)
	if e1 != nil || e2 != nil || end.Before(start) || days > 365 || x.Recurrence.Pattern.DaysOfWeek != start.Weekday().String() {
		return badWindow()
	}
	zone, err := time.LoadLocation("America/Denver")
	if err != nil {
		return badWindow()
	}
	first := time.Date(start.Year(), start.Month(), start.Day(), 0, 0, 0, 0, zone)
	last := time.Date(end.Year(), end.Month(), end.Day()+1, 0, 0, 0, 0, zone)
	return first.Format(time.RFC3339), last.Format(time.RFC3339), days/7 + 1, nil
}

// SafeWeeklyDeletion requires a complete expansion with no cancellations or
// meeting/online occurrences, including exceptions. It cannot make the
// separate server check and delete atomic.
func (v Occurrences) SafeWeeklyDeletion(masterID string, expected int) bool {
	var master struct {
		ItemID struct {
			ID string `json:"Id"`
		} `json:"ItemId"`
	}
	if !validID(masterID) || expected < 1 || json.Unmarshal(v.Master, &master) != nil || master.ItemID.ID != masterID || len(v.Cancellations) != 0 || len(v.Occurrences) != expected || len(v.Regular)+len(v.Exceptions) != expected {
		return false
	}
	decode := func(raw json.RawMessage) (string, string, bool) {
		var x struct {
			ItemID struct {
				ID string `json:"Id"`
			} `json:"ItemId"`
			Master struct {
				ID string `json:"Id"`
			} `json:"SeriesMasterItemId"`
			IsMeeting, IsRecurring, IsOnlineMeeting, IsOrganizer, IsAllDayEvent, MeetingRequestWasSent *bool
			CalendarItemType                                                                           string
			RequiredAttendees, OptionalAttendees, Resources                                            []json.RawMessage
		}
		if json.Unmarshal(raw, &x) != nil || !validID(x.ItemID.ID) || x.Master.ID != masterID || x.IsAllDayEvent == nil || *x.IsAllDayEvent {
			return "", "", false
		}
		e := Event{IsMeeting: x.IsMeeting, IsRecurring: x.IsRecurring, IsOnlineMeeting: x.IsOnlineMeeting, IsOrganizer: x.IsOrganizer, IsAllDayEvent: x.IsAllDayEvent, MeetingRequestWasSent: x.MeetingRequestWasSent, CalendarItemType: x.CalendarItemType, RequiredAttendees: x.RequiredAttendees, OptionalAttendees: x.OptionalAttendees, Resources: x.Resources, Raw: raw}
		return x.ItemID.ID, x.CalendarItemType, e.SafeOccurrence()
	}
	ids := map[string]string{}
	for _, raw := range v.Occurrences {
		id, kind, ok := decode(raw)
		if !ok || ids[id] != "" {
			return false
		}
		ids[id] = kind
	}
	seen := map[string]bool{}
	for _, group := range []struct {
		items []json.RawMessage
		kind  string
	}{{v.Regular, "Occurrence"}, {v.Exceptions, "Exception"}} {
		for _, raw := range group.items {
			id, kind, ok := decode(raw)
			if !ok || kind != group.kind || ids[id] != kind || seen[id] {
				return false
			}
			seen[id] = true
		}
	}
	return len(seen) == expected && strings.TrimSpace(masterID) != ""
}
