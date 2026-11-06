// Package calendar implements captured personal-appointment Outlook RPCs,
// including the observed weekly recurrence and occurrence-subject shapes.
package calendar

import (
	"crypto/rand"
	"encoding/json"
	"fmt"
	"html"
	"strings"
	"time"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/webcapture"
)

type Params struct {
	ID, CalendarID, Subject, Start, End              string
	StartDate, EndDate                               string
	TimeZone                                         string
	WeeklyUntil                                      string
	BodyText, AddLocation                            *string
	ReplaceLocation, WithLocation, RemoveLocation    *string
	Current                                          *Event
	Occurrence, Series, ResetExceptions, SelfMeeting bool
}

func obj(v any) map[string]any { x, _ := v.(map[string]any); return x }
func validID(s string) bool    { return s != "" && len(s) <= 4096 && !strings.ContainsAny(s, "\x00\r\n") }
func arrayEmpty(v any) bool {
	if v == nil {
		return true
	}
	a, ok := v.([]any)
	return ok && len(a) == 0
}
func bad() error {
	return fmt.Errorf("calendar request does not match the captured personal-appointment contract")
}
func localTime(s string) (string, error) {
	return localTimeInZone(s, "Mountain Standard Time")
}

// AllDayRange converts inclusive calendar dates to the observed local-midnight
// request range, whose end is exclusive.
func AllDayRange(startDate, endDate string) (string, string, error) {
	start, e1 := time.Parse("2006-01-02", startDate)
	end, e2 := time.Parse("2006-01-02", endDate)
	if e1 != nil || e2 != nil || start.Year() < 1 || end.Year() < 1 || end.Before(start) || start.Format("2006-01-02") != startDate || end.Format("2006-01-02") != endDate {
		return "", "", fmt.Errorf("all-day dates must be valid YYYY-MM-DD values with the inclusive end on or after the start")
	}
	exclusive := end.AddDate(0, 0, 1)
	if exclusive.Year() > 9999 {
		return "", "", fmt.Errorf("all-day end date is out of range")
	}
	return start.Format("2006-01-02") + "T00:00:00.000", exclusive.Format("2006-01-02") + "T00:00:00.000", nil
}
func objectID(id string) map[string]any {
	return map[string]any{"__type": "ItemId:#Exchange", "Id": id}
}

func bodyHTML(text string) string {
	return "<div>" + strings.ReplaceAll(html.EscapeString(text), "\n", "<br>") + "</div>"
}
func typedLocation(name string) map[string]any {
	return map[string]any{"DisplayName": name, "Id": name, "PostalAddress": map[string]any{"LocationSource": "None", "Type": "Business"}}
}
func locationSourceURI() map[string]any {
	return map[string]any{"__type": "ExtendedPropertyUri:#Exchange", "PropertySetId": "00020329-0000-0000-C000-000000000046", "PropertyName": "LocationSelectionSourceMap", "PropertyType": "String"}
}
func locationProperty(sources map[string]string) (map[string]any, error) {
	encoded, err := json.Marshal(sources)
	if err != nil {
		return nil, err
	}
	return map[string]any{"__type": "ExtendedPropertyType:#Exchange", "ExtendedFieldURI": locationSourceURI(), "Value": string(encoded)}, nil
}

func validLocationName(name string) bool {
	return strings.TrimSpace(name) != "" && len(name) <= 1024 && !strings.ContainsAny(name, "\r\n\x00")
}

// locationEdit derives a complete new value for the observed Locations field
// and preserves unrelated entries and source metadata. Replacement/removal
// use the observed SetItemField array shape and have passed live readback.
func locationEdit(e *Event, id string, p Params) ([]any, map[string]string, error) {
	if e == nil || e.ID != id || !e.SafePersonal() {
		return nil, nil, bad()
	}
	ops := 0
	for _, v := range []*string{p.AddLocation, p.ReplaceLocation, p.RemoveLocation} {
		if v != nil {
			ops++
			if !validLocationName(*v) {
				return nil, nil, bad()
			}
		}
	}
	if ops != 1 || (p.ReplaceLocation != nil) != (p.WithLocation != nil) || p.WithLocation != nil && !validLocationName(*p.WithLocation) {
		return nil, nil, bad()
	}
	var item map[string]json.RawMessage
	if json.Unmarshal(e.Raw, &item) != nil {
		return nil, nil, bad()
	}
	var locations []any
	if len(item["Locations"]) == 0 || string(item["Locations"]) == "null" || json.Unmarshal(item["Locations"], &locations) != nil {
		return nil, nil, bad()
	}
	var properties []struct {
		Field struct {
			PropertyName  string `json:"PropertyName"`
			PropertySetID string `json:"PropertySetId"`
			Distinguished string `json:"DistinguishedPropertySetId"`
		} `json:"ExtendedFieldURI"`
		Value string `json:"Value"`
	}
	if v := item["ExtendedProperty"]; len(v) > 0 && string(v) != "null" && json.Unmarshal(v, &properties) != nil {
		return nil, nil, bad()
	}
	sources := map[string]string{}
	found := false
	for _, p := range properties {
		if p.Field.PropertyName != "LocationSelectionSourceMap" {
			continue
		}
		if p.Field.PropertySetID != "00020329-0000-0000-C000-000000000046" && p.Field.Distinguished != "PublicStrings" {
			return nil, nil, bad()
		}
		if found || json.Unmarshal([]byte(p.Value), &sources) != nil || sources == nil {
			return nil, nil, bad()
		}
		found = true
	}
	if len(locations) > 0 && !found {
		return nil, nil, fmt.Errorf("existing location source metadata is unavailable; no change made")
	}
	indices := map[string]int{}
	for i, v := range locations {
		loc := obj(v)
		display, ok := loc["DisplayName"].(string)
		if !ok || display == "" || sources[display] == "" {
			return nil, nil, bad()
		}
		if _, exists := indices[display]; exists {
			return nil, nil, fmt.Errorf("location names are ambiguous; no change made")
		}
		indices[display] = i
	}
	switch {
	case p.AddLocation != nil:
		name := *p.AddLocation
		if _, exists := indices[name]; exists {
			return nil, nil, fmt.Errorf("location is already present")
		}
		sources[name] = "ExplicitTyping"
		locations = append(locations, typedLocation(name))
	case p.ReplaceLocation != nil:
		old, newName := *p.ReplaceLocation, *p.WithLocation
		index, exists := indices[old]
		if !exists {
			return nil, nil, fmt.Errorf("location to replace was not found")
		}
		if _, exists = indices[newName]; exists {
			return nil, nil, fmt.Errorf("replacement location is already present")
		}
		delete(sources, old)
		sources[newName] = "ExplicitTyping"
		locations[index] = typedLocation(newName)
	case p.RemoveLocation != nil:
		old := *p.RemoveLocation
		index, exists := indices[old]
		if !exists {
			return nil, nil, fmt.Errorf("location to remove was not found")
		}
		delete(sources, old)
		locations = append(locations[:index], locations[index+1:]...)
	}
	return locations, sources, nil
}

// Prepare uses exact private UI request templates and replaces every fixture ID
// or user field used by the selected operation. It accepts no arbitrary RPC.
func Prepare(action, detail, request string, current transport.Session, p Params) (transport.Session, transport.Operation, error) {
	fail := func() (transport.Session, transport.Operation, error) {
		return transport.Session{}, transport.Operation{}, bad()
	}
	switch action {
	case "GetCalendarEvent", "CreateCalendarEvent", "UpdateCalendarEvent", "DeleteCalendarEvent":
	default:
		return fail()
	}
	d, err := webcapture.Load(detail, request, current)
	if err != nil {
		return transport.Session{}, transport.Operation{}, err
	}
	if d.Method != "POST" || d.URL.Path != "/owa/service.svc" || d.URL.Query().Get("action") != action || d.URL.Query().Get("app") != "Calendar" || d.Session.Headers.Get("Action") != action {
		return fail()
	}
	var root map[string]any
	if json.Unmarshal(d.Body, &root) != nil {
		return fail()
	}
	b := obj(root["Body"])
	h := obj(root["Header"])
	tz := obj(obj(h["TimeZoneContext"])["TimeZoneDefinition"])
	zoneID, err := CanonicalTimeZone(p.TimeZone)
	if err != nil {
		return fail()
	}
	if action == "UpdateCalendarEvent" && p.TimeZone == "" && (p.Start != "" || p.End != "") && p.Current != nil {
		if currentZone, ok := p.Current.SafeTimedZone(); ok {
			zoneID = currentZone
		}
	}
	if tz["Id"] != zoneID || b == nil {
		return fail()
	}
	var class transport.Class
	var validate func([]byte) transport.EnvelopeResult
	switch action {
	case "GetCalendarEvent":
		if root["__type"] != "GetCalendarEventJsonRequest:#Exchange" || !validID(p.ID) {
			return fail()
		}
		ids, ok := b["EventIds"].([]any)
		if !ok || len(ids) != 1 {
			return fail()
		}
		ids[0] = objectID(p.ID)
		class = transport.Read
		validate = validateEvent
	case "CreateCalendarEvent":
		if root["__type"] != "CreateItemJsonRequest:#Exchange" || !validID(p.CalendarID) || strings.TrimSpace(p.Subject) == "" || b["SendMeetingInvitations"] != "SendToNone" || p.ReplaceLocation != nil || p.WithLocation != nil || p.RemoveLocation != nil {
			return fail()
		}
		allDay := p.StartDate != "" || p.EndDate != ""
		if p.SelfMeeting && (allDay || zoneID != "Mountain Standard Time" || p.WeeklyUntil != "" || p.BodyText != nil || p.AddLocation != nil || p.Occurrence || p.Series) {
			return fail()
		}
		if zoneID != "Mountain Standard Time" && (allDay || p.WeeklyUntil != "") {
			return fail()
		}
		var a time.Time
		var start, end string
		if allDay {
			if p.Start != "" || p.End != "" || p.WeeklyUntil != "" {
				return fail()
			}
			start, end, err = AllDayRange(p.StartDate, p.EndDate)
			if err != nil {
				return transport.Session{}, transport.Operation{}, err
			}
		} else {
			var z time.Time
			var e1, e2 error
			a, e1 = time.Parse(time.RFC3339, p.Start)
			z, e2 = time.Parse(time.RFC3339, p.End)
			if e1 != nil || e2 != nil || !a.Before(z) {
				return fail()
			}
			start, e1 = localTimeInZone(p.Start, zoneID)
			end, e2 = localTimeInZone(p.End, zoneID)
			if e1 != nil {
				return transport.Session{}, transport.Operation{}, e1
			}
			if e2 != nil {
				return transport.Session{}, transport.Operation{}, e2
			}
		}
		items, ok := b["Items"].([]any)
		if !ok || len(items) != 1 {
			return fail()
		}
		item := obj(items[0])
		if item["__type"] != "CalendarItem:#Exchange" || item["IsAllDayEvent"] != allDay || allDay && (item["StartTimeZoneId"] != "Mountain Standard Time" || item["EndTimeZoneId"] != "Mountain Standard Time") {
			return fail()
		}
		for _, key := range []string{"StartTimeZoneId", "EndTimeZoneId"} {
			value, present := item[key]
			if present && value != zoneID || !present && zoneID != "Mountain Standard Time" {
				return fail()
			}
		}
		if p.SelfMeeting {
			online := obj(item["OnlineMeeting"])
			if item["IsOnlineMeeting"] != true || item["OnlineMeetingProvider"] != "Unknown" || item["IsOrganizer"] != true || item["IsResponseRequested"] != true || online == nil || len(online) != 0 {
				return fail()
			}
			if err := bindSelfMeetingAttendee(item, current, strings.HasPrefix(detail, "runtime:")); err != nil {
				return transport.Session{}, transport.Operation{}, err
			}
		} else if item["IsOnlineMeeting"] == true {
			return fail()
		}
		if p.WeeklyUntil == "" {
			if item["Recurrence"] != nil {
				return fail()
			}
		} else {
			if p.BodyText != nil || p.AddLocation != nil {
				return fail()
			}
			rec := obj(item["Recurrence"])
			pattern := obj(rec["RecurrencePattern"])
			dateRange := obj(rec["RecurrenceRange"])
			if rec["__type"] != "RecurrenceType:#Exchange" || pattern["__type"] != "WeeklyRecurrence:#Exchange" || pattern["Interval"] != float64(1) || pattern["FirstDayOfWeek"] != "Sunday" || dateRange["__type"] != "EndDateRecurrence:#Exchange" {
				return fail()
			}
			zone, err := time.LoadLocation("America/Denver")
			if err != nil {
				return fail()
			}
			localStart := a.In(zone)
			first := localStart.Format("2006-01-02")
			until, err := time.Parse("2006-01-02", p.WeeklyUntil)
			if err != nil || until.Format("2006-01-02") != p.WeeklyUntil || p.WeeklyUntil < first {
				return fail()
			}
			pattern["DaysOfWeek"] = localStart.Weekday().String()
			dateRange["StartDate"] = first
			dateRange["EndDate"] = p.WeeklyUntil
		}
		for _, k := range []string{"OptionalAttendees", "Resources", "Attachments"} {
			if !arrayEmpty(item[k]) {
				return fail()
			}
		}
		if !p.SelfMeeting && !arrayEmpty(item["RequiredAttendees"]) {
			return fail()
		}
		folder := obj(obj(b["SavedItemFolderId"])["BaseFolderId"])
		parent := obj(item["ParentFolderId"])
		if folder == nil || parent == nil {
			return fail()
		}
		folder["Id"] = p.CalendarID
		parent["Id"] = p.CalendarID
		var random [16]byte
		if _, err := rand.Read(random[:]); err != nil {
			return fail()
		}
		random[6] = (random[6] & 0x0f) | 0x40
		random[8] = (random[8] & 0x3f) | 0x80
		local := fmt.Sprintf("localevent:%x-%x-%x-%x-%x", random[:4], random[4:6], random[6:8], random[8:10], random[10:])
		item["ItemId"] = objectID(local)
		item["id"] = local
		item["ClientSeriesId"] = local
		item["Subject"] = p.Subject
		item["Start"] = start
		item["End"] = end
		// The original create fixture has no body or locations. Rewrite both
		// so no content from an operation template can become user data.
		if !arrayEmpty(item["Locations"]) || !arrayEmpty(item["ExtendedProperty"]) {
			return fail()
		}
		body := obj(item["Body"])
		if body == nil || body["BodyType"] != "HTML" {
			return fail()
		}
		body["Value"] = ""
		if p.BodyText != nil {
			if strings.TrimSpace(*p.BodyText) == "" {
				return fail()
			}
			body["Value"] = bodyHTML(*p.BodyText)
		}
		item["Locations"] = []any{}
		item["ExtendedProperty"] = []any{}
		if p.AddLocation != nil {
			name := *p.AddLocation
			if strings.TrimSpace(name) == "" || len(name) > 1024 || strings.ContainsAny(name, "\r\n\x00") {
				return fail()
			}
			property, err := locationProperty(map[string]string{name: "ExplicitTyping"})
			if err != nil {
				return fail()
			}
			item["Locations"] = []any{typedLocation(name)}
			item["ExtendedProperty"] = []any{property}
		}
		class = transport.Write
		validate = validateEvent
	case "UpdateCalendarEvent":
		if root["__type"] != "UpdateCalendarEventJsonRequest:#Exchange" || !validID(p.ID) || b["EventScope"] != float64(0) || b["ShouldSendUpdateToAttendees"] != false {
			return fail()
		}
		allDayEdit := p.StartDate != "" || p.EndDate != ""
		seriesTime := p.Start != "" || p.End != ""
		seriesUntil := p.WeeklyUntil != ""
		locationChange := p.AddLocation != nil || p.ReplaceLocation != nil || p.WithLocation != nil || p.RemoveLocation != nil
		combined := seriesTime && (p.BodyText != nil || locationChange)
		if p.SelfMeeting && (p.Series || p.Occurrence || p.Current == nil || p.Current.ID != p.ID || !p.Current.safeSelfOnlineEvent(current) || p.Subject == "" || seriesTime || seriesUntil || allDayEdit || p.BodyText != nil || locationChange || p.ResetExceptions) {
			return fail()
		}
		if combined && (p.Series || p.Occurrence || zoneID != "Mountain Standard Time" || p.Subject != "") {
			return fail()
		}
		if p.Series {
			changes := 0
			if p.Subject != "" {
				changes++
			}
			if seriesTime {
				changes++
			}
			if seriesUntil {
				changes++
			}
			if p.Occurrence || p.Current == nil || p.Current.ID != p.ID || !p.Current.SafeWeeklyMountain() || zoneID != "Mountain Standard Time" || allDayEdit || p.BodyText != nil || p.AddLocation != nil || p.ReplaceLocation != nil || p.WithLocation != nil || p.RemoveLocation != nil || changes != 1 || p.ResetExceptions != (seriesTime || seriesUntil) {
				return fail()
			}
		} else if seriesUntil || p.ResetExceptions {
			return fail()
		}
		if p.Occurrence && (p.Current == nil || p.Current.ID != p.ID || !p.Current.SafeOccurrence() || p.Subject == "" || p.Start != "" || p.End != "" || allDayEdit || p.BodyText != nil || p.AddLocation != nil || p.ReplaceLocation != nil || p.WithLocation != nil || p.RemoveLocation != nil) {
			return fail()
		}
		if allDayEdit && (zoneID != "Mountain Standard Time" || p.Current == nil || p.Current.ID != p.ID || !p.Current.SafeAllDayMountain() || p.Start != "" || p.End != "" || p.Subject != "" || p.BodyText != nil || p.AddLocation != nil || p.ReplaceLocation != nil || p.WithLocation != nil || p.RemoveLocation != nil) {
			return fail()
		}
		change := obj(b["ItemChange"])
		updates, ok := change["Updates"].([]any)
		if !ok || len(updates) == 0 {
			return fail()
		}
		var bodyField, locationsField, sourceField, recurrenceField map[string]any
		subjectObserved := false
		timeFields := map[string]map[string]any{}
		for _, v := range updates {
			field := obj(v)
			path := obj(field["Path"])
			item := obj(field["Item"])
			if field["__type"] != "SetItemField:#Exchange" {
				return fail()
			}
			switch {
			case path["FieldURI"] == "Subject" && item["__type"] == "CalendarItem:#Exchange":
				if subjectObserved {
					return fail()
				}
				subjectObserved = true
			case (path["FieldURI"] == "Start" || path["FieldURI"] == "End" || path["FieldURI"] == "StartTimeZoneId" || path["FieldURI"] == "EndTimeZoneId") && path["__type"] == "PropertyUri:#Exchange" && item["__type"] == "CalendarItem:#Exchange":
				name, _ := path["FieldURI"].(string)
				if timeFields[name] != nil {
					return fail()
				}
				timeFields[name] = field
			case path["FieldURI"] == "OnlineMeetingProvider" && item["__type"] == "CalendarItem:#Exchange" && item["OnlineMeetingProvider"] == nil: // Unrelated UI bookkeeping, omitted.
			case path["FieldURI"] == "Recurrence" && path["__type"] == "PropertyUri:#Exchange" && item["__type"] == "CalendarItem:#Exchange":
				if !p.Series || recurrenceField != nil {
					return fail()
				}
				recurrenceField = field
			case path["FieldURI"] == "Body" && item["__type"] == "CalendarItem:#Exchange":
				if bodyField != nil {
					return fail()
				}
				bodyField = field
			case path["FieldURI"] == "Locations" && item["__type"] == "CalendarItem:#Exchange":
				if locationsField != nil {
					return fail()
				}
				locationsField = field
			case path["PropertyName"] == "LocationSelectionSourceMap" && path["PropertySetId"] == "00020329-0000-0000-C000-000000000046" && item["__type"] == "Message:#Exchange":
				if sourceField != nil {
					return fail()
				}
				sourceField = field
			case path["PropertyName"] == "IsDefaultPreferences": // Captured UI bookkeeping; intentionally omitted.
			default:
				return fail()
			}
		}
		if !subjectObserved && (bodyField == nil || locationsField == nil || sourceField == nil) && len(timeFields) != 4 && !(allDayEdit && len(timeFields) == 3) && !(p.Series && seriesUntil && recurrenceField != nil) {
			return fail()
		}
		if recurrenceField != nil && !(p.Series && (seriesTime || seriesUntil)) {
			return fail()
		}
		selected := []any{}
		if p.Subject != "" {
			if strings.TrimSpace(p.Subject) == "" {
				return fail()
			}
			selected = append(selected, map[string]any{"__type": "SetItemField:#Exchange", "Path": map[string]any{"__type": "PropertyUri:#Exchange", "FieldURI": "Subject"}, "Item": map[string]any{"__type": "CalendarItem:#Exchange", "Subject": p.Subject}})
		}
		if p.BodyText != nil {
			if strings.TrimSpace(*p.BodyText) == "" || bodyField == nil {
				return fail()
			}
			body := obj(obj(bodyField["Item"])["Body"])
			if body["BodyType"] != "HTML" {
				return fail()
			}
			body["Value"] = bodyHTML(*p.BodyText)
			selected = append(selected, bodyField)
		}
		if combined && locationChange {
			if locationsField == nil || sourceField != nil {
				return fail()
			}
			locations, err := combinedLocations(p.Current, p.ID, p)
			if err != nil {
				return transport.Session{}, transport.Operation{}, err
			}
			obj(locationsField["Item"])["Locations"] = locations
		}
		if p.Start != "" || p.End != "" {
			if p.Start == "" || p.End == "" || p.Current == nil || p.Current.ID != p.ID {
				return fail()
			}
			if !p.Series {
				currentZone, ok := p.Current.SafeTimedZone()
				if !ok || currentZone != zoneID {
					return fail()
				}
			}
			startTime, e1 := time.Parse(time.RFC3339, p.Start)
			endTime, e2 := time.Parse(time.RFC3339, p.End)
			if e1 != nil || e2 != nil || !startTime.Before(endTime) {
				return fail()
			}
			start, e1 := localTimeInZone(p.Start, zoneID)
			end, e2 := localTimeInZone(p.End, zoneID)
			if e1 != nil {
				return transport.Session{}, transport.Operation{}, e1
			}
			if e2 != nil {
				return transport.Session{}, transport.Operation{}, e2
			}
			if p.Series {
				if recurrenceField == nil {
					return fail()
				}
				rec, err := seriesRecurrence(*p.Current, recurrenceField, "", p.Start, p.End)
				if err != nil {
					return transport.Session{}, transport.Operation{}, err
				}
				obj(recurrenceField["Item"])["Recurrence"] = rec
			}
			// Preserve captured ordering and omit unrelated UI bookkeeping.
			for _, field := range []struct{ name, value string }{{"End", end}, {"EndTimeZoneId", zoneID}, {"Start", start}, {"StartTimeZoneId", zoneID}} {
				if p.Series && field.name == "Start" {
					selected = append(selected, recurrenceField)
				}
				if combined && locationChange && field.name == "Start" {
					selected = append(selected, locationsField)
				}
				f := timeFields[field.name]
				if f == nil {
					return fail()
				}
				item := obj(f["Item"])
				if strings.HasSuffix(field.name, "TimeZoneId") && item[field.name] != zoneID {
					return fail()
				}
				item[field.name] = field.value
				selected = append(selected, f)
			}
		}
		if p.Series && seriesUntil {
			if recurrenceField == nil || len(timeFields) != 0 {
				return fail()
			}
			rec, err := seriesRecurrence(*p.Current, recurrenceField, p.WeeklyUntil, "", "")
			if err != nil {
				return transport.Session{}, transport.Operation{}, err
			}
			obj(recurrenceField["Item"])["Recurrence"] = rec
			selected = append(selected, recurrenceField)
		}
		if allDayEdit {
			start, end, err := AllDayRange(p.StartDate, p.EndDate)
			if err != nil {
				return transport.Session{}, transport.Operation{}, err
			}
			if len(timeFields) != 3 || timeFields["StartTimeZoneId"] != nil {
				return fail()
			}
			// This all-day edit captured End, EndTimeZoneId, Start, in that order.
			for _, field := range []struct{ name, value string }{{"End", end}, {"EndTimeZoneId", "Mountain Standard Time"}, {"Start", start}} {
				f := timeFields[field.name]
				if f == nil {
					return fail()
				}
				item := obj(f["Item"])
				if field.name == "EndTimeZoneId" && item[field.name] != "Mountain Standard Time" {
					return fail()
				}
				item[field.name] = field.value
				selected = append(selected, f)
			}
		}
		if locationChange && !combined {
			if locationsField == nil || sourceField == nil {
				return fail()
			}
			locations, sources, err := locationEdit(p.Current, p.ID, p)
			if err != nil {
				return transport.Session{}, transport.Operation{}, err
			}
			property, err := locationProperty(sources)
			if err != nil {
				return fail()
			}
			// Rebuild only the observed source-map property, discarding the
			// fixture's content and the unrelated IsDefaultPreferences update.
			sourceField = map[string]any{"__type": "SetItemField:#Exchange", "Path": locationSourceURI(), "Item": map[string]any{"__type": "Message:#Exchange", "ExtendedProperty": []any{property}}}
			obj(locationsField["Item"])["Locations"] = locations
			selected = append(selected, sourceField, locationsField)
		}
		if len(selected) == 0 {
			return fail()
		}
		change["Updates"] = selected
		b["EventId"] = objectID(p.ID)
		change["ItemId"] = objectID(p.ID)
		class = transport.Write
		validate = validateEvent
	case "DeleteCalendarEvent":
		if root["__type"] != "DeleteCalendarEventJsonRequest:#Exchange" || !validID(p.ID) || b["EventScope"] != float64(0) {
			return fail()
		}
		if p.SelfMeeting && (p.Series || p.Current == nil || p.Current.ID != p.ID || !p.Current.safeSelfOnlineEvent(current)) {
			return fail()
		}
		if p.Series && (p.Current == nil || p.Current.ID != p.ID || !p.Current.SafeWeeklySeries()) {
			return fail()
		}
		b["EventId"] = objectID(p.ID)
		class = transport.Write
		validate = validateDelete
	}
	data, err := json.Marshal(root)
	if err != nil {
		return fail()
	}
	return d.Session, d.Operation("POST", d.URL.Path, data, class, validate), nil
}

type Event struct {
	ID                    string            `json:"id"`
	ChangeKey             string            `json:"change_key"`
	Subject               string            `json:"subject"`
	Start                 string            `json:"start"`
	End                   string            `json:"end"`
	StartTimeZone         string            `json:"start_time_zone,omitempty"`
	EndTimeZone           string            `json:"end_time_zone,omitempty"`
	EventStart            string            `json:"event_start,omitempty"`
	EventEnd              string            `json:"event_end,omitempty"`
	StartDate             string            `json:"start_date,omitempty"`
	EndDate               string            `json:"end_date,omitempty"`
	IsMeeting             *bool             `json:"is_meeting"`
	IsRecurring           *bool             `json:"is_recurring"`
	IsOnlineMeeting       *bool             `json:"is_online_meeting"`
	IsOrganizer           *bool             `json:"is_organizer"`
	IsAllDayEvent         *bool             `json:"is_all_day"`
	MeetingRequestWasSent *bool             `json:"meeting_request_was_sent"`
	CalendarItemType      string            `json:"calendar_item_type"`
	RequiredAttendees     []json.RawMessage `json:"required_attendees"`
	OptionalAttendees     []json.RawMessage `json:"optional_attendees"`
	Resources             []json.RawMessage `json:"resources"`
	Raw                   json.RawMessage   `json:"raw"`
}

func ParseEvent(data []byte) (Event, transport.EnvelopeResult) {
	var root struct {
		Body struct {
			ResponseMessages struct {
				Items []struct {
					ResponseClass, ResponseCode string
					Items                       []json.RawMessage
				}
			}
		}
	}
	if json.Unmarshal(data, &root) != nil || len(root.Body.ResponseMessages.Items) != 1 {
		return Event{}, transport.EnvelopeError
	}
	r := root.Body.ResponseMessages.Items[0]
	if r.ResponseClass != "Success" || r.ResponseCode != "NoError" {
		if strings.Contains(strings.ToLower(r.ResponseCode), "change") || strings.Contains(strings.ToLower(r.ResponseCode), "conflict") {
			return Event{}, transport.EnvelopeConflict
		}
		return Event{}, transport.EnvelopeRejected
	}
	if len(r.Items) != 1 {
		return Event{}, transport.EnvelopeError
	}
	var x struct {
		Type   string `json:"__type"`
		ItemID struct {
			ID, ChangeKey string `json:",omitempty"`
		} `json:"ItemId"`
		Subject, Start, End, StartTimeZoneId, EndTimeZoneId                                        string
		IsMeeting, IsRecurring, IsOnlineMeeting, IsOrganizer, IsAllDayEvent, MeetingRequestWasSent *bool
		CalendarItemType                                                                           string
		RequiredAttendees, OptionalAttendees, Resources                                            []json.RawMessage
	}
	if json.Unmarshal(r.Items[0], &x) != nil || x.Type != "CalendarItem:#Exchange" || !validID(x.ItemID.ID) || x.ItemID.ChangeKey == "" {
		return Event{}, transport.EnvelopeError
	}
	event := Event{ID: x.ItemID.ID, ChangeKey: x.ItemID.ChangeKey, Subject: x.Subject, Start: x.Start, End: x.End, StartTimeZone: x.StartTimeZoneId, EndTimeZone: x.EndTimeZoneId, IsMeeting: x.IsMeeting, IsRecurring: x.IsRecurring, IsOnlineMeeting: x.IsOnlineMeeting, IsOrganizer: x.IsOrganizer, IsAllDayEvent: x.IsAllDayEvent, MeetingRequestWasSent: x.MeetingRequestWasSent, CalendarItemType: x.CalendarItemType, RequiredAttendees: x.RequiredAttendees, OptionalAttendees: x.OptionalAttendees, Resources: x.Resources, Raw: r.Items[0]}
	if x.IsAllDayEvent != nil && !*x.IsAllDayEvent {
		event.EventStart = eventLocalTime(x.Start, x.StartTimeZoneId)
		event.EventEnd = eventLocalTime(x.End, x.EndTimeZoneId)
	}
	if start, end, ok := event.allDayDates(); ok {
		event.StartDate, event.EndDate = start, end
	}
	return event, transport.EnvelopeOK
}
func validateEvent(b []byte) transport.EnvelopeResult { _, s := ParseEvent(b); return s }
func validateDelete(data []byte) transport.EnvelopeResult {
	var r struct {
		Body struct{ ResponseClass, ResponseCode string }
	}
	if json.Unmarshal(data, &r) != nil {
		return transport.EnvelopeError
	}
	if r.Body.ResponseClass == "Success" && r.Body.ResponseCode == "NoError" {
		return transport.EnvelopeOK
	}
	return transport.EnvelopeRejected
}
func (e Event) SafePersonal() bool {
	return e.IsMeeting != nil && !*e.IsMeeting && e.IsRecurring != nil && !*e.IsRecurring && e.IsOnlineMeeting != nil && !*e.IsOnlineMeeting && e.IsOrganizer != nil && *e.IsOrganizer && e.MeetingRequestWasSent != nil && !*e.MeetingRequestWasSent && e.CalendarItemType == "Single" && e.RequiredAttendees != nil && e.OptionalAttendees != nil && e.Resources != nil && len(e.RequiredAttendees) == 0 && len(e.OptionalAttendees) == 0 && len(e.Resources) == 0
}

func (e Event) SafeTimedMountain() bool {
	zone, ok := e.SafeTimedZone()
	return ok && zone == "Mountain Standard Time"
}

// SafeTimedZone returns the common observed zone for a single personal event.
func (e Event) SafeTimedZone() (string, bool) {
	if !e.SafePersonal() || e.IsAllDayEvent == nil || *e.IsAllDayEvent {
		return "", false
	}
	var item struct {
		StartZone string `json:"StartTimeZoneId"`
		EndZone   string `json:"EndTimeZoneId"`
	}
	if json.Unmarshal(e.Raw, &item) != nil || item.StartZone == "" || item.StartZone != item.EndZone {
		return "", false
	}
	zone, err := CanonicalTimeZone(item.StartZone)
	return zone, err == nil
}

func (e Event) allDayDates() (string, string, bool) {
	if !e.SafePersonal() || e.IsAllDayEvent == nil || !*e.IsAllDayEvent {
		return "", "", false
	}
	var item struct {
		StartZone string `json:"StartTimeZoneId"`
		EndZone   string `json:"EndTimeZoneId"`
	}
	if json.Unmarshal(e.Raw, &item) != nil || item.StartZone != "Mountain Standard Time" || item.EndZone != "Mountain Standard Time" {
		return "", "", false
	}
	zone, err := time.LoadLocation("America/Denver")
	if err != nil {
		return "", "", false
	}
	start, e1 := time.Parse(time.RFC3339, e.Start)
	end, e2 := time.Parse(time.RFC3339, e.End)
	if e1 != nil || e2 != nil || !start.Before(end) {
		return "", "", false
	}
	start, end = start.In(zone), end.In(zone)
	if start.Hour() != 0 || start.Minute() != 0 || start.Second() != 0 || start.Nanosecond() != 0 || end.Hour() != 0 || end.Minute() != 0 || end.Second() != 0 || end.Nanosecond() != 0 {
		return "", "", false
	}
	inclusive := end.AddDate(0, 0, -1)
	if inclusive.Before(start) {
		return "", "", false
	}
	return start.Format("2006-01-02"), inclusive.Format("2006-01-02"), true
}

func (e Event) SafeAllDayMountain() bool {
	_, _, ok := e.allDayDates()
	return ok
}

// SafeOccurrence is restricted to observed, attendee-free timed Mountain
// occurrences and exceptions with an explicit master reference. It does not
// authorize a series update or any meeting notification.
func (e Event) SafeOccurrence() bool {
	if e.IsMeeting == nil || *e.IsMeeting || e.IsRecurring == nil || !*e.IsRecurring || e.IsOnlineMeeting == nil || *e.IsOnlineMeeting || e.IsOrganizer == nil || !*e.IsOrganizer || e.IsAllDayEvent == nil || *e.IsAllDayEvent || e.MeetingRequestWasSent == nil || *e.MeetingRequestWasSent || (e.CalendarItemType != "Occurrence" && e.CalendarItemType != "Exception") || e.RequiredAttendees == nil || e.OptionalAttendees == nil || e.Resources == nil || len(e.RequiredAttendees) != 0 || len(e.OptionalAttendees) != 0 || len(e.Resources) != 0 {
		return false
	}
	var item struct {
		StartZone string `json:"StartTimeZoneId"`
		EndZone   string `json:"EndTimeZoneId"`
		Master    struct {
			ID string `json:"Id"`
		} `json:"SeriesMasterItemId"`
	}
	return json.Unmarshal(e.Raw, &item) == nil && item.StartZone == "Mountain Standard Time" && item.EndZone == "Mountain Standard Time" && validID(item.Master.ID)
}

// SafeWeeklySeries only authorizes deletion of the observed kind of personal
// series master, never a meeting or a single occurrence.
func (e Event) SafeWeeklySeries() bool {
	if e.IsMeeting == nil || *e.IsMeeting || e.IsRecurring == nil || !*e.IsRecurring || e.IsOnlineMeeting == nil || *e.IsOnlineMeeting || e.IsOrganizer == nil || !*e.IsOrganizer || e.MeetingRequestWasSent == nil || *e.MeetingRequestWasSent || e.IsAllDayEvent == nil || *e.IsAllDayEvent || e.CalendarItemType != "RecurringMaster" || e.RequiredAttendees == nil || e.OptionalAttendees == nil || e.Resources == nil || len(e.RequiredAttendees) != 0 || len(e.OptionalAttendees) != 0 || len(e.Resources) != 0 {
		return false
	}
	var item struct {
		Recurrence struct {
			Pattern struct {
				Type           string `json:"__type"`
				DaysOfWeek     string
				FirstDayOfWeek string
				Interval       int
			} `json:"RecurrencePattern"`
			Range struct {
				Type string `json:"__type"`
			} `json:"RecurrenceRange"`
		} `json:"Recurrence"`
	}
	if json.Unmarshal(e.Raw, &item) != nil {
		return false
	}
	p, r := item.Recurrence.Pattern, item.Recurrence.Range
	return p.Type == "WeeklyRecurrence:#Exchange" && p.Interval == 1 && p.DaysOfWeek != "" && p.FirstDayOfWeek == "Sunday" && r.Type == "EndDateRecurrence:#Exchange"
}
