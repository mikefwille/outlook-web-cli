package calendar

import (
	"encoding/json"
	"fmt"
	"time"
)

// SafeWeeklyMountain limits series edits to the event zone in the captures.
func (e Event) SafeWeeklyMountain() bool {
	if !e.SafeWeeklySeries() {
		return false
	}
	var zones struct {
		Start string `json:"StartTimeZoneId"`
		End   string `json:"EndTimeZoneId"`
	}
	return json.Unmarshal(e.Raw, &zones) == nil && zones.Start == "Mountain Standard Time" && zones.End == "Mountain Standard Time"
}

// seriesRecurrence rebuilds the captured weekly recurrence from the current
// master instead of retaining the fixture's dates or weekday.
func seriesRecurrence(e Event, field map[string]any, until, start, end string) (map[string]any, error) {
	fail := func() (map[string]any, error) { return nil, fmt.Errorf("unsupported weekly series edit") }
	if !e.SafeWeeklyMountain() {
		return fail()
	}
	if _, _, _, err := e.WeeklyDeleteWindow(); err != nil {
		return fail()
	}
	var current struct {
		Recurrence struct {
			Pattern struct{ DaysOfWeek string }         `json:"RecurrencePattern"`
			Range   struct{ StartDate, EndDate string } `json:"RecurrenceRange"`
		} `json:"Recurrence"`
	}
	var zones struct {
		Start string `json:"StartTimeZoneId"`
		End   string `json:"EndTimeZoneId"`
	}
	if json.Unmarshal(e.Raw, &current) != nil || json.Unmarshal(e.Raw, &zones) != nil || zones.Start != "Mountain Standard Time" || zones.End != "Mountain Standard Time" || len(current.Recurrence.Range.StartDate) < 10 || len(current.Recurrence.Range.EndDate) < 10 {
		return fail()
	}
	first := current.Recurrence.Range.StartDate[:10]
	last := current.Recurrence.Range.EndDate[:10]
	firstDate, e1 := time.Parse("2006-01-02", first)
	if e1 != nil {
		return fail()
	}
	if until != "" {
		lastDate, err := time.Parse("2006-01-02", until)
		if err != nil || lastDate.Format("2006-01-02") != until || lastDate.Before(firstDate) || lastDate.Sub(firstDate).Hours()/24 > 365 {
			return fail()
		}
		last = until
	}
	if start != "" || end != "" {
		a, e1 := time.Parse(time.RFC3339, start)
		b, e2 := time.Parse(time.RFC3339, end)
		zone, e3 := time.LoadLocation("America/Denver")
		if e1 != nil || e2 != nil || e3 != nil || !a.Before(b) || a.In(zone).Format("2006-01-02") != first || b.In(zone).Format("2006-01-02") != first {
			return fail()
		}
	}
	rec := obj(obj(field["Item"])["Recurrence"])
	pattern := obj(rec["RecurrencePattern"])
	rangeItem := obj(rec["RecurrenceRange"])
	if rec["__type"] != "RecurrenceType:#Exchange" || pattern["__type"] != "WeeklyRecurrence:#Exchange" || pattern["Interval"] != float64(1) || pattern["FirstDayOfWeek"] != "Sunday" || rangeItem["__type"] != "EndDateRecurrence:#Exchange" {
		return fail()
	}
	return map[string]any{"__type": "RecurrenceType:#Exchange", "RecurrencePattern": map[string]any{"__type": "WeeklyRecurrence:#Exchange", "__typename": "WeeklyRecurrence", "DaysOfWeek": current.Recurrence.Pattern.DaysOfWeek, "FirstDayOfWeek": "Sunday", "Interval": 1}, "RecurrenceRange": map[string]any{"__type": "EndDateRecurrence:#Exchange", "__typename": "EndDateRecurrence", "StartDate": first, "EndDate": last}}, nil
}
