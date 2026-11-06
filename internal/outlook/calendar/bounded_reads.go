package calendar

import (
	"encoding/json"
	"fmt"
	"strings"
	"time"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/exchange"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

// SearchAgendaSubjects searches only the subjects present in a single
// GetCalendarView response. It cannot establish that the server returned every
// event or that a missing subject could not have matched.
func SearchAgendaSubjects(data []byte, query string, limit int) ([]json.RawMessage, int, int, bool, error) {
	if strings.TrimSpace(query) == "" || len(query) > 256 || limit < 1 || limit > 200 {
		return nil, 0, 0, false, fmt.Errorf("invalid calendar subject query or result limit")
	}
	if exchange.ValidateRead("GetCalendarView", data) != transport.EnvelopeOK {
		return nil, 0, 0, false, fmt.Errorf("invalid calendar agenda response")
	}
	var root struct {
		Body struct {
			Items []json.RawMessage `json:"Items"`
		} `json:"Body"`
	}
	if json.Unmarshal(data, &root) != nil {
		return nil, 0, 0, false, fmt.Errorf("invalid calendar agenda response")
	}
	results := make([]json.RawMessage, 0)
	missing := 0
	matches := 0
	needle := strings.ToLower(query)
	for _, raw := range root.Body.Items {
		var item map[string]json.RawMessage
		if json.Unmarshal(raw, &item) != nil || item == nil {
			return nil, 0, 0, false, fmt.Errorf("invalid calendar agenda item")
		}
		var subject string
		value, exists := item["Subject"]
		if !exists || string(value) == "null" || json.Unmarshal(value, &subject) != nil {
			missing++
			continue
		}
		if strings.Contains(strings.ToLower(subject), needle) {
			matches++
			if len(results) < limit {
				results = append(results, raw)
			}
		}
	}
	return results, matches, missing, matches > limit, nil
}

type SuggestedTime struct {
	Start time.Time `json:"start"`
	End   time.Time `json:"end"`
}

// FindSelfTimes treats only self-view code 0 as free. Other codes are never
// suggested. Suggestions overlap and start on the observed 15-minute grid.
func FindSelfTimes(from, to, view string, durationMinutes, limit int) ([]SuggestedTime, bool, bool, error) {
	start, _, slots, err := AvailabilityRange(from, to)
	if err != nil {
		return nil, false, false, err
	}
	if durationMinutes < 15 || durationMinutes%15 != 0 || durationMinutes/15 > slots || limit < 1 || limit > 200 || len(view) != slots {
		return nil, false, false, fmt.Errorf("invalid duration, result limit, or availability view")
	}
	unknown := false
	for _, code := range view {
		if code < '0' || code > '9' {
			return nil, false, false, fmt.Errorf("invalid availability code")
		}
		if code != '0' && code != '2' {
			unknown = true
		}
	}
	results := make([]SuggestedTime, 0)
	needed := durationMinutes / 15
	freeRun := 0
	truncated := false
	for i, code := range view {
		if code != '0' {
			freeRun = 0
			continue
		}
		freeRun++
		if freeRun < needed {
			continue
		}
		if len(results) == limit {
			truncated = true
			break
		}
		end := start.Add(time.Duration(i+1) * 15 * time.Minute)
		results = append(results, SuggestedTime{Start: end.Add(-time.Duration(durationMinutes) * time.Minute), End: end})
	}
	return results, truncated, unknown, nil
}
