package calendar

import (
	"encoding/json"
	"strings"
	"testing"
	"time"
)

func TestSearchAgendaSubjectsBoundedAndIncomplete(t *testing.T) {
	data := []byte(`{"Body":{"ResponseClass":"Success","ResponseCode":"NoError","Items":[{"Subject":"Design REVIEW","ItemId":{"Id":"one"}},{"Subject":"review again","ItemId":{"Id":"two"}},{"Subject":"other"},{"Location":"review"},{"Subject":null}]}}`)
	items, matches, missing, truncated, err := SearchAgendaSubjects(data, "ReViEw", 1)
	if err != nil || len(items) != 1 || matches != 2 || missing != 2 || !truncated {
		t.Fatalf("unexpected bounded search: items=%d matches=%d missing=%d truncated=%v err=%v", len(items), matches, missing, truncated, err)
	}
	var first struct{ Subject string }
	if json.Unmarshal(items[0], &first) != nil || first.Subject != "Design REVIEW" {
		t.Fatal("search did not preserve the first matching raw event")
	}
	items, matches, missing, truncated, err = SearchAgendaSubjects(data, "not present", 200)
	if err != nil || items == nil || len(items) != 0 || matches != 0 || missing != 2 || truncated {
		t.Fatalf("unexpected empty result: %v %v %v %v %v", items, matches, missing, truncated, err)
	}
}

func TestSearchAgendaRejectsInvalidInputAndResponse(t *testing.T) {
	valid := `{"Body":{"ResponseClass":"Success","ResponseCode":"NoError","Items":[]}}`
	for _, tc := range []struct {
		data, query string
		limit       int
	}{
		{valid, " ", 1}, {valid, strings.Repeat("a", 257), 1}, {valid, "x", 0}, {valid, "x", 201},
		{`{"Body":{"ResponseClass":"Error","ResponseCode":"ErrorAccessDenied","Items":[]}}`, "x", 1},
		{`{"Body":{"ResponseClass":"Success","ResponseCode":"NoError"}}`, "x", 1},
		{`{"Body":{"ResponseClass":"Success","ResponseCode":"NoError","Items":[null]}}`, "x", 1},
	} {
		if _, _, _, _, err := SearchAgendaSubjects([]byte(tc.data), tc.query, tc.limit); err == nil {
			t.Fatalf("accepted invalid search input or envelope: %+v", tc)
		}
	}
}

func TestFindSelfTimesOnlyConsecutiveFreeSlots(t *testing.T) {
	from, to := "2026-11-06T09:00:00-07:00", "2026-11-06T11:00:00-07:00"
	items, truncated, unknown, err := FindSelfTimes(from, to, "00020120", 30, 20)
	if err != nil || truncated || !unknown || len(items) != 2 {
		t.Fatalf("unexpected suggestions: %+v truncated=%v unknown=%v err=%v", items, truncated, unknown, err)
	}
	for i, expected := range []string{"2026-11-06T16:00:00Z", "2026-11-06T16:15:00Z"} {
		if items[i].Start.Format(time.RFC3339) != expected || items[i].End.Sub(items[i].Start) != 30*time.Minute {
			t.Fatalf("unexpected suggestion %d: %+v", i, items[i])
		}
	}
	items, truncated, unknown, err = FindSelfTimes(from, to, "00020120", 30, 1)
	if err != nil || !truncated || !unknown || len(items) != 1 {
		t.Fatalf("limit not reported: %+v %v %v %v", items, truncated, unknown, err)
	}
	items, truncated, unknown, err = FindSelfTimes(from, to, "22222222", 30, 10)
	if err != nil || truncated || unknown || items == nil || len(items) != 0 {
		t.Fatalf("busy view produced suggestions: %+v %v %v %v", items, truncated, unknown, err)
	}
}

func TestFindSelfTimesValidation(t *testing.T) {
	for _, tc := range []struct {
		from, to, view  string
		duration, limit int
	}{
		{"2026-11-06T16:00:00Z", "2026-11-06T17:00:00Z", "0000", 20, 1},
		{"2026-11-06T16:00:00Z", "2026-11-06T17:00:00Z", "0000", 75, 1},
		{"2026-11-06T16:00:00Z", "2026-11-06T17:00:00Z", "000", 30, 1},
		{"2026-11-06T16:00:00Z", "2026-11-06T17:00:00Z", "00x0", 30, 1},
		{"2026-11-06T16:01:00Z", "2026-11-06T17:00:00Z", "0000", 30, 1},
		{"2026-11-06T16:00:00Z", "2026-11-06T17:00:00Z", "0000", 30, 201},
	} {
		if _, _, _, err := FindSelfTimes(tc.from, tc.to, tc.view, tc.duration, tc.limit); err == nil {
			t.Fatalf("accepted invalid availability input: %+v", tc)
		}
	}
}
