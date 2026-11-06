package contacts

import (
	"encoding/json"
	"testing"
)

func TestFilterPersonalContactsAndCompleteness(t *testing.T) {
	page := ListPage{Results: []json.RawMessage{
		json.RawMessage(`{"id":"one","changeKey":"a","name":{"displayName":"Ada Lovelace"},"emails":[{"address":"ada@example.invalid"}],"phones":[{"number":"+12025550100"}],"positions":[{"detail":{"jobTitle":"Analyst","company":{"displayName":"Analytical Engines"}}}]}`),
		json.RawMessage(`{"id":"two","changeKey":"b","names":[{"first":"Grace","last":"Hopper"}],"emails":[{"address":"grace@example.invalid"}]}`),
		json.RawMessage(`{"id":"three","changeKey":"c","name":{"displayName":"Another Ada"}}`),
	}}
	for _, tc := range []struct {
		query string
		count int
	}{{"ADA", 2}, {"grace hopper", 1}, {"EXAMPLE.INVALID", 2}, {"5550100", 1}, {"engines", 1}, {"analyst", 1}, {"missing", 0}} {
		got, err := FilterPersonalContacts(page, tc.query)
		if err != nil || len(got.Results) != tc.count {
			t.Fatalf("query %q yielded %d, %v", tc.query, len(got.Results), err)
		}
	}
	first, err := PersonalSearchWindow(page, "ada", 0, 1, "", "", 30000)
	if err != nil || len(first.Results) != 1 || first.HasMore == nil || !*first.HasMore || first.NextOffset == nil || *first.NextOffset != 1 || first.NextSnapshot == "" || first.Complete {
		t.Fatal("first search page incorrect")
	}
	second, err := PersonalSearchWindow(page, "ada", 1, 1, "", first.NextSnapshot, 30000)
	if err != nil || len(second.Results) != 1 || second.HasMore == nil || *second.HasMore || second.Complete {
		t.Fatal("second search page incorrect")
	}
	if _, err = PersonalSearchWindow(page, "grace", 1, 1, "", first.NextSnapshot, 30000); err == nil {
		t.Fatal("snapshot from another query accepted")
	}
	capped, err := PersonalSearchWindow(page, "grace", 0, 5, "", "", 3)
	if err != nil || capped.HasMore != nil || capped.Complete || !capped.ServerCapReached {
		t.Fatal("capped source presented as complete")
	}
	nextPage := page
	nextPage.NextCursor = "server-next"
	next, err := PersonalSearchWindow(nextPage, "missing", 0, 5, "", "", 30000)
	if err != nil || next.HasMore == nil || !*next.HasMore || next.NextCursor != "server-next" {
		t.Fatal("empty matching page lost server continuation")
	}
	empty, err := PersonalSearchWindow(page, "missing", 0, 5, "", "", 30000)
	if err != nil || len(empty.Results) != 0 || empty.Results == nil || !empty.Complete || empty.HasMore == nil || *empty.HasMore {
		t.Fatal("empty search result incorrect")
	}
}
