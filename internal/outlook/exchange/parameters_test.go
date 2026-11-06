package exchange

import (
	"encoding/json"
	"net/http"
	"net/url"
	"strings"
	"testing"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func headerOp(action, body string) transport.Operation {
	return transport.Operation{Class: transport.Read, Query: url.Values{"action": {action}}, Headers: http.Header{"X-Owa-Urlpostdata": {url.QueryEscape(body)}}}
}
func decoded(t *testing.T, op transport.Operation) map[string]any {
	t.Helper()
	v, err := url.QueryUnescape(op.Headers.Get("X-Owa-Urlpostdata"))
	if err != nil {
		t.Fatal(err)
	}
	var x map[string]any
	if json.Unmarshal([]byte(v), &x) != nil {
		t.Fatal("bad JSON")
	}
	return x
}
func intp(v int) *int { return &v }

func TestItemListParametersUseObservedFolderAndPagingFields(t *testing.T) {
	op := headerOp("FindItem", `{"Body":{"ParentFolderIds":[{"Id":"old"}],"Paging":{"__type":"IndexedPageView:#Exchange","Offset":0,"MaxEntriesReturned":100}}}`)
	got, err := ParameterizeRead(op, "FindItem", ReadParameters{FolderID: "new-id", Offset: intp(25), Limit: intp(10)})
	if err != nil {
		t.Fatal(err)
	}
	b := decoded(t, got)["Body"].(map[string]any)
	f := b["ParentFolderIds"].([]any)[0].(map[string]any)
	page := b["Paging"].(map[string]any)
	if f["Id"] != "new-id" || page["Offset"] != float64(25) || page["MaxEntriesReturned"] != float64(10) {
		t.Fatal("parameters lost")
	}
	if strings.Contains(op.Headers.Get("X-Owa-Urlpostdata"), "new-id") {
		t.Fatal("mutated original request")
	}
}

func TestConversationParametersRequireEmptySyncState(t *testing.T) {
	body := []byte(`{"Body":{"Conversations":[{"ConversationId":{"Id":"old"},"SyncState":""}],"MaxItemsToReturn":20}}`)
	op := transport.Operation{Class: transport.Read, Query: url.Values{"action": {"GetConversationItems"}}, Body: body}
	got, err := ParameterizeRead(op, "GetConversationItems", ReadParameters{ConversationID: "new", Limit: intp(50)})
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(got.Body), `"Id":"new"`) || !strings.Contains(string(got.Body), `"MaxItemsToReturn":50`) {
		t.Fatal("parameters lost")
	}
	op.Body = []byte(strings.Replace(string(body), `"SyncState":""`, `"SyncState":"old-state"`, 1))
	if _, err := ParameterizeRead(op, "GetConversationItems", ReadParameters{ConversationID: "new"}); err == nil {
		t.Fatal("accepted stale sync state")
	}
}

func TestCalendarOffsetsAndAmbiguousDST(t *testing.T) {
	op := headerOp("GetCalendarView", `{"Header":{"TimeZoneContext":{"TimeZoneDefinition":{"Id":"Mountain Standard Time"}}},"Body":{"CalendarId":{"BaseFolderId":{"Id":"old"}},"RangeStart":"","RangeEnd":""}}`)
	got, err := ParameterizeRead(op, "GetCalendarView", ReadParameters{CalendarID: "calendar", Start: "2026-03-08T08:30:00Z", End: "2026-03-08T10:30:00Z"})
	if err != nil {
		t.Fatal(err)
	}
	b := decoded(t, got)["Body"].(map[string]any)
	if b["RangeStart"] != "2026-03-08T01:30:00.000" || b["RangeEnd"] != "2026-03-08T04:30:00.000" {
		t.Fatalf("wrong DST conversion: %v %v", b["RangeStart"], b["RangeEnd"])
	}
	if _, err := ParameterizeRead(op, "GetCalendarView", ReadParameters{Start: "2026-11-01T07:30:00Z", End: "2026-11-01T10:30:00Z"}); err == nil {
		t.Fatal("accepted ambiguous repeated local time")
	}
}

func TestRejectsFreeformAndInvalidParameters(t *testing.T) {
	op := headerOp("FindConversation", `{"Body":{"Paging":{"__type":"IndexedPageView:#Exchange"}}}`)
	if _, err := ParameterizeRead(op, "FindConversation", ReadParameters{FolderID: "other"}); err == nil {
		t.Fatal("accepted unproven folder switch")
	}
	if _, err := ParameterizeRead(op, "FindConversation", ReadParameters{Limit: intp(0)}); err == nil {
		t.Fatal("accepted invalid limit")
	}
	if _, err := ParameterizeRead(op, "DeleteItem", ReadParameters{}); err == nil {
		t.Fatal("accepted arbitrary action")
	}
}
