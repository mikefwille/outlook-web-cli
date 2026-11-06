package webcapture_test

import (
	"encoding/json"
	"net/url"
	people "github.com/mikefwille/outlook-web-cli/internal/outlook/contacts"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"testing"
)

func TestObservedPeopleList(t *testing.T) {
	filter := "graphRelationships/any(t: t/relationshipType eq 'x')"
	q := url.Values{"$filter": []string{filter}, "$expand": []string{"names,emails,personSources"}, "$select": []string{"id,changeKey"}, "$top": []string{"30000"}, "n": []string{"1"}}
	d, r := captureFile(t, "GET", "/PeopleGraphVx/v1.0/people?"+q.Encode(), "", nil, false)
	if r != "" {
		t.Fatal("GET fixture unexpectedly has body")
	}
	session, op, err := people.PrepareList(d, accountSession(), 25, "")
	if err != nil {
		t.Fatal(err)
	}
	if op.Method != "GET" || op.Class != transport.Read || op.Query.Get("$top") != "30000" || session.Headers.Get("Cookie") != "fresh=synthetic" {
		t.Fatal("initial page request incorrect")
	}
	q.Set("$top", "25")
	q.Set("$skiptoken", "page2")
	next := "https://outlook.cloud.microsoft/PeopleGraphVx/v1.0/people?" + q.Encode()
	_, op, err = people.PrepareList(d, accountSession(), 25, next)
	if err != nil || op.Query.Get("$skiptoken") != "page2" {
		t.Fatal("valid next link rejected")
	}
	badFilter := q.Encode()
	q.Set("$filter", "different")
	wrongScope := "https://outlook.cloud.microsoft/PeopleGraphVx/v1.0/people?" + q.Encode()
	q.Set("$filter", filter)
	q.Set("$deltatoken", "checkpoint")
	delta := "https://outlook.cloud.microsoft/PeopleGraphVx/v1.0/people?" + q.Encode()
	for _, bad := range []string{"https://evil.invalid/PeopleGraphVx/v1.0/people?" + badFilter, wrongScope, delta} {
		if _, _, err = people.PrepareList(d, accountSession(), 25, bad); err == nil {
			t.Fatal("unsafe or delta continuation accepted")
		}
	}
	for _, v := range []string{`{"value":[],"@odata.deltaLink":"https://example.invalid/checkpoint"}`, `{"value":[{"id":"synthetic"}],"@odata.nextLink":"https://example.invalid/next"}`} {
		page, state := people.ParseList([]byte(v))
		if state != transport.EnvelopeOK || page.Results == nil {
			t.Fatal("valid page rejected")
		}
	}
	for _, v := range []string{`{"value":[]}`, `{"value":null,"@odata.deltaLink":"x"}`, `{"value":[{"id":"gone","reason":"deleted"}],"@odata.deltaLink":"x"}`, `{"value":[],"@odata.nextLink":"x","@odata.deltaLink":"y"}`} {
		if _, state := people.ParseList([]byte(v)); state == transport.EnvelopeOK {
			t.Fatal("invalid or incremental page accepted")
		}
	}
}

func TestPeopleListLocalWindowsAndSnapshotGuard(t *testing.T) {
	page := people.ListPage{Results: []json.RawMessage{json.RawMessage(`{"id":"one","changeKey":"v1"}`), json.RawMessage(`{"id":"two","changeKey":"v2"}`)}, SyncCursor: "sync-checkpoint"}
	first, err := people.WindowList(page, 0, 1, "", "", 30000)
	if err != nil || len(first.Results) != 1 || first.HasMore == nil || !*first.HasMore || first.NextOffset == nil || *first.NextOffset != 1 || first.NextSnapshot == "" || first.SyncCursor != "" || first.Complete {
		t.Fatal("first local page incorrectly reported")
	}
	second, err := people.WindowList(page, 1, 1, "", first.NextSnapshot, 30000)
	if err != nil || len(second.Results) != 1 || second.HasMore == nil || *second.HasMore || second.Complete {
		t.Fatal("second local page incorrectly reported")
	}
	if _, err = people.WindowList(page, 1, 1, "", "", 30000); err == nil {
		t.Fatal("offset without snapshot accepted")
	}
	changed := people.ListPage{Results: []json.RawMessage{page.Results[1], page.Results[0]}}
	if _, err = people.WindowList(changed, 1, 1, "", first.NextSnapshot, 30000); err == nil {
		t.Fatal("changed ordering accepted")
	}
	changed = people.ListPage{Results: []json.RawMessage{page.Results[0], json.RawMessage(`{"id":"two","changeKey":"new"}`)}}
	if _, err = people.WindowList(changed, 1, 1, "", first.NextSnapshot, 30000); err == nil {
		t.Fatal("changed item accepted")
	}
	full, err := people.WindowList(page, 0, 10, "", "", 30000)
	if err != nil || !full.Complete || full.HasMore == nil || *full.HasMore || full.SyncCursor != "sync-checkpoint" {
		t.Fatal("complete initial snapshot incorrectly reported")
	}
	capped, err := people.WindowList(page, 0, 10, "", "", 2)
	if err != nil || capped.Complete || !capped.ServerCapReached || capped.HasMore != nil || capped.SyncCursor != "" {
		t.Fatal("server cap was treated as complete")
	}
	unknown, err := people.WindowList(page, 0, 10, "opaque-next", "", 0)
	if err != nil || unknown.HasMore != nil || unknown.Complete || unknown.ServerLimit != 0 {
		t.Fatal("unknown server cap was treated as complete")
	}
	page.NextCursor = "server-next"
	next, err := people.WindowList(page, 0, 10, "", "", 2)
	if err != nil || next.HasMore == nil || !*next.HasMore || next.NextCursor != "server-next" || next.NextOffset == nil || *next.NextOffset != 0 || next.NextSnapshot != "" {
		t.Fatal("server continuation incorrectly reported")
	}
}
