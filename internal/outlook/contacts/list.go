package contacts

import (
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"net/url"
	"strconv"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/webcapture"
)

const personalFilter = "graphRelationships/any(t: t/relationshipType eq 'x')"
const peopleListPath = "/PeopleGraphVx/v1.0/people"

// PrepareList replays the observed initial Your Contacts query, or a validated
// server-provided next link. A delta link is never accepted as a list page.
func PrepareList(detail string, current transport.Session, pageSize int, cursor string) (transport.Session, transport.Operation, error) {
	d, err := webcapture.Load(detail, "", current)
	if err != nil {
		return transport.Session{}, transport.Operation{}, err
	}
	if d.Method != "GET" || d.URL.Path != peopleListPath || len(d.Body) != 0 || d.URL.Query().Get("$filter") != personalFilter || d.URL.Query().Has("$deltatoken") || d.URL.Query().Get("$top") != "30000" {
		return transport.Session{}, transport.Operation{}, bad()
	}
	query := d.URL.Query()
	if query.Get("$expand") == "" || query.Get("$select") == "" {
		return transport.Session{}, transport.Operation{}, bad()
	}
	if cursor == "" {
		if pageSize < 1 || pageSize > 1000 {
			return transport.Session{}, transport.Operation{}, fmt.Errorf("page size must be between 1 and 1000")
		}
		// The observed service treats $top as a hard cap and can return a delta
		// link without a next link even when more contacts exist. Keep the
		// UI's observed cap and paginate the returned snapshot locally.
		query.Set("$top", "30000")
	} else {
		query, err = validatedPeopleLink(cursor, query, false)
		if err != nil {
			return transport.Session{}, transport.Operation{}, err
		}
	}
	op := d.Operation("GET", peopleListPath, nil, transport.Read, validateList)
	op.Query = query
	return d.Session, op, nil
}
func validatedPeopleLink(raw string, expected url.Values, delta bool) (url.Values, error) {
	if len(raw) == 0 || len(raw) > 16000 {
		return nil, fmt.Errorf("invalid People continuation")
	}
	u, err := url.Parse(raw)
	if err != nil || u.Scheme != "https" || u.Host != "outlook.cloud.microsoft" || u.User != nil || u.Path != peopleListPath || u.Fragment != "" {
		return nil, fmt.Errorf("invalid People continuation")
	}
	q, err := url.ParseQuery(u.RawQuery)
	if err != nil {
		return nil, fmt.Errorf("invalid People continuation")
	}
	for _, key := range []string{"$filter", "$expand", "$select"} {
		if len(q[key]) != 1 || q.Get(key) != expected.Get(key) {
			return nil, fmt.Errorf("People continuation has a different scope")
		}
	}
	if !delta && q.Has("$deltatoken") {
		return nil, fmt.Errorf("a delta cursor cannot be used as a list continuation")
	}
	if len(q["$top"]) > 1 {
		return nil, fmt.Errorf("invalid People continuation")
	}
	if v := q.Get("$top"); v != "" {
		n, e := strconv.Atoi(v)
		if e != nil || n < 1 || n > 30000 {
			return nil, fmt.Errorf("invalid People continuation")
		}
	}
	return q, nil
}

type ListPage struct {
	Results    []json.RawMessage `json:"results"`
	NextCursor string            `json:"next_cursor,omitempty"`
	SyncCursor string            `json:"sync_cursor,omitempty"`
}

// ListWindow is a view over one re-fetched server page. A nil HasMore means
// the server cap was reached or is unknown, so completeness is unknown.
type ListWindow struct {
	Results          []json.RawMessage `json:"results"`
	NextCursor       string            `json:"next_cursor,omitempty"`
	NextOffset       *int              `json:"next_offset,omitempty"`
	Snapshot         string            `json:"snapshot"`
	NextSnapshot     string            `json:"next_snapshot,omitempty"`
	SyncCursor       string            `json:"sync_cursor,omitempty"`
	HasMore          *bool             `json:"has_more"`
	Complete         bool              `json:"complete"`
	ServerCapReached bool              `json:"server_cap_reached"`
	ServerLimit      int               `json:"server_limit"`
	FetchedCount     int               `json:"fetched_count"`
	Offset           int               `json:"offset"`
}

func listSnapshot(page ListPage, currentCursor string) (string, error) {
	h := sha256.New()
	fmt.Fprintf(h, "people-list-v1:%d:", len(currentCursor))
	_, _ = h.Write([]byte(currentCursor))
	for _, raw := range page.Results {
		var item struct {
			ID        string `json:"id"`
			ChangeKey string `json:"changeKey"`
		}
		if json.Unmarshal(raw, &item) != nil || item.ID == "" {
			return "", fmt.Errorf("invalid People snapshot")
		}
		fmt.Fprintf(h, "%d:%s:%d:%s:", len(item.ID), item.ID, len(item.ChangeKey), item.ChangeKey)
		if item.ChangeKey == "" {
			var data any
			if json.Unmarshal(raw, &data) != nil {
				return "", fmt.Errorf("invalid People snapshot")
			}
			normalized, err := json.Marshal(data)
			if err != nil {
				return "", fmt.Errorf("invalid People snapshot")
			}
			fmt.Fprintf(h, "%d:", len(normalized))
			_, _ = h.Write(normalized)
		}
	}
	return hex.EncodeToString(h.Sum(nil)), nil
}
func WindowList(page ListPage, offset, limit int, currentCursor, expectedSnapshot string, serverLimit int) (ListWindow, error) {
	if offset < 0 || offset > 1_000_000 || limit < 1 || limit > 1000 || serverLimit < 0 || serverLimit > 30000 {
		return ListWindow{}, fmt.Errorf("invalid list offset, limit, or server limit")
	}
	snapshot, err := listSnapshot(page, currentCursor)
	if err != nil {
		return ListWindow{}, err
	}
	if offset > 0 && expectedSnapshot == "" {
		return ListWindow{}, fmt.Errorf("--snapshot is required when using --offset")
	}
	if expectedSnapshot != "" && expectedSnapshot != snapshot {
		return ListWindow{}, fmt.Errorf("contacts changed between pages; restart the list")
	}
	n := len(page.Results)
	start := offset
	if start > n {
		start = n
	}
	end := start + limit
	if end > n {
		end = n
	}
	out := ListWindow{Results: append([]json.RawMessage{}, page.Results[start:end]...), FetchedCount: n, Offset: offset, Snapshot: snapshot, ServerLimit: serverLimit, ServerCapReached: serverLimit > 0 && n >= serverLimit && page.NextCursor == ""}
	switch {
	case end < n:
		v := true
		next := end
		out.HasMore = &v
		out.NextOffset = &next
		out.NextCursor = currentCursor
		out.NextSnapshot = snapshot
	case page.NextCursor != "":
		v := true
		next := 0
		out.HasMore = &v
		out.NextOffset = &next
		out.NextCursor = page.NextCursor
	case serverLimit == 0 || n >= serverLimit:
		out.ServerCapReached = true
	default:
		v := false
		out.HasMore = &v
		out.Complete = offset == 0 && currentCursor == ""
		if out.Complete {
			out.SyncCursor = page.SyncCursor
		}
	}
	return out, nil
}

func ParseList(data []byte) (ListPage, transport.EnvelopeResult) {
	var root struct {
		Value json.RawMessage `json:"value"`
		Next  string          `json:"@odata.nextLink"`
		Delta string          `json:"@odata.deltaLink"`
	}
	if json.Unmarshal(data, &root) != nil || len(root.Value) == 0 || string(root.Value) == "null" || (root.Next == "") == (root.Delta == "") {
		return ListPage{}, transport.EnvelopeError
	}
	var values []json.RawMessage
	if json.Unmarshal(root.Value, &values) != nil {
		return ListPage{}, transport.EnvelopeError
	}
	for _, v := range values {
		var item map[string]json.RawMessage
		var id string
		if json.Unmarshal(v, &item) != nil || item == nil || json.Unmarshal(item["id"], &id) != nil || id == "" || len(item["reason"]) != 0 {
			return ListPage{}, transport.EnvelopeError
		}
	}
	if values == nil {
		values = []json.RawMessage{}
	}
	// Link validation is repeated against the captured request before callers
	// follow a cursor; these strings are explicit output, never followed here.
	return ListPage{Results: values, NextCursor: root.Next, SyncCursor: root.Delta}, transport.EnvelopeOK
}
func validateList(b []byte) transport.EnvelopeResult { _, s := ParseList(b); return s }
