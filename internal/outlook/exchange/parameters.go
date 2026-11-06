package exchange

import (
	"encoding/json"
	"errors"
	"net/url"
	"strings"
	"time"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

var ErrParameters = errors.New("unsupported or invalid Outlook query parameters")

type ReadParameters struct {
	FolderID       string
	ConversationID string
	CalendarID     string
	Offset         *int
	Limit          *int
	Start          string
	End            string
}

// ParameterizeRead modifies only fields observed in the operation's own
// captured request. It never accepts a free-form action, method, URL or body.
func ParameterizeRead(op transport.Operation, action string, p ReadParameters) (transport.Operation, error) {
	if op.Class != transport.Read || op.Query.Get("action") != action {
		return op, ErrParameters
	}
	for _, id := range []string{p.FolderID, p.ConversationID, p.CalendarID} {
		if len(id) > 4096 || strings.ContainsAny(id, "\r\n\x00") {
			return op, ErrParameters
		}
	}
	if p.Offset != nil && (*p.Offset < 0 || *p.Offset > 1_000_000) || p.Limit != nil && (*p.Limit < 1 || *p.Limit > 1000) {
		return op, ErrParameters
	}
	if action == "GetCalendarFoldersV2" || action == "GetAccountInformation" {
		if p != (ReadParameters{}) {
			return op, ErrParameters
		}
		return op, nil
	}
	header := op.Headers.Get("X-Owa-Urlpostdata")
	var encoded []byte
	var err error
	if header != "" {
		var s string
		s, err = url.QueryUnescape(header)
		encoded = []byte(s)
	} else {
		encoded = op.Body
	}
	var root map[string]any
	if err != nil || json.Unmarshal(encoded, &root) != nil {
		return op, ErrParameters
	}
	body, ok := root["Body"].(map[string]any)
	if !ok {
		return op, ErrParameters
	}
	setID := func(parent map[string]any, key, id string) bool {
		v, ok := parent[key].(map[string]any)
		if !ok || id == "" {
			return false
		}
		v["Id"] = id
		return true
	}
	switch action {
	case "FindConversation", "FindItem":
		if p.ConversationID != "" || p.CalendarID != "" || p.Start != "" || p.End != "" {
			return op, ErrParameters
		}
		if p.FolderID != "" {
			if action != "FindItem" {
				return op, ErrParameters
			} // FindConversation also carries a separate search-folder ID.
			folders, ok := body["ParentFolderIds"].([]any)
			if !ok || len(folders) != 1 {
				return op, ErrParameters
			}
			f, ok := folders[0].(map[string]any)
			if !ok {
				return op, ErrParameters
			}
			f["Id"] = p.FolderID
		}
		if p.Offset != nil || p.Limit != nil {
			paging, ok := body["Paging"].(map[string]any)
			if !ok || paging["__type"] != "IndexedPageView:#Exchange" {
				return op, ErrParameters
			}
			if p.Offset != nil {
				paging["Offset"] = *p.Offset
			}
			if p.Limit != nil {
				paging["MaxEntriesReturned"] = *p.Limit
			}
		}
	case "GetConversationItems":
		if p.FolderID != "" || p.CalendarID != "" || p.Offset != nil || p.Start != "" || p.End != "" {
			return op, ErrParameters
		}
		if p.ConversationID != "" {
			convs, ok := body["Conversations"].([]any)
			if !ok || len(convs) != 1 {
				return op, ErrParameters
			}
			c, ok := convs[0].(map[string]any)
			if !ok || c["SyncState"] != "" || !setID(c, "ConversationId", p.ConversationID) {
				return op, ErrParameters
			}
		}
		if p.Limit != nil {
			body["MaxItemsToReturn"] = *p.Limit
		}
	case "GetCalendarView":
		if p.FolderID != "" || p.ConversationID != "" || p.Offset != nil || p.Limit != nil {
			return op, ErrParameters
		}
		if p.CalendarID != "" {
			c, ok := body["CalendarId"].(map[string]any)
			if !ok || !setID(c, "BaseFolderId", p.CalendarID) {
				return op, ErrParameters
			}
		}
		if p.Start != "" || p.End != "" {
			start, e1 := time.Parse(time.RFC3339, p.Start)
			end, e2 := time.Parse(time.RFC3339, p.End)
			if e1 != nil || e2 != nil || !start.Before(end) {
				return op, ErrParameters
			}
			h, ok := root["Header"].(map[string]any)
			if !ok {
				return op, ErrParameters
			}
			tc, ok := h["TimeZoneContext"].(map[string]any)
			if !ok {
				return op, ErrParameters
			}
			def, ok := tc["TimeZoneDefinition"].(map[string]any)
			if !ok {
				return op, ErrParameters
			}
			// CLDR maps the observed Windows zone to this canonical IANA zone.
			if def["Id"] != "Mountain Standard Time" {
				return op, ErrParameters
			}
			zone, e := time.LoadLocation("America/Denver")
			if e != nil {
				return op, ErrParameters
			}
			format := func(t time.Time) (string, bool) {
				const layout = "2006-01-02T15:04:05.000"
				v := t.In(zone).Format(layout)
				if t.Add(time.Hour).In(zone).Format(layout) == v || t.Add(-time.Hour).In(zone).Format(layout) == v {
					return "", false
				}
				return v, true
			}
			from, ok1 := format(start)
			to, ok2 := format(end)
			if !ok1 || !ok2 {
				return op, ErrParameters
			}
			body["RangeStart"], body["RangeEnd"] = from, to
		}
	default:
		return op, ErrParameters
	}
	updated, err := json.Marshal(root)
	if err != nil {
		return op, ErrParameters
	}
	if header != "" {
		op.Headers = op.Headers.Clone()
		op.Headers.Set("X-Owa-Urlpostdata", strings.ReplaceAll(url.QueryEscape(string(updated)), "+", "%20"))
	} else {
		op.Body = updated
	}
	return op, nil
}
