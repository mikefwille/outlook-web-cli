// Package conversation implements the observed whole-conversation mail actions.
package conversation

import (
	"encoding/json"
	"fmt"
	"strings"
	"time"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/webcapture"
)

type Params struct {
	ConversationID string
	FolderID       string
	DestinationID  string
	Through        string
	CompletedAt    string
}

func validID(s string) bool { return s != "" && len(s) <= 4096 && !strings.ContainsAny(s, "\x00\r\n") }
func timestamp(s string) (string, bool) {
	t, err := time.Parse(time.RFC3339Nano, s)
	if err != nil {
		return "", false
	}
	return t.UTC().Format("2006-01-02T15:04:05.000Z"), true
}
func object(v any) map[string]any { x, _ := v.(map[string]any); return x }
func folder(id string) map[string]any {
	return map[string]any{"__type": "TargetFolderId:#Exchange", "BaseFolderId": map[string]any{"__type": "FolderId:#Exchange", "Id": id}}
}

// Prepare accepts one fixed observed action and replaces the captured fixture
// identifiers and timestamps. A move is scoped to the supplied folder.
func Prepare(kind, detail string, current transport.Session, p Params) (transport.Session, transport.Operation, error) {
	fail := func() (transport.Session, transport.Operation, error) {
		return transport.Session{}, transport.Operation{}, fmt.Errorf("conversation request does not match the observed action contract")
	}
	want := map[string]string{"mark-read": "SetReadState", "mark-unread": "SetReadState", "flag": "Flag", "complete": "Flag", "move": "Move", "restore": "Restore"}[kind]
	if want == "" || !validID(p.ConversationID) || !validID(p.FolderID) {
		return fail()
	}
	d, err := webcapture.Load(detail, "", current)
	if err != nil {
		return fail()
	}
	if d.Method != "POST" || d.URL.Path != "/owa/service.svc" || d.URL.Query().Get("action") != "ApplyConversationAction" || d.URL.Query().Get("app") != "Mail" || d.Session.Headers.Get("Action") != "ApplyConversationAction" || !d.HeaderBody {
		return fail()
	}
	var root map[string]any
	if json.Unmarshal(d.Body, &root) != nil || root["__type"] != "ApplyConversationActionJsonRequest:#Exchange" {
		return fail()
	}
	b := object(root["Body"])
	if b == nil || b["__type"] != "ApplyConversationActionRequest:#Exchange" {
		return fail()
	}
	items, ok := b["ConversationActions"].([]any)
	if !ok || len(items) != 1 {
		return fail()
	}
	old := object(items[0])
	if old == nil || old["__type"] != "ConversationAction:#Exchange" || old["Action"] != want {
		return fail()
	}
	if kind == "mark-read" && old["IsRead"] != true || kind == "mark-unread" && old["IsRead"] != false {
		return fail()
	}
	if kind == "flag" && object(old["Flag"])["FlagStatus"] != "Flagged" || kind == "complete" && object(old["Flag"])["FlagStatus"] != "Complete" {
		return fail()
	}
	a := map[string]any{"__type": "ConversationAction:#Exchange", "Action": want, "ConversationId": map[string]any{"__type": "ItemId:#Exchange", "Id": p.ConversationID}, "ContextFolderId": folder(p.FolderID)}
	if kind != "move" {
		through, ok := timestamp(p.Through)
		if !ok {
			return fail()
		}
		a["ConversationLastSyncTime"] = through
	}
	switch kind {
	case "mark-read", "mark-unread":
		a["IsRead"] = kind == "mark-read"
		a["SuppressReadReceipts"] = true
		b["ReturnMovedItemIds"] = false
	case "flag":
		a["Flag"] = map[string]any{"__type": "FlagType:#Exchange", "FlagStatus": "Flagged"}
		delete(b, "ReturnMovedItemIds")
	case "complete":
		at, ok := timestamp(p.CompletedAt)
		if !ok {
			return fail()
		}
		a["Flag"] = map[string]any{"__type": "FlagType:#Exchange", "FlagStatus": "Complete", "CompleteDate": at}
		delete(b, "ReturnMovedItemIds")
	case "restore":
		b["ReturnMovedItemIds"] = false
	case "move":
		if !validID(p.DestinationID) || p.DestinationID == p.FolderID {
			return fail()
		}
		a["DestinationFolderId"] = folder(p.DestinationID)
		b["ReturnMovedItemIds"] = true
	}
	b["ConversationActions"] = []any{a}
	data, err := json.Marshal(root)
	if err != nil {
		return fail()
	}
	return d.Session, d.Operation("POST", "/owa/service.svc", data, transport.Write, Validate), nil
}

func Validate(data []byte) transport.EnvelopeResult {
	var root struct {
		Body struct {
			ResponseMessages struct {
				Items []struct {
					Type  string `json:"__type"`
					Class string `json:"ResponseClass"`
					Code  string `json:"ResponseCode"`
				} `json:"Items"`
			} `json:"ResponseMessages"`
		} `json:"Body"`
	}
	if json.Unmarshal(data, &root) != nil || len(root.Body.ResponseMessages.Items) != 1 {
		return transport.EnvelopeError
	}
	r := root.Body.ResponseMessages.Items[0]
	if r.Type != "ApplyConversationActionResponseMessage:#Exchange" || r.Class != "Success" || r.Code != "NoError" {
		return transport.EnvelopeError
	}
	return transport.EnvelopeOK
}
