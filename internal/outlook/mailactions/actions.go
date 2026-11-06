// Package mailactions builds observed single-message Outlook actions.
package mailactions

import (
	"encoding/json"
	"fmt"
	"strings"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/webcapture"
)

func object(v any) map[string]any { x, _ := v.(map[string]any); return x }
func validID(s string) bool       { return s != "" && len(s) <= 4096 && !strings.ContainsAny(s, "\x00\r\n") }

// Prepare replaces every fixture ID and rebuilds a single observed action.
// The UI used AlwaysOverwrite: callers must disclose that a preceding version
// check is not atomic. Mark-read suppresses read receipts; trash is soft-delete.
func Prepare(kind, detail string, current transport.Session, id string) (transport.Session, transport.Operation, error) {
	fail := func() (transport.Session, transport.Operation, error) {
		return transport.Session{}, transport.Operation{}, fmt.Errorf("message request does not match the observed action")
	}
	action := map[string]string{"mark-read": "UpdateItem", "mark-unread": "UpdateItem", "flag": "UpdateItem", "unflag": "UpdateItem", "complete": "UpdateItem", "trash": "DeleteItem"}[kind]
	if action == "" || !validID(id) {
		return fail()
	}
	d, err := webcapture.Load(detail, "", current)
	if err != nil || d.Method != "POST" || d.URL.Path != "/owa/service.svc" || d.URL.Query().Get("action") != action || d.URL.Query().Get("app") != "Mail" || d.Session.Headers.Get("Action") != action || !d.HeaderBody {
		return fail()
	}
	var root map[string]any
	if json.Unmarshal(d.Body, &root) != nil || root["__type"] != action+"JsonRequest:#Exchange" {
		return fail()
	}
	old := object(root["Body"])
	if old["__type"] != action+"Request:#Exchange" {
		return fail()
	}
	body := map[string]any{"__type": action + "Request:#Exchange"}
	if kind == "trash" {
		ids, ok := old["ItemIds"].([]any)
		if !ok || len(ids) != 1 || object(ids[0])["__type"] != "ItemId:#Exchange" || old["DeleteType"] != "MoveToDeletedItems" || old["SendMeetingCancellations"] != "SendToNone" || old["SuppressReadReceipts"] != true || old["ReturnMovedItemIds"] != true || old["AffectedTaskOccurrences"] != "AllOccurrences" {
			return fail()
		}
		body["ItemIds"] = []any{map[string]any{"__type": "ItemId:#Exchange", "Id": id}}
		body["DeleteType"], body["SendMeetingCancellations"], body["AffectedTaskOccurrences"] = "MoveToDeletedItems", "SendToNone", "AllOccurrences"
		body["SuppressReadReceipts"], body["ReturnMovedItemIds"] = true, true
	} else {
		changes, ok := old["ItemChanges"].([]any)
		if !ok || len(changes) != 1 {
			return fail()
		}
		change := object(changes[0])
		updates, ok := change["Updates"].([]any)
		if !ok || len(updates) != 1 || change["__type"] != "ItemChange:#Exchange" || object(change["ItemId"])["__type"] != "ItemId:#Exchange" || old["ConflictResolution"] != "AlwaysOverwrite" || old["MessageDisposition"] != "SaveOnly" {
			return fail()
		}
		u := object(updates[0])
		path := object(u["Path"])
		item := object(u["Item"])
		field := "IsRead"
		if kind == "flag" || kind == "unflag" {
			field = "Flag"
		}
		if u["__type"] != "SetItemField:#Exchange" || path["__type"] != "PropertyUri:#Exchange" || path["FieldURI"] != field || item["__type"] != "Message:#Exchange" {
			return fail()
		}
		replacement := map[string]any{"__type": "Message:#Exchange"}
		if field == "IsRead" {
			want := kind == "mark-read"
			if item["IsRead"] != want || old["SuppressReadReceipts"] != true {
				return fail()
			}
			replacement[field] = want
			body["SuppressReadReceipts"] = true
		} else {
			status := "Flagged"
			if kind == "unflag" {
				status = "NotFlagged"
			}
			if object(item["Flag"])["FlagStatus"] != status || object(item["Flag"])["__type"] != "FlagType:#Exchange" {
				return fail()
			}
			replacement[field] = map[string]any{"__type": "FlagType:#Exchange", "FlagStatus": status}
		}
		body["ItemChanges"] = []any{map[string]any{"__type": "ItemChange:#Exchange", "ItemId": map[string]any{"__type": "ItemId:#Exchange", "Id": id}, "Updates": []any{map[string]any{"__type": "SetItemField:#Exchange", "Path": map[string]any{"__type": "PropertyUri:#Exchange", "FieldURI": field}, "Item": replacement}}}}
		body["ConflictResolution"], body["MessageDisposition"] = "AlwaysOverwrite", "SaveOnly"
	}
	root["Body"] = body
	encoded, err := json.Marshal(root)
	if err != nil {
		return fail()
	}
	return d.Session, d.Operation("POST", "/owa/service.svc", encoded, transport.Write, func(data []byte) transport.EnvelopeResult { _, state := ParseResult(kind, data); return state }), nil
}

type Result struct {
	ID        string `json:"id"`
	ChangeKey string `json:"change_key"`
}

func ParseResult(kind string, data []byte) (Result, transport.EnvelopeResult) {
	action := map[string]string{"mark-read": "UpdateItem", "mark-unread": "UpdateItem", "flag": "UpdateItem", "unflag": "UpdateItem", "complete": "UpdateItem", "trash": "DeleteItem"}[kind]
	var root struct {
		Body struct {
			Type      string `json:"__type"`
			Responses struct {
				Items []struct {
					Type  string `json:"__type"`
					Class string `json:"ResponseClass"`
					Code  string `json:"ResponseCode"`
					Items []struct {
						ID struct{ ID, ChangeKey string } `json:"ItemId"`
					} `json:"Items"`
					Moved struct{ ID, ChangeKey string } `json:"MovedItemId"`
				} `json:"Items"`
			} `json:"ResponseMessages"`
		} `json:"Body"`
	}
	if action == "" || json.Unmarshal(data, &root) != nil || root.Body.Type != action+"Response:#Exchange" || len(root.Body.Responses.Items) != 1 {
		return Result{}, transport.EnvelopeError
	}
	x := root.Body.Responses.Items[0]
	if x.Class != "Success" || x.Code != "NoError" {
		return Result{}, transport.EnvelopeRejected
	}
	var result Result
	if kind == "trash" {
		if x.Type != "DeleteItemResponseMessage:#Exchange" {
			return Result{}, transport.EnvelopeError
		}
		result = Result{ID: x.Moved.ID, ChangeKey: x.Moved.ChangeKey}
	} else {
		if x.Type != "UpdateItemResponseMessage:#Exchange" || len(x.Items) != 1 {
			return Result{}, transport.EnvelopeError
		}
		result = Result{ID: x.Items[0].ID.ID, ChangeKey: x.Items[0].ID.ChangeKey}
	}
	if !validID(result.ID) || !validID(result.ChangeKey) {
		return Result{}, transport.EnvelopeError
	}
	return result, transport.EnvelopeOK
}
