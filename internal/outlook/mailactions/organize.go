package mailactions

import (
	"encoding/json"
	"fmt"
	"net/url"
	"strings"
	"time"
	"unicode/utf8"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/webcapture"
)

// PrepareMove replaces the captured item and destination. Outlook returns a
// replacement item ID and change key; callers must use those for later reads.
func PrepareMove(detail string, current transport.Session, id, destination string) (transport.Session, transport.Operation, error) {
	fail := func() (transport.Session, transport.Operation, error) {
		return transport.Session{}, transport.Operation{}, fmt.Errorf("message move does not match the observed request")
	}
	if !validID(id) || !validID(destination) {
		return fail()
	}
	d, err := webcapture.Load(detail, "", current)
	if err != nil || d.Method != "POST" || d.URL.Path != "/owa/service.svc" || d.URL.Query().Get("action") != "MoveItem" || d.URL.Query().Get("app") != "Mail" || d.Session.Headers.Get("Action") != "MoveItem" || !d.HeaderBody {
		return fail()
	}
	var root map[string]any
	if json.Unmarshal(d.Body, &root) != nil || root["__type"] != "MoveItemJsonRequest:#Exchange" {
		return fail()
	}
	body := object(root["Body"])
	ids, ok := body["ItemIds"].([]any)
	target := object(object(body["ToFolderId"])["BaseFolderId"])
	if body["__type"] != "MoveItemRequest:#Exchange" || !ok || len(ids) != 1 || object(ids[0])["__type"] != "ItemId:#Exchange" || body["ReturnNewItemIds"] != true || object(body["ToFolderId"])["__type"] != "TargetFolderId:#Exchange" || target["__type"] != "FolderId:#Exchange" {
		return fail()
	}
	root["Body"] = map[string]any{"__type": "MoveItemRequest:#Exchange", "ItemIds": []any{map[string]any{"__type": "ItemId:#Exchange", "Id": id}}, "ReturnNewItemIds": true, "ToFolderId": map[string]any{"__type": "TargetFolderId:#Exchange", "BaseFolderId": map[string]any{"__type": "FolderId:#Exchange", "Id": destination}}}
	encoded, err := json.Marshal(root)
	if err != nil {
		return fail()
	}
	return d.Session, d.Operation("POST", "/owa/service.svc", encoded, transport.Write, func(data []byte) transport.EnvelopeResult { _, state := ParseMove(data); return state }), nil
}

func ParseMove(data []byte) (Result, transport.EnvelopeResult) {
	var root struct {
		Body struct {
			Type      string `json:"__type"`
			Responses struct {
				Items []struct {
					Type  string `json:"__type"`
					Class string `json:"ResponseClass"`
					Code  string `json:"ResponseCode"`
					Items []struct {
						Type string                         `json:"__type"`
						ID   struct{ ID, ChangeKey string } `json:"ItemId"`
					} `json:"Items"`
				} `json:"Items"`
			} `json:"ResponseMessages"`
		} `json:"Body"`
	}
	if json.Unmarshal(data, &root) != nil || root.Body.Type != "MoveItemResponse:#Exchange" || len(root.Body.Responses.Items) != 1 {
		return Result{}, transport.EnvelopeError
	}
	response := root.Body.Responses.Items[0]
	if response.Class != "Success" || response.Code != "NoError" {
		return Result{}, transport.EnvelopeRejected
	}
	if response.Type != "ItemInfoResponseMessage:#Exchange" || len(response.Items) != 1 || response.Items[0].Type != "Message:#Exchange" {
		return Result{}, transport.EnvelopeError
	}
	result := Result{ID: response.Items[0].ID.ID, ChangeKey: response.Items[0].ID.ChangeKey}
	if !validID(result.ID) || !validID(result.ChangeKey) {
		return Result{}, transport.EnvelopeError
	}
	return result, transport.EnvelopeOK
}

// PrepareCategories replaces the complete category array on one message. The
// caller must first read and preserve the message's existing categories.
func PrepareCategories(detail string, current transport.Session, id string, categories []string) (transport.Session, transport.Operation, error) {
	fail := func() (transport.Session, transport.Operation, error) {
		return transport.Session{}, transport.Operation{}, fmt.Errorf("message categories do not match the observed request")
	}
	if !validID(id) || categories == nil || len(categories) > 128 {
		return fail()
	}
	seen := map[string]bool{}
	for _, name := range categories {
		if name == "" || len(name) > 255 || !utf8.ValidString(name) || strings.ContainsAny(name, "\r\n\x00") || seen[strings.ToLower(name)] {
			return fail()
		}
		seen[strings.ToLower(name)] = true
	}
	d, err := webcapture.Load(detail, "", current)
	if err != nil || d.Method != "POST" || d.URL.Path != "/owa/service.svc" || d.URL.Query().Get("action") != "UpdateItem" || d.URL.Query().Get("app") != "Mail" || d.Session.Headers.Get("Action") != "UpdateItem" || !d.HeaderBody {
		return fail()
	}
	var root map[string]any
	if json.Unmarshal(d.Body, &root) != nil || root["__type"] != "UpdateItemJsonRequest:#Exchange" {
		return fail()
	}
	old := object(root["Body"])
	changes, ok := old["ItemChanges"].([]any)
	if old["__type"] != "UpdateItemRequest:#Exchange" || old["ConflictResolution"] != "AlwaysOverwrite" || old["MessageDisposition"] != "SaveOnly" || !ok || len(changes) != 1 {
		return fail()
	}
	change := object(changes[0])
	updates, ok := change["Updates"].([]any)
	if change["__type"] != "ItemChange:#Exchange" || object(change["ItemId"])["__type"] != "ItemId:#Exchange" || !ok || len(updates) != 1 {
		return fail()
	}
	update := object(updates[0])
	path, item := object(update["Path"]), object(update["Item"])
	if update["__type"] != "SetItemField:#Exchange" || path["__type"] != "PropertyUri:#Exchange" || path["FieldURI"] != "Categories" || item["__type"] != "Message:#Exchange" {
		return fail()
	}
	if _, ok := item["Categories"].([]any); !ok {
		return fail()
	}
	root["Body"] = map[string]any{"__type": "UpdateItemRequest:#Exchange", "ConflictResolution": "AlwaysOverwrite", "MessageDisposition": "SaveOnly", "ItemChanges": []any{map[string]any{"__type": "ItemChange:#Exchange", "ItemId": map[string]any{"__type": "ItemId:#Exchange", "Id": id}, "Updates": []any{map[string]any{"__type": "SetItemField:#Exchange", "Path": map[string]any{"__type": "PropertyUri:#Exchange", "FieldURI": "Categories"}, "Item": map[string]any{"__type": "Message:#Exchange", "Categories": categories}}}}}}
	encoded, err := json.Marshal(root)
	if err != nil {
		return fail()
	}
	return d.Session, d.Operation("POST", "/owa/service.svc", encoded, transport.Write, func(data []byte) transport.EnvelopeResult { _, state := ParseCategoriesResult(data); return state }), nil
}

func ParseCategoriesResult(data []byte) (Result, transport.EnvelopeResult) {
	return parseUpdateResult(data)
}

func parseUpdateResult(data []byte) (Result, transport.EnvelopeResult) {
	// The category update uses the same success envelope as flag updates.
	return ParseResult("flag", data)
}

// PrepareComplete composes the observed single-message Flag update with the
// Complete flag shape observed on a conversation. A live item-level proof is
// still required; callers must supply the reviewed completion timestamp.
func PrepareComplete(detail string, current transport.Session, id, completed string) (transport.Session, transport.Operation, error) {
	at, err := time.Parse(time.RFC3339Nano, completed)
	if err != nil {
		return transport.Session{}, transport.Operation{}, fmt.Errorf("completion time must be RFC3339")
	}
	session, op, err := Prepare("flag", detail, current, id)
	if err != nil {
		return transport.Session{}, transport.Operation{}, err
	}
	encoded, err := url.QueryUnescape(op.Headers.Get("X-Owa-Urlpostdata"))
	if err != nil {
		return transport.Session{}, transport.Operation{}, fmt.Errorf("invalid flag update")
	}
	var root map[string]any
	if json.Unmarshal([]byte(encoded), &root) != nil {
		return transport.Session{}, transport.Operation{}, fmt.Errorf("invalid flag update")
	}
	body := object(root["Body"])
	changes, ok := body["ItemChanges"].([]any)
	if !ok || len(changes) != 1 {
		return transport.Session{}, transport.Operation{}, fmt.Errorf("invalid flag update")
	}
	updates, ok := object(changes[0])["Updates"].([]any)
	if !ok || len(updates) != 1 {
		return transport.Session{}, transport.Operation{}, fmt.Errorf("invalid flag update")
	}
	item := object(object(updates[0])["Item"])
	flag := object(item["Flag"])
	if flag["__type"] != "FlagType:#Exchange" || flag["FlagStatus"] != "Flagged" {
		return transport.Session{}, transport.Operation{}, fmt.Errorf("invalid flag update")
	}
	item["Flag"] = map[string]any{"__type": "FlagType:#Exchange", "FlagStatus": "Complete", "CompleteDate": at.UTC().Format("2006-01-02T15:04:05.000Z")}
	bodyBytes, err := json.Marshal(root)
	if err != nil {
		return transport.Session{}, transport.Operation{}, fmt.Errorf("invalid flag update")
	}
	op.Headers.Set("X-Owa-Urlpostdata", strings.ReplaceAll(url.QueryEscape(string(bodyBytes)), "+", "%20"))
	return session, op, nil
}
