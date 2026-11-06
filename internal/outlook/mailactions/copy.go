package mailactions

import (
	"encoding/json"
	"fmt"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/webcapture"
)

// PrepareCopy reproduces the observed item-scoped CopyItem request. The
// source and destination are caller-selected; no captured fixture ID survives.
func PrepareCopy(detail string, current transport.Session, id, destination string) (transport.Session, transport.Operation, error) {
	fail := func() (transport.Session, transport.Operation, error) {
		return transport.Session{}, transport.Operation{}, fmt.Errorf("message copy does not match the observed request")
	}
	if !validID(id) || !validID(destination) {
		return fail()
	}
	d, err := webcapture.Load(detail, "", current)
	if err != nil || d.Method != "POST" || d.URL.Path != "/owa/service.svc" || d.Session.Headers.Get("Action") != "CopyItem" || !d.HeaderBody {
		return fail()
	}
	q := d.URL.Query()
	counter := q.Get("n")
	if len(q) != 3 || len(q["action"]) != 1 || len(q["app"]) != 1 || len(q["n"]) != 1 || q.Get("action") != "CopyItem" || q.Get("app") != "Mail" || len(counter) == 0 || len(counter) > 12 {
		return fail()
	}
	for _, ch := range counter {
		if ch < '0' || ch > '9' {
			return fail()
		}
	}
	var root map[string]any
	if json.Unmarshal(d.Body, &root) != nil || len(root) != 3 || root["__type"] != "CopyItemJsonRequest:#Exchange" {
		return fail()
	}
	header := object(root["Header"])
	zone := object(header["TimeZoneContext"])
	definition := object(zone["TimeZoneDefinition"])
	zoneID, ok := definition["Id"].(string)
	if len(header) != 3 || header["__type"] != "JsonRequestHeaders:#Exchange" || header["RequestServerVersion"] != "V2018_01_08" || len(zone) != 2 || zone["__type"] != "TimeZoneContext:#Exchange" || len(definition) != 2 || definition["__type"] != "TimeZoneDefinitionType:#Exchange" || !ok || !validID(zoneID) || len(zoneID) > 128 {
		return fail()
	}
	body := object(root["Body"])
	ids, ok := body["ItemIds"].([]any)
	target := object(object(body["ToFolderId"])["BaseFolderId"])
	if body["__type"] != "CopyItemRequest:#Exchange" || len(body) != 5 || !ok || len(ids) != 1 || object(ids[0])["__type"] != "ItemId:#Exchange" || body["ReturnNewItemIds"] != true || body["RemoteExecute"] != false || object(body["ToFolderId"])["__type"] != "TargetFolderId:#Exchange" || target["__type"] != "FolderId:#Exchange" {
		return fail()
	}
	root["Body"] = map[string]any{"__type": "CopyItemRequest:#Exchange", "ItemIds": []any{map[string]any{"__type": "ItemId:#Exchange", "Id": id}}, "ReturnNewItemIds": true, "RemoteExecute": false, "ToFolderId": map[string]any{"__type": "TargetFolderId:#Exchange", "BaseFolderId": map[string]any{"__type": "FolderId:#Exchange", "Id": destination}}}
	encoded, err := json.Marshal(root)
	if err != nil {
		return fail()
	}
	return d.Session, d.Operation("POST", "/owa/service.svc", encoded, transport.Write, func(data []byte) transport.EnvelopeResult { _, state := ParseCopy(data); return state }), nil
}

// ParseCopy requires one new message identity and its change key. A response
// that merely says HTTP 200 without them is an unknown write outcome.
func ParseCopy(data []byte) (Result, transport.EnvelopeResult) {
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
	if json.Unmarshal(data, &root) != nil || root.Body.Type != "CopyItemResponse:#Exchange" || len(root.Body.Responses.Items) != 1 {
		return Result{}, transport.EnvelopeError
	}
	r := root.Body.Responses.Items[0]
	if r.Class != "Success" || r.Code != "NoError" {
		return Result{}, transport.EnvelopeRejected
	}
	if r.Type != "ItemInfoResponseMessage:#Exchange" || len(r.Items) != 1 || r.Items[0].Type != "Message:#Exchange" {
		return Result{}, transport.EnvelopeError
	}
	result := Result{ID: r.Items[0].ID.ID, ChangeKey: r.Items[0].ID.ChangeKey}
	if !validID(result.ID) || !validID(result.ChangeKey) {
		return Result{}, transport.EnvelopeError
	}
	return result, transport.EnvelopeOK
}
