package folders

import (
	"encoding/json"
	"fmt"
	"strings"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/webcapture"
)

type Params struct {
	ID, ParentID, Name string
}

func validID(s string) bool { return s != "" && len(s) <= 4096 && !strings.ContainsAny(s, "\x00\r\n") }
func validName(s string) bool {
	return strings.TrimSpace(s) == s && s != "" && len(s) <= 255 && !strings.ContainsAny(s, "\x00\r\n")
}
func object(v any) map[string]any { m, _ := v.(map[string]any); return m }
func folderID(id string) map[string]any {
	return map[string]any{"__type": "FolderId:#Exchange", "Id": id}
}
func target(id string) map[string]any {
	return map[string]any{"__type": "TargetFolderId:#Exchange", "BaseFolderId": folderID(id)}
}

// Prepare replaces all captured fixture values using the observed folder action.
// Delete only requests a move to Deleted Items; it never requests hard deletion.
func Prepare(kind, detail string, current transport.Session, p Params) (transport.Session, transport.Operation, error) {
	fail := func() (transport.Session, transport.Operation, error) {
		return transport.Session{}, transport.Operation{}, fmt.Errorf("folder request does not match the observed action contract")
	}
	action := map[string]string{"create": "CreateFolder", "rename": "UpdateFolder", "move": "MoveFolder", "delete": "DeleteFolder"}[kind]
	if action == "" || kind != "create" && !validID(p.ID) || (kind == "create" || kind == "move") && !validID(p.ParentID) || (kind == "create" || kind == "rename") && !validName(p.Name) || kind == "move" && p.ID == p.ParentID {
		return fail()
	}
	d, err := webcapture.Load(detail, "", current)
	if err != nil || d.Method != "POST" || d.URL.Path != "/owa/service.svc" || d.URL.Query().Get("action") != action || d.URL.Query().Get("app") != "Mail" || d.Session.Headers.Get("Action") != action || !d.HeaderBody {
		return fail()
	}
	var root map[string]any
	if json.Unmarshal(d.Body, &root) != nil || root["__type"] != action+"JsonRequest:#Exchange" || object(root["Body"])["__type"] != action+"Request:#Exchange" {
		return fail()
	}
	old := object(root["Body"])
	body := map[string]any{"__type": action + "Request:#Exchange"}
	switch kind {
	case "create":
		items, ok := old["Folders"].([]any)
		if !ok || len(items) != 1 || object(items[0])["__type"] != "Folder:#Exchange" || object(items[0])["FolderClass"] != "IPF.Note" || object(object(old["ParentFolderId"])["BaseFolderId"])["__type"] != "FolderId:#Exchange" {
			return fail()
		}
		body["ParentFolderId"] = target(p.ParentID)
		body["Folders"] = []any{map[string]any{"__type": "Folder:#Exchange", "DisplayName": p.Name, "FolderClass": "IPF.Note"}}
	case "rename":
		changes, ok := old["FolderChanges"].([]any)
		if !ok || len(changes) != 1 {
			return fail()
		}
		change := object(changes[0])
		updates, ok := change["Updates"].([]any)
		if !ok || len(updates) != 1 || object(updates[0])["__type"] != "SetFolderField:#Exchange" || object(object(updates[0])["Path"])["FieldURI"] != "FolderDisplayName" {
			return fail()
		}
		body["FolderChanges"] = []any{map[string]any{"FolderId": folderID(p.ID), "Updates": []any{map[string]any{"__type": "SetFolderField:#Exchange", "Path": map[string]any{"__type": "PropertyUri:#Exchange", "FieldURI": "FolderDisplayName"}, "Folder": map[string]any{"__type": "Folder:#Exchange", "DisplayName": p.Name}}}}}
	case "move":
		items, ok := old["FolderIds"].([]any)
		if !ok || len(items) != 1 || object(items[0])["__type"] != "FolderId:#Exchange" || object(object(old["ToFolderId"])["BaseFolderId"])["__type"] != "FolderId:#Exchange" {
			return fail()
		}
		body["FolderIds"], body["ToFolderId"] = []any{folderID(p.ID)}, target(p.ParentID)
	case "delete":
		items, ok := old["FolderIds"].([]any)
		if !ok || len(items) != 1 || object(items[0])["__type"] != "FolderId:#Exchange" || old["DeleteType"] != "MoveToDeletedItems" {
			return fail()
		}
		body["FolderIds"], body["DeleteType"] = []any{folderID(p.ID)}, "MoveToDeletedItems"
	}
	root["Body"] = body
	data, err := json.Marshal(root)
	if err != nil {
		return fail()
	}
	return d.Session, d.Operation("POST", "/owa/service.svc", data, transport.Write, func(b []byte) transport.EnvelopeResult { _, state := ParseResult(kind, b); return state }), nil
}

type Result struct {
	ID        string `json:"id,omitempty"`
	ChangeKey string `json:"change_key,omitempty"`
}

func ParseResult(kind string, data []byte) (Result, transport.EnvelopeResult) {
	var root struct {
		Body struct {
			Type             string `json:"__type"`
			ResponseMessages struct {
				Items []struct {
					Type    string `json:"__type"`
					Class   string `json:"ResponseClass"`
					Code    string `json:"ResponseCode"`
					Folders []struct {
						ID struct{ ID, ChangeKey string } `json:"FolderId"`
					} `json:"Folders"`
				} `json:"Items"`
			} `json:"ResponseMessages"`
		} `json:"Body"`
	}
	action := map[string]string{"create": "CreateFolder", "rename": "UpdateFolder", "move": "MoveFolder", "delete": "DeleteFolder"}[kind]
	if action == "" || json.Unmarshal(data, &root) != nil || root.Body.Type != action+"Response:#Exchange" || len(root.Body.ResponseMessages.Items) != 1 {
		return Result{}, transport.EnvelopeError
	}
	item := root.Body.ResponseMessages.Items[0]
	if item.Class != "Success" || item.Code != "NoError" {
		return Result{}, transport.EnvelopeRejected
	}
	if kind == "delete" {
		if item.Type != "SingleResponseMessage:#Exchange" {
			return Result{}, transport.EnvelopeError
		}
		return Result{}, transport.EnvelopeOK
	}
	if item.Type != "FolderInfoResponseMessage:#Exchange" || len(item.Folders) != 1 || !validID(item.Folders[0].ID.ID) || !validID(item.Folders[0].ID.ChangeKey) {
		return Result{}, transport.EnvelopeError
	}
	return Result{ID: item.Folders[0].ID.ID, ChangeKey: item.Folders[0].ID.ChangeKey}, transport.EnvelopeOK
}
