package folders

import (
	"encoding/json"
	"fmt"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/exchange"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/webcapture"
)

// PrepareList uses Outlook's observed deep folder query including names,
// parent IDs, distinguished folders and counts. Paging is explicit.
func PrepareList(detail string, current transport.Session, offset, limit int) (transport.Session, transport.Operation, error) {
	fail := func() (transport.Session, transport.Operation, error) {
		return transport.Session{}, transport.Operation{}, fmt.Errorf("folder listing does not match the observed hierarchy query")
	}
	if offset < 0 || offset > 1_000_000 || limit < 1 || limit > 1000 {
		return fail()
	}
	d, err := webcapture.Load(detail, "", current)
	if err != nil {
		return fail()
	}
	if d.Method != "POST" || d.URL.Path != "/owa/service.svc" || d.URL.Query().Get("action") != "FindFolder" || d.URL.Query().Get("app") != "Mail" || d.Session.Headers.Get("Action") != "FindFolder" || !d.HeaderBody {
		return fail()
	}
	var root map[string]any
	if json.Unmarshal(d.Body, &root) != nil || root["__type"] != "FindFolderJsonRequest:#Exchange" {
		return fail()
	}
	body, ok := root["Body"].(map[string]any)
	if !ok || body["__type"] != "FindFolderRequest:#Exchange" || body["Traversal"] != "Deep" || body["ShapeName"] != "Folder" {
		return fail()
	}
	parents, ok := body["ParentFolderIds"].([]any)
	if !ok || len(parents) != 1 {
		return fail()
	}
	p, ok := parents[0].(map[string]any)
	if !ok || p["__type"] != "DistinguishedFolderId:#Exchange" || p["Id"] != "msgfolderroot" {
		return fail()
	}
	paging, ok := body["Paging"].(map[string]any)
	if !ok || paging["__type"] != "IndexedPageView:#Exchange" || paging["BasePoint"] != "Beginning" {
		return fail()
	}
	shape, ok := body["FolderShape"].(map[string]any)
	if !ok || shape["__type"] != "FolderResponseShape:#Exchange" {
		return fail()
	}
	props, ok := shape["AdditionalProperties"].([]any)
	if !ok {
		return fail()
	}
	names := map[string]bool{}
	for _, v := range props {
		p, ok := v.(map[string]any)
		if !ok || p["__type"] != "PropertyUri:#Exchange" {
			return fail()
		}
		s, _ := p["FieldURI"].(string)
		names[s] = true
	}
	for _, name := range []string{"DistinguishedFolderId", "FolderClass", "FolderDisplayName", "ParentFolderId", "TotalCount", "UnreadCount"} {
		if !names[name] {
			return fail()
		}
	}
	paging["Offset"], paging["MaxEntriesReturned"] = offset, limit
	b, err := json.Marshal(root)
	if err != nil {
		return fail()
	}
	return d.Session, d.Operation("POST", d.URL.Path, b, transport.Read, exchange.ValidateFindFolder), nil
}
