package mailactions

import (
	"encoding/base64"
	"encoding/json"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/capture"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func TestCopyUsesOneCallerItemAndDestination(t *testing.T) {
	claims, _ := json.Marshal(map[string]any{"exp": time.Now().Add(time.Hour).Unix(), "tid": "tenant", "oid": "user", "aud": "audience"})
	current := transport.Session{AccountKey: "fixture-account", Origin: "https://outlook.cloud.microsoft", ExpiresAt: time.Now().Add(time.Hour), Headers: http.Header{"Authorization": {"Bearer h." + base64.RawURLEncoding.EncodeToString(claims) + ".s"}, "Cookie": {"fixture-cookie"}, "X-Anchormailbox": {"fixture-anchor"}}}
	bundle, err := capture.NewRuntimeBundle(current)
	if err != nil {
		t.Fatal(err)
	}
	session := current
	session.Headers = current.Headers.Clone()
	session.Headers.Set("Action", "CopyItem")
	template := `{"__type":"CopyItemJsonRequest:#Exchange","Header":{"__type":"JsonRequestHeaders:#Exchange","RequestServerVersion":"V2018_01_08","TimeZoneContext":{"__type":"TimeZoneContext:#Exchange","TimeZoneDefinition":{"__type":"TimeZoneDefinitionType:#Exchange","Id":"UTC"}}},"Body":{"__type":"CopyItemRequest:#Exchange","ItemIds":[{"__type":"ItemId:#Exchange","Id":"captured-item"}],"ToFolderId":{"__type":"TargetFolderId:#Exchange","BaseFolderId":{"__type":"FolderId:#Exchange","Id":"captured-folder"}},"ReturnNewItemIds":true,"RemoteExecute":false}}`
	op := transport.Operation{Method: "POST", Path: "/owa/service.svc", Query: url.Values{"action": {"CopyItem"}, "app": {"Mail"}, "n": {"1"}}, Headers: http.Header{"X-Owa-Urlpostdata": {url.QueryEscape(template)}}, Class: transport.Write}
	if err = bundle.AddWeb("mail_copy", session, op); err != nil {
		t.Fatal(err)
	}
	encoded, _ := json.Marshal(bundle)
	path := filepath.Join(t.TempDir(), "bundle.json")
	if err = os.WriteFile(path, encoded, 0600); err != nil {
		t.Fatal(err)
	}
	t.Setenv("OUTLOOK_WEB_OPERATIONS_FILE", path)
	_, got, err := PrepareCopy("runtime:mail_copy", current, "selected-item", "selected-folder")
	if err != nil || got.Class != transport.Write || got.Query.Get("action") != "CopyItem" {
		t.Fatal("copy was not prepared")
	}
	body, err := url.QueryUnescape(got.Headers.Get("X-Owa-Urlpostdata"))
	if err != nil || strings.Contains(body, "captured-item") || strings.Contains(body, "captured-folder") {
		t.Fatal("captured identifiers survived")
	}
	var v struct {
		Body struct {
			ItemIDs []struct {
				ID string `json:"Id"`
			} `json:"ItemIds"`
			Return bool `json:"ReturnNewItemIds"`
			Remote bool `json:"RemoteExecute"`
			Target struct {
				Base struct {
					ID string `json:"Id"`
				} `json:"BaseFolderId"`
			} `json:"ToFolderId"`
		} `json:"Body"`
	}
	if json.Unmarshal([]byte(body), &v) != nil || len(v.Body.ItemIDs) != 1 || v.Body.ItemIDs[0].ID != "selected-item" || v.Body.Target.Base.ID != "selected-folder" || !v.Body.Return || v.Body.Remote {
		t.Fatal("copy scope or destination changed")
	}
	if _, _, err = PrepareCopy("runtime:mail_copy", current, "selected-item", "bad\nfolder"); err == nil {
		t.Fatal("unsafe folder accepted")
	}
	// A private template must not smuggle unobserved top-level request fields.
	tampered := bundle.WebOperations["mail_copy"]
	tampered.Body = json.RawMessage(strings.Replace(template, `"Body":`, `"UnobservedOperation":true,"Body":`, 1))
	bundle.WebOperations["mail_copy"] = tampered
	encoded, _ = json.Marshal(bundle)
	if err = os.WriteFile(path, encoded, 0600); err != nil {
		t.Fatal(err)
	}
	if _, _, err = PrepareCopy("runtime:mail_copy", current, "selected-item", "selected-folder"); err == nil {
		t.Fatal("unobserved request fields accepted")
	}
}

func TestCopyRequiresCompleteNewMessageResult(t *testing.T) {
	valid := `{"Body":{"__type":"CopyItemResponse:#Exchange","ResponseMessages":{"Items":[{"__type":"ItemInfoResponseMessage:#Exchange","ResponseClass":"Success","ResponseCode":"NoError","Items":[{"__type":"Message:#Exchange","ItemId":{"Id":"new-id","ChangeKey":"new-key"}}]}]}}}`
	result, state := ParseCopy([]byte(valid))
	if state != transport.EnvelopeOK || result.ID != "new-id" || result.ChangeKey != "new-key" {
		t.Fatal("copy identity was not preserved")
	}
	for _, bad := range []string{`{}`, strings.Replace(valid, "CopyItemResponse", "MoveItemResponse", 1), strings.Replace(valid, `"ChangeKey":"new-key"`, `"ChangeKey":""`, 1), strings.Replace(valid, `"ResponseClass":"Success"`, `"ResponseClass":"Error"`, 1)} {
		if _, state := ParseCopy([]byte(bad)); state == transport.EnvelopeOK {
			t.Fatal("incomplete or rejected copy was accepted")
		}
	}
}
