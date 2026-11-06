// Package attachments implements observed Outlook file-attachment operations.
package attachments

import (
	"crypto/rand"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"net/http"
	"net/url"
	"strings"
	"unicode/utf8"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/webcapture"
)

const MaxUploadBytes = 4 << 20

type Uploaded struct {
	ID            string `json:"id"`
	RootItemID    string `json:"root_item_id"`
	RootChangeKey string `json:"root_change_key"`
}

func newID() (string, error) {
	var b [16]byte
	if _, err := rand.Read(b[:]); err != nil {
		return "", err
	}
	b[6] = (b[6] & 0x0f) | 0x40
	b[8] = (b[8] & 0x3f) | 0x80
	v := hex.EncodeToString(b[:])
	return v[:8] + "-" + v[8:12] + "-" + v[12:16] + "-" + v[16:20] + "-" + v[20:], nil
}

func PrepareUpload(detail string, current transport.Session, id, key, name, mime string, data []byte) (transport.Session, transport.Operation, error) {
	fail := func() (transport.Session, transport.Operation, error) {
		return transport.Session{}, transport.Operation{}, fmt.Errorf("attachment request does not match the observed upload contract")
	}
	if id == "" || key == "" || len(id) > 4096 || len(key) > 4096 || name == "" || len(name) > 255 || !utf8.ValidString(name) || strings.ContainsAny(name, "/\\\x00\r\n") || mime == "" || len(mime) > 255 || strings.ContainsAny(mime, "\x00\r\n") || len(data) > MaxUploadBytes {
		return fail()
	}
	d, err := webcapture.Load(detail, "", current)
	if err != nil {
		return fail()
	}
	if d.Method != "POST" || d.URL.Path != "/owa/service.svc/CreateAttachmentFromLocalFile" || len(d.URL.Query()) != 0 || !d.HeaderBody || d.Session.Headers.Get("X-Owa-Action") != "CreateAttachmentFromLocalFile" {
		return fail()
	}
	var root map[string]any
	if json.Unmarshal(d.Body, &root) != nil || root["__type"] != "CreateAttachmentJsonRequest:#Exchange" {
		return fail()
	}
	b, ok := root["Body"].(map[string]any)
	if !ok || b["__type"] != "CreateAttachmentRequest:#Exchange" || b["SliceNumber"] != float64(0) || b["TotalSlices"] != float64(1) {
		return fail()
	}
	parent, ok := b["ParentItemId"].(map[string]any)
	if !ok || parent["__type"] != "ItemId:#Exchange" {
		return fail()
	}
	b["ParentItemId"] = map[string]any{"__type": "ItemId:#Exchange", "Id": id, "ChangeKey": key}
	items, ok := b["Attachments"].([]any)
	if !ok || len(items) != 1 {
		return fail()
	}
	a, ok := items[0].(map[string]any)
	if !ok || a["__type"] != "FileAttachment:#Exchange" || a["IsInline"] != false || a["Content"] != "" {
		return fail()
	}
	aid, ok := a["AttachmentId"].(map[string]any)
	if !ok {
		return fail()
	}
	if _, ok := mailboxRoutingForCurrent(aid["mailboxInfo"], current); !ok {
		return fail()
	}
	attachmentID, e1 := newID()
	cancelID, e2 := newID()
	if e1 != nil || e2 != nil {
		return fail()
	}
	aid["Id"] = attachmentID
	a["Name"] = name
	a["ContentType"] = mime
	a["Size"] = len(data)
	a["Content"] = ""
	b["CancellationId"] = cancelID
	envelope, err := json.Marshal(root)
	if err != nil {
		return fail()
	}
	op := transport.Operation{Method: "POST", Path: d.URL.Path, Query: d.URL.Query(), Headers: http.Header{"X-Owa-Urlpostdata": []string{strings.ReplaceAll(url.QueryEscape(string(envelope)), "+", "%20")}}, Body: data, Class: transport.Write, Validate: func(b []byte) transport.EnvelopeResult { _, s := ParseUploaded(b); return s }}
	return d.Session, op, nil
}

func ParseUploaded(data []byte) (Uploaded, transport.EnvelopeResult) {
	var root struct {
		Body *struct {
			ResponseMessages *struct {
				Items []struct {
					ResponseClass string `json:"ResponseClass"`
					ResponseCode  string `json:"ResponseCode"`
					Attachments   []struct {
						Type         string `json:"__type"`
						AttachmentID struct {
							ID      string `json:"Id"`
							RootID  string `json:"RootItemId"`
							RootKey string `json:"RootItemChangeKey"`
						} `json:"AttachmentId"`
					} `json:"Attachments"`
				} `json:"Items"`
			} `json:"ResponseMessages"`
		} `json:"Body"`
	}
	if json.Unmarshal(data, &root) != nil || root.Body == nil || root.Body.ResponseMessages == nil || len(root.Body.ResponseMessages.Items) != 1 {
		return Uploaded{}, transport.EnvelopeError
	}
	r := root.Body.ResponseMessages.Items[0]
	if r.ResponseClass != "Success" || r.ResponseCode != "NoError" || len(r.Attachments) != 1 || r.Attachments[0].Type != "FileAttachment:#Exchange" {
		return Uploaded{}, transport.EnvelopeError
	}
	a := r.Attachments[0].AttachmentID
	if a.ID == "" || a.RootID == "" || a.RootKey == "" {
		return Uploaded{}, transport.EnvelopeError
	}
	return Uploaded{ID: a.ID, RootItemID: a.RootID, RootChangeKey: a.RootKey}, transport.EnvelopeOK
}
