package attachments

import (
	"encoding/base64"
	"encoding/json"
	"fmt"
	"strings"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/webcapture"
)

func mailboxRoutingForCurrent(value any, current transport.Session) (map[string]any, bool) {
	m, ok := value.(map[string]any)
	if !ok || len(m) != 7 {
		return nil, false
	}
	for _, key := range []string{"sourceId", "type", "mailboxSmtpAddress", "userIdentity", "mailboxRank", "diagnosticData", "mailboxProvider"} {
		v, ok := m[key].(string)
		if !ok || v == "" || len(v) > 1024 || strings.ContainsAny(v, "\x00\r\n") {
			return nil, false
		}
	}
	auth := strings.SplitN(current.Headers.Get("Authorization"), " ", 2)
	if len(auth) != 2 || !strings.EqualFold(auth[0], "Bearer") {
		return nil, false
	}
	parts := strings.Split(auth[1], ".")
	if len(parts) != 3 {
		return nil, false
	}
	data, err := base64.RawURLEncoding.DecodeString(parts[1])
	if err != nil {
		return nil, false
	}
	var c struct {
		Tenant      string `json:"tid"`
		User        string `json:"upn"`
		Alternative string `json:"unique_name"`
	}
	if json.Unmarshal(data, &c) != nil || c.Tenant == "" {
		return nil, false
	}
	if c.User == "" {
		c.User = c.Alternative
	}
	identity := m["userIdentity"].(string)
	source := m["sourceId"].(string)
	smtp := m["mailboxSmtpAddress"].(string)
	at := strings.LastIndex(source, "@")
	if !validMailbox(c.User) || !strings.EqualFold(identity, c.User) || !validMailbox(smtp) || at <= 0 || !strings.EqualFold(source[at+1:], c.Tenant) {
		return nil, false
	}
	return m, true
}

// PrepareRemove builds the observed one-attachment DeleteAttachment request.
// Mailbox routing metadata remains in the owner-only, account-bound template;
// attachment and root IDs are always replaced with the selected values.
func PrepareRemove(detail string, current transport.Session, draftID, key, attachmentID string) (transport.Session, transport.Operation, error) {
	fail := func() (transport.Session, transport.Operation, error) {
		return transport.Session{}, transport.Operation{}, fmt.Errorf("attachment removal does not match the observed request")
	}
	for _, v := range []string{draftID, key, attachmentID} {
		if v == "" || len(v) > 4096 || strings.ContainsAny(v, "\x00\r\n") {
			return fail()
		}
	}
	d, err := webcapture.Load(detail, "", current)
	if err != nil || d.Method != "POST" || d.URL.Path != "/owa/service.svc" || d.URL.Query().Get("action") != "DeleteAttachment" || d.URL.Query().Get("app") != "Mail" || d.Session.Headers.Get("Action") != "DeleteAttachment" || !d.HeaderBody {
		return fail()
	}
	var root map[string]any
	if json.Unmarshal(d.Body, &root) != nil || root["__type"] != "DeleteAttachmentJsonRequest:#Exchange" {
		return fail()
	}
	body, ok := root["Body"].(map[string]any)
	if !ok || body["__type"] != "DeleteAttachmentRequest:#Exchange" || body["ClientSupportsIrm"] != true {
		return fail()
	}
	ids, ok := body["AttachmentIds"].([]any)
	if !ok || len(ids) != 1 {
		return fail()
	}
	captured, ok := ids[0].(map[string]any)
	if !ok || len(captured) != 4 {
		return fail()
	}
	for _, field := range []string{"Id", "RootItemId", "RootItemChangeKey"} {
		if _, ok := captured[field].(string); !ok {
			return fail()
		}
	}
	routing, ok := mailboxRoutingForCurrent(captured["mailboxInfo"], current)
	if !ok {
		return fail()
	}
	root["Body"] = map[string]any{"__type": "DeleteAttachmentRequest:#Exchange", "ClientSupportsIrm": true, "AttachmentIds": []any{map[string]any{"Id": attachmentID, "RootItemId": draftID, "RootItemChangeKey": key, "mailboxInfo": routing}}}
	encoded, err := json.Marshal(root)
	if err != nil {
		return fail()
	}
	return d.Session, d.Operation("POST", "/owa/service.svc", encoded, transport.Write, func(data []byte) transport.EnvelopeResult { _, state := ParseRemoved(data); return state }), nil
}

type Removed struct {
	RootItemID    string `json:"root_item_id"`
	RootChangeKey string `json:"root_change_key"`
}

func ParseRemoved(data []byte) (Removed, transport.EnvelopeResult) {
	var root struct {
		Body struct {
			Type      string `json:"__type"`
			Responses struct {
				Items []struct {
					Type  string `json:"__type"`
					Class string `json:"ResponseClass"`
					Code  string `json:"ResponseCode"`
					Root  struct {
						Type string `json:"__type"`
						ID   string `json:"RootItemId"`
						Key  string `json:"RootItemChangeKey"`
					} `json:"RootItemId"`
				} `json:"Items"`
			} `json:"ResponseMessages"`
		} `json:"Body"`
	}
	if json.Unmarshal(data, &root) != nil || root.Body.Type != "DeleteAttachmentResponse:#Exchange" || len(root.Body.Responses.Items) != 1 {
		return Removed{}, transport.EnvelopeError
	}
	response := root.Body.Responses.Items[0]
	if response.Class != "Success" || response.Code != "NoError" {
		return Removed{}, transport.EnvelopeRejected
	}
	if response.Type != "DeleteAttachmentResponseMessage:#Exchange" || response.Root.Type != "RootItemId:#Exchange" || response.Root.ID == "" || response.Root.Key == "" {
		return Removed{}, transport.EnvelopeError
	}
	return Removed{RootItemID: response.Root.ID, RootChangeKey: response.Root.Key}, transport.EnvelopeOK
}
