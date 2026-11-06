package drafts

import (
	"encoding/json"
	"html"
	"strings"
)

// PrepareResponseDraft uses an observed ReplyToItem or ForwardItem SaveOnly
// request. The caller supplies every recipient; captured fixture addresses,
// reference identifiers, text and extended properties are never reused.
func PrepareResponseDraft(template []byte, kind, id, changeKey, subject, text string, recipients Recipients) ([]byte, error) {
	want := map[string]string{"reply": "ReplyToItem:#Exchange", "reply-all": "ReplyAllToItem:#Exchange", "forward": "ForwardItem:#Exchange"}[kind]
	if want == "" || id == "" || changeKey == "" || len(id) > 4096 || len(changeKey) > 4096 || len(subject) > 998 || len(text) > 1<<20 || strings.ContainsAny(subject, "\r\n\x00") || len(recipients.To)+len(recipients.Cc)+len(recipients.Bcc) == 0 || len(recipients.To)+len(recipients.Cc)+len(recipients.Bcc) > 500 {
		return nil, ErrInvalid
	}
	var root map[string]any
	if json.Unmarshal(template, &root) != nil || root["__type"] != "CreateItemJsonRequest:#Exchange" {
		return nil, ErrInvalid
	}
	body, ok := root["Body"].(map[string]any)
	if !ok || body["__type"] != "CreateItemRequest:#Exchange" || body["MessageDisposition"] != "SaveOnly" {
		return nil, ErrInvalid
	}
	items, ok := body["Items"].([]any)
	if !ok || len(items) != 1 {
		return nil, ErrInvalid
	}
	item, ok := items[0].(map[string]any)
	if !ok || item["__type"] != want || item["MessageDisposition"] != "SaveOnly" {
		return nil, ErrInvalid
	}
	ref, ok := item["ReferenceItemId"].(map[string]any)
	if !ok || ref["__type"] != "ItemId:#Exchange" {
		return nil, ErrInvalid
	}
	item["ReferenceItemId"] = map[string]any{"__type": "ItemId:#Exchange", "Id": id, "ChangeKey": changeKey}
	content, ok := item["NewBodyContent"].(map[string]any)
	if !ok || content["BodyType"] != "HTML" {
		return nil, ErrInvalid
	}
	item["NewBodyContent"] = map[string]any{"BodyType": "HTML", "Value": "<html><body><div>" + strings.ReplaceAll(html.EscapeString(text), "\n", "<br>") + "</div></body></html>"}
	for name, values := range map[string][]Recipient{"ToRecipients": recipients.To, "CcRecipients": recipients.Cc, "BccRecipients": recipients.Bcc} {
		for _, v := range values {
			if !validRecipient(v) {
				return nil, ErrInvalid
			}
		}
		if values == nil {
			values = []Recipient{}
		}
		item[name] = values
	}
	if from, ok := item["From"].(map[string]any); !ok || len(from) != 0 {
		return nil, ErrInvalid
	}
	item["Subject"] = subject
	item["ReplyTo"] = []any{}
	delete(item, "ReferenceItemDocumentId")
	delete(item, "ExtendedProperty")
	item["ShouldIgnoreChangeKey"] = false
	body["SendOnNotFoundError"] = false
	body["RemoteExecute"] = false
	body["SuppressMarkAsReadOnReplyOrForward"] = true
	body["ShouldSuppressReadReceipt"] = true
	return json.Marshal(root)
}
