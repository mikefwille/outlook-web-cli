// Package drafts builds narrowly scoped SaveOnly draft operations from observed
// Outlook request templates. It does not implement sending or meeting actions.
package drafts

import (
	"encoding/json"
	"errors"
	"html"
	"strings"
)

var ErrInvalid = errors.New("invalid draft request")

// PrepareGet retains the observed GetItem shape while requesting exactly the
// item selected by the caller. No content or identifier from the fixture is
// carried forward.
func PrepareGet(template []byte, id string) ([]byte, error) {
	if id == "" || len(id) > 4096 {
		return nil, ErrInvalid
	}
	var root map[string]any
	if json.Unmarshal(template, &root) != nil || root["__type"] != "GetItemJsonRequest:#Exchange" {
		return nil, ErrInvalid
	}
	body, ok := root["Body"].(map[string]any)
	if !ok || body["__type"] != "GetItemRequest:#Exchange" {
		return nil, ErrInvalid
	}
	ids, ok := body["ItemIds"].([]any)
	if !ok || len(ids) != 1 {
		return nil, ErrInvalid
	}
	item, ok := ids[0].(map[string]any)
	if !ok || item["__type"] != "ItemId:#Exchange" {
		return nil, ErrInvalid
	}
	body["ItemIds"] = []any{map[string]any{"__type": "ItemId:#Exchange", "Id": id}}
	return json.Marshal(root)
}

// PrepareCreate replaces only the fixture's subject and body and refuses any
// recipient or pre-existing deferred send content. The server request remains
// SaveOnly and SendOnNotFoundError is explicitly disabled.
func PrepareCreate(template []byte, subject, text string) ([]byte, error) {
	return PrepareCreateWithRecipients(template, subject, text, Recipients{})
}

func PrepareCreateWithRecipients(template []byte, subject, text string, recipients Recipients) ([]byte, error) {
	if len(recipients.To)+len(recipients.Cc)+len(recipients.Bcc) > 500 || len(subject) > 998 || len(text) > 1<<20 || strings.ContainsAny(subject, "\r\n\x00") {
		return nil, ErrInvalid
	}
	var root map[string]any
	if json.Unmarshal(template, &root) != nil || root["__type"] != "CreateItemJsonRequest:#Exchange" {
		return nil, ErrInvalid
	}
	body, ok := root["Body"].(map[string]any)
	if !ok || body["__type"] != "CreateItemRequest:#Exchange" || body["MessageDisposition"] != "SaveOnly" || body["ComposeOperation"] != "newMail" {
		return nil, ErrInvalid
	}
	items, ok := body["Items"].([]any)
	if !ok || len(items) != 1 {
		return nil, ErrInvalid
	}
	item, ok := items[0].(map[string]any)
	if !ok || item["__type"] != "Message:#Exchange" || item["MessageDisposition"] != "SaveOnly" {
		return nil, ErrInvalid
	}
	for _, key := range []string{"ToRecipients", "CcRecipients", "BccRecipients", "ReplyTo", "appendOnSend", "prependOnSend", "atMentionedRecipients", "recipientsAddedViaAtMention", "internetHeaders", "internetHeadersToRemove"} {
		v, ok := item[key].([]any)
		if !ok || len(v) != 0 {
			return nil, ErrInvalid
		}
	}
	// The template itself must be recipient-free so fixture recipients can
	// never leak. Only the caller-supplied, parsed addresses are inserted.
	for key, values := range map[string][]Recipient{"ToRecipients": recipients.To, "CcRecipients": recipients.Cc, "BccRecipients": recipients.Bcc} {
		if values == nil {
			values = []Recipient{}
		}
		for _, value := range values {
			if !validRecipient(value) {
				return nil, ErrInvalid
			}
		}
		item[key] = values
	}
	content, ok := item["Body"].(map[string]any)
	if !ok || content["BodyType"] != "HTML" {
		return nil, ErrInvalid
	}
	item["Subject"] = subject
	content["Value"] = "<html><body><div>" + strings.ReplaceAll(html.EscapeString(text), "\n", "<br>") + "</div></body></html>"
	item["plainText"] = text
	body["SendOnNotFoundError"] = false
	body["RemoteExecute"] = false
	return json.Marshal(root)
}

// PrepareUpdate keeps only the observed subject/body field updates. It never
// changes recipients or sends a message. Callers must supply the current ID and
// change key from a previously verified draft read/create.
func PrepareUpdate(template []byte, id, changeKey, subject, text string) ([]byte, error) {
	return PrepareUpdateWithRecipients(template, id, changeKey, subject, text, nil)
}

func PrepareUpdateWithRecipients(template []byte, id, changeKey, subject, text string, recipients *RecipientChanges) ([]byte, error) {
	if id == "" || changeKey == "" || len(id) > 4096 || len(changeKey) > 4096 || len(subject) > 998 || len(text) > 1<<20 || strings.ContainsAny(subject, "\r\n\x00") {
		return nil, ErrInvalid
	}
	var root map[string]any
	if json.Unmarshal(template, &root) != nil || root["__type"] != "UpdateItemJsonRequest:#Exchange" {
		return nil, ErrInvalid
	}
	body, ok := root["Body"].(map[string]any)
	if !ok || body["MessageDisposition"] != "SaveOnly" || body["SendCalendarInvitationsOrCancellations"] != "SendToNone" {
		return nil, ErrInvalid
	}
	changes, ok := body["ItemChanges"].([]any)
	if !ok || len(changes) != 1 {
		return nil, ErrInvalid
	}
	change, ok := changes[0].(map[string]any)
	if !ok {
		return nil, ErrInvalid
	}
	itemID, ok := change["ItemId"].(map[string]any)
	if !ok || itemID["__type"] != "ItemId:#Exchange" {
		return nil, ErrInvalid
	}
	itemID["Id"], itemID["ChangeKey"] = id, changeKey
	updates, ok := change["Updates"].([]any)
	if !ok {
		return nil, ErrInvalid
	}
	chosen := make([]any, 0, 2)
	seen := map[string]bool{}
	for _, raw := range updates {
		u, ok := raw.(map[string]any)
		if !ok {
			return nil, ErrInvalid
		}
		path, ok := u["Path"].(map[string]any)
		if !ok {
			continue
		}
		field, _ := path["FieldURI"].(string)
		var list *[]Recipient
		if recipients != nil {
			switch field {
			case "ToRecipients":
				list = recipients.To
			case "CcRecipients":
				list = recipients.Cc
			case "BccRecipients":
				list = recipients.Bcc
			}
		}
		if field != "Subject" && field != "Body" && list == nil {
			continue
		}
		if seen[field] {
			return nil, ErrInvalid
		}
		seen[field] = true
		item, ok := u["Item"].(map[string]any)
		if !ok || item["__type"] != "Message:#Exchange" {
			return nil, ErrInvalid
		}
		if list != nil {
			if len(*list) > 500 {
				return nil, ErrInvalid
			}
			for _, value := range *list {
				if !validRecipient(value) {
					return nil, ErrInvalid
				}
			}
			values := *list
			if values == nil {
				values = []Recipient{}
			}
			item[field] = values
		} else if field == "Subject" {
			item["Subject"] = subject
		} else {
			content, ok := item["Body"].(map[string]any)
			if !ok || content["BodyType"] != "HTML" {
				return nil, ErrInvalid
			}
			content["Value"] = "<html><body><div>" + strings.ReplaceAll(html.EscapeString(text), "\n", "<br>") + "</div></body></html>"
		}
		chosen = append(chosen, u)
	}
	if !seen["Subject"] || !seen["Body"] {
		return nil, ErrInvalid
	}
	if recipients != nil && (recipients.To != nil && !seen["ToRecipients"] || recipients.Cc != nil && !seen["CcRecipients"] || recipients.Bcc != nil && !seen["BccRecipients"]) {
		return nil, ErrInvalid
	}
	change["Updates"] = chosen
	body["SendOnNotFoundError"] = false
	// Microsoft's UpdateItem contract defines NeverOverwrite as failing on
	// conflicts. It must still be verified against this observed OWA route.
	body["ConflictResolution"] = "NeverOverwrite"
	return json.Marshal(root)
}

func PrepareDiscard(template []byte, id, changeKey string) ([]byte, error) {
	if id == "" || changeKey == "" || len(id) > 4096 || len(changeKey) > 4096 {
		return nil, ErrInvalid
	}
	var root map[string]any
	if json.Unmarshal(template, &root) != nil || root["__type"] != "DeleteItemJsonRequest:#Exchange" {
		return nil, ErrInvalid
	}
	body, ok := root["Body"].(map[string]any)
	if !ok || body["DeleteType"] != "MoveToDeletedItems" || body["SendMeetingCancellations"] != "SendToNone" {
		return nil, ErrInvalid
	}
	ids, ok := body["ItemIds"].([]any)
	if !ok || len(ids) != 1 {
		return nil, ErrInvalid
	}
	item, ok := ids[0].(map[string]any)
	if !ok || item["__type"] != "ItemId:#Exchange" {
		return nil, ErrInvalid
	}
	item["Id"], item["ChangeKey"] = id, changeKey
	return json.Marshal(root)
}

// PrepareSend uses the observed SendAndSaveCopy UpdateItem request, retaining
// only a subject update that matches the freshly read draft. Other existing
// fields, including recipients, body and attachments, are not replaced by
// fixture data. The caller must verify and preview the current draft first.
func PrepareSend(template []byte, id, changeKey, subject string) ([]byte, error) {
	if id == "" || changeKey == "" || len(id) > 4096 || len(changeKey) > 4096 || len(subject) > 998 || strings.ContainsAny(subject, "\r\n\x00") {
		return nil, ErrInvalid
	}
	var root map[string]any
	if json.Unmarshal(template, &root) != nil || root["__type"] != "UpdateItemJsonRequest:#Exchange" {
		return nil, ErrInvalid
	}
	body, ok := root["Body"].(map[string]any)
	if !ok || body["__type"] != "UpdateItemRequest:#Exchange" || body["MessageDisposition"] != "SendAndSaveCopy" || body["SendCalendarInvitationsOrCancellations"] != "SendToNone" {
		return nil, ErrInvalid
	}
	changes, ok := body["ItemChanges"].([]any)
	if !ok || len(changes) != 1 {
		return nil, ErrInvalid
	}
	change, ok := changes[0].(map[string]any)
	if !ok {
		return nil, ErrInvalid
	}
	itemID, ok := change["ItemId"].(map[string]any)
	if !ok || itemID["__type"] != "ItemId:#Exchange" {
		return nil, ErrInvalid
	}
	itemID["Id"], itemID["ChangeKey"] = id, changeKey
	updates, ok := change["Updates"].([]any)
	if !ok {
		return nil, ErrInvalid
	}
	var selected map[string]any
	for _, raw := range updates {
		u, ok := raw.(map[string]any)
		if !ok {
			return nil, ErrInvalid
		}
		path, ok := u["Path"].(map[string]any)
		if !ok || path["FieldURI"] != "Subject" {
			continue
		}
		if selected != nil || u["__type"] != "SetItemField:#Exchange" {
			return nil, ErrInvalid
		}
		item, ok := u["Item"].(map[string]any)
		if !ok || item["__type"] != "Message:#Exchange" {
			return nil, ErrInvalid
		}
		// Discard every fixture item field except the intended subject.
		u["Item"] = map[string]any{"__type": "Message:#Exchange", "Subject": subject}
		selected = u
	}
	if selected == nil {
		return nil, ErrInvalid
	}
	change["Updates"] = []any{selected}
	body["ConflictResolution"] = "NeverOverwrite"
	body["SendOnNotFoundError"] = false
	return json.Marshal(root)
}
