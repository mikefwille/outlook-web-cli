package attachments

import (
	"encoding/json"
	"strings"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

// Metadata describes fields returned by Outlook. ReportedSize is the server's
// value; it is not a verified byte count of the downloaded attachment.
type Metadata struct {
	ID           string `json:"id"`
	Type         string `json:"type"`
	Name         string `json:"name"`
	ContentType  string `json:"content_type,omitempty"`
	ContentID    string `json:"content_id,omitempty"`
	ReportedSize *int64 `json:"reported_size,omitempty"`
	IsInline     *bool  `json:"is_inline,omitempty"`
}
type MessageAttachments struct {
	MessageID      string     `json:"message_id"`
	ChangeKey      string     `json:"change_key"`
	HasAttachments bool       `json:"has_attachments"`
	Attachments    []Metadata `json:"attachments"`
}

// ParseMessage consumes the actual GetItem message, not a captured fixture.
// It fails rather than silently omitting a malformed or unidentifiable entry.
func ParseMessage(item []byte) (MessageAttachments, transport.EnvelopeResult) {
	var x struct {
		Type   string `json:"__type"`
		ItemID struct {
			ID        string `json:"Id"`
			ChangeKey string `json:"ChangeKey"`
		} `json:"ItemId"`
		HasAttachments *bool `json:"HasAttachments"`
		Attachments    []struct {
			Type string `json:"__type"`
			ID   struct {
				ID string `json:"Id"`
			} `json:"AttachmentId"`
			Name        string `json:"Name"`
			ContentType string `json:"ContentType"`
			ContentID   string `json:"ContentId"`
			Size        *int64 `json:"Size"`
			IsInline    *bool  `json:"IsInline"`
		} `json:"Attachments"`
	}
	if json.Unmarshal(item, &x) != nil || x.Type != "Message:#Exchange" || x.ItemID.ID == "" || x.ItemID.ChangeKey == "" || x.HasAttachments == nil || len(x.Attachments) > 500 || (*x.HasAttachments && len(x.Attachments) == 0) {
		return MessageAttachments{}, transport.EnvelopeError
	}
	result := MessageAttachments{MessageID: x.ItemID.ID, ChangeKey: x.ItemID.ChangeKey, HasAttachments: *x.HasAttachments, Attachments: make([]Metadata, 0, len(x.Attachments))}
	seen := map[string]bool{}
	for _, a := range x.Attachments {
		if a.ID.ID == "" || len(a.ID.ID) > 4096 || seen[a.ID.ID] || a.Type == "" || a.Name == "" || len(a.Name) > 1024 || strings.ContainsAny(a.ID.ID+a.Type+a.Name+a.ContentType+a.ContentID, "\x00\r\n") || a.Size != nil && *a.Size < 0 {
			return MessageAttachments{}, transport.EnvelopeError
		}
		seen[a.ID.ID] = true
		result.Attachments = append(result.Attachments, Metadata{ID: a.ID.ID, Type: a.Type, Name: a.Name, ContentType: a.ContentType, ContentID: a.ContentID, ReportedSize: a.Size, IsInline: a.IsInline})
	}
	return result, transport.EnvelopeOK
}
