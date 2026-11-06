package drafts

import (
	"encoding/json"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

type Created struct {
	ID        string `json:"id"`
	ChangeKey string `json:"change_key"`
	IsDraft   bool   `json:"is_draft"`
}

// Draft is the returned server representation. Body may be truncated by the
// observed GetItem shape; callers must not treat it as a full export.
type Draft struct {
	ID             string          `json:"id"`
	ChangeKey      string          `json:"change_key"`
	IsDraft        bool            `json:"is_draft"`
	Subject        string          `json:"subject"`
	ParentFolderID string          `json:"parent_folder_id,omitempty"`
	HasAttachments *bool           `json:"has_attachments,omitempty"`
	Body           json.RawMessage `json:"normalized_body,omitempty"`
}

type itemMessage struct {
	ResponseClass string `json:"ResponseClass"`
	ResponseCode  string `json:"ResponseCode"`
	Items         []struct {
		Type   string `json:"__type"`
		ItemID struct {
			ID        string `json:"Id"`
			ChangeKey string `json:"ChangeKey"`
		} `json:"ItemId"`
		ParentFolderID struct {
			ID string `json:"Id"`
		} `json:"ParentFolderId"`
		HasAttachments *bool           `json:"HasAttachments"`
		IsDraft        *bool           `json:"IsDraft"`
		Subject        string          `json:"Subject"`
		NormalizedBody json.RawMessage `json:"NormalizedBody"`
	} `json:"Items"`
	MovedItemID struct {
		ID string `json:"Id"`
	} `json:"MovedItemId"`
}

func singleMessage(data []byte) (itemMessage, bool) {
	var root struct {
		Body *struct {
			ResponseMessages *struct {
				Items []itemMessage `json:"Items"`
			} `json:"ResponseMessages"`
		} `json:"Body"`
	}
	if json.Unmarshal(data, &root) != nil || root.Body == nil || root.Body.ResponseMessages == nil || len(root.Body.ResponseMessages.Items) != 1 {
		return itemMessage{}, false
	}
	r := root.Body.ResponseMessages.Items[0]
	return r, r.ResponseClass == "Success" && r.ResponseCode == "NoError"
}

func ParseGet(data []byte) (Draft, transport.EnvelopeResult) {
	r, ok := singleMessage(data)
	if !ok || len(r.Items) != 1 {
		return Draft{}, transport.EnvelopeError
	}
	i := r.Items[0]
	if i.Type != "Message:#Exchange" || i.IsDraft == nil || !*i.IsDraft || i.ItemID.ID == "" || i.ItemID.ChangeKey == "" {
		return Draft{}, transport.EnvelopeError
	}
	return Draft{ID: i.ItemID.ID, ChangeKey: i.ItemID.ChangeKey, IsDraft: true, Subject: i.Subject, ParentFolderID: i.ParentFolderID.ID, HasAttachments: i.HasAttachments, Body: i.NormalizedBody}, transport.EnvelopeOK
}

func ParseUpdate(data []byte) (Created, transport.EnvelopeResult) {
	r, ok := singleMessage(data)
	if !ok && r.ResponseClass == "Error" && r.ResponseCode != "" {
		switch r.ResponseCode {
		case "ErrorIrresolvableConflict", "ErrorChangeKeyRequired", "ErrorInvalidChangeKey", "ErrorStaleObject":
			return Created{}, transport.EnvelopeConflict
		default:
			return Created{}, transport.EnvelopeRejected
		}
	}
	if !ok || len(r.Items) != 1 || r.Items[0].Type != "Message:#Exchange" || r.Items[0].ItemID.ID == "" || r.Items[0].ItemID.ChangeKey == "" {
		return Created{}, transport.EnvelopeError
	}
	return Created{ID: r.Items[0].ItemID.ID, ChangeKey: r.Items[0].ItemID.ChangeKey, IsDraft: true}, transport.EnvelopeOK
}

func ParseDiscard(data []byte) transport.EnvelopeResult {
	r, ok := singleMessage(data)
	if !ok && r.ResponseClass == "Error" && r.ResponseCode != "" {
		switch r.ResponseCode {
		case "ErrorIrresolvableConflict", "ErrorChangeKeyRequired", "ErrorInvalidChangeKey", "ErrorStaleObject":
			return transport.EnvelopeConflict
		default:
			return transport.EnvelopeRejected
		}
	}
	if !ok || r.MovedItemID.ID == "" {
		return transport.EnvelopeError
	}
	return transport.EnvelopeOK
}

// ParseSend confirms only that the server accepted the send request; it does
// not imply that any recipient received the message.
func ParseSend(data []byte) transport.EnvelopeResult {
	r, ok := singleMessage(data)
	if !ok && r.ResponseClass == "Error" && r.ResponseCode != "" {
		switch r.ResponseCode {
		case "ErrorIrresolvableConflict", "ErrorChangeKeyRequired", "ErrorInvalidChangeKey", "ErrorStaleObject":
			return transport.EnvelopeConflict
		default:
			return transport.EnvelopeRejected
		}
	}
	if !ok || len(r.Items) != 0 {
		return transport.EnvelopeError
	}
	return transport.EnvelopeOK
}

func ParseCreate(data []byte) (Created, transport.EnvelopeResult) {
	var root struct {
		Body *struct {
			ResponseMessages *struct {
				Items []struct {
					ResponseClass string `json:"ResponseClass"`
					ResponseCode  string `json:"ResponseCode"`
					Items         []struct {
						ItemID struct {
							ID        string `json:"Id"`
							ChangeKey string `json:"ChangeKey"`
						} `json:"ItemId"`
						IsDraft *bool `json:"IsDraft"`
					} `json:"Items"`
				} `json:"Items"`
			} `json:"ResponseMessages"`
		} `json:"Body"`
	}
	if json.Unmarshal(data, &root) != nil || root.Body == nil || root.Body.ResponseMessages == nil || len(root.Body.ResponseMessages.Items) != 1 {
		return Created{}, transport.EnvelopeError
	}
	r := root.Body.ResponseMessages.Items[0]
	if r.ResponseClass != "Success" || r.ResponseCode != "NoError" || len(r.Items) != 1 || r.Items[0].IsDraft == nil || !*r.Items[0].IsDraft || r.Items[0].ItemID.ID == "" || r.Items[0].ItemID.ChangeKey == "" {
		return Created{}, transport.EnvelopeError
	}
	return Created{ID: r.Items[0].ItemID.ID, ChangeKey: r.Items[0].ItemID.ChangeKey, IsDraft: true}, transport.EnvelopeOK
}
