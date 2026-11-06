package exchange

import (
	"encoding/json"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

// ParseMessageGet retains the observed item, including any body truncation
// indicators, and rejects a success envelope without a single message.
func ParseMessageGet(data []byte) (string, json.RawMessage, transport.EnvelopeResult) {
	var root struct {
		Body *struct {
			ResponseMessages *struct {
				Items []struct {
					ResponseClass string            `json:"ResponseClass"`
					ResponseCode  string            `json:"ResponseCode"`
					Items         []json.RawMessage `json:"Items"`
				} `json:"Items"`
			} `json:"ResponseMessages"`
		} `json:"Body"`
	}
	if json.Unmarshal(data, &root) != nil || root.Body == nil || root.Body.ResponseMessages == nil || len(root.Body.ResponseMessages.Items) != 1 {
		return "", nil, transport.EnvelopeError
	}
	r := root.Body.ResponseMessages.Items[0]
	if r.ResponseClass != "Success" || r.ResponseCode != "NoError" || len(r.Items) != 1 {
		return "", nil, transport.EnvelopeError
	}
	var item struct {
		Type   string `json:"__type"`
		ItemID struct {
			ID string `json:"Id"`
		} `json:"ItemId"`
	}
	if json.Unmarshal(r.Items[0], &item) != nil || item.Type != "Message:#Exchange" || item.ItemID.ID == "" {
		return "", nil, transport.EnvelopeError
	}
	return item.ItemID.ID, r.Items[0], transport.EnvelopeOK
}

func ValidateMessageGet(data []byte) transport.EnvelopeResult {
	_, _, result := ParseMessageGet(data)
	return result
}
