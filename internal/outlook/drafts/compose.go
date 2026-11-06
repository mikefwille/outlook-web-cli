package drafts

import (
	"encoding/json"
	"strings"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

type ComposeRecipient struct {
	Name        string `json:"name"`
	Address     string `json:"address"`
	RoutingType string `json:"routing_type"`
}

type ComposeSnapshot struct {
	ID            string             `json:"id"`
	ChangeKey     string             `json:"change_key"`
	Subject       string             `json:"subject"`
	BodyType      string             `json:"body_type"`
	Body          string             `json:"body"`
	BodyTruncated bool               `json:"body_truncated"`
	To            []ComposeRecipient `json:"to"`
	Cc            []ComposeRecipient `json:"cc"`
	Bcc           []ComposeRecipient `json:"bcc"`
}

// ParseCompose verifies the captured MailCompose shape and returns only the
// fields needed for an exact recipient and body preview. Missing or truncated
// bodies, unresolved recipients and non-draft items fail closed.
func ParseCompose(data []byte) (ComposeSnapshot, transport.EnvelopeResult) {
	var root struct {
		Body *struct {
			ResponseMessages *struct {
				Items []struct {
					ResponseClass string `json:"ResponseClass"`
					ResponseCode  string `json:"ResponseCode"`
					Items         []struct {
						Type   string `json:"__type"`
						ItemID struct {
							ID        string `json:"Id"`
							ChangeKey string `json:"ChangeKey"`
						} `json:"ItemId"`
						IsDraft *bool  `json:"IsDraft"`
						Subject string `json:"Subject"`
						Body    *struct {
							BodyType    string `json:"BodyType"`
							Value       string `json:"Value"`
							IsTruncated *bool  `json:"IsTruncated"`
						} `json:"Body"`
						To []struct {
							Name         string `json:"Name"`
							EmailAddress string `json:"EmailAddress"`
							RoutingType  string `json:"RoutingType"`
						} `json:"ToRecipients"`
						Cc []struct {
							Name         string `json:"Name"`
							EmailAddress string `json:"EmailAddress"`
							RoutingType  string `json:"RoutingType"`
						} `json:"CcRecipients"`
						Bcc []struct {
							Name         string `json:"Name"`
							EmailAddress string `json:"EmailAddress"`
							RoutingType  string `json:"RoutingType"`
						} `json:"BccRecipients"`
					} `json:"Items"`
				} `json:"Items"`
			} `json:"ResponseMessages"`
		} `json:"Body"`
	}
	if json.Unmarshal(data, &root) != nil || root.Body == nil || root.Body.ResponseMessages == nil || len(root.Body.ResponseMessages.Items) != 1 {
		return ComposeSnapshot{}, transport.EnvelopeError
	}
	r := root.Body.ResponseMessages.Items[0]
	if r.ResponseClass != "Success" || r.ResponseCode != "NoError" || len(r.Items) != 1 {
		return ComposeSnapshot{}, transport.EnvelopeError
	}
	i := r.Items[0]
	if i.Type != "Message:#Exchange" || i.ItemID.ID == "" || i.ItemID.ChangeKey == "" || i.IsDraft == nil || !*i.IsDraft || i.Body == nil || i.Body.IsTruncated == nil || *i.Body.IsTruncated || (i.Body.BodyType != "HTML" && i.Body.BodyType != "Text") {
		return ComposeSnapshot{}, transport.EnvelopeError
	}
	if len(i.To)+len(i.Cc)+len(i.Bcc) == 0 || len(i.To)+len(i.Cc)+len(i.Bcc) > 500 {
		return ComposeSnapshot{}, transport.EnvelopeError
	}
	result := ComposeSnapshot{ID: i.ItemID.ID, ChangeKey: i.ItemID.ChangeKey, Subject: i.Subject, BodyType: i.Body.BodyType, Body: i.Body.Value, BodyTruncated: false, To: []ComposeRecipient{}, Cc: []ComposeRecipient{}, Bcc: []ComposeRecipient{}}
	appendRecipient := func(name, address, routing string, out *[]ComposeRecipient) bool {
		if address == "" || strings.ContainsAny(name+address+routing, "\r\n\x00") || !strings.EqualFold(routing, "SMTP") || !strings.Contains(address, "@") {
			return false
		}
		*out = append(*out, ComposeRecipient{Name: name, Address: address, RoutingType: routing})
		return true
	}
	for _, r := range i.To {
		if !appendRecipient(r.Name, r.EmailAddress, r.RoutingType, &result.To) {
			return ComposeSnapshot{}, transport.EnvelopeError
		}
	}
	for _, r := range i.Cc {
		if !appendRecipient(r.Name, r.EmailAddress, r.RoutingType, &result.Cc) {
			return ComposeSnapshot{}, transport.EnvelopeError
		}
	}
	for _, r := range i.Bcc {
		if !appendRecipient(r.Name, r.EmailAddress, r.RoutingType, &result.Bcc) {
			return ComposeSnapshot{}, transport.EnvelopeError
		}
	}
	return result, transport.EnvelopeOK
}
