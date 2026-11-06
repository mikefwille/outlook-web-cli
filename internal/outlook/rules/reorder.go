package rules

import (
	"encoding/json"
	"fmt"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/webcapture"
)

type orderEntry struct {
	Type     string    `json:"__type"`
	Header   header    `json:"Header"`
	Identity *Identity `json:"Identity"`
	Enabled  *bool     `json:"IsEnabled"`
}

type orderRequest struct {
	Type   string       `json:"__type"`
	Header header       `json:"Header"`
	Rules  []orderEntry `json:"EnableDisableInboxRules"`
}

type OrderMember struct {
	ID              string
	IdentityDisplay string
	Enabled         bool
}

// PrepareReorder replays the observed list-shaped request, replacing every
// fixture identity and retaining each freshly read enabled state. Because the
// service accepts a list of states, ordering and state reapplication are one
// non-atomic write, not an order-only operation.
func PrepareReorder(detail string, current transport.Session, members []OrderMember) (transport.Session, transport.Operation, error) {
	bad := func() (transport.Session, transport.Operation, error) {
		return transport.Session{}, transport.Operation{}, fmt.Errorf("rule ordering request does not match the observed mailbox contract")
	}
	d, err := webcapture.Load(detail, "", current)
	if err != nil {
		return bad()
	}
	q := d.URL.Query()
	if d.Method != "POST" || d.URL.Path != "/owa/service.svc" || !d.HeaderBody || d.Session.Headers.Get("Action") != "SetInboxAndSweepRules" || len(q) != 3 || q.Get("action") != "SetInboxAndSweepRules" || q.Get("app") != "Mail" || !number.MatchString(q.Get("n")) {
		return bad()
	}
	for _, v := range q {
		if len(v) != 1 {
			return bad()
		}
	}
	var request orderRequest
	if !strict(d.Body, &request) || request.Type != "SetInboxAndSweepRulesRequest:#Exchange" || !validHeader(request.Header) || len(request.Rules) != 2 {
		return bad()
	}
	for _, entry := range request.Rules {
		if entry.Type != "EnableDisableInboxRuleRequest:#Exchange" || !validHeader(entry.Header) || entry.Header != request.Header || !validIdentity(entry.Identity) || entry.Enabled == nil {
			return bad()
		}
	}
	if len(members) < 2 || len(members) > 1000 {
		return bad()
	}
	seen := map[string]bool{}
	request.Rules = make([]orderEntry, 0, len(members))
	for _, member := range members {
		if !validText(member.ID, 4096) || !validText(member.IdentityDisplay, 4096) || seen[member.ID] {
			return bad()
		}
		seen[member.ID] = true
		enabled := member.Enabled
		request.Rules = append(request.Rules, orderEntry{Type: "EnableDisableInboxRuleRequest:#Exchange", Header: request.Header, Identity: &Identity{DisplayName: member.IdentityDisplay, RawIdentity: member.ID}, Enabled: &enabled})
	}
	body, err := json.Marshal(request)
	if err != nil {
		return bad()
	}
	return d.Session, d.Operation("POST", d.URL.Path, body, transport.Write, func(b []byte) transport.EnvelopeResult {
		var response struct {
			Success   *bool `json:"WasSuccessful"`
			ErrorCode *int  `json:"ErrorCode"`
			UserError *bool `json:"IsUserError"`
		}
		if json.Unmarshal(b, &response) != nil || response.Success == nil || !*response.Success || response.ErrorCode == nil || *response.ErrorCode != 0 || response.UserError == nil || *response.UserError {
			return transport.EnvelopeError
		}
		return transport.EnvelopeOK
	}), nil
}
