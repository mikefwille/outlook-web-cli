// Package rules implements only the captured narrow personal inbox-rule shapes.
package rules

import (
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"io"
	"regexp"
	"strings"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/webcapture"
)

type zone struct {
	Type string `json:"__type"`
	ID   string `json:"Id"`
}
type zoneContext struct {
	Type       string `json:"__type"`
	Definition zone   `json:"TimeZoneDefinition"`
}
type header struct {
	Type    string      `json:"__type"`
	Version string      `json:"RequestServerVersion"`
	Zone    zoneContext `json:"TimeZoneContext"`
}
type Identity struct {
	DisplayName string `json:"DisplayName"`
	RawIdentity string `json:"RawIdentity"`
}
type narrowRule struct {
	Name     string    `json:"Name"`
	Subject  []string  `json:"SubjectContainsWords"`
	Category []string  `json:"ApplyCategory"`
	Stop     *bool     `json:"StopProcessingRules"`
	Identity *Identity `json:"Identity,omitempty"`
}
type ruleRequest struct {
	Type     string      `json:"__type"`
	Header   header      `json:"Header"`
	Rule     *narrowRule `json:"InboxRule,omitempty"`
	Identity *Identity   `json:"Identity,omitempty"`
	Force    *bool       `json:"Force,omitempty"`
}
type Params struct{ ID, IdentityDisplay, Name, Subject, Category string }

func validText(s string, max int) bool {
	return len(s) > 0 && len(s) <= max && strings.TrimSpace(s) == s && !strings.ContainsAny(s, "\x00\r\n")
}
func strict(b []byte, v any) bool {
	d := json.NewDecoder(bytes.NewReader(b))
	d.DisallowUnknownFields()
	if d.Decode(v) != nil {
		return false
	}
	var extra any
	return d.Decode(&extra) == io.EOF
}
func validHeader(h header) bool {
	return h.Type == "JsonRequestHeaders:#Exchange" && h.Version == "V2018_01_08" && h.Zone.Type == "TimeZoneContext:#Exchange" && h.Zone.Definition.Type == "TimeZoneDefinitionType:#Exchange" && validText(h.Zone.Definition.ID, 128)
}
func validIdentity(i *Identity) bool {
	return i != nil && validText(i.DisplayName, 4096) && validText(i.RawIdentity, 4096)
}
func validNarrow(r *narrowRule) bool {
	return r != nil && validText(r.Name, 255) && len(r.Subject) == 1 && validText(r.Subject[0], 255) && len(r.Category) == 1 && validText(r.Category[0], 255) && r.Stop != nil && !*r.Stop
}

var number = regexp.MustCompile(`^[0-9]{1,12}$`)

func fail() error {
	return fmt.Errorf("inbox rule request does not match the observed narrow rule contract")
}

// Prepare replaces fixture identities and fields. Creation is enabled by the
// service; requiring explicit acknowledgement belongs to the CLI caller.
func Prepare(kind, detail string, current transport.Session, p Params) (transport.Session, transport.Operation, error) {
	bad := func() (transport.Session, transport.Operation, error) {
		return transport.Session{}, transport.Operation{}, fail()
	}
	action := map[string]string{"create": "NewInboxRule", "enable": "EnableInboxRule", "disable": "DisableInboxRule", "update": "SetInboxRule", "delete": "RemoveInboxRule"}[kind]
	if action == "" {
		return bad()
	}
	d, err := webcapture.Load(detail, "", current)
	if err != nil {
		return bad()
	}
	q := d.URL.Query()
	if d.Method != "POST" || d.URL.Path != "/owa/service.svc" || !d.HeaderBody || d.Session.Headers.Get("Action") != action || len(q) != 3 || q.Get("action") != action || q.Get("app") != "Mail" || !number.MatchString(q.Get("n")) {
		return bad()
	}
	for _, v := range q {
		if len(v) != 1 {
			return bad()
		}
	}
	var x ruleRequest
	if !strict(d.Body, &x) || x.Type != action+"Request:#Exchange" || !validHeader(x.Header) {
		return bad()
	}
	switch kind {
	case "create":
		if !validNarrow(x.Rule) || x.Rule.Identity != nil || x.Identity != nil || x.Force != nil || !validText(p.Name, 255) || !validText(p.Subject, 255) || !validText(p.Category, 255) {
			return bad()
		}
		x.Rule.Name = p.Name
		x.Rule.Subject = []string{p.Subject}
		x.Rule.Category = []string{p.Category}
	case "update":
		if !validNarrow(x.Rule) || !validIdentity(x.Rule.Identity) || x.Identity != nil || x.Force == nil || *x.Force || !validText(p.ID, 4096) || !validText(p.Name, 255) || !validText(p.Subject, 255) || !validText(p.Category, 255) {
			return bad()
		}
		x.Rule.Name = p.Name
		x.Rule.Subject = []string{p.Subject}
		x.Rule.Category = []string{p.Category}
		x.Rule.Identity = &Identity{DisplayName: p.Name, RawIdentity: p.ID}
	case "enable", "disable", "delete":
		if x.Rule != nil || x.Force != nil || !validIdentity(x.Identity) || !validText(p.ID, 4096) || !validText(p.IdentityDisplay, 4096) {
			return bad()
		}
		x.Identity = &Identity{DisplayName: p.IdentityDisplay, RawIdentity: p.ID}
	}
	body, err := json.Marshal(x)
	if err != nil {
		return bad()
	}
	return d.Session, d.Operation("POST", d.URL.Path, body, transport.Write, func(b []byte) transport.EnvelopeResult {
		var response struct {
			Success   *bool           `json:"WasSuccessful"`
			ErrorCode *int            `json:"ErrorCode"`
			UserError *bool           `json:"IsUserError"`
			Rule      json.RawMessage `json:"InboxRule"`
		}
		if json.Unmarshal(b, &response) != nil || response.Success == nil || !*response.Success || response.ErrorCode == nil || *response.ErrorCode != 0 || response.UserError == nil || *response.UserError {
			return transport.EnvelopeError
		}
		if kind == "create" {
			rule, err := Parse(response.Rule)
			if err != nil || rule.Name != p.Name || !rule.Enabled || len(rule.Subject) != 1 || rule.Subject[0] != p.Subject || len(rule.Category) != 1 || rule.Category[0] != p.Category {
				return transport.EnvelopeError
			}
		}
		return transport.EnvelopeOK
	}), nil
}

type Rule struct {
	ID              string          `json:"id"`
	IdentityDisplay string          `json:"identity_display"`
	Name            string          `json:"name"`
	Enabled         bool            `json:"enabled"`
	Version         string          `json:"version"`
	Subject         []string        `json:"subject_contains"`
	Category        []string        `json:"apply_category"`
	Raw             json.RawMessage `json:"raw"`
}

func Parse(raw []byte) (Rule, error) {
	var x struct {
		Type     string    `json:"__type"`
		Identity *Identity `json:"Identity"`
		Name     string    `json:"Name"`
		Enabled  *bool     `json:"Enabled"`
		Subject  []string  `json:"SubjectContainsWords"`
		Category []string  `json:"ApplyCategory"`
	}
	if json.Unmarshal(raw, &x) != nil || x.Type != "InboxRule:#Exchange" || !validIdentity(x.Identity) || !validText(x.Name, 255) || x.Enabled == nil {
		return Rule{}, fail()
	}
	var obj any
	decoder := json.NewDecoder(bytes.NewReader(raw))
	decoder.UseNumber()
	if decoder.Decode(&obj) != nil {
		return Rule{}, fail()
	}
	canonical, err := json.Marshal(obj)
	if err != nil {
		return Rule{}, fail()
	}
	sum := sha256.Sum256(canonical)
	return Rule{ID: x.Identity.RawIdentity, IdentityDisplay: x.Identity.DisplayName, Name: x.Name, Enabled: *x.Enabled, Version: hex.EncodeToString(sum[:]), Subject: x.Subject, Category: x.Category, Raw: append(json.RawMessage(nil), raw...)}, nil
}

// SafeToRename rejects unobserved/extra conditions, actions and exceptions.
// The captured SetInboxRule sends only name, one subject, one category and stop.
func (r Rule) SafeToRename() bool {
	if r.Enabled || len(r.Subject) != 1 || !validText(r.Subject[0], 255) || len(r.Category) != 1 || !validText(r.Category[0], 255) {
		return false
	}
	var x map[string]json.RawMessage
	if json.Unmarshal(r.Raw, &x) != nil {
		return false
	}
	metadata := map[string]bool{"__type": true, "Description": true, "DescriptionTimeFormat": true, "DescriptionTimeZone": true, "RuleProvider": true, "WarningMessages": true, "Identity": true, "Name": true, "Priority": true, "SubjectContainsWords": true, "ApplyCategory": true}
	defaults := map[string]string{
		"Enabled": "false", "InError": "false", "ErrorType": "0", "SupportedByTask": "true", "StopProcessingRules": "false",
		"AnyCategory": "false", "ApplySystemCategory": "null", "AssignedCategories": "null", "ClearCategories": "false", "ClearFlag": "false", "CompleteFlag": "false", "BodyContainsWords": "null", "BulkCategory": "null", "CopyToFolder": "null", "DeleteMessage": "false", "DeleteSystemCategory": "null",
		"ExceptIfAnyCategory": "false", "ExceptIfAssignedCategories": "null", "ExceptIfBodyContainsWords": "null", "ExceptIfBulkCategory": "null", "ExceptIfFlaggedForAction": `"NullInboxRuleMessageFlag"`, "ExceptIfRequestedAction": `"NullInboxRuleMessageFlag"`, "ExceptIfFrom": "null", "ExceptIfFromAddressContainsWords": "null", "ExceptIfHasAttachment": "false", "ExceptIfHasClassification": "null", "ExceptIfIsExternal": "false", "ExceptIfIsMentioned": "false", "ExceptIfHeaderContainsWords": "null", "ExceptIfMessageTypeMatches": `"NullInboxRuleMessageType"`, "ExceptIfMyNameInCcBox": "false", "ExceptIfMyNameInToBox": "false", "ExceptIfMyNameInToOrCcBox": "false", "ExceptIfMyNameNotInToBox": "false", "ExceptIfReceivedAfterDate": "null", "ExceptIfReceivedBeforeDate": "null", "ExceptIfRecipientAddressContainsWords": "null", "ExceptIfSentOnlyToMe": "false", "ExceptIfSentTo": "null", "ExceptIfSubjectContainsWords": "null", "ExceptIfSubjectOrBodyContainsWords": "null", "ExceptIfWithImportance": `"NullImportance"`, "ExceptIfWithSensitivity": `"NullSensitivity"`,
		"FlaggedForAction": `"NullInboxRuleMessageFlag"`, "FlagMessage": "null", "ForwardAsAttachmentTo": "null", "ForwardTo": "null", "From": "null", "FromAddressContainsWords": "null", "HasAttachment": "false", "HasClassification": "null", "HeaderContainsWords": "null", "IsExternal": "false", "IsMentioned": "false", "MarkAsRead": "false", "MarkImportance": `"NullImportance"`, "MessageTypeMatches": `"NullInboxRuleMessageType"`, "MyNameInCcBox": "false", "MyNameInToBox": "false", "MyNameInToOrCcBox": "false", "MyNameNotInToBox": "false", "MoveToFolder": "null", "PermanentDelete": "false", "PinMessage": "false", "AutoReplyWithTemplate": "null", "ReceivedAfterDate": "null", "ReceivedBeforeDate": "null", "RecipientAddressContainsWords": "null", "RedirectTo": "null", "RequestedAction": `"NullInboxRuleMessageFlag"`, "SendTextMessageNotificationTo": "null", "SentOnlyToMe": "false", "SentTo": "null", "SoftDeleteMessage": "false", "SubjectOrBodyContainsWords": "null", "WithImportance": `"NullImportance"`, "WithSensitivity": `"NullSensitivity"`, "DisplayAlert": "null", "PlaySound": "null",
	}
	for key, value := range x {
		if metadata[key] {
			continue
		}
		expected, ok := defaults[key]
		if !ok || strings.TrimSpace(string(value)) != expected {
			return false
		}
	}
	for key := range defaults {
		if _, ok := x[key]; !ok {
			return false
		}
	}
	return true
}
