package account

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"regexp"
	"strconv"
	"strings"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/webcapture"
)

type prefZone struct {
	Type string `json:"__type"`
	ID   string `json:"Id"`
}
type prefZoneContext struct {
	Type       string   `json:"__type"`
	Definition prefZone `json:"TimeZoneDefinition"`
}
type prefHeader struct {
	Type     string          `json:"__type"`
	Version  string          `json:"RequestServerVersion"`
	TimeZone prefZoneContext `json:"TimeZoneContext"`
}
type prefValue struct {
	TypeName string   `json:"__type"`
	Type     string   `json:"Type"`
	Value    []string `json:"Value"`
}
type prefEntry struct {
	Type  string    `json:"__type"`
	Key   prefValue `json:"DictionaryKey"`
	Value prefValue `json:"DictionaryValue"`
}
type prefFolder struct {
	Type string `json:"__type"`
	ID   string `json:"Id"`
}
type prefConfigName struct {
	Type   string     `json:"__type"`
	Folder prefFolder `json:"BaseFolderId"`
	Name   string     `json:"Name"`
}
type prefConfig struct {
	Type       string         `json:"__type"`
	Name       prefConfigName `json:"UserConfigurationName"`
	Dictionary []prefEntry    `json:"Dictionary"`
}
type prefBody struct {
	Type   string     `json:"__type"`
	Config prefConfig `json:"UserConfiguration"`
}
type focusedRequest struct {
	Type   string     `json:"__type"`
	Header prefHeader `json:"Header"`
	Body   prefBody   `json:"Body"`
}
type rulesRequest struct {
	Type                 string     `json:"__type"`
	Header               prefHeader `json:"Header"`
	UseServerRulesLoader bool       `json:"UseServerRulesLoader"`
}

func strictJSON(b []byte, target any) bool {
	d := json.NewDecoder(bytes.NewReader(b))
	d.DisallowUnknownFields()
	if d.Decode(target) != nil {
		return false
	}
	var trailing any
	return d.Decode(&trailing) == io.EOF
}
func validHeader(h prefHeader) bool {
	return h.Type == "JsonRequestHeaders:#Exchange" && h.Version == "V2018_01_08" && h.TimeZone.Type == "TimeZoneContext:#Exchange" && h.TimeZone.Definition.Type == "TimeZoneDefinitionType:#Exchange" && len(h.TimeZone.Definition.ID) > 0 && len(h.TimeZone.Definition.ID) <= 128 && !strings.ContainsAny(h.TimeZone.Definition.ID, "\x00\r\n")
}

var prefCounter = regexp.MustCompile(`^[0-9]{1,12}$`)

// PreparePreference reuses only observed rule and automatic-reply reads, or the
// single-key Focused Inbox write. Unknown fields and other settings fail closed.
func PreparePreference(kind, detail string, current transport.Session, focused bool) (transport.Session, transport.Operation, error) {
	fail := func() (transport.Session, transport.Operation, error) {
		return transport.Session{}, transport.Operation{}, fmt.Errorf("preference request does not match the observed mailbox contract")
	}
	action := map[string]string{"rules": "GetInboxRule", "auto-replies": "GetMailboxAutoReplyConfiguration", "focused-update": "UpdateUserConfiguration"}[kind]
	if action == "" {
		return fail()
	}
	d, err := webcapture.Load(detail, "", current)
	if err != nil {
		return fail()
	}
	q := d.URL.Query()
	if d.Method != "POST" || d.URL.Path != "/owa/service.svc" || q.Get("action") != action || q.Get("app") != "Mail" || !prefCounter.MatchString(q.Get("n")) || len(q) != 3 || d.Session.Headers.Get("Action") != action || !d.HeaderBody {
		return fail()
	}
	for _, v := range q {
		if len(v) != 1 {
			return fail()
		}
	}
	body := []byte(`{}`)
	class := transport.Read
	validator := validateAutoReplies
	switch kind {
	case "auto-replies":
		var x map[string]any
		if json.Unmarshal(d.Body, &x) != nil || x == nil || len(x) != 0 {
			return fail()
		}
	case "rules":
		var x rulesRequest
		if !strictJSON(d.Body, &x) || x.Type != "GetInboxRuleRequest:#Exchange" || !x.UseServerRulesLoader || !validHeader(x.Header) {
			return fail()
		}
		body, err = json.Marshal(x)
		if err != nil {
			return fail()
		}
		validator = validateRules
	case "focused-update":
		var x focusedRequest
		if !strictJSON(d.Body, &x) || x.Type != "UpdateUserConfigurationJsonRequest:#Exchange" || !validHeader(x.Header) || x.Body.Type != "UpdateUserConfigurationOwaRequest:#Exchange" || x.Body.Config.Type != "ServiceUserConfiguration:#Exchange" || x.Body.Config.Name.Type != "UserConfigurationNameType:#Exchange" || x.Body.Config.Name.Name != "OWA.UserOptions" || x.Body.Config.Name.Folder.Type != "DistinguishedFolderId:#Exchange" || x.Body.Config.Name.Folder.ID != "root" || len(x.Body.Config.Dictionary) != 1 {
			return fail()
		}
		e := &x.Body.Config.Dictionary[0]
		if e.Type != "UserConfigurationDictionaryEntry:#Exchange" || e.Key.TypeName != "UserConfigurationDictionaryObject:#Exchange" || e.Key.Type != "String" || len(e.Key.Value) != 1 || e.Key.Value[0] != "IsFocusedInboxEnabled" || e.Value.TypeName != "UserConfigurationDictionaryObject:#Exchange" || e.Value.Type != "Boolean" || len(e.Value.Value) != 1 || (e.Value.Value[0] != "true" && e.Value.Value[0] != "false") {
			return fail()
		}
		e.Value.Value[0] = strconv.FormatBool(focused)
		body, err = json.Marshal(x)
		if err != nil {
			return fail()
		}
		class = transport.Write
		validator = validateFocusedUpdate
	}
	return d.Session, d.Operation("POST", d.URL.Path, body, class, validator), nil
}

func ParseAutoReplies(data []byte) (json.RawMessage, transport.EnvelopeResult) {
	var x struct {
		Success   *bool           `json:"WasSuccessful"`
		ErrorCode *int            `json:"ErrorCode"`
		UserError *bool           `json:"IsUserError"`
		Options   json.RawMessage `json:"Options"`
	}
	if json.Unmarshal(data, &x) != nil || x.Success == nil || !*x.Success || x.ErrorCode == nil || *x.ErrorCode != 0 || x.UserError == nil || *x.UserError {
		return nil, transport.EnvelopeError
	}
	var opts map[string]json.RawMessage
	if json.Unmarshal(x.Options, &opts) != nil || opts == nil || len(opts["AutoReplyState"]) == 0 || len(opts["ExternalAudience"]) == 0 {
		return nil, transport.EnvelopeError
	}
	return append(json.RawMessage(nil), x.Options...), transport.EnvelopeOK
}
func validateAutoReplies(b []byte) transport.EnvelopeResult { _, s := ParseAutoReplies(b); return s }

func ParseRules(data []byte) ([]json.RawMessage, transport.EnvelopeResult) {
	var x struct {
		Success    *bool `json:"WasSuccessful"`
		ErrorCode  *int  `json:"ErrorCode"`
		UserError  *bool `json:"IsUserError"`
		Collection *struct {
			Rules json.RawMessage `json:"InboxRules"`
		} `json:"InboxRuleCollection"`
	}
	if json.Unmarshal(data, &x) != nil || x.Success == nil || !*x.Success || x.ErrorCode == nil || *x.ErrorCode != 0 || x.UserError == nil || *x.UserError || x.Collection == nil || len(x.Collection.Rules) == 0 {
		return nil, transport.EnvelopeError
	}
	items := []json.RawMessage{}
	if string(x.Collection.Rules) != "null" && json.Unmarshal(x.Collection.Rules, &items) != nil {
		return nil, transport.EnvelopeError
	}
	if len(items) > 10000 {
		return nil, transport.EnvelopeError
	}
	for _, v := range items {
		var obj map[string]json.RawMessage
		if json.Unmarshal(v, &obj) != nil || obj == nil {
			return nil, transport.EnvelopeError
		}
	}
	return items, transport.EnvelopeOK
}
func validateRules(b []byte) transport.EnvelopeResult { _, s := ParseRules(b); return s }
func validateFocusedUpdate(b []byte) transport.EnvelopeResult {
	var x struct {
		Body struct {
			Messages struct {
				Items []struct {
					Class string `json:"ResponseClass"`
					Code  string `json:"ResponseCode"`
				} `json:"Items"`
			} `json:"ResponseMessages"`
		} `json:"Body"`
	}
	if json.Unmarshal(b, &x) != nil || len(x.Body.Messages.Items) != 1 || x.Body.Messages.Items[0].Class != "Success" || x.Body.Messages.Items[0].Code != "NoError" {
		return transport.EnvelopeError
	}
	return transport.EnvelopeOK
}
