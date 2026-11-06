package capture

import (
	"crypto/rand"
	"encoding/base64"
	"encoding/json"
	"net/http"
	"net/url"
	"strings"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

type RuntimeWebTemplate struct {
	Method        string          `json:"method"`
	Path          string          `json:"path"`
	Query         url.Values      `json:"query"`
	HeaderCarrier bool            `json:"header_carrier"`
	Headers       http.Header     `json:"headers"`
	Body          json.RawMessage `json:"body,omitempty"`
}

var webCalendarActions = map[string]string{"calendar_get": "GetCalendarEvent", "calendar_create_self_meeting": "CreateCalendarEvent", "calendar_update_self_meeting": "UpdateCalendarEvent", "calendar_delete_self_meeting": "DeleteCalendarEvent", "calendar_create": "CreateCalendarEvent", "calendar_create_series": "CreateCalendarEvent", "calendar_create_all_day": "CreateCalendarEvent", "calendar_create_eastern": "CreateCalendarEvent", "calendar_update": "UpdateCalendarEvent", "calendar_update_fields": "UpdateCalendarEvent", "calendar_update_time": "UpdateCalendarEvent", "calendar_update_combined": "UpdateCalendarEvent", "calendar_update_all_day": "UpdateCalendarEvent", "calendar_update_eastern": "UpdateCalendarEvent", "calendar_update_series_subject": "UpdateCalendarEvent", "calendar_update_series_time": "UpdateCalendarEvent", "calendar_update_series_recurrence": "UpdateCalendarEvent", "calendar_update_occurrence": "UpdateCalendarEvent", "calendar_occurrences": "ExpandCalendarEvent", "calendar_delete": "DeleteCalendarEvent", "calendar_delete_series": "DeleteCalendarEvent"}
var webThreadActions = map[string]bool{"thread_mark_read": true, "thread_mark_unread": true, "thread_flag": true, "thread_complete": true, "thread_move": true, "thread_restore": true}
var webMailActions = map[string]string{"mail_move": "MoveItem", "mail_copy": "CopyItem", "mail_categories": "UpdateItem", "mail_mark_read": "UpdateItem", "mail_mark_unread": "UpdateItem", "mail_flag": "UpdateItem", "mail_unflag": "UpdateItem", "mail_trash": "DeleteItem"}
var webFolderActions = map[string]string{"folder_create": "CreateFolder", "folder_rename": "UpdateFolder", "folder_move": "MoveFolder", "folder_delete": "DeleteFolder"}

func validWebTemplate(name string, t RuntimeWebTemplate) bool {
	if len(t.Path) > 4096 || len(t.Query) > 8 {
		return false
	}
	for k, values := range t.Query {
		if len(values) != 1 || len(values[0]) > 8192 || strings.ContainsAny(k+values[0], "\x00\r\n") {
			return false
		}
	}
	if handled, valid := validContactListTemplate(name, t); handled {
		return valid
	}
	if handled, valid := validPreferenceTemplate(name, t); handled {
		return valid
	}
	if action := webCalendarActions[name]; action != "" {
		for k := range t.Query {
			if k != "action" && k != "app" && k != "n" && !((name == "calendar_delete" || name == "calendar_delete_series") && k == "UA") {
				return false
			}
		}
		return t.Method == "POST" && t.Path == "/owa/service.svc" && t.Query.Get("action") == action && t.Query.Get("app") == "Calendar" && digits.MatchString(t.Query.Get("n")) && t.HeaderCarrier == (name == "calendar_update_self_meeting" || name == "calendar_delete_self_meeting" || name == "calendar_update" || name == "calendar_delete" || name == "calendar_delete_series" || name == "calendar_update_occurrence" || name == "calendar_update_series_subject" || name == "calendar_occurrences") && json.Valid(t.Body)
	}
	if name == "calendar_availability" {
		for k := range t.Query {
			if k != "n" && k != "cv" {
				return false
			}
		}
		return t.Method == "POST" && t.Path == "/outlookgatewayb2/graphql" && digits.MatchString(t.Query.Get("n")) && t.Query.Get("cv") != "" && !t.HeaderCarrier && json.Valid(t.Body)
	}
	if name == "draft_remove_attachment" {
		for k := range t.Query {
			if k != "action" && k != "app" && k != "n" {
				return false
			}
		}
		return t.Method == "POST" && t.Path == "/owa/service.svc" && t.Query.Get("action") == "DeleteAttachment" && t.Query.Get("app") == "Mail" && digits.MatchString(t.Query.Get("n")) && t.HeaderCarrier && json.Valid(t.Body)
	}
	if name == "draft_attach" {
		return t.Method == "POST" && t.Path == "/owa/service.svc/CreateAttachmentFromLocalFile" && len(t.Query) == 0 && t.HeaderCarrier && json.Valid(t.Body)
	}
	if name == "account_aliases" {
		for k := range t.Query {
			if k != "action" && k != "app" && k != "n" {
				return false
			}
		}
		return t.Method == "POST" && t.Path == "/owa/service.svc" && t.Query.Get("action") == "GetOwaUserConfiguration" && (t.Query.Get("app") == "Mail" || t.Query.Get("app") == "People") && digits.MatchString(t.Query.Get("n")) && t.HeaderCarrier && string(t.Body) == "{}"
	}
	if action := webMailActions[name]; action != "" {
		for k := range t.Query {
			if k != "action" && k != "app" && k != "n" {
				return false
			}
		}
		return t.Method == "POST" && t.Path == "/owa/service.svc" && t.Query.Get("action") == action && t.Query.Get("app") == "Mail" && digits.MatchString(t.Query.Get("n")) && t.HeaderCarrier && json.Valid(t.Body)
	}
	if name == "folders_list" {
		for k := range t.Query {
			if k != "action" && k != "app" && k != "n" {
				return false
			}
		}
		return t.Method == "POST" && t.Path == "/owa/service.svc" && t.Query.Get("action") == "FindFolder" && t.Query.Get("app") == "Mail" && digits.MatchString(t.Query.Get("n")) && t.HeaderCarrier && json.Valid(t.Body)
	}
	if action := webFolderActions[name]; action != "" {
		for k := range t.Query {
			if k != "action" && k != "app" && k != "n" {
				return false
			}
		}
		return t.Method == "POST" && t.Path == "/owa/service.svc" && t.Query.Get("action") == action && t.Query.Get("app") == "Mail" && digits.MatchString(t.Query.Get("n")) && t.HeaderCarrier && json.Valid(t.Body)
	}
	if name == "contacts_list" {
		for k := range t.Query {
			if k != "$expand" && k != "$top" && k != "$select" && k != "$filter" && k != "n" && k != "cri" && k != "cv" {
				return false
			}
		}
		return t.Method == "GET" && t.Path == "/PeopleGraphVx/v1.0/people" && t.Query.Get("$top") == "30000" && t.Query.Get("$filter") == "graphRelationships/any(t: t/relationshipType eq 'x')" && t.Query.Get("$expand") != "" && t.Query.Get("$select") != "" && !t.HeaderCarrier && len(t.Body) == 0
	}
	if webThreadActions[name] {
		for k := range t.Query {
			if k != "action" && k != "app" && k != "n" {
				return false
			}
		}
		return t.Method == "POST" && t.Path == "/owa/service.svc" && t.Query.Get("action") == "ApplyConversationAction" && t.Query.Get("app") == "Mail" && digits.MatchString(t.Query.Get("n")) && t.HeaderCarrier && json.Valid(t.Body)
	}
	if name == "search_all" || name == "search_folder" {
		for k := range t.Query {
			if k != "n" && k != "cv" {
				return false
			}
		}
		return t.Method == "POST" && t.Path == "/searchservice/api/v2/query" && digits.MatchString(t.Query.Get("n")) && t.Query.Get("cv") != "" && !t.HeaderCarrier && json.Valid(t.Body)
	}
	if name == "contacts_search" {
		for k := range t.Query {
			if k != "n" && k != "cv" && k != "domain" {
				return false
			}
		}
		return t.Method == "POST" && t.Path == "/search/api/v1/suggestions" && digits.MatchString(t.Query.Get("n")) && t.Query.Get("cv") != "" && t.Query.Get("domain") == "People" && !t.HeaderCarrier && json.Valid(t.Body)
	}
	paths := map[string]struct{ method, path string }{
		"contacts_get":        {"POST", "/PeopleGraphVx/v1.0/peopleLookup"},
		"people_lookup_email": {"POST", "/PeopleGraphVx/v1.0/peopleLookup"},
		"contacts_create":     {"POST", "/PeopleGraphVx/v1.0/contacts"},
		"contacts_update":     {"POST", "/PeopleGraphVx/v1.0/contacts('template-id')/UpdateCollections"},
		"contacts_delete":     {"DELETE", "/PeopleGraphVx/v1.0/contacts('template-id')"},
	}
	expected, ok := paths[name]
	if !ok {
		return false
	}
	for k := range t.Query {
		if k != "$expand" && k != "$select" && k != "cri" && k != "cv" && k != "n" {
			return false
		}
	}
	return t.Method == expected.method && t.Path == expected.path && !t.HeaderCarrier && (name == "contacts_delete" && (len(t.Body) == 0 || string(t.Body) == "null") || name != "contacts_delete" && json.Valid(t.Body))
}

// AddWeb stores a sanitized, fixed-path operation without credentials.
func (b *RuntimeBundle) AddWeb(name string, session transport.Session, op transport.Operation) error {
	if b == nil || session.AccountKey != b.AccountKey || session.Origin != "https://outlook.cloud.microsoft" {
		return &transport.Error{Kind: transport.KindAuth}
	}
	aud, ok := jwtAudience(session.Headers.Get("Authorization"))
	if !ok || digest(aud) != b.AudienceHash || digest(session.Headers.Get("X-AnchorMailbox")) != b.AnchorHash {
		return &transport.Error{Kind: transport.KindAuth}
	}
	body := op.Body
	header := op.Headers.Get("X-Owa-Urlpostdata") != ""
	if header {
		decoded, err := url.QueryUnescape(op.Headers.Get("X-Owa-Urlpostdata"))
		if err != nil {
			return &transport.Error{Kind: transport.KindInput}
		}
		body = []byte(decoded)
	}
	t := RuntimeWebTemplate{Method: op.Method, Path: op.Path, Query: op.Query, HeaderCarrier: header, Headers: make(http.Header), Body: append(json.RawMessage(nil), body...)}
	for k := range runtimeHeaders {
		if v := session.Headers.Get(k); v != "" {
			t.Headers.Set(k, v)
		}
	}
	if !validWebTemplate(name, t) {
		return &transport.Error{Kind: transport.KindInput}
	}
	if b.WebOperations == nil {
		b.WebOperations = map[string]RuntimeWebTemplate{}
	}
	if _, exists := b.WebOperations[name]; exists {
		return &transport.Error{Kind: transport.KindInput}
	}
	b.WebOperations[name] = t
	return nil
}

func LoadRuntimeWeb(name string, current transport.Session) (transport.Session, RuntimeWebTemplate, error) {
	fail := func() (transport.Session, RuntimeWebTemplate, error) {
		return transport.Session{}, RuntimeWebTemplate{}, &transport.Error{Kind: transport.KindAuth}
	}
	b, err := ReadBoundRuntimeBundle(current)
	if err != nil {
		return fail()
	}
	t, ok := b.WebOperations[name]
	if !ok || !validWebTemplate(name, t) {
		return fail()
	}
	if name == "calendar_availability" {
		var nonce [16]byte
		if _, err := rand.Read(nonce[:]); err != nil {
			return fail()
		}
		t.Query.Set("cv", base64.RawStdEncoding.EncodeToString(nonce[:])+".0")
	}
	current.Headers = current.Headers.Clone()
	current.Headers.Del("X-Owa-Urlpostdata")
	current.Headers.Del("Action")
	current.Headers.Del("X-Owa-Canary")
	if name == "draft_remove_attachment" {
		current.Headers.Set("Action", "DeleteAttachment")
	}
	if action := webCalendarActions[name]; action != "" {
		current.Headers.Set("Action", action)
	}
	if webThreadActions[name] {
		current.Headers.Set("Action", "ApplyConversationAction")
	}
	if name == "account_aliases" {
		current.Headers.Set("Action", "GetOwaUserConfiguration")
	}
	if name == "folders_list" {
		current.Headers.Set("Action", "FindFolder")
	}
	if action := webMailActions[name]; action != "" {
		current.Headers.Set("Action", action)
	}
	if action := webFolderActions[name]; action != "" {
		current.Headers.Set("Action", action)
	}
	if action := webPreferenceActions[name]; action != "" {
		current.Headers.Set("Action", action)
	}
	for k, values := range t.Headers {
		canonical := http.CanonicalHeaderKey(k)
		if !runtimeHeaders[canonical] || len(values) != 1 || strings.ContainsAny(values[0], "\r\n\x00") {
			return fail()
		}
		current.Headers.Set(canonical, values[0])
	}
	return current, t, nil
}
