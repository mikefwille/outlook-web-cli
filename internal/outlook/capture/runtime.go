package capture

import (
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/drafts"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/exchange"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

// RuntimeBundle stores account-scoped, credential-free operation templates.
// It is private configuration, not raw network evidence or a token store.
type RuntimeBundle struct {
	Version       int                           `json:"version"`
	AccountKey    string                        `json:"account_key"`
	AudienceHash  string                        `json:"audience_hash"`
	AnchorHash    string                        `json:"anchor_hash"`
	Operations    map[string]RuntimeTemplate    `json:"operations"`
	WebOperations map[string]RuntimeWebTemplate `json:"web_operations,omitempty"`
}

type RuntimeTemplate struct {
	Action        string          `json:"action"`
	Query         url.Values      `json:"query"`
	HeaderCarrier bool            `json:"header_carrier"`
	Headers       http.Header     `json:"headers"`
	Body          json.RawMessage `json:"body"`
}

var runtimeActions = map[string]string{
	"draft_create": "CreateItem", "draft_get": "GetItem", "mail_get": "GetItem", "draft_compose": "GetItem",
	"draft_update": "UpdateItem", "draft_discard": "DeleteItem", "draft_send": "UpdateItem",
	"reply_create": "CreateItem", "reply_all_create": "CreateItem", "forward_create": "CreateItem",
	"threads_list": "FindConversation", "threads_get": "GetConversationItems", "mail_list": "FindItem",
	"calendars_list": "GetCalendarFoldersV2", "calendar_agenda": "GetCalendarView", "account_info": "GetAccountInformation",
}

func runtimeHeaderCarrier(name string) bool {
	switch name {
	case "draft_discard", "threads_list", "mail_list", "calendars_list", "calendar_agenda", "account_info":
		return true
	}
	return false
}

var runtimeHeaders = map[string]bool{
	"Accept": true, "Content-Type": true, "Prefer": true, "X-Owa-Actionsource": true,
	"X-Owa-Hosted-Ux": true, "X-Owa-Priority": true, "X-Req-Source": true, "X-Client-Version": true,
	"Owaappid": true, "Scenariotag": true, "Workload-Type": true, "X-Ms-Appname": true, "X-Search-Griffin-Version": true, "X-Client-Flights": true,
	"X-Owa-Action": true,
}

func digest(s string) string { h := sha256.Sum256([]byte(s)); return hex.EncodeToString(h[:]) }

func NewRuntimeBundle(current transport.Session) (RuntimeBundle, error) {
	aud, ok := jwtAudience(current.Headers.Get("Authorization"))
	anchor := current.Headers.Get("X-AnchorMailbox")
	if !ok || anchor == "" || current.AccountKey == "" || current.Origin != "https://outlook.cloud.microsoft" || !time.Now().Before(current.ExpiresAt) {
		return RuntimeBundle{}, &transport.Error{Kind: transport.KindAuth}
	}
	return RuntimeBundle{Version: 1, AccountKey: current.AccountKey, AudienceHash: digest(aud), AnchorHash: digest(anchor), Operations: map[string]RuntimeTemplate{}}, nil
}

// Add records an already sanitized operation constructed by the fixed-shape
// draft builders. Authentication and cookies are deliberately omitted.
func (b *RuntimeBundle) Add(name string, session transport.Session, op transport.Operation) error {
	want := runtimeActions[name]
	if b == nil || want == "" || session.AccountKey != b.AccountKey || session.Origin != "https://outlook.cloud.microsoft" || op.Method != "POST" || op.Path != "/owa/service.svc" || op.Query.Get("action") != want || session.Headers.Get("Action") != want {
		return &transport.Error{Kind: transport.KindInput}
	}
	aud, ok := jwtAudience(session.Headers.Get("Authorization"))
	if !ok || digest(aud) != b.AudienceHash || digest(session.Headers.Get("X-AnchorMailbox")) != b.AnchorHash {
		return &transport.Error{Kind: transport.KindAuth}
	}
	body := op.Body
	header := op.Headers.Get("X-Owa-Urlpostdata") != ""
	if header != runtimeHeaderCarrier(name) {
		return &transport.Error{Kind: transport.KindInput}
	}
	if header {
		decoded, err := url.QueryUnescape(op.Headers.Get("X-Owa-Urlpostdata"))
		if err != nil {
			return &transport.Error{Kind: transport.KindInput}
		}
		body = []byte(decoded)
	}
	if !json.Valid(body) || len(body) > 512<<10 {
		return &transport.Error{Kind: transport.KindInput}
	}
	headers := make(http.Header)
	for k := range runtimeHeaders {
		if value := session.Headers.Get(k); value != "" {
			headers.Set(k, value)
		}
	}
	if b.Operations == nil {
		b.Operations = map[string]RuntimeTemplate{}
	}
	if _, exists := b.Operations[name]; exists {
		return &transport.Error{Kind: transport.KindInput}
	}
	b.Operations[name] = RuntimeTemplate{Action: want, Query: op.Query, HeaderCarrier: header, Headers: headers, Body: append(json.RawMessage(nil), body...)}
	return nil
}

func runtimePath() (string, error) {
	if p := os.Getenv("OUTLOOK_WEB_OPERATIONS_FILE"); p != "" {
		return p, nil
	}
	if p := os.Getenv("OUTLOOK_WEB_SESSION_POINTER"); p != "" {
		return filepath.Join(filepath.Dir(p), "outlook-operations-v1.json"), nil
	}
	home, err := os.UserHomeDir()
	if err != nil {
		return "", err
	}
	return filepath.Join(home, ".local", "share", "outlook-web-dev", "private", "outlook-operations-v1.json"), nil
}

// RuntimeBundlePathForAccount is derived from the full private binding;
// callers never supply a path or identity string from a CLI/MCP argument.
func RuntimeBundlePathForAccount(pointer string, current transport.Session) (string, error) {
	if len(current.AccountKey) != 64 {
		return "", &transport.Error{Kind: transport.KindAuth}
	}
	if _, err := hex.DecodeString(current.AccountKey); err != nil || strings.ToLower(current.AccountKey) != current.AccountKey {
		return "", &transport.Error{Kind: transport.KindAuth}
	}
	b, err := NewRuntimeBundle(current)
	if err != nil {
		return "", err
	}
	key := digest(b.AccountKey + "\x00" + b.AudienceHash + "\x00" + b.AnchorHash)
	return filepath.Join(filepath.Dir(pointer), "outlook-operations-"+key+".json"), nil
}

// ReadBoundRuntimeBundle reads only a bundle with the same account, origin,
// audience and mailbox anchor as the selected session. The current shared
// bundle wins when it matches, so a later template refresh is not shadowed by
// an older saved profile. Explicit overrides never fall back to another file.
func ReadBoundRuntimeBundle(current transport.Session) (RuntimeBundle, error) {
	fail := func() (RuntimeBundle, error) { return RuntimeBundle{}, &transport.Error{Kind: transport.KindAuth} }
	if current.Origin != "https://outlook.cloud.microsoft" || current.AccountKey == "" || !time.Now().Before(current.ExpiresAt) || current.Headers.Get("Cookie") == "" || current.Headers.Get("X-AnchorMailbox") == "" {
		return fail()
	}
	aud, ok := jwtAudience(current.Headers.Get("Authorization"))
	if !ok {
		return fail()
	}
	p, err := runtimePath()
	if err != nil {
		return fail()
	}
	paths := []string{p}
	if os.Getenv("OUTLOOK_WEB_OPERATIONS_FILE") == "" {
		accountPath, err := RuntimeBundlePathForAccount(filepath.Join(filepath.Dir(p), "current-outlook.json"), current)
		if err != nil {
			return fail()
		}
		paths = append(paths, accountPath)
	}
	for _, path := range paths {
		data, err := readPrivateFile(path, 8<<20)
		if err != nil {
			continue
		}
		var b RuntimeBundle
		if json.Unmarshal(data, &b) == nil && b.Version == 1 && b.AccountKey == current.AccountKey && b.AudienceHash == digest(aud) && b.AnchorHash == digest(current.Headers.Get("X-AnchorMailbox")) {
			return b, nil
		}
	}
	return fail()
}

func loadRuntime(name string, current transport.Session) (transport.Session, RuntimeTemplate, error) {
	fail := func() (transport.Session, RuntimeTemplate, error) {
		return transport.Session{}, RuntimeTemplate{}, &transport.Error{Kind: transport.KindAuth}
	}
	b, err := ReadBoundRuntimeBundle(current)
	if err != nil {
		return fail()
	}
	t, ok := b.Operations[name]
	want := runtimeActions[name]
	if !ok || want == "" || t.Action != want || t.HeaderCarrier != runtimeHeaderCarrier(name) || t.Query.Get("action") != want || len(t.Query) != 3 || len(t.Query["action"]) != 1 || len(t.Query["app"]) != 1 || t.Query.Get("app") != "Mail" || len(t.Query["n"]) != 1 || !digits.MatchString(t.Query.Get("n")) || !json.Valid(t.Body) {
		return fail()
	}
	current.Headers = current.Headers.Clone()
	current.Headers.Del("X-Owa-Urlpostdata")
	current.Headers.Set("Action", want)
	for k, values := range t.Headers {
		canonical := http.CanonicalHeaderKey(k)
		if !runtimeHeaders[canonical] || len(values) != 1 || strings.ContainsAny(values[0], "\r\n\x00") {
			return fail()
		}
		current.Headers.Set(canonical, values[0])
	}
	return current, t, nil
}

func PrepareRuntimeRead(name string, current transport.Session) (transport.Session, transport.Operation, error) {
	action := map[string]string{"threads_list": "FindConversation", "threads_get": "GetConversationItems", "mail_list": "FindItem", "calendars_list": "GetCalendarFoldersV2", "calendar_agenda": "GetCalendarView", "account_info": "GetAccountInformation"}[name]
	if action == "" {
		return transport.Session{}, transport.Operation{}, &transport.Error{Kind: transport.KindInput}
	}
	s, t, err := loadRuntime(name, current)
	if err != nil {
		return transport.Session{}, transport.Operation{}, err
	}
	op := transport.Operation{Method: "POST", Path: "/owa/service.svc", Query: t.Query, Class: transport.Read, Validate: func(b []byte) transport.EnvelopeResult { return exchange.ValidateRead(action, b) }}
	if t.HeaderCarrier {
		op.Headers = http.Header{"X-Owa-Urlpostdata": []string{strings.ReplaceAll(url.QueryEscape(string(t.Body)), "+", "%20")}}
	} else {
		op.Body = append([]byte(nil), t.Body...)
	}
	return s, op, nil
}

// PrepareRuntimeDraft reconstructs only the named, fixed Outlook operation;
// IDs and message content always come from the caller, not the stored fixture.
func PrepareRuntimeDraft(name string, current transport.Session, id, key, subject, text string, recipients drafts.Recipients, changes *drafts.RecipientChanges) (transport.Session, transport.Operation, error) {
	s, t, err := loadRuntime(name, current)
	if err != nil {
		return transport.Session{}, transport.Operation{}, err
	}
	var body []byte
	op := transport.Operation{Method: "POST", Path: "/owa/service.svc", Query: t.Query}
	switch name {
	case "draft_create":
		body, err = drafts.PrepareCreateWithRecipients(t.Body, subject, text, recipients)
		op.Class = transport.Write
		op.Validate = func(b []byte) transport.EnvelopeResult { _, v := drafts.ParseCreate(b); return v }
	case "draft_get", "mail_get", "draft_compose":
		body, err = drafts.PrepareGet(t.Body, id)
		op.Class = transport.Read
		op.Validate = func(b []byte) transport.EnvelopeResult { _, v := drafts.ParseGet(b); return v }
	case "draft_update":
		body, err = drafts.PrepareUpdateWithRecipients(t.Body, id, key, subject, text, changes)
		op.Class = transport.Write
		op.Validate = func(b []byte) transport.EnvelopeResult { _, v := drafts.ParseUpdate(b); return v }
	case "draft_discard":
		body, err = drafts.PrepareDiscard(t.Body, id, key)
		op.Class = transport.Write
		op.Validate = drafts.ParseDiscard
	case "draft_send":
		body, err = drafts.PrepareSend(t.Body, id, key, subject)
		op.Class = transport.Write
		op.Validate = drafts.ParseSend
	case "reply_create", "reply_all_create", "forward_create":
		kind := "reply"
		if name == "reply_all_create" {
			kind = "reply-all"
		}
		if name == "forward_create" {
			kind = "forward"
		}
		body, err = drafts.PrepareResponseDraft(t.Body, kind, id, key, subject, text, recipients)
		op.Class = transport.Write
		op.Validate = func(b []byte) transport.EnvelopeResult { _, v := drafts.ParseCreate(b); return v }
	default:
		return transport.Session{}, transport.Operation{}, &transport.Error{Kind: transport.KindInput}
	}
	if err != nil {
		return transport.Session{}, transport.Operation{}, &transport.Error{Kind: transport.KindInput}
	}
	if t.HeaderCarrier {
		op.Headers = http.Header{"X-Owa-Urlpostdata": []string{strings.ReplaceAll(url.QueryEscape(string(body)), "+", "%20")}}
	} else {
		op.Body = body
	}
	return s, op, nil
}
