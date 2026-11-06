package capture

import (
	"bufio"
	"net/http"
	"net/url"
	"path/filepath"
	"strings"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/drafts"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

// PrepareDraftOperation binds one of the three independently observed draft
// operations to a fresh same-account session. It does not accept other RPCs.
// Mutations must be preceded by a verified draft read in the caller.
func PrepareDraftOperation(detailPath, requestPath, action string, current transport.Session, id, key, subject, text string) (transport.Session, transport.Operation, error) {
	return PrepareDraftOperationWithRecipients(detailPath, requestPath, action, current, id, key, subject, text, nil)
}

func PrepareDraftOperationWithRecipients(detailPath, requestPath, action string, current transport.Session, id, key, subject, text string, recipients *drafts.RecipientChanges) (transport.Session, transport.Operation, error) {
	return prepareDraftOperation(detailPath, requestPath, action, current, id, key, subject, text, recipients, false)
}

// PrepareDraftSend is a distinct, explicitly invoked delivery operation. Its
// caller must separately verify the current draft and preview recipients.
func PrepareDraftSend(detailPath, requestPath string, current transport.Session, id, key, subject string) (transport.Session, transport.Operation, error) {
	return prepareDraftOperation(detailPath, requestPath, "UpdateItem", current, id, key, subject, "", nil, true)
}

func prepareDraftOperation(detailPath, requestPath, action string, current transport.Session, id, key, subject, text string, recipients *drafts.RecipientChanges, send bool) (transport.Session, transport.Operation, error) {
	fail := func() (transport.Session, transport.Operation, error) {
		return transport.Session{}, transport.Operation{}, &transport.Error{Kind: transport.KindInput}
	}
	if action != "GetItem" && action != "UpdateItem" && action != "DeleteItem" {
		return fail()
	}
	if action != "UpdateItem" && recipients != nil {
		return fail()
	}
	if requestPath != "" && filepath.Dir(filepath.Clean(detailPath)) != filepath.Dir(filepath.Clean(requestPath)) {
		return fail()
	}
	detail, err := readPrivateFile(detailPath, 2<<20)
	if err != nil {
		return fail()
	}
	match := requestLine.FindStringSubmatch(string(detail))
	if len(match) != 2 {
		return fail()
	}
	u, err := url.Parse(match[1])
	if err != nil || u.Scheme != "https" || u.Host != "outlook.cloud.microsoft" || u.User != nil || u.Path != "/owa/service.svc" || u.Fragment != "" {
		return fail()
	}
	q := u.Query()
	if len(q) != 3 || len(q["action"]) != 1 || q.Get("action") != action || len(q["app"]) != 1 || q.Get("app") != "Mail" || len(q["n"]) != 1 || !digits.MatchString(q.Get("n")) {
		return fail()
	}
	blocks := strings.SplitN(string(detail), "### Request Headers\n", 2)
	if len(blocks) != 2 {
		return fail()
	}
	block := strings.SplitN(blocks[1], "\n### ", 2)[0]
	headers := make(http.Header)
	scanner := bufio.NewScanner(strings.NewReader(block))
	scanner.Buffer(make([]byte, 1024), 1<<20)
	for scanner.Scan() {
		line := scanner.Text()
		if !strings.HasPrefix(line, "- ") {
			continue
		}
		k, v, ok := strings.Cut(strings.TrimPrefix(line, "- "), ":")
		if !ok || k == "" || strings.HasPrefix(k, ":") {
			continue
		}
		switch strings.ToLower(k) {
		case "host", "connection", "content-length", "accept-encoding", "transfer-encoding", "proxy-authorization", "proxy-connection", "upgrade", "te", "trailer":
			continue
		}
		v = strings.TrimSpace(v)
		if strings.ContainsAny(v, "\r\n\x00") {
			return fail()
		}
		headers.Set(k, v)
	}
	if scanner.Err() != nil || headers.Get("Action") != action || headers.Get("Cookie") == "" || headers.Get("X-AnchorMailbox") == "" {
		return fail()
	}
	expires, account, ok := jwtIdentity(headers.Get("Authorization"))
	if !ok {
		return fail()
	}
	observed := transport.Session{AccountKey: account, Origin: "https://outlook.cloud.microsoft", ExpiresAt: expires, Headers: headers}
	observed, err = RebindRead(observed, current)
	if err != nil {
		return fail()
	}
	var template []byte
	headerCarrier := observed.Headers.Get("X-Owa-Urlpostdata") != ""
	if headerCarrier {
		if action != "DeleteItem" || requestPath != "" {
			return fail()
		}
		value, e := url.QueryUnescape(observed.Headers.Get("X-Owa-Urlpostdata"))
		if e != nil {
			return fail()
		}
		template = []byte(value)
		observed.Headers.Del("X-Owa-Urlpostdata")
	} else {
		if action == "DeleteItem" || requestPath == "" {
			return fail()
		}
		template, err = readPrivateFile(requestPath, 256_000)
		if err != nil {
			return fail()
		}
	}
	var body []byte
	op := transport.Operation{Method: "POST", Path: u.Path, Query: q}
	switch action {
	case "GetItem":
		body, err = drafts.PrepareGet(template, id)
		op.Class = transport.Read
		op.Validate = func(b []byte) transport.EnvelopeResult { _, state := drafts.ParseGet(b); return state }
	case "UpdateItem":
		if send {
			body, err = drafts.PrepareSend(template, id, key, subject)
		} else {
			body, err = drafts.PrepareUpdateWithRecipients(template, id, key, subject, text, recipients)
		}
		op.Class = transport.Write
		if send {
			op.Validate = drafts.ParseSend
		} else {
			op.Validate = func(b []byte) transport.EnvelopeResult { _, state := drafts.ParseUpdate(b); return state }
		}
	case "DeleteItem":
		body, err = drafts.PrepareDiscard(template, id, key)
		op.Class = transport.Write
		op.Validate = drafts.ParseDiscard
	}
	if err != nil {
		return fail()
	}
	if headerCarrier {
		// Outlook's captured header uses percent-encoded spaces (%20), not
		// form-encoded plus signs. The service interprets it as a URI component.
		encoded := strings.ReplaceAll(url.QueryEscape(string(body)), "+", "%20")
		op.Headers = http.Header{"X-Owa-Urlpostdata": []string{encoded}}
	} else {
		op.Body = body
	}
	return observed, op, nil
}
