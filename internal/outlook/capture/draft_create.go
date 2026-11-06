package capture

import (
	"bufio"
	"encoding/json"
	"net/http"
	"net/url"
	"path/filepath"
	"strings"
	"time"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/drafts"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

// PrepareDraftCreate binds a captured CreateItem template to the currently
// selected Outlook account. It only permits a recipient-free SaveOnly draft.
func PrepareDraftCreate(detailPath, requestPath string, current transport.Session, subject, text string) (transport.Session, transport.Operation, error) {
	return PrepareDraftCreateWithRecipients(detailPath, requestPath, current, subject, text, drafts.Recipients{})
}

func PrepareDraftCreateWithRecipients(detailPath, requestPath string, current transport.Session, subject, text string, recipients drafts.Recipients) (transport.Session, transport.Operation, error) {
	return prepareDraftCreation(detailPath, requestPath, current, subject, text, recipients, nil)
}

type responseDraft struct{ kind, id, changeKey string }

func PrepareResponseDraft(detailPath, requestPath string, current transport.Session, kind, id, changeKey, subject, text string, recipients drafts.Recipients) (transport.Session, transport.Operation, error) {
	return prepareDraftCreation(detailPath, requestPath, current, subject, text, recipients, &responseDraft{kind: kind, id: id, changeKey: changeKey})
}

func prepareDraftCreation(detailPath, requestPath string, current transport.Session, subject, text string, recipients drafts.Recipients, response *responseDraft) (transport.Session, transport.Operation, error) {
	fail := func() (transport.Session, transport.Operation, error) {
		return transport.Session{}, transport.Operation{}, &transport.Error{Kind: transport.KindInput}
	}
	if current.Origin != "https://outlook.cloud.microsoft" || current.AccountKey == "" || !time.Now().Before(current.ExpiresAt) || current.Headers.Get("Authorization") == "" || current.Headers.Get("Cookie") == "" {
		return fail()
	}
	if filepath.Dir(filepath.Clean(detailPath)) != filepath.Dir(filepath.Clean(requestPath)) {
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
	if len(q) != 3 || len(q["action"]) != 1 || q.Get("action") != "CreateItem" || len(q["app"]) != 1 || q.Get("app") != "Mail" || len(q["n"]) != 1 || !digits.MatchString(q.Get("n")) {
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
	if scanner.Err() != nil || headers.Get("Action") != "CreateItem" || headers.Get("X-Owa-Urlpostdata") != "" || headers.Get("Cookie") == "" {
		return fail()
	}
	_, account, ok := jwtIdentity(headers.Get("Authorization"))
	observedAudience, audienceOK := jwtAudience(headers.Get("Authorization"))
	currentAudience, currentAudienceOK := jwtAudience(current.Headers.Get("Authorization"))
	if !ok || !audienceOK || !currentAudienceOK || observedAudience != currentAudience || account != current.AccountKey || headers.Get("X-AnchorMailbox") != current.Headers.Get("X-AnchorMailbox") {
		return fail()
	}
	headers.Set("Authorization", current.Headers.Get("Authorization"))
	headers.Set("Cookie", current.Headers.Get("Cookie"))
	template, err := readPrivateFile(requestPath, 256_000)
	if err != nil {
		return fail()
	}
	var body []byte
	if response == nil {
		body, err = drafts.PrepareCreateWithRecipients(template, subject, text, recipients)
	} else {
		body, err = drafts.PrepareResponseDraft(template, response.kind, response.id, response.changeKey, subject, text, recipients)
	}
	if err != nil {
		return fail()
	}
	var check struct {
		Body struct {
			MessageDisposition  string `json:"MessageDisposition"`
			SendOnNotFoundError bool   `json:"SendOnNotFoundError"`
		} `json:"Body"`
	}
	if json.Unmarshal(body, &check) != nil || check.Body.MessageDisposition != "SaveOnly" || check.Body.SendOnNotFoundError {
		return fail()
	}
	current.Headers = headers
	op := transport.Operation{Method: "POST", Path: u.Path, Query: q, Body: body, Class: transport.Write,
		Validate: func(b []byte) transport.EnvelopeResult { _, state := drafts.ParseCreate(b); return state }}
	return current, op, nil
}
