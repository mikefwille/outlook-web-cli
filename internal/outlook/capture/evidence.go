package capture

import (
	"bufio"
	"encoding/json"
	"io"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"syscall"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/exchange"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

var observedTypes = map[string]string{
	"FindConversation":      "FindConversationJsonRequest:#Exchange",
	"FindItem":              "FindItemJsonRequest:#Exchange",
	"GetConversationItems":  "GetConversationItemsJsonRequest:#Exchange",
	"GetCalendarFoldersV2":  "GetCalendarFoldersRequestV2:#Exchange",
	"GetCalendarView":       "GetCalendarViewJsonRequest:#Exchange",
	"GetAccountInformation": "GetAccountInformationRequest:#Exchange",
}

var requestLine = regexp.MustCompile(`(?m)^## Request (https://\S+)\r?\nStatus: 200\r?$`)

type evidenceRecord struct {
	Method          string `json:"method"`
	Operation       string `json:"operation"`
	DetailText      string `json:"detail_text"`
	FullRequestPath string `json:"full_request_path"`
}

// LoadObservedRead replays only an exact, privately captured, allowlisted read.
// It is a development bridge pending the shared session adapter; it does not
// accept arbitrary routes, methods or actions from callers.
func LoadObservedRead(path, action string) (transport.Session, transport.Operation, error) {
	data, err := readPrivateFile(path, 2<<20)
	if err != nil {
		return transport.Session{}, transport.Operation{}, &transport.Error{Kind: transport.KindAuth}
	}
	var record evidenceRecord
	if json.Unmarshal(data, &record) != nil {
		return transport.Session{}, transport.Operation{}, &transport.Error{Kind: transport.KindAuth}
	}
	return loadObservedRecord(path, action, record)
}

// LoadObservedReadDetail accepts the same private captured request details
// without copying credential-bearing content into a second evidence file.
func LoadObservedReadDetail(path, action string) (transport.Session, transport.Operation, error) {
	data, err := readPrivateFile(path, 2<<20)
	if err != nil {
		return transport.Session{}, transport.Operation{}, &transport.Error{Kind: transport.KindAuth}
	}
	return loadObservedRecord(path, action, evidenceRecord{Method: "POST", Operation: action, DetailText: string(data)})
}

func loadObservedRecord(path, action string, record evidenceRecord) (transport.Session, transport.Operation, error) {
	fail := func() (transport.Session, transport.Operation, error) {
		return transport.Session{}, transport.Operation{}, &transport.Error{Kind: transport.KindAuth}
	}
	wantType, ok := observedTypes[action]
	if !ok || record.Method != "POST" || record.Operation != action {
		return fail()
	}
	match := requestLine.FindStringSubmatch(record.DetailText)
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
	headerText := strings.SplitN(record.DetailText, "### Request Headers\n", 2)
	if len(headerText) != 2 {
		return fail()
	}
	headerText = strings.SplitN(headerText[1], "\n### ", 2)
	headers := make(http.Header)
	scanner := bufio.NewScanner(strings.NewReader(headerText[0]))
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
	var envelope []byte
	opHeaders := make(http.Header)
	if h := headers.Get("X-Owa-Urlpostdata"); h != "" {
		value, err := url.QueryUnescape(h)
		if err != nil {
			return fail()
		}
		envelope = []byte(value)
		headers.Del("X-Owa-Urlpostdata")
		opHeaders.Set("X-Owa-Urlpostdata", h)
	} else {
		if action != "GetConversationItems" || record.FullRequestPath == "" || filepath.Dir(filepath.Clean(record.FullRequestPath)) != filepath.Dir(filepath.Clean(path)) {
			return fail()
		}
		envelope, err = readPrivateFile(record.FullRequestPath, 256_000)
		if err != nil {
			return fail()
		}
	}
	var request struct {
		Type string `json:"__type"`
	}
	if json.Unmarshal(envelope, &request) != nil || request.Type != wantType {
		return fail()
	}
	op := transport.Operation{Method: "POST", Path: u.Path, Query: q, Headers: opHeaders, Class: transport.Read,
		Validate: func(b []byte) transport.EnvelopeResult { return exchange.ValidateRead(action, b) }}
	if action == "GetConversationItems" {
		op.Body = envelope
	}
	return transport.Session{AccountKey: account, Origin: "https://outlook.cloud.microsoft", ExpiresAt: expires, Headers: headers}, op, nil
}

// ReadPrivateEvidence reads one bounded owner-only capture for the local
// session-template converter. Capture contents must never be logged.
func ReadPrivateEvidence(path string) ([]byte, error) {
	return readPrivateFile(path, 1<<20)
}

func readPrivateFile(path string, max int64) ([]byte, error) {
	info, err := os.Lstat(path)
	if err != nil || !info.Mode().IsRegular() || info.Mode().Perm()&0o077 != 0 {
		return nil, os.ErrPermission
	}
	stat, ok := info.Sys().(*syscall.Stat_t)
	if !ok || stat.Uid != uint32(os.Getuid()) {
		return nil, os.ErrPermission
	}
	f, err := os.Open(path)
	if err != nil {
		return nil, os.ErrPermission
	}
	defer f.Close()
	opened, err := f.Stat()
	if err != nil || !os.SameFile(info, opened) || !opened.Mode().IsRegular() || opened.Mode().Perm()&0o077 != 0 {
		return nil, os.ErrPermission
	}
	data, err := io.ReadAll(io.LimitReader(f, max+1))
	if err != nil || int64(len(data)) > max {
		return nil, os.ErrPermission
	}
	return data, nil
}
