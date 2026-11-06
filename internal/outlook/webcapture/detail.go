// Package webcapture binds an operation's privately captured UI request to the
// current Outlook session. Capture files are templates, never credential stores
// used as a substitute for the current session.
package webcapture

import (
	"bufio"
	"crypto/sha256"
	"encoding/base64"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"syscall"
	"time"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/capture"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

var linePattern = regexp.MustCompile(`(?m)^## Request (https://\S+)\r?\nStatus: (200|201|204)\r?$`)

type Detail struct {
	Method     string
	Session    transport.Session
	URL        *url.URL
	HeaderBody bool
	Body       []byte
}

func privateFile(path string, max int64) ([]byte, error) {
	info, err := os.Lstat(path)
	if err != nil || !info.Mode().IsRegular() || info.Mode().Perm()&0077 != 0 {
		return nil, fmt.Errorf("private capture is unavailable or has unsafe permissions")
	}
	st, ok := info.Sys().(*syscall.Stat_t)
	if !ok || st.Uid != uint32(os.Getuid()) {
		return nil, fmt.Errorf("private capture is unavailable or has unsafe permissions")
	}
	f, err := os.Open(path)
	if err != nil {
		return nil, fmt.Errorf("private capture is unavailable")
	}
	defer f.Close()
	opened, err := f.Stat()
	if err != nil || !os.SameFile(info, opened) || opened.Mode().Perm()&0077 != 0 {
		return nil, fmt.Errorf("private capture changed while opening")
	}
	data, err := io.ReadAll(io.LimitReader(f, max+1))
	if err != nil || int64(len(data)) > max {
		return nil, fmt.Errorf("private capture exceeds size limit")
	}
	return data, nil
}

func identity(h string) (time.Time, string, bool) {
	p := strings.SplitN(h, " ", 2)
	if len(p) != 2 || !strings.EqualFold(p[0], "Bearer") {
		return time.Time{}, "", false
	}
	parts := strings.Split(p[1], ".")
	if len(parts) != 3 {
		return time.Time{}, "", false
	}
	raw, err := base64.RawURLEncoding.DecodeString(parts[1])
	if err != nil {
		return time.Time{}, "", false
	}
	var c struct {
		Exp    int64  `json:"exp"`
		Tenant string `json:"tid"`
		Object string `json:"oid"`
	}
	if json.Unmarshal(raw, &c) != nil || c.Exp <= 0 || c.Tenant == "" || c.Object == "" {
		return time.Time{}, "", false
	}
	s := sha256.Sum256([]byte(c.Tenant + "\x00" + c.Object))
	return time.Unix(c.Exp, 0), hex.EncodeToString(s[:]), true
}

// Load accepts only a private exact-origin capture. The caller must validate
// operation-specific method, URL, request type and fields before sending.
func Load(detailPath, requestPath string, current transport.Session) (Detail, error) {
	fail := func() (Detail, error) {
		return Detail{}, fmt.Errorf("invalid or incompatible private Outlook operation capture")
	}
	if strings.HasPrefix(detailPath, "runtime:") {
		if requestPath != "" {
			return fail()
		}
		session, template, err := capture.LoadRuntimeWeb(strings.TrimPrefix(detailPath, "runtime:"), current)
		if err != nil {
			return fail()
		}
		u := &url.URL{Scheme: "https", Host: "outlook.cloud.microsoft", Path: template.Path, RawQuery: template.Query.Encode()}
		return Detail{Method: template.Method, Session: session, URL: u, HeaderBody: template.HeaderCarrier, Body: template.Body}, nil
	}
	if detailPath == "" || (requestPath != "" && filepath.Dir(filepath.Clean(detailPath)) != filepath.Dir(filepath.Clean(requestPath))) {
		return fail()
	}
	data, err := privateFile(detailPath, 2<<20)
	if err != nil {
		return fail()
	}
	m := linePattern.FindStringSubmatch(string(data))
	if len(m) != 3 {
		return fail()
	}
	u, err := url.Parse(m[1])
	if err != nil || u.Scheme != "https" || u.Host != "outlook.cloud.microsoft" || u.User != nil || u.Fragment != "" {
		return fail()
	}
	parts := strings.SplitN(string(data), "### Request Headers\n", 2)
	if len(parts) != 2 {
		return fail()
	}
	block := strings.SplitN(parts[1], "\n### ", 2)[0]
	h := http.Header{}
	method := ""
	scan := bufio.NewScanner(strings.NewReader(block))
	scan.Buffer(make([]byte, 1024), 1<<20)
	for scan.Scan() {
		line := scan.Text()
		if strings.HasPrefix(line, "- :method:") {
			method = strings.TrimSpace(strings.TrimPrefix(line, "- :method:"))
			continue
		}
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
		h.Set(k, v)
	}
	if scan.Err() != nil {
		return fail()
	}
	exp, account, ok := identity(h.Get("Authorization"))
	if !ok || h.Get("Cookie") == "" || h.Get("X-AnchorMailbox") == "" {
		return fail()
	}
	session, err := capture.RebindRead(transport.Session{AccountKey: account, Origin: "https://outlook.cloud.microsoft", ExpiresAt: exp, Headers: h}, current)
	if err != nil {
		return fail()
	}
	result := Detail{Method: method, Session: session, URL: u}
	if v := session.Headers.Get("X-Owa-Urlpostdata"); v != "" {
		if requestPath != "" {
			return fail()
		}
		s, e := url.QueryUnescape(v)
		if e != nil {
			return fail()
		}
		result.HeaderBody = true
		result.Body = []byte(s)
		session.Headers.Del("X-Owa-Urlpostdata")
		result.Session = session
	} else if requestPath != "" {
		result.Body, err = privateFile(requestPath, 512<<10)
		if err != nil {
			return fail()
		}
	}
	return result, nil
}

// Operation creates an operation with the same observed routing and header/body
// carrier, but never retains the captured ID after a caller has rewritten it.
func (d Detail) Operation(method string, path string, body []byte, class transport.Class, validate func([]byte) transport.EnvelopeResult) transport.Operation {
	op := transport.Operation{Method: method, Path: path, Query: d.URL.Query(), Class: class, Validate: validate}
	if d.HeaderBody {
		op.Headers = http.Header{"X-Owa-Urlpostdata": []string{strings.ReplaceAll(url.QueryEscape(string(body)), "+", "%20")}}
	} else {
		op.Body = body
	}
	return op
}
