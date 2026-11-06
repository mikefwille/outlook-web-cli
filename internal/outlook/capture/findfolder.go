// Package capture converts a private, exact observed FindFolder capture into a
// narrowly scoped replay. It is a development probe, not a general login flow.
package capture

import (
	"crypto/sha256"
	"encoding/base64"
	"encoding/hex"
	"encoding/json"
	"io"
	"net/http"
	"net/url"
	"os"
	"regexp"
	"strings"
	"syscall"
	"time"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/exchange"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

type rawCapture struct {
	Service        string            `json:"service"`
	Method         string            `json:"method"`
	Operation      string            `json:"operation"`
	ObservedStatus string            `json:"observed_status"`
	URL            string            `json:"url"`
	Headers        map[string]string `json:"headers"`
}

var digits = regexp.MustCompile(`^[0-9]{1,6}$`)

// LoadFindFolder accepts only a same-user, mode-0600 regular file with the
// known Outlook helper shape, exact service/action and observed HTTPS host.
// It never returns file contents or credential bytes in an error.
func LoadFindFolder(path string) (transport.Session, transport.Operation, error) {
	invalid := func() (transport.Session, transport.Operation, error) {
		return transport.Session{}, transport.Operation{}, &transport.Error{Kind: transport.KindAuth}
	}
	info, err := os.Lstat(path)
	if err != nil || !info.Mode().IsRegular() || info.Mode().Perm()&0o077 != 0 {
		return invalid()
	}
	stat, ok := info.Sys().(*syscall.Stat_t)
	if !ok || stat.Uid != uint32(os.Getuid()) {
		return invalid()
	}
	f, err := os.Open(path)
	if err != nil {
		return invalid()
	}
	defer f.Close()
	opened, err := f.Stat()
	if err != nil || !os.SameFile(info, opened) || !opened.Mode().IsRegular() || opened.Mode().Perm()&0o077 != 0 {
		return invalid()
	}
	data, err := io.ReadAll(io.LimitReader(f, 2<<20+1))
	if err != nil || len(data) > 2<<20 {
		return invalid()
	}
	var raw rawCapture
	if json.Unmarshal(data, &raw) != nil || raw.Service != "outlook" || raw.Method != "POST" || raw.Operation != "FindFolder" || raw.ObservedStatus != "200" {
		return invalid()
	}
	u, err := url.Parse(raw.URL)
	if err != nil || u.Scheme != "https" || u.Host != "outlook.cloud.microsoft" || u.User != nil || u.Path != "/owa/service.svc" || u.Fragment != "" {
		return invalid()
	}
	query := u.Query()
	if len(query) != 3 || len(query["action"]) != 1 || query.Get("action") != "FindFolder" || len(query["app"]) != 1 || query.Get("app") != "Mail" || len(query["n"]) != 1 || !digits.MatchString(query.Get("n")) {
		return invalid()
	}
	headers := make(http.Header)
	for k, v := range raw.Headers {
		key := strings.ToLower(k)
		switch key {
		case "host", "connection", "content-length", "accept-encoding", "transfer-encoding", "proxy-authorization", "proxy-connection", "upgrade", "te", "trailer":
			continue
		}
		if strings.HasPrefix(key, ":") || strings.ContainsAny(v, "\r\n\x00") {
			return invalid()
		}
		headers.Set(k, v)
	}
	if headers.Get("Action") != "FindFolder" || headers.Get("Cookie") == "" {
		return invalid()
	}
	if headers.Get("X-AnchorMailbox") == "" {
		return invalid()
	}
	envelope, err := url.QueryUnescape(headers.Get("X-Owa-Urlpostdata"))
	if err != nil {
		return invalid()
	}
	var request struct {
		Type string `json:"__type"`
		Body struct {
			Type string `json:"__type"`
		} `json:"Body"`
	}
	if json.Unmarshal([]byte(envelope), &request) != nil || request.Type != "FindFolderJsonRequest:#Exchange" || request.Body.Type != "FindFolderRequest:#Exchange" {
		return invalid()
	}
	expires, account, ok := jwtIdentity(headers.Get("Authorization"))
	if !ok || !time.Now().Before(expires) {
		return invalid()
	}
	return transport.Session{AccountKey: account, Origin: "https://outlook.cloud.microsoft", ExpiresAt: expires, Headers: headers},
		transport.Operation{Method: "POST", Path: u.Path, Query: query, Class: transport.Read, Validate: exchange.ValidateFindFolder}, nil
}

// Decoding a JWT supplies an expiry hint and local partition key only. It is
// not a signature or identity verification; the service validates the token.
func jwtIdentity(authorization string) (time.Time, string, bool) {
	parts := strings.SplitN(authorization, " ", 2)
	if len(parts) != 2 || !strings.EqualFold(parts[0], "Bearer") {
		return time.Time{}, "", false
	}
	segments := strings.Split(parts[1], ".")
	if len(segments) != 3 {
		return time.Time{}, "", false
	}
	payload, err := base64.RawURLEncoding.DecodeString(segments[1])
	if err != nil {
		return time.Time{}, "", false
	}
	var claims struct {
		Exp    int64  `json:"exp"`
		Tenant string `json:"tid"`
		Object string `json:"oid"`
	}
	if json.Unmarshal(payload, &claims) != nil || claims.Exp <= 0 || claims.Tenant == "" || claims.Object == "" {
		return time.Time{}, "", false
	}
	sum := sha256.Sum256([]byte(claims.Tenant + "\x00" + claims.Object))
	return time.Unix(claims.Exp, 0), hex.EncodeToString(sum[:]), true
}
