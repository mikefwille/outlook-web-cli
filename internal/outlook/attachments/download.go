package attachments

import (
	"bufio"
	"context"
	"crypto/sha256"
	"encoding/base64"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"io"
	"mime"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"syscall"
	"time"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

const MaxDownloadBytes int64 = 16 << 20

var detailLine = regexp.MustCompile(`(?m)^## Request (https://\S+)\r?\nStatus: (200|201|204)\r?$`)

// DownloadSession is a short-lived, private attachment-service credential.
// It is separate from the normal Outlook service session and templates.
type DownloadSession struct {
	Version       int        `json:"version"`
	AccountKey    string     `json:"account_key"`
	Mailbox       string     `json:"mailbox"`
	Authorization string     `json:"authorization"`
	ExpiresAt     time.Time  `json:"expires_at"`
	Query         url.Values `json:"query"`
}

// DownloadRoute contains account-bound, non-credential routing values only.
type DownloadRoute struct {
	Version    int        `json:"version"`
	AccountKey string     `json:"account_key"`
	Mailbox    string     `json:"mailbox"`
	Query      url.Values `json:"query"`
}
type downloadClaims struct {
	Tenant string `json:"tid"`
	Object string `json:"oid"`
	User   string `json:"preferred_username"`
	Exp    int64  `json:"exp"`
}

func claims(auth string) (downloadClaims, string, bool) {
	var c downloadClaims
	if !strings.HasPrefix(auth, "Bearer ") || strings.ContainsAny(auth, "\r\n\x00") {
		return c, "", false
	}
	p := strings.Split(strings.TrimPrefix(auth, "Bearer "), ".")
	if len(p) != 3 {
		return c, "", false
	}
	b, err := base64.RawURLEncoding.DecodeString(p[1])
	if err != nil || json.Unmarshal(b, &c) != nil || c.Tenant == "" || c.Object == "" || c.User == "" || c.Exp <= 0 {
		return c, "", false
	}
	h := sha256.Sum256([]byte(c.Tenant + "\x00" + c.Object))
	return c, hex.EncodeToString(h[:]), true
}
func validMailbox(s string) bool {
	return s != "" && len(s) <= 320 && strings.Contains(s, "@") && !strings.ContainsAny(s, "/\\?#\x00\r\n")
}
func validDownloadQuery(q url.Values) bool {
	if len(q) != 6 {
		return false
	}
	for _, k := range []string{"owa", "scriptVer", "clientId", "origin", "isDownload", "animation"} {
		if len(q[k]) != 1 || q.Get(k) == "" || len(q.Get(k)) > 1024 || strings.ContainsAny(q.Get(k), "\r\n\x00") {
			return false
		}
	}
	return q.Get("origin") == "2" && q.Get("isDownload") == "true" && q.Get("animation") == "true"
}
func validateDownloadSession(s DownloadSession, current transport.Session) bool {
	c, key, ok := claims(s.Authorization)
	return ok && s.Version == 1 && key == current.AccountKey && s.AccountKey == current.AccountKey && current.Origin == "https://outlook.cloud.microsoft" && time.Now().Before(current.ExpiresAt) && time.Now().Before(time.Unix(c.Exp, 0)) && s.ExpiresAt.Equal(time.Unix(c.Exp, 0)) && validMailbox(s.Mailbox) && strings.EqualFold(s.Mailbox, c.User) && validDownloadQuery(s.Query)
}
func privateRead(path string, limit int64) ([]byte, error) {
	i, err := os.Lstat(path)
	if err != nil || !i.Mode().IsRegular() || i.Mode().Perm()&0077 != 0 {
		return nil, fmt.Errorf("private attachment session is unavailable")
	}
	st, ok := i.Sys().(*syscall.Stat_t)
	if !ok || st.Uid != uint32(os.Getuid()) {
		return nil, fmt.Errorf("private attachment session is unavailable")
	}
	f, err := os.Open(path)
	if err != nil {
		return nil, fmt.Errorf("private attachment session is unavailable")
	}
	defer f.Close()
	j, err := f.Stat()
	if err != nil || !os.SameFile(i, j) || j.Mode().Perm()&0077 != 0 {
		return nil, fmt.Errorf("private attachment session changed")
	}
	b, err := io.ReadAll(io.LimitReader(f, limit+1))
	if err != nil || int64(len(b)) > limit {
		return nil, fmt.Errorf("private attachment session is invalid")
	}
	return b, nil
}

// RouteFromCapture removes the fixture ID, bearer and canary. The captured
// token's tenant/object only bind the observed route to the current account.
func RouteFromCapture(path string, current transport.Session) (DownloadRoute, error) {
	fail := func() (DownloadRoute, error) {
		return DownloadRoute{}, fmt.Errorf("invalid or incompatible attachment-service capture")
	}
	b, err := privateRead(path, 2<<20)
	if err != nil {
		return fail()
	}
	m := detailLine.FindStringSubmatch(string(b))
	if len(m) != 3 {
		return fail()
	}
	u, err := url.Parse(m[1])
	if err != nil || u.Scheme != "https" || u.Host != "attachments.office.net" || u.User != nil || u.Fragment != "" {
		return fail()
	}
	const prefix = "/owa/"
	const suffix = "/service.svc/s/GetFileAttachment"
	if !strings.HasPrefix(u.Path, prefix) || !strings.HasSuffix(u.Path, suffix) {
		return fail()
	}
	mailbox := strings.TrimSuffix(strings.TrimPrefix(u.Path, prefix), suffix)
	if !validMailbox(mailbox) {
		return fail()
	}
	q := u.Query()
	if len(q["id"]) != 1 || q.Get("id") == "" {
		return fail()
	}
	q.Del("id")
	q.Del("X-OWA-CANARY")
	if !validDownloadQuery(q) {
		return fail()
	}
	parts := strings.SplitN(string(b), "### Request Headers\n", 2)
	if len(parts) != 2 {
		return fail()
	}
	block := strings.SplitN(parts[1], "\n### ", 2)[0]
	scan := bufio.NewScanner(strings.NewReader(block))
	scan.Buffer(make([]byte, 1024), 1<<20)
	auth, method := "", ""
	for scan.Scan() {
		l := scan.Text()
		if strings.HasPrefix(l, "- :method:") {
			method = strings.TrimSpace(strings.TrimPrefix(l, "- :method:"))
			continue
		}
		if strings.HasPrefix(strings.ToLower(l), "- authorization:") {
			auth = strings.TrimSpace(l[len("- authorization:"):])
		}
	}
	c, key, ok := claims(auth)
	if scan.Err() != nil || method != "GET" || !ok || key != current.AccountKey || !strings.EqualFold(mailbox, c.User) {
		return fail()
	}
	if current.Origin != "https://outlook.cloud.microsoft" || !time.Now().Before(current.ExpiresAt) {
		return fail()
	}
	return DownloadRoute{Version: 1, AccountKey: key, Mailbox: mailbox, Query: q}, nil
}

// LoadDownloadSession reads the current separately scoped MSAL token each
// time, so the browser-session helper can refresh it without touching routes.
func LoadDownloadSession(routePath, credentialPath string, current transport.Session) (DownloadSession, error) {
	fail := func() (DownloadSession, error) { return DownloadSession{}, &transport.Error{Kind: transport.KindAuth} }
	b, err := privateRead(routePath, 128<<10)
	if err != nil {
		return fail()
	}
	var route DownloadRoute
	if json.Unmarshal(b, &route) != nil || route.Version != 1 || route.AccountKey != current.AccountKey || !validMailbox(route.Mailbox) || !validDownloadQuery(route.Query) {
		return fail()
	}
	b, err = privateRead(credentialPath, 128<<10)
	if err != nil {
		return fail()
	}
	var token struct {
		CredentialType string `json:"credentialType"`
		TokenType      string `json:"tokenType"`
		Secret         string `json:"secret"`
		Target         string `json:"target"`
		ExpiresOn      string `json:"expiresOn"`
	}
	if json.Unmarshal(b, &token) != nil || token.CredentialType != "AccessToken" || !strings.EqualFold(token.TokenType, "Bearer") || !strings.Contains(strings.ToLower(token.Target), "/owaattachments.read") {
		return fail()
	}
	exp, err := strconv.ParseInt(token.ExpiresOn, 10, 64)
	if err != nil || !time.Now().Before(time.Unix(exp, 0)) {
		return fail()
	}
	auth := "Bearer " + token.Secret
	c, _, ok := claims(auth)
	if !ok {
		return fail()
	}
	s := DownloadSession{Version: 1, AccountKey: route.AccountKey, Mailbox: route.Mailbox, Authorization: auth, ExpiresAt: time.Unix(c.Exp, 0), Query: route.Query}
	if !validateDownloadSession(s, current) {
		return fail()
	}
	return s, nil
}

// SaveDownloadRoute atomically replaces the sanitized route in a private directory.
func SaveDownloadRoute(path string, s DownloadRoute) error {
	dir := filepath.Dir(path)
	i, err := os.Stat(dir)
	if err != nil || !i.IsDir() || i.Mode().Perm()&0077 != 0 {
		return fmt.Errorf("private session directory is unavailable")
	}
	st, ok := i.Sys().(*syscall.Stat_t)
	if !ok || st.Uid != uint32(os.Getuid()) {
		return fmt.Errorf("private session directory is unavailable")
	}
	b, err := json.Marshal(s)
	if err != nil {
		return fmt.Errorf("cannot encode attachment session")
	}
	f, err := os.CreateTemp(dir, ".attachment-session-*")
	if err != nil {
		return fmt.Errorf("cannot save attachment session")
	}
	name := f.Name()
	defer os.Remove(name)
	if f.Chmod(0600) != nil {
		f.Close()
		return fmt.Errorf("cannot secure attachment session")
	}
	if _, err = f.Write(b); err != nil {
		f.Close()
		return fmt.Errorf("cannot save attachment session")
	}
	if err = f.Sync(); err != nil {
		f.Close()
		return fmt.Errorf("cannot save attachment session")
	}
	if err = f.Close(); err != nil {
		return fmt.Errorf("cannot save attachment session")
	}
	if err = os.Rename(name, path); err != nil {
		return fmt.Errorf("cannot save attachment session")
	}
	return nil
}

// Download uses only the attachment-service bearer, never the Outlook cookie
// or bearer. It follows no redirects and requires the observed attachment
// disposition before returning bytes.
func Download(ctx context.Context, current transport.Session, s DownloadSession, id string, rt http.RoundTripper, maxBytes int64) ([]byte, error) {
	if !validateDownloadSession(s, current) {
		return nil, &transport.Error{Kind: transport.KindAuth}
	}
	if id == "" || len(id) > 4096 || strings.ContainsAny(id, "\r\n\x00") || maxBytes < 1 || maxBytes > MaxDownloadBytes {
		return nil, &transport.Error{Kind: transport.KindInput}
	}
	u := url.URL{Scheme: "https", Host: "attachments.office.net", Path: "/owa/" + s.Mailbox + "/service.svc/s/GetFileAttachment"}
	q := url.Values{}
	for k, v := range s.Query {
		q[k] = append([]string(nil), v...)
	}
	q.Set("id", id)
	u.RawQuery = q.Encode()
	req, err := http.NewRequestWithContext(ctx, "GET", u.String(), nil)
	if err != nil {
		return nil, &transport.Error{Kind: transport.KindInput}
	}
	req.Header.Set("Authorization", s.Authorization)
	req.Header.Set("Accept", "*/*")
	req.Header.Set("Origin", "https://outlook.cloud.microsoft")
	req.Header.Set("Referer", "https://outlook.cloud.microsoft/")
	if rt == nil {
		rt = http.DefaultTransport
	}
	client := &http.Client{Transport: rt, Timeout: 30 * time.Second, CheckRedirect: func(*http.Request, []*http.Request) error { return http.ErrUseLastResponse }}
	res, err := client.Do(req)
	if err != nil {
		if ctx.Err() != nil {
			return nil, &transport.Error{Kind: transport.KindCancelled}
		}
		return nil, &transport.Error{Kind: transport.KindNetwork}
	}
	defer res.Body.Close()
	if res.StatusCode != http.StatusOK {
		kind := transport.KindUpstream
		if res.StatusCode == 401 {
			kind = transport.KindAuth
		}
		if res.StatusCode == 403 {
			kind = transport.KindPermission
		}
		return nil, &transport.Error{Kind: kind, Status: res.StatusCode}
	}
	disposition, _, err := mime.ParseMediaType(res.Header.Get("Content-Disposition"))
	if err != nil || !strings.EqualFold(disposition, "attachment") {
		return nil, &transport.Error{Kind: transport.KindUpstream, Status: res.StatusCode}
	}
	b, err := io.ReadAll(io.LimitReader(res.Body, maxBytes+1))
	if err != nil {
		return nil, &transport.Error{Kind: transport.KindNetwork}
	}
	if int64(len(b)) > maxBytes {
		return nil, &transport.Error{Kind: transport.KindLimit, Status: res.StatusCode}
	}
	return b, nil
}
