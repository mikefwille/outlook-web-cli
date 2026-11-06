package attachments

import (
	"context"
	"encoding/base64"
	"encoding/json"
	"io"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"testing"
	"time"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

type roundTripFunc func(*http.Request) (*http.Response, error)

func (f roundTripFunc) RoundTrip(r *http.Request) (*http.Response, error) { return f(r) }

func TestAttachmentDownloadUsesScopedTokenAndNoCanary(t *testing.T) {
	exp := time.Now().Add(time.Hour).Unix()
	payload, _ := json.Marshal(map[string]any{"tid": "tenant", "oid": "person", "preferred_username": "person@example.invalid", "exp": exp})
	token := "header." + base64.RawURLEncoding.EncodeToString(payload) + ".signature"
	_, key, _ := claims("Bearer " + token)
	current := transport.Session{AccountKey: key, Origin: "https://outlook.cloud.microsoft", ExpiresAt: time.Now().Add(time.Hour), Headers: http.Header{"Cookie": []string{"private-cookie"}, "Authorization": []string{"Bearer outlook-token"}}}
	q := url.Values{"owa": []string{"app"}, "scriptVer": []string{"1"}, "clientId": []string{"client"}, "origin": []string{"2"}, "isDownload": []string{"true"}, "animation": []string{"true"}}
	dir := t.TempDir()
	if err := os.Chmod(dir, 0700); err != nil {
		t.Fatal(err)
	}
	routePath := filepath.Join(dir, "route.json")
	tokenPath := filepath.Join(dir, "token.json")
	if err := SaveDownloadRoute(routePath, DownloadRoute{Version: 1, AccountKey: key, Mailbox: "person@example.invalid", Query: q}); err != nil {
		t.Fatal(err)
	}
	credential, _ := json.Marshal(map[string]string{"credentialType": "AccessToken", "tokenType": "Bearer", "secret": token, "target": "api://resource/owaattachments.read", "expiresOn": strconv.FormatInt(exp, 10)})
	if err := os.WriteFile(tokenPath, credential, 0600); err != nil {
		t.Fatal(err)
	}
	s, err := LoadDownloadSession(routePath, tokenPath, current)
	if err != nil {
		t.Fatal(err)
	}
	rt := roundTripFunc(func(r *http.Request) (*http.Response, error) {
		if r.URL.Host != "attachments.office.net" || r.URL.Query().Get("id") != "attachment-id" || r.URL.Query().Has("X-OWA-CANARY") || r.Header.Get("Cookie") != "" || r.Header.Get("Authorization") != "Bearer "+token {
			t.Error("incorrect download scope")
		}
		return &http.Response{StatusCode: 200, Header: http.Header{"Content-Disposition": []string{`attachment; filename="file.txt"`}}, Body: io.NopCloser(strings.NewReader("fixture"))}, nil
	})
	b, err := Download(context.Background(), current, s, "attachment-id", rt, 20)
	if err != nil || string(b) != "fixture" {
		t.Fatal("valid attachment download failed")
	}
	bad := current
	bad.AccountKey = "another-account"
	if _, err = LoadDownloadSession(routePath, tokenPath, bad); err == nil {
		t.Fatal("cross-account token accepted")
	}
	if _, err = Download(context.Background(), current, s, "attachment-id", rt, 2); err == nil {
		t.Fatal("response cap ignored")
	}
	redirect := roundTripFunc(func(*http.Request) (*http.Response, error) {
		return &http.Response{StatusCode: 302, Header: http.Header{"Location": []string{"https://example.invalid/"}}, Body: io.NopCloser(strings.NewReader(""))}, nil
	})
	if _, err = Download(context.Background(), current, s, "attachment-id", redirect, 20); err == nil {
		t.Fatal("redirect accepted")
	}
}
