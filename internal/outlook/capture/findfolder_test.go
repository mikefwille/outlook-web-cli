package capture

import (
	"encoding/base64"
	"encoding/json"
	"net/url"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"
)

func syntheticCapture(t *testing.T, expiry time.Time) map[string]any {
	t.Helper()
	claims, _ := json.Marshal(map[string]any{"exp": expiry.Unix(), "tid": "synthetic-tenant", "oid": "synthetic-user", "aud": "synthetic-audience"})
	token := "header." + base64.RawURLEncoding.EncodeToString(claims) + ".signature"
	envelope := url.QueryEscape(`{"__type":"FindFolderJsonRequest:#Exchange","Body":{"__type":"FindFolderRequest:#Exchange"}}`)
	return map[string]any{"service": "outlook", "method": "POST", "operation": "FindFolder", "observed_status": "200",
		"url":     "https://outlook.cloud.microsoft/owa/service.svc?action=FindFolder&app=Mail&n=1",
		"headers": map[string]string{"authorization": "Bearer " + token, "cookie": "session=synthetic-secret", "action": "FindFolder", "x-owa-urlpostdata": envelope, "x-anchormailbox": "synthetic-account", "x-tenantid": "synthetic-tenant"}}
}

func writeCapture(t *testing.T, v map[string]any, mode os.FileMode) string {
	t.Helper()
	p := filepath.Join(t.TempDir(), "capture.json")
	b, _ := json.Marshal(v)
	if err := os.WriteFile(p, b, mode); err != nil {
		t.Fatal(err)
	}
	return p
}

func TestLoadExactCapturedRead(t *testing.T) {
	p := writeCapture(t, syntheticCapture(t, time.Now().Add(time.Hour)), 0o600)
	s, op, err := LoadFindFolder(p)
	if err != nil || s.Origin != "https://outlook.cloud.microsoft" || s.AccountKey == "" || op.Query.Get("action") != "FindFolder" || op.Validate == nil {
		t.Fatalf("load failed: %v", err)
	}
	if strings.Contains(s.AccountKey, "synthetic-account") {
		t.Fatal("partition key leaked account value")
	}
}

func TestRejectsWrongOriginActionExpiredAndWorldReadable(t *testing.T) {
	for _, change := range []func(map[string]any){
		func(v map[string]any) {
			v["url"] = "https://outlook.cloud.microsoft.evil.example/owa/service.svc?action=FindFolder&app=Mail&n=1"
		},
		func(v map[string]any) {
			v["url"] = "https://outlook.cloud.microsoft/owa/service.svc?action=DeleteFolder&app=Mail&n=1"
		},
		func(v map[string]any) {
			v["url"] = "https://outlook.cloud.microsoft/owa/service.svc?action=FindFolder&app=Mail&n=1&token=synthetic-secret"
		},
	} {
		v := syntheticCapture(t, time.Now().Add(time.Hour))
		change(v)
		_, _, err := LoadFindFolder(writeCapture(t, v, 0o600))
		if err == nil || strings.Contains(err.Error(), "synthetic-secret") {
			t.Fatalf("unsafe capture accepted or leaked: %v", err)
		}
	}
	_, _, err := LoadFindFolder(writeCapture(t, syntheticCapture(t, time.Now().Add(-time.Second)), 0o600))
	if err == nil {
		t.Fatal("expired capture accepted")
	}
	_, _, err = LoadFindFolder(writeCapture(t, syntheticCapture(t, time.Now().Add(time.Hour)), 0o644))
	if err == nil {
		t.Fatal("world-readable capture accepted")
	}
}

func TestRejectsSymlink(t *testing.T) {
	file := writeCapture(t, syntheticCapture(t, time.Now().Add(time.Hour)), 0o600)
	link := filepath.Join(t.TempDir(), "link.json")
	if err := os.Symlink(file, link); err != nil {
		t.Fatal(err)
	}
	if _, _, err := LoadFindFolder(link); err == nil {
		t.Fatal("symlink accepted")
	}
}
