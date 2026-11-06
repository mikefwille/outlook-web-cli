package capture

import (
	"encoding/json"
	"net/url"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func TestRuntimeWebBundleRejectsForeignRoutesAndDropsCanary(t *testing.T) {
	current, _, err := LoadFindFolder(writeCapture(t, syntheticCapture(t, time.Now().Add(time.Hour)), 0600))
	if err != nil {
		t.Fatal(err)
	}
	current.Headers.Set("X-Owa-Canary", "synthetic-canary-secret")
	b, err := NewRuntimeBundle(current)
	if err != nil {
		t.Fatal(err)
	}
	op := transport.Operation{Method: "POST", Path: "/searchservice/api/v2/query", Query: url.Values{"n": {"1"}, "cv": {"synthetic"}}, Body: []byte(`{}`), Class: transport.Read}
	if err = b.AddWeb("search_all", current, op); err != nil {
		t.Fatal(err)
	}
	data, err := json.Marshal(b)
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(string(data), "synthetic-canary-secret") || strings.Contains(string(data), "synthetic-secret") {
		t.Fatal("credential retained")
	}
	file := filepath.Join(t.TempDir(), "ops.json")
	if err = os.WriteFile(file, data, 0600); err != nil {
		t.Fatal(err)
	}
	t.Setenv("OUTLOOK_WEB_OPERATIONS_FILE", file)
	s, template, err := LoadRuntimeWeb("search_all", current)
	if err != nil || template.Path != "/searchservice/api/v2/query" || s.Headers.Get("X-Owa-Canary") != "" || s.Headers.Get("Authorization") == "" {
		t.Fatal("runtime web operation invalid")
	}
	wrong := current
	wrong.AccountKey = "foreign"
	if _, _, err = LoadRuntimeWeb("search_all", wrong); err == nil {
		t.Fatal("foreign account accepted")
	}
	if _, _, err = LoadRuntimeWeb("unlisted", current); err == nil {
		t.Fatal("unlisted operation accepted")
	}
	op.Path = "/elsewhere"
	if err = b.AddWeb("search_folder", current, op); err == nil {
		t.Fatal("foreign route accepted")
	}
}
