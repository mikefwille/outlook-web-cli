package capture

import (
	"encoding/json"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func TestRuntimeRuleTemplateBindsAccountAndDiscardsCredentials(t *testing.T) {
	current, _, err := LoadFindFolder(writeCapture(t, syntheticCapture(t, time.Now().Add(time.Hour)), 0600))
	if err != nil {
		t.Fatal(err)
	}
	bundle, err := NewRuntimeBundle(current)
	if err != nil {
		t.Fatal(err)
	}
	op := transport.Operation{Method: "POST", Path: "/owa/service.svc", Query: url.Values{"action": {"NewInboxRule"}, "app": {"Mail"}, "n": {"1"}}, Headers: http.Header{"X-Owa-Urlpostdata": {url.QueryEscape(`{"synthetic":true}`)}}, Class: transport.Write}
	if err = bundle.AddWeb("rule_create", current, op); err != nil {
		t.Fatal(err)
	}
	data, err := json.Marshal(bundle)
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(string(data), "synthetic-secret") {
		t.Fatal("credentials retained")
	}
	file := filepath.Join(t.TempDir(), "operations.json")
	if os.WriteFile(file, data, 0600) != nil {
		t.Fatal("cannot write synthetic bundle")
	}
	t.Setenv("OUTLOOK_WEB_OPERATIONS_FILE", file)
	s, _, err := LoadRuntimeWeb("rule_create", current)
	if err != nil || s.Headers.Get("Action") != "NewInboxRule" {
		t.Fatal("rule action not rebound")
	}
	foreign := current
	foreign.AccountKey = "another-account"
	if _, _, err = LoadRuntimeWeb("rule_create", foreign); err == nil {
		t.Fatal("foreign account accepted")
	}
	op.Query.Set("action", "RemoveInboxRule")
	if err = bundle.AddWeb("rule_disable", current, op); err == nil {
		t.Fatal("mismatched action accepted")
	}
}

func TestRuntimeRuleExpansionActionMapping(t *testing.T) {
	current, _, err := LoadFindFolder(writeCapture(t, syntheticCapture(t, time.Now().Add(time.Hour)), 0600))
	if err != nil {
		t.Fatal(err)
	}
	bundle, err := NewRuntimeBundle(current)
	if err != nil {
		t.Fatal(err)
	}
	for name, action := range map[string]string{"rule_enable": "EnableInboxRule", "rules_reorder": "SetInboxAndSweepRules"} {
		op := transport.Operation{Method: "POST", Path: "/owa/service.svc", Query: url.Values{"action": {action}, "app": {"Mail"}, "n": {"1"}}, Headers: http.Header{"X-Owa-Urlpostdata": {url.QueryEscape(`{"synthetic":true}`)}}, Class: transport.Write}
		if err = bundle.AddWeb(name, current, op); err != nil {
			t.Fatal("valid expansion template rejected")
		}
	}
	data, err := json.Marshal(bundle)
	if err != nil {
		t.Fatal(err)
	}
	file := filepath.Join(t.TempDir(), "operations.json")
	if os.WriteFile(file, data, 0600) != nil {
		t.Fatal("cannot write synthetic bundle")
	}
	t.Setenv("OUTLOOK_WEB_OPERATIONS_FILE", file)
	for name, action := range map[string]string{"rule_enable": "EnableInboxRule", "rules_reorder": "SetInboxAndSweepRules"} {
		s, _, err := LoadRuntimeWeb(name, current)
		if err != nil || s.Headers.Get("Action") != action {
			t.Fatal("expansion action not rebound")
		}
	}
}
