package cli

import (
	"bytes"
	"encoding/base64"
	"encoding/json"
	"net/url"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/capture"
)

func accountFixture(t *testing.T, dir, id, account, audience, anchor string) {
	t.Helper()
	claims, _ := json.Marshal(map[string]any{"exp": time.Now().Add(time.Hour).Unix(), "tid": "tenant-" + account, "oid": "user-" + account, "aud": audience})
	token := "header." + base64.RawURLEncoding.EncodeToString(claims) + ".signature"
	envelope := url.QueryEscape(`{"__type":"FindFolderJsonRequest:#Exchange","Body":{"__type":"FindFolderRequest:#Exchange"}}`)
	data, _ := json.Marshal(map[string]any{"service": "outlook", "method": "POST", "operation": "FindFolder", "observed_status": "200", "url": "https://outlook.cloud.microsoft/owa/service.svc?action=FindFolder&app=Mail&n=1", "headers": map[string]string{"authorization": "Bearer " + token, "cookie": "session=fixture-secret", "action": "FindFolder", "x-owa-urlpostdata": envelope, "x-anchormailbox": anchor}})
	path := filepath.Join(dir, "session-outlook-"+id+".json")
	if err := os.WriteFile(path, data, 0600); err != nil {
		t.Fatal(err)
	}
	s, _, err := capture.LoadFindFolder(path)
	if err != nil {
		t.Fatal(err)
	}
	bundle, err := capture.NewRuntimeBundle(s)
	if err != nil {
		t.Fatal(err)
	}
	bundle.WebOperations = map[string]capture.RuntimeWebTemplate{"folders_list": {Method: "POST", Path: "/owa/service.svc", Query: url.Values{"action": {"FindFolder"}, "app": {"Mail"}, "n": {"1"}}, HeaderCarrier: true, Body: json.RawMessage(`{}`)}}
	bundle.Operations = map[string]capture.RuntimeTemplate{"mail_list": {Action: "FindItem", Query: url.Values{"action": {"FindItem"}, "app": {"Mail"}, "n": {"1"}}, HeaderCarrier: true, Body: json.RawMessage(`{}`)}}
	data, _ = json.Marshal(bundle)
	if err := os.WriteFile(filepath.Join(dir, "outlook-operations-v1.json"), data, 0600); err != nil {
		t.Fatal(err)
	}
	pointAtAccount(t, dir, id)
}

func pointAtAccount(t *testing.T, dir, id string) {
	t.Helper()
	b, _ := json.Marshal(map[string]any{"version": 1, "service": "outlook", "export_id": id})
	if err := os.WriteFile(filepath.Join(dir, "current-outlook.json"), b, 0600); err != nil {
		t.Fatal(err)
	}
}

func runAccountCommand(t *testing.T, args ...string) (map[string]any, error) {
	t.Helper()
	var out bytes.Buffer
	cmd := newOutlookAccountsCmd(&rootFlags{asJSON: true})
	cmd.SilenceErrors, cmd.SilenceUsage = true, true
	cmd.SetOut(&out)
	cmd.SetArgs(args)
	err := cmd.Execute()
	if err != nil {
		if strings.Contains(err.Error(), "fixture-secret") || strings.Contains(out.String(), "fixture-secret") {
			t.Fatal("private credential appeared in output")
		}
		return nil, err
	}
	var value map[string]any
	if err := json.Unmarshal(out.Bytes(), &value); err != nil {
		t.Fatalf("invalid JSON: %v: %s", err, out.String())
	}
	if strings.Contains(out.String(), "fixture-secret") || strings.Contains(out.String(), "account_key") || strings.Contains(out.String(), "anchor_hash") {
		t.Fatal("private account binding appeared in output")
	}
	return value, nil
}

func TestAccountAliasesBindAndSwitchWithoutCrossAccountTemplates(t *testing.T) {
	dir := t.TempDir()
	if err := os.Chmod(dir, 0700); err != nil {
		t.Fatal(err)
	}
	pointer := filepath.Join(dir, "current-outlook.json")
	t.Setenv("OUTLOOK_WEB_SESSION_POINTER", pointer)
	t.Setenv("OUTLOOK_WEB_SESSION_FILE", "")
	t.Setenv("OUTLOOK_WEB_OPERATIONS_FILE", "")
	t.Setenv("PRINTING_PRESS_VERIFY", "")
	t.Setenv("PRINTING_PRESS_DOGFOOD", "")
	accountFixture(t, dir, "exporta-aaaaaaaa", "alpha", "aud-a", "anchor-a")
	if v, err := runAccountCommand(t, "register-current", "alpha"); err != nil || v["preview"] != true {
		t.Fatalf("registration preview: %v %#v", err, v)
	}
	if _, err := os.Stat(filepath.Join(dir, "outlook-accounts-v1.json")); !os.IsNotExist(err) {
		t.Fatal("registration preview changed local state")
	}
	if _, err := runAccountCommand(t, "register-current", "alpha", "--execute"); err != nil {
		t.Fatal(err)
	}
	alpha, _, err := capture.LoadCurrentFindFolder(pointer)
	if err != nil {
		t.Fatal(err)
	}
	// This simulates a different already-imported account. It performs no login
	// or network request and does not claim a second real account was tested.
	accountFixture(t, dir, "exportb-bbbbbbbb", "beta", "aud-b", "anchor-b")
	if _, err := runAccountCommand(t, "register-current", "beta", "--execute"); err != nil {
		t.Fatal(err)
	}
	beta, _, _ := capture.LoadCurrentFindFolder(pointer)
	if _, err := runAccountCommand(t, "register-current", "alpha", "--execute"); err == nil {
		t.Fatal("accepted an alias rebound to another account")
	}
	before, _ := os.ReadFile(pointer)
	if v, err := runAccountCommand(t, "use", "alpha"); err != nil || v["preview"] != true || v["already_active"] != false {
		t.Fatalf("switch preview: %v %#v", err, v)
	}
	if now, _ := os.ReadFile(pointer); !bytes.Equal(before, now) {
		t.Fatal("switch preview changed the active session")
	}
	if _, err := runAccountCommand(t, "use", "alpha", "--execute"); err != nil {
		t.Fatal(err)
	}
	selected, _, err := capture.LoadCurrentFindFolder(pointer)
	if err != nil || selected.AccountKey != alpha.AccountKey || selected.AccountKey == beta.AccountKey {
		t.Fatal("selected the wrong account")
	}
	// The shared default bundle belongs to beta; alpha must use its own
	// binding-specific snapshot instead of receiving beta's templates.
	if _, _, err := capture.LoadRuntimeWeb("folders_list", selected); err != nil {
		t.Fatal("selected account could not load its own snapshot")
	}
	if _, _, err := capture.PrepareRuntimeRead("mail_list", selected); err != nil {
		t.Fatal("selected account could not load its own mail template")
	}
	v, err := runAccountCommand(t, "list")
	if err != nil {
		t.Fatal(err)
	}
	items, ok := v["accounts"].([]any)
	if !ok || len(items) != 2 || items[0].(map[string]any)["alias"] != "alpha" || items[0].(map[string]any)["active"] != true || items[1].(map[string]any)["active"] != false {
		t.Fatalf("incorrect account listing: %#v", v)
	}
	before, _ = os.ReadFile(pointer)
	if _, err := runAccountCommand(t, "use", "unknown", "--execute"); err == nil {
		t.Fatal("unknown alias accepted")
	}
	if now, _ := os.ReadFile(pointer); !bytes.Equal(before, now) {
		t.Fatal("unknown alias changed active account")
	}
	// Replacing a registered export with another audience is detected before
	// pointer publication, even though it has the same account and anchor.
	accountFixture(t, dir, "exportb-bbbbbbbb", "beta", "aud-changed", "anchor-b")
	pointAtAccount(t, dir, "exporta-aaaaaaaa")
	before, _ = os.ReadFile(pointer)
	if _, err := runAccountCommand(t, "use", "beta", "--execute"); err == nil {
		t.Fatal("changed token audience accepted")
	}
	if now, _ := os.ReadFile(pointer); !bytes.Equal(before, now) {
		t.Fatal("mismatched profile changed active account")
	}
}

func TestAccountSelectionRespectsLockAndPinsInvocation(t *testing.T) {
	dir := t.TempDir()
	if err := os.Chmod(dir, 0700); err != nil {
		t.Fatal(err)
	}
	pointer := filepath.Join(dir, "current-outlook.json")
	t.Setenv("OUTLOOK_WEB_SESSION_POINTER", pointer)
	t.Setenv("OUTLOOK_WEB_SESSION_FILE", "")
	t.Setenv("OUTLOOK_WEB_OPERATIONS_FILE", "")
	t.Setenv("PRINTING_PRESS_VERIFY", "")
	t.Setenv("PRINTING_PRESS_DOGFOOD", "")
	accountFixture(t, dir, "exporta-aaaaaaaa", "alpha", "aud-a", "anchor-a")
	if _, err := runAccountCommand(t, "register-current", "alpha", "--execute"); err != nil {
		t.Fatal(err)
	}
	accountFixture(t, dir, "exportb-bbbbbbbb", "beta", "aud-b", "anchor-b")
	if _, err := runAccountCommand(t, "register-current", "beta", "--execute"); err != nil {
		t.Fatal(err)
	}
	finish := beginOutlookInvocation()
	first, _, err := currentSession()
	if err != nil {
		t.Fatal(err)
	}
	first.Headers.Set("Cookie", "modified-locally")
	if err := os.Mkdir(filepath.Join(dir, ".m365-session-outlook.lock"), 0700); err != nil {
		t.Fatal(err)
	}
	if _, err := runAccountCommand(t, "use", "alpha", "--execute"); err == nil {
		t.Fatal("switched while shared helper lock held")
	}
	if err := os.Remove(filepath.Join(dir, ".m365-session-outlook.lock")); err != nil {
		t.Fatal(err)
	}
	if _, err := runAccountCommand(t, "use", "alpha", "--execute"); err != nil {
		t.Fatal(err)
	}
	pinned, _, err := currentSession()
	if err != nil || pinned.AccountKey != first.AccountKey || pinned.Headers.Get("Cookie") != "session=fixture-secret" {
		t.Fatal("in-flight invocation changed account or shared mutable headers")
	}
	finish()
	fresh, _, err := currentSession()
	if err != nil || fresh.AccountKey == first.AccountKey {
		t.Fatal("next invocation did not see new account")
	}
}
