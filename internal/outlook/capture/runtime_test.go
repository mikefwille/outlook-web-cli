package capture

import (
	"encoding/json"
	"net/url"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/drafts"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func TestRuntimeBundleOmitsCredentialsAndBindsAccount(t *testing.T) {
	current, _, err := LoadFindFolder(writeCapture(t, syntheticCapture(t, time.Now().Add(time.Hour)), 0600))
	if err != nil {
		t.Fatal(err)
	}
	bundle, err := NewRuntimeBundle(current)
	if err != nil {
		t.Fatal(err)
	}
	session := current
	session.Headers = current.Headers.Clone()
	session.Headers.Del("X-Owa-Urlpostdata")
	session.Headers.Set("Action", "CreateItem")
	body := `{"__type":"CreateItemJsonRequest:#Exchange","Body":{"__type":"CreateItemRequest:#Exchange","ComposeOperation":"newMail","MessageDisposition":"SaveOnly","Items":[{"__type":"Message:#Exchange","MessageDisposition":"SaveOnly","ToRecipients":[],"CcRecipients":[],"BccRecipients":[],"ReplyTo":[],"appendOnSend":[],"prependOnSend":[],"atMentionedRecipients":[],"recipientsAddedViaAtMention":[],"internetHeaders":[],"internetHeadersToRemove":[],"Body":{"BodyType":"HTML","Value":"template"},"Subject":"template","plainText":"template"}]}}`
	op := transport.Operation{Method: "POST", Path: "/owa/service.svc", Query: url.Values{"action": {"CreateItem"}, "app": {"Mail"}, "n": {"1"}}, Body: []byte(body), Class: transport.Write}
	if err := bundle.Add("draft_create", session, op); err != nil {
		t.Fatal(err)
	}
	data, err := json.Marshal(bundle)
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(string(data), "synthetic-secret") || strings.Contains(string(data), "synthetic-account") || strings.Contains(string(data), "Authorization") || strings.Contains(string(data), "Cookie") {
		t.Fatal("bundle retained credentials or identity")
	}
	file := filepath.Join(t.TempDir(), "operations.json")
	if err := os.WriteFile(file, data, 0600); err != nil {
		t.Fatal(err)
	}
	t.Setenv("OUTLOOK_WEB_OPERATIONS_FILE", file)
	r, err := drafts.ParseRecipients([]string{"to@example.invalid"}, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	rebound, prepared, err := PrepareRuntimeDraft("draft_create", current, "", "", "new subject", "new text", r, nil)
	if err != nil || rebound.Headers.Get("Authorization") == "" || rebound.Headers.Get("Action") != "CreateItem" || !strings.Contains(string(prepared.Body), "to@example.invalid") || !strings.Contains(string(prepared.Body), "new subject") {
		t.Fatal("runtime operation not reconstructed")
	}
	other := current
	other.AccountKey = "other"
	if _, _, err := PrepareRuntimeDraft("draft_create", other, "", "", "x", "y", r, nil); err == nil {
		t.Fatal("accepted other account")
	}
	other = current
	other.Headers = current.Headers.Clone()
	other.Headers.Set("X-AnchorMailbox", "other")
	if _, _, err := PrepareRuntimeDraft("draft_create", other, "", "", "x", "y", r, nil); err == nil {
		t.Fatal("accepted other mailbox")
	}
	if _, _, err := PrepareRuntimeDraft("unlisted_action", current, "", "", "x", "y", r, nil); err == nil {
		t.Fatal("accepted unlisted action")
	}
}
