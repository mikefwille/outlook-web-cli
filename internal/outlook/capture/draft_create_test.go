package capture

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func TestPrepareDraftCreateAllowsOnlyMatchingAccountAndSaveOnly(t *testing.T) {
	dir := t.TempDir()
	v := syntheticCapture(t, time.Now().Add(time.Hour))
	b, _ := json.Marshal(v)
	export := filepath.Join(dir, "capture.json")
	if err := os.WriteFile(export, b, 0o600); err != nil {
		t.Fatal(err)
	}
	current, _, err := LoadFindFolder(export)
	if err != nil {
		t.Fatal(err)
	}
	h := v["headers"].(map[string]string)
	detail := fmt.Sprintf("## Request https://outlook.cloud.microsoft/owa/service.svc?action=CreateItem&app=Mail&n=1\nStatus: 200\n### Request Headers\n- authorization:%s\n- cookie:%s\n- action:CreateItem\n- x-anchormailbox:synthetic-account\n### Response Headers\n", h["authorization"], h["cookie"])
	detailPath := filepath.Join(dir, "details.txt")
	if err := os.WriteFile(detailPath, []byte(detail), 0o600); err != nil {
		t.Fatal(err)
	}
	template := `{"__type":"CreateItemJsonRequest:#Exchange","Body":{"__type":"CreateItemRequest:#Exchange","ComposeOperation":"newMail","MessageDisposition":"SaveOnly","SendOnNotFoundError":true,"Items":[{"__type":"Message:#Exchange","MessageDisposition":"SaveOnly","ToRecipients":[],"CcRecipients":[],"BccRecipients":[],"ReplyTo":[],"appendOnSend":[],"prependOnSend":[],"atMentionedRecipients":[],"recipientsAddedViaAtMention":[],"internetHeaders":[],"internetHeadersToRemove":[],"Body":{"BodyType":"HTML","Value":"old"},"Subject":"old","plainText":"old"}]}}`
	requestPath := filepath.Join(dir, "request.json")
	if err := os.WriteFile(requestPath, []byte(template), 0o600); err != nil {
		t.Fatal(err)
	}
	s, op, err := PrepareDraftCreate(detailPath, requestPath, current, "fixture", "text")
	if err != nil || op.Class != transport.Write || op.Query.Get("action") != "CreateItem" || s.Headers.Get("Action") != "CreateItem" || s.Headers.Get("X-Owa-Urlpostdata") != "" {
		t.Fatalf("safe create rejected: %v", err)
	}
	if strings.Contains(string(op.Body), `"SendOnNotFoundError":true`) || !strings.Contains(string(op.Body), `"SaveOnly"`) {
		t.Fatal("unsafe disposition")
	}
	current.AccountKey = "different-account"
	if _, _, err := PrepareDraftCreate(detailPath, requestPath, current, "fixture", "text"); err == nil {
		t.Fatal("accepted different account")
	}
}
