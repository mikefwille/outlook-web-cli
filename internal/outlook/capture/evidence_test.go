package capture

import (
	"encoding/json"
	"fmt"
	"net/url"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func evidenceFile(t *testing.T, action, typ string, bodyCarrier bool) string {
	t.Helper()
	dir := t.TempDir()
	v := syntheticCapture(t, time.Now().Add(time.Hour))
	headers := v["headers"].(map[string]string)
	request := `{"__type":"` + typ + `","Body":{}}`
	hdr := fmt.Sprintf("- authorization:%s\n- cookie:%s\n- action:%s\n- x-anchormailbox:synthetic-account\n- x-tenantid:synthetic-tenant\n", headers["authorization"], headers["cookie"], action)
	requestPath := ""
	if bodyCarrier {
		requestPath = filepath.Join(dir, "request.json")
		if err := os.WriteFile(requestPath, []byte(request), 0o600); err != nil {
			t.Fatal(err)
		}
	} else {
		hdr += "- x-owa-urlpostdata:" + url.QueryEscape(request) + "\n"
	}
	record := map[string]any{"method": "POST", "operation": action, "full_request_path": requestPath,
		"detail_text": fmt.Sprintf("## Request https://outlook.cloud.microsoft/owa/service.svc?action=%s&app=Mail&n=1\nStatus: 200\n### Request Headers\n%s### Response Headers\n", action, hdr)}
	b, _ := json.Marshal(record)
	p := filepath.Join(dir, "evidence.json")
	if err := os.WriteFile(p, b, 0o600); err != nil {
		t.Fatal(err)
	}
	return p
}

func TestLoadObservedHeaderAndBodyReads(t *testing.T) {
	for _, tc := range []struct {
		action, typ string
		body        bool
	}{
		{"FindConversation", "FindConversationJsonRequest:#Exchange", false},
		{"GetCalendarView", "GetCalendarViewJsonRequest:#Exchange", false},
		{"GetConversationItems", "GetConversationItemsJsonRequest:#Exchange", true},
	} {
		s, op, err := LoadObservedRead(evidenceFile(t, tc.action, tc.typ, tc.body), tc.action)
		if err != nil || s.AccountKey == "" || op.Class != transport.Read || op.Query.Get("action") != tc.action {
			t.Fatalf("%s: %v", tc.action, err)
		}
		if tc.body != (len(op.Body) > 0) {
			t.Fatalf("%s: wrong carrier", tc.action)
		}
	}
}

func TestLoadObservedRejectsActionAndEnvelopeMismatch(t *testing.T) {
	p := evidenceFile(t, "FindConversation", "FindConversationJsonRequest:#Exchange", false)
	for _, action := range []string{"DeleteFolder", "GetCalendarView"} {
		_, _, err := LoadObservedRead(p, action)
		if err == nil || strings.Contains(err.Error(), "synthetic-secret") {
			t.Fatalf("unsafe result: %v", err)
		}
	}
	p = evidenceFile(t, "FindConversation", "DeleteFolderJsonRequest:#Exchange", false)
	if _, _, err := LoadObservedRead(p, "FindConversation"); err == nil {
		t.Fatal("accepted mismatched envelope")
	}
}
