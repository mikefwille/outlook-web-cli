package cli

import (
	"bytes"
	"strings"
	"testing"
)

func TestProvisionalCommandsReplaceGenericRPCHandlers(t *testing.T) {
	root := RootCmd()
	for _, words := range [][]string{{"folders", "list"}, {"threads", "list"}, {"threads", "get"}, {"mail", "list"}, {"calendars", "list"}, {"calendar", "agenda"}, {"auth", "status"}} {
		cmd, _, err := root.Find(words)
		if err != nil || cmd == root || cmd.Name() != words[len(words)-1] {
			t.Fatalf("missing %v", words)
		}
		if cmd.Flags().Lookup("action") != nil {
			t.Fatalf("%v exposes free-form action", words)
		}
	}
}

func TestCapturedReadDryRunDoesNotNeedPrivateState(t *testing.T) {
	root := RootCmd()
	var out bytes.Buffer
	root.SetOut(&out)
	root.SetErr(&out)
	root.SetArgs([]string{"threads", "list", "--dry-run", "--no-learn"})
	if err := root.Execute(); err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(out.String(), `"dry_run":true`) || !strings.Contains(out.String(), `"FindConversation"`) {
		t.Fatalf("unexpected preview %s", out.String())
	}
}

func TestDraftCreateDefaultsToPreview(t *testing.T) {
	root := RootCmd()
	var out bytes.Buffer
	root.SetOut(&out)
	root.SetErr(&out)
	root.SetArgs([]string{"drafts", "create", "--subject", "synthetic", "--text", "no recipients", "--no-learn"})
	if err := root.Execute(); err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(out.String(), `"preview":true`) || !strings.Contains(out.String(), `"sends_mail":false`) {
		t.Fatalf("unexpected preview %s", out.String())
	}
}
