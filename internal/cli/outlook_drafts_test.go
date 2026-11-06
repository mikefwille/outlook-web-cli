package cli

import (
	"bytes"
	"encoding/json"
	"strings"
	"testing"
)

func TestOutlookDraftMutationsPreviewWithoutSession(t *testing.T) {
	t.Setenv("OUTLOOK_WEB_DRAFT_GET_DETAILS", "")
	for _, args := range [][]string{
		{"update", "synthetic-id", "--change-key", "key", "--subject", "new subject", "--text", "new text"},
		{"discard", "synthetic-id", "--change-key", "key"},
	} {
		cmd := newDraftsCmd(&rootFlags{})
		var out bytes.Buffer
		cmd.SetOut(&out)
		cmd.SetArgs(args)
		if err := cmd.Execute(); err != nil {
			t.Fatalf("preview unexpectedly used session: %v", err)
		}
		var result map[string]any
		if json.Unmarshal(out.Bytes(), &result) != nil || result["preview"] != true || result["sends_mail"] != false {
			t.Fatal("unsafe or invalid preview")
		}
	}
}

func TestOutlookDraftWritesRefuseVerificationHarness(t *testing.T) {
	t.Setenv("PRINTING_PRESS_DOGFOOD", "1")
	for _, args := range [][]string{
		{"create", "--subject", "fixture", "--execute"},
		{"update", "id", "--change-key", "key", "--subject", "subject", "--text", "body", "--execute"},
		{"discard", "id", "--change-key", "key", "--execute"},
		{"send", "id", "--change-key", "key", "--execute"},
	} {
		cmd := newDraftsCmd(&rootFlags{})
		cmd.SetOut(&bytes.Buffer{})
		cmd.SetErr(&bytes.Buffer{})
		cmd.SetArgs(args)
		if err := cmd.Execute(); err == nil || !strings.Contains(err.Error(), "disabled under the verification harness") {
			t.Errorf("harness did not block %s", args[0])
		}
	}
}

func TestOutlookDraftDryRunAllowsHarnessIntrospection(t *testing.T) {
	t.Setenv("PRINTING_PRESS_DOGFOOD", "1")
	for _, name := range []string{"create", "get", "update", "discard", "send"} {
		cmd := newDraftsCmd(&rootFlags{dryRun: true})
		cmd.SetOut(&bytes.Buffer{})
		cmd.SetErr(&bytes.Buffer{})
		cmd.SetArgs([]string{name})
		if err := cmd.Execute(); err != nil {
			t.Errorf("dry-run introspection %s: %v", name, err)
		}
	}
}

func TestOutlookDraftMutationRequiresVersion(t *testing.T) {
	cmd := newDraftsCmd(&rootFlags{})
	cmd.SetOut(&bytes.Buffer{})
	cmd.SetErr(&bytes.Buffer{})
	cmd.SetArgs([]string{"discard", "synthetic-id", "--execute"})
	if cmd.Execute() == nil {
		t.Fatal("discard accepted without version")
	}
}

func TestOutlookAddressedDraftPreviewIsExactAndDoesNotSend(t *testing.T) {
	cmd := newDraftsCmd(&rootFlags{})
	var out bytes.Buffer
	cmd.SetOut(&out)
	cmd.SetArgs([]string{"create", "--subject", "subject", "--text", "body", "--to", "Alice <alice@example.invalid>", "--bcc", "blind@example.invalid"})
	if err := cmd.Execute(); err != nil {
		t.Fatal(err)
	}
	var result struct {
		Preview    bool `json:"preview"`
		Sends      bool `json:"sends_mail"`
		Recipients struct {
			To []struct {
				Email string `json:"EmailAddress"`
			} `json:"to"`
			Bcc []struct {
				Email string `json:"EmailAddress"`
			} `json:"bcc"`
		} `json:"recipients"`
	}
	if json.Unmarshal(out.Bytes(), &result) != nil || !result.Preview || result.Sends || len(result.Recipients.To) != 1 || result.Recipients.To[0].Email != "alice@example.invalid" || len(result.Recipients.Bcc) != 1 || result.Recipients.Bcc[0].Email != "blind@example.invalid" {
		t.Fatal("recipient preview mismatch")
	}
}
