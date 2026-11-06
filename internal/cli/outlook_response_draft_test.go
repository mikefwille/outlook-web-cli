package cli

import (
	"bytes"
	"strings"
	"testing"
)

func TestOutlookResponseDraftHarnessAndRequiredRecipients(t *testing.T) {
	t.Setenv("PRINTING_PRESS_DOGFOOD", "1")
	for _, kind := range []string{"reply", "forward"} {
		preview := newResponseDraftCmd(&rootFlags{dryRun: true}, kind)
		preview.SetOut(&bytes.Buffer{})
		preview.SetErr(&bytes.Buffer{})
		preview.SetArgs([]string{})
		if err := preview.Execute(); err != nil {
			t.Errorf("%s dry-run: %v", kind, err)
		}
		write := newResponseDraftCmd(&rootFlags{}, kind)
		write.SetOut(&bytes.Buffer{})
		write.SetErr(&bytes.Buffer{})
		write.SetArgs([]string{"id", "--to", "self@example.invalid", "--execute"})
		if err := write.Execute(); err == nil || !strings.Contains(err.Error(), "disabled under the verification harness") {
			t.Errorf("%s write reached network under harness", kind)
		}
		missing := newResponseDraftCmd(&rootFlags{}, kind)
		missing.SetOut(&bytes.Buffer{})
		missing.SetErr(&bytes.Buffer{})
		missing.SetArgs([]string{"id"})
		if err := missing.Execute(); err == nil {
			t.Errorf("%s accepted missing recipients", kind)
		}
	}
}
