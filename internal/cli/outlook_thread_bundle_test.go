package cli

import (
	"bytes"
	"encoding/json"
	"testing"
)

func TestThreadBundleDryRunValidatesInputs(t *testing.T) {
	for _, tc := range []struct {
		args []string
		ok   bool
	}{
		{[]string{"synthetic-id", "--limit", "10"}, true},
		{[]string{"synthetic-id", "--limit", "0"}, false},
		{[]string{"synthetic-id", "--limit", "101"}, false},
		{[]string{"synthetic-id", "--execute"}, false},
		{[]string{}, false},
	} {
		cmd := newThreadBundleCmd(&rootFlags{dryRun: true})
		var out bytes.Buffer
		cmd.SetOut(&out)
		cmd.SetErr(&bytes.Buffer{})
		cmd.SetArgs(tc.args)
		err := cmd.Execute()
		if (err == nil) != tc.ok {
			t.Fatalf("unexpected bundle validation: %v", err)
		}
		if tc.ok {
			var p map[string]any
			if json.Unmarshal(out.Bytes(), &p) != nil || p["dry_run"] != true || p["operation"] != "threads bundle" {
				t.Fatal("invalid bundle preview")
			}
		}
	}
}
