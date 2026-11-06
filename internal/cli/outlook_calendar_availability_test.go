package cli

import (
	"bytes"
	"encoding/json"
	"testing"

	"github.com/spf13/cobra"
)

func TestCalendarAvailabilityPreviewAndValidation(t *testing.T) {
	for _, tc := range []struct {
		args  []string
		valid bool
	}{
		{[]string{"availability", "--from", "2026-11-06T04:00:00-07:00", "--to", "2026-11-06T05:00:00-07:00"}, true},
		{[]string{"availability", "--from", "2026-11-06T04:00:00-07:00", "--to", "2026-11-06T05:00:00-07:00", "--principal", "person@example.invalid"}, true},
		{[]string{"availability", "--from", "2026-11-06T04:01:00-07:00", "--to", "2026-11-06T05:00:00-07:00"}, false},
		{[]string{"availability", "--from", "2026-11-06T04:00:00-07:00", "--to", "2026-11-06T05:00:00-07:00", "--principal", "invalid"}, false},
	} {
		parent := &cobra.Command{Use: "calendar"}
		addOutlookAvailability(parent, &rootFlags{dryRun: true})
		var out bytes.Buffer
		parent.SetOut(&out)
		parent.SetErr(&bytes.Buffer{})
		parent.SetArgs(tc.args)
		err := parent.Execute()
		if (err == nil) != tc.valid {
			t.Fatalf("unexpected result for %v: %v", tc.args, err)
		}
		if tc.valid {
			var result map[string]any
			if json.Unmarshal(out.Bytes(), &result) != nil || result["dry_run"] != true {
				t.Fatal("availability preview missing")
			}
		}
	}
}
