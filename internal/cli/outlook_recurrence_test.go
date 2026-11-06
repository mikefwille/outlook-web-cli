package cli

import (
	"bytes"
	"encoding/json"
	"testing"

	"github.com/spf13/cobra"
)

func TestRecurrencePreviewsAndValidation(t *testing.T) {
	for _, tc := range []struct {
		args  []string
		valid bool
		scope string
	}{
		{[]string{"create", "--calendar-id", "folder", "--subject", "Weekly", "--start", "2026-11-06T09:00:00-07:00", "--end", "2026-11-06T09:30:00-07:00", "--weekly-until", "2026-12-04"}, true, ""},
		{[]string{"create", "--calendar-id", "folder", "--subject", "Weekly", "--start", "2026-11-06T09:00:00-07:00", "--end", "2026-11-06T09:30:00-07:00", "--weekly-until", "2026-11-05"}, false, ""},
		{[]string{"update", "instance", "--change-key", "v1", "--subject", "New", "--occurrence"}, true, "occurrence"},
		{[]string{"update", "instance", "--change-key", "v1", "--body", "New", "--occurrence"}, false, ""},
		{[]string{"delete", "master", "--change-key", "v1", "--series"}, true, "entire series"},
	} {
		parent := &cobra.Command{Use: "calendar"}
		addOutlookAppointments(parent, &rootFlags{})
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
			if json.Unmarshal(out.Bytes(), &result) != nil || result["preview"] != true {
				t.Fatal("preview missing")
			}
			if tc.scope != "" && result["scope"] != tc.scope {
				t.Fatal("scope missing")
			}
		}
	}
	parent := &cobra.Command{Use: "calendar"}
	addOutlookOccurrences(parent, &rootFlags{dryRun: true})
	var out bytes.Buffer
	parent.SetOut(&out)
	parent.SetErr(&bytes.Buffer{})
	parent.SetArgs([]string{"occurrences", "master", "--from", "2026-11-01T00:00:00-06:00", "--to", "2026-11-15T00:00:00-07:00"})
	if err := parent.Execute(); err != nil {
		t.Fatal(err)
	}
	var result map[string]any
	if json.Unmarshal(out.Bytes(), &result) != nil || result["dry_run"] != true {
		t.Fatal("expansion preview missing")
	}
}
