package cli

import (
	"bytes"
	"encoding/json"
	"testing"

	"github.com/spf13/cobra"
)

func TestCalendarBoundedReadPreviewValidation(t *testing.T) {
	for _, tc := range []struct {
		args  []string
		valid bool
	}{
		{[]string{"search", "Planning", "--calendar-id", "calendar-id", "--from", "2026-11-06T09:00:00-07:00", "--to", "2026-11-07T09:00:00-07:00"}, true},
		{[]string{"search", "Planning", "--from", "2026-11-06T09:00:00-07:00", "--to", "2026-11-07T09:00:00-07:00"}, false},
		{[]string{"search", "Planning", "--calendar-id", "calendar-id", "--from", "2026-11-06T09:00:00-07:00", "--to", "2027-01-07T09:00:00-07:00"}, false},
		{[]string{"search", "", "--calendar-id", "calendar-id", "--from", "2026-11-06T09:00:00-07:00", "--to", "2026-11-07T09:00:00-07:00"}, false},
		{[]string{"find-times", "--from", "2026-11-06T09:00:00-07:00", "--to", "2026-11-06T11:00:00-07:00", "--duration-minutes", "45"}, true},
		{[]string{"find-times", "--from", "2026-11-06T09:01:00-07:00", "--to", "2026-11-06T11:00:00-07:00"}, false},
		{[]string{"find-times", "--from", "2026-11-06T09:00:00-07:00", "--to", "2026-11-06T11:00:00-07:00", "--duration-minutes", "20"}, false},
		{[]string{"find-times", "--from", "2026-11-06T09:00:00-07:00", "--to", "2026-11-06T11:00:00-07:00", "--principal", "person@example.invalid"}, false},
	} {
		parent := &cobra.Command{Use: "calendar", SilenceUsage: true, SilenceErrors: true}
		addOutlookCalendarSubjectSearch(parent, &rootFlags{dryRun: true})
		addOutlookCalendarFindTimes(parent, &rootFlags{dryRun: true})
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
			if json.Unmarshal(out.Bytes(), &result) != nil || result["dry_run"] != true || result["state_change"] != false || result["complete"] != false {
				t.Fatalf("invalid preview for %v: %s", tc.args, out.String())
			}
		}
	}
}

func TestCalendarBoundedCommandsAreReadOnlyLive(t *testing.T) {
	parent := &cobra.Command{Use: "calendar"}
	addOutlookCalendarSubjectSearch(parent, &rootFlags{})
	addOutlookCalendarFindTimes(parent, &rootFlags{})
	for _, name := range []string{"search", "find-times"} {
		cmd := findSubcommand(parent, name)
		if cmd == nil || cmd.Annotations["mcp:read-only"] != "true" || cmd.Annotations["pp:data-source"] != "live" {
			t.Fatalf("%s missing live/read-only metadata", name)
		}
	}
}
