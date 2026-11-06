package cli

import (
	"bytes"
	"encoding/json"
	"testing"

	"github.com/spf13/cobra"
)

func TestAllDayCalendarPreviewsAndValidation(t *testing.T) {
	for _, tc := range []struct {
		args      []string
		valid     bool
		operation string
	}{
		{[]string{"create", "--calendar-id", "folder", "--subject", "All day", "--start-date", "2026-11-06", "--end-date", "2026-11-07"}, true, "calendar create"},
		{[]string{"create", "--calendar-id", "folder", "--subject", "All day", "--start-date", "2026-11-06", "--end-date", "2026-11-07", "--body", "note", "--location", "room"}, true, "calendar create"},
		{[]string{"create", "--calendar-id", "folder", "--subject", "All day", "--start-date", "2026-11-06"}, false, ""},
		{[]string{"create", "--calendar-id", "folder", "--subject", "All day", "--start-date", "2026-11-06", "--end-date", "2026-11-07", "--start", "2026-11-06T09:00:00-07:00"}, false, ""},
		{[]string{"create", "--calendar-id", "folder", "--subject", "All day", "--start-date", "2026-11-06", "--end-date", "2026-11-07", "--weekly-until", "2026-12-04"}, false, ""},
		{[]string{"update", "event", "--change-key", "v1", "--start-date", "2026-11-06", "--end-date", "2026-11-08"}, true, "calendar update"},
		{[]string{"update", "event", "--change-key", "v1", "--start-date", "2026-11-06", "--end-date", "2026-11-05"}, false, ""},
		{[]string{"update", "event", "--change-key", "v1", "--start-date", "2026-11-06", "--end-date", "2026-11-08", "--subject", "Mixed"}, false, ""},
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
			if json.Unmarshal(out.Bytes(), &result) != nil || result["preview"] != true || result["operation"] != tc.operation {
				t.Fatal("all-day preview missing")
			}
		}
	}
}
