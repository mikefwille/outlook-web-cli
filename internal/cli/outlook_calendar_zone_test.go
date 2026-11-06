package cli

import (
	"bytes"
	"encoding/json"
	"testing"

	"github.com/spf13/cobra"
)

func TestCalendarZonePreviewsAndValidation(t *testing.T) {
	for _, tc := range []struct {
		args  []string
		valid bool
	}{
		{[]string{"create", "--calendar-id", "folder", "--subject", "Eastern", "--start", "2026-11-06T04:00:00-05:00", "--end", "2026-11-06T04:30:00-05:00", "--time-zone", "America/New_York"}, true},
		{[]string{"create", "--calendar-id", "folder", "--subject", "Eastern", "--start", "2026-11-06T04:00:00-05:00", "--end", "2026-11-06T04:30:00-05:00", "--time-zone", "America/New_York", "--body", "note", "--location", "room"}, true},
		{[]string{"create", "--calendar-id", "folder", "--subject", "Eastern", "--start", "2026-11-06T04:00:00-05:00", "--end", "2026-11-06T04:30:00-05:00", "--time-zone", "America/Los_Angeles"}, false},
		{[]string{"create", "--calendar-id", "folder", "--subject", "All day", "--start-date", "2026-11-06", "--end-date", "2026-11-07", "--time-zone", "America/New_York"}, false},
		{[]string{"update", "event", "--change-key", "v1", "--start", "2026-11-06T05:00:00-05:00", "--end", "2026-11-06T05:30:00-05:00", "--time-zone", "America/New_York"}, true},
		{[]string{"update", "event", "--change-key", "v1", "--subject", "New", "--time-zone", "America/New_York"}, false},
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
				t.Fatal("zone preview missing")
			}
		}
	}
}
