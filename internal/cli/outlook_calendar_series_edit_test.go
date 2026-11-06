package cli

import (
	"bytes"
	"encoding/json"
	"testing"

	"github.com/spf13/cobra"
)

func TestCalendarSeriesEditPreviewsAndResetAcknowledgment(t *testing.T) {
	for _, tc := range []struct {
		args         []string
		valid, reset bool
	}{
		{[]string{"update", "master", "--change-key", "v1", "--series", "--subject", "Renamed"}, true, false},
		{[]string{"update", "master", "--change-key", "v1", "--series", "--start", "2026-11-06T06:00:00-07:00", "--end", "2026-11-06T06:30:00-07:00"}, false, false},
		{[]string{"update", "master", "--change-key", "v1", "--series", "--start", "2026-11-06T06:00:00-07:00", "--end", "2026-11-06T06:30:00-07:00", "--reset-exceptions"}, true, true},
		{[]string{"update", "master", "--change-key", "v1", "--series", "--weekly-until", "2026-11-27", "--reset-exceptions"}, true, true},
		{[]string{"update", "master", "--change-key", "v1", "--series", "--weekly-until", "2026-11-27"}, false, false},
		{[]string{"update", "master", "--change-key", "v1", "--series", "--subject", "Renamed", "--reset-exceptions"}, false, false},
		{[]string{"update", "master", "--change-key", "v1", "--series", "--occurrence", "--subject", "Renamed"}, false, false},
		{[]string{"update", "master", "--change-key", "v1", "--weekly-until", "2026-11-27", "--reset-exceptions"}, false, false},
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
			if json.Unmarshal(out.Bytes(), &result) != nil || result["preview"] != true || result["scope"] != "entire series" {
				t.Fatal("series preview missing")
			}
			if tc.reset && result["exception_effect"] != "existing exceptions will be reset" {
				t.Fatal("reset effect not explicit")
			}
		}
	}
}
