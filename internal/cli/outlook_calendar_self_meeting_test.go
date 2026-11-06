package cli

import (
	"bytes"
	"encoding/json"
	"testing"

	"github.com/spf13/cobra"
)

func TestSelfOnlineEventPreviewsAndScope(t *testing.T) {
	for _, tc := range []struct {
		args  []string
		valid bool
	}{
		{[]string{"create", "--calendar-id", "folder", "--subject", "Self", "--start", "2026-11-06T04:30:00-07:00", "--end", "2026-11-06T05:00:00-07:00", "--self-meeting"}, true},
		{[]string{"create", "--calendar-id", "folder", "--subject", "Self", "--start-date", "2026-11-06", "--end-date", "2026-11-06", "--self-meeting"}, false},
		{[]string{"update", "event", "--change-key", "v1", "--self-meeting", "--subject", "New"}, true},
		{[]string{"update", "event", "--change-key", "v1", "--self-meeting", "--body", "New"}, false},
		{[]string{"update", "event", "--change-key", "v1", "--self-meeting", "--series", "--subject", "New"}, false},
		{[]string{"delete", "event", "--change-key", "v1", "--self-meeting"}, true},
		{[]string{"delete", "event", "--change-key", "v1", "--self-meeting", "--series"}, false},
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
			var v map[string]any
			if json.Unmarshal(out.Bytes(), &v) != nil || v["preview"] != true {
				t.Fatal("self event preview missing")
			}
		}
	}
}
