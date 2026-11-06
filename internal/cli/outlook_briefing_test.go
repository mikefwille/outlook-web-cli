package cli

import (
	"bytes"
	"encoding/json"
	"testing"
)

func TestBriefingValidatesScopeAndOutput(t *testing.T) {
	base := []string{"--folder-id", "folder", "--calendar-id", "calendar", "--from", "2026-11-06T09:00:00-07:00", "--to", "2026-11-07T09:00:00-07:00"}
	for _, tc := range []struct {
		args []string
		ok   bool
	}{
		{base, true},
		{[]string{"--calendar-id", "calendar", "--from", "2026-11-06T09:00:00-07:00", "--to", "2026-11-07T09:00:00-07:00"}, false},
		{append(append([]string{}, base...), "--limit", "51"), false},
		{append(append([]string{}, base...), "--execute"), false},
		{[]string{"--folder-id", "folder", "--calendar-id", "calendar", "--from", "2026-11-06T09:00:00Z", "--to", "2027-01-06T09:00:00Z"}, false},
	} {
		cmd := newOutlookBriefingCmd(&rootFlags{dryRun: true})
		var out bytes.Buffer
		cmd.SetOut(&out)
		cmd.SetErr(&bytes.Buffer{})
		cmd.SetArgs(tc.args)
		err := cmd.Execute()
		if (err == nil) != tc.ok {
			t.Fatalf("unexpected briefing validation: %v", err)
		}
		if tc.ok {
			var v map[string]any
			if json.Unmarshal(out.Bytes(), &v) != nil || v["dry_run"] != true || v["calendar_complete"] != false {
				t.Fatal("briefing scope missing")
			}
		}
	}
}
