package calendar

import "testing"

func TestCapturedZoneConversionAndAmbiguousTimes(t *testing.T) {
	for _, tc := range []struct {
		zone, input, expected string
		valid                 bool
	}{
		{"America/New_York", "2026-11-06T02:00:00-07:00", "2026-11-06T04:00:00.000", true},
		{"Eastern Standard Time", "2026-07-06T04:00:00-04:00", "2026-07-06T04:00:00.000", true},
		{"America/Denver", "2026-11-06T04:00:00-05:00", "2026-11-06T02:00:00.000", true},
		{"America/New_York", "2026-11-01T01:30:00-04:00", "", false},
		{"America/New_York", "2026-11-01T01:30:00-05:00", "", false},
		{"America/Los_Angeles", "2026-11-06T04:00:00-08:00", "", false},
	} {
		got, err := localTimeInZone(tc.input, tc.zone)
		if (err == nil) != tc.valid || tc.valid && got != tc.expected {
			t.Fatalf("wrong conversion for %q: %q %v", tc.zone, got, err)
		}
	}
	if got := eventLocalTime("2026-11-06T02:00:00.123-07:00", "Eastern Standard Time"); got != "2026-11-06T04:00:00.123-05:00" {
		t.Fatal("event-local timestamp lost fractional seconds")
	}
}
