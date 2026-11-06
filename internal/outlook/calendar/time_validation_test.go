package calendar

import "testing"

func TestCapturedMountainTimeConversion(t *testing.T) {
	cases := []struct {
		input, want string
		invalid     bool
	}{
		{"2026-07-01T15:00:00Z", "2026-07-01T09:00:00.000", false},
		{"2026-03-08T03:30:00-06:00", "2026-03-08T03:30:00.000", false},
		{"2026-11-01T01:30:00-06:00", "", true},
		{"2026-11-01T01:30:00-07:00", "", true},
		{"2026-11-02T09:00:00", "", true},
	}
	for _, tc := range cases {
		got, err := localTime(tc.input)
		if tc.invalid {
			if err == nil {
				t.Fatalf("ambiguous or unzoned input accepted: %s", tc.input)
			}
			continue
		}
		if err != nil || got != tc.want {
			t.Fatalf("time conversion %q got %q, error %v", tc.input, got, err)
		}
	}
}
