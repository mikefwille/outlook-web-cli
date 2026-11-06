package cli

import (
	"bytes"
	"encoding/json"
	"strings"
	"testing"

	"github.com/spf13/cobra"
)

func TestOutlookCalendarContactPreviewsAreOffline(t *testing.T) {
	tests := []struct {
		group string
		args  []string
	}{
		{"calendar", []string{"create", "--calendar-id", "synthetic-folder", "--subject", "Appointment", "--start", "2026-11-09T09:00:00-07:00", "--end", "2026-11-09T10:00:00-07:00"}},
		{"calendar", []string{"update", "synthetic-event", "--change-key", "version", "--subject", "New"}},
		{"calendar", []string{"delete", "synthetic-event", "--change-key", "version"}},
		{"contacts", []string{"create", "--first", "Ada", "--last", "Lovelace"}},
		{"contacts", []string{"update", "synthetic-contact", "--change-key", "version", "--last", "Byron"}},
		{"contacts", []string{"delete", "synthetic-contact", "--change-key", "version"}},
	}
	for _, tc := range tests {
		t.Run(strings.Join(tc.args, "_"), func(t *testing.T) {
			flags := &rootFlags{}
			var cmd *cobra.Command
			if tc.group == "calendar" {
				cmd = &cobra.Command{Use: "calendar"}
				addOutlookAppointments(cmd, flags)
			} else {
				cmd = newOutlookPersonalContacts(flags)
			}
			var out bytes.Buffer
			cmd.SetOut(&out)
			cmd.SetErr(&bytes.Buffer{})
			cmd.SetArgs(tc.args)
			if err := cmd.Execute(); err != nil {
				t.Fatal(err)
			}
			var result map[string]any
			if json.Unmarshal(out.Bytes(), &result) != nil || result["preview"] != true {
				t.Fatalf("expected JSON preview, got %s", out.String())
			}
		})
	}
}
func TestOutlookCalendarContactRejectUnsafeExecution(t *testing.T) {
	for _, tc := range []struct {
		group string
		args  []string
	}{{"calendar", []string{"delete", "event", "--execute"}}, {"contacts", []string{"delete", "contact", "--execute"}}, {"contacts", []string{"update", "contact", "--change-key", "v1", "--execute"}}} {
		flags := &rootFlags{}
		var cmd *cobra.Command
		if tc.group == "calendar" {
			cmd = &cobra.Command{Use: "calendar"}
			addOutlookAppointments(cmd, flags)
		} else {
			cmd = newOutlookPersonalContacts(flags)
		}
		cmd.SetOut(&bytes.Buffer{})
		cmd.SetErr(&bytes.Buffer{})
		cmd.SetArgs(tc.args)
		if cmd.Execute() == nil {
			t.Fatal("invalid write accepted")
		}
	}
	t.Setenv("PRINTING_PRESS_VERIFY", "1")
	cmd := newOutlookPersonalContacts(&rootFlags{})
	cmd.SetOut(&bytes.Buffer{})
	cmd.SetErr(&bytes.Buffer{})
	cmd.SetArgs([]string{"create", "--first", "Ada", "--execute"})
	if err := cmd.Execute(); err == nil || !strings.Contains(err.Error(), "verification harness") {
		t.Fatalf("harness write not blocked: %v", err)
	}
}
func TestOutlookExtensionsRegisteredOnce(t *testing.T) {
	root := newRootCmd(&rootFlags{})
	calendar := findSubcommand(root, "calendar")
	contacts := findSubcommand(root, "contacts")
	if calendar == nil || contacts == nil || findSubcommand(calendar, "agenda") == nil || findSubcommand(calendar, "get") == nil || findSubcommand(contacts, "get") == nil {
		t.Fatal("expected observed calendar and new extension commands")
	}
}

func TestCalendarBodyAndLocationPreview(t *testing.T) {
	parent := &cobra.Command{Use: "calendar"}
	addOutlookAppointments(parent, &rootFlags{})
	var out bytes.Buffer
	parent.SetOut(&out)
	parent.SetErr(&bytes.Buffer{})
	parent.SetArgs([]string{"update", "event", "--change-key", "v1", "--body", "plain text", "--add-location", "Second room"})
	if err := parent.Execute(); err != nil {
		t.Fatal(err)
	}
	var v struct {
		Preview bool              `json:"preview"`
		Changes map[string]string `json:"changes"`
	}
	if json.Unmarshal(out.Bytes(), &v) != nil || !v.Preview || v.Changes["body"] != "plain text" || v.Changes["add_location"] != "Second room" {
		t.Fatalf("incorrect preview: %s", out.String())
	}
}

func TestCalendarTimeUpdatePreview(t *testing.T) {
	for _, tc := range []struct {
		args  []string
		valid bool
	}{
		{[]string{"update", "event", "--change-key", "v1", "--start", "2026-11-09T09:00:00-07:00", "--end", "2026-11-09T10:00:00-07:00"}, true},
		{[]string{"update", "event", "--change-key", "v1", "--start", "2026-11-09T09:00:00-07:00"}, false},
		{[]string{"update", "event", "--change-key", "v1", "--start", "2026-11-09T09:00:00-07:00", "--end", "2026-11-09T10:00:00-07:00", "--body", "text"}, true},
	} {
		parent := &cobra.Command{Use: "calendar"}
		addOutlookAppointments(parent, &rootFlags{})
		var out bytes.Buffer
		parent.SetOut(&out)
		parent.SetErr(&bytes.Buffer{})
		parent.SetArgs(tc.args)
		err := parent.Execute()
		if (err == nil) != tc.valid {
			t.Fatalf("unexpected result %v for %v", err, tc.args)
		}
		if tc.valid {
			var x struct {
				Preview bool              `json:"preview"`
				Changes map[string]string `json:"changes"`
			}
			if json.Unmarshal(out.Bytes(), &x) != nil || !x.Preview || x.Changes["start"] == "" || x.Changes["end"] == "" {
				t.Fatal("time preview incorrect")
			}
		}
	}
}

func TestCalendarLocationEditPreview(t *testing.T) {
	for _, tc := range []struct {
		args  []string
		valid bool
	}{
		{[]string{"update", "event", "--change-key", "v1", "--replace-location", "Old room", "--with-location", "New room"}, true},
		{[]string{"update", "event", "--change-key", "v1", "--remove-location", "Old room"}, true},
		{[]string{"update", "event", "--change-key", "v1", "--replace-location", "Old room"}, false},
		{[]string{"update", "event", "--change-key", "v1", "--with-location", "New room"}, false},
		{[]string{"update", "event", "--change-key", "v1", "--remove-location", "Old room", "--add-location", "New room"}, false},
	} {
		parent := &cobra.Command{Use: "calendar"}
		addOutlookAppointments(parent, &rootFlags{})
		var out bytes.Buffer
		parent.SetOut(&out)
		parent.SetErr(&bytes.Buffer{})
		parent.SetArgs(tc.args)
		err := parent.Execute()
		if (err == nil) != tc.valid {
			t.Fatalf("unexpected result %v for %v", err, tc.args)
		}
		if tc.valid && !strings.Contains(out.String(), `"preview"`) {
			t.Fatal("preview omitted its operation details")
		}
	}
}
