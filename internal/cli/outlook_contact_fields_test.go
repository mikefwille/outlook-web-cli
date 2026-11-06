package cli

import (
	"bytes"
	"encoding/json"
	"testing"
)

func TestOutlookContactClearPreview(t *testing.T) {
	cmd := newOutlookPersonalContacts(&rootFlags{})
	var out bytes.Buffer
	cmd.SetOut(&out)
	cmd.SetErr(&bytes.Buffer{})
	cmd.SetArgs([]string{"update", "contact", "--change-key", "v1", "--clear-email", "--clear-notes"})
	if err := cmd.Execute(); err != nil {
		t.Fatal(err)
	}
	var data struct {
		Preview bool           `json:"preview"`
		Changes map[string]any `json:"changes"`
	}
	if json.Unmarshal(out.Bytes(), &data) != nil || !data.Preview || data.Changes["clear_email"] != true || data.Changes["clear_notes"] != true {
		t.Fatal("invalid clear preview")
	}
	cmd = newOutlookPersonalContacts(&rootFlags{})
	cmd.SetOut(&bytes.Buffer{})
	cmd.SetErr(&bytes.Buffer{})
	cmd.SetArgs([]string{"update", "contact", "--change-key", "v1", "--email", "a@example.invalid", "--clear-email"})
	if cmd.Execute() == nil {
		t.Fatal("conflicting set and clear accepted")
	}
}
