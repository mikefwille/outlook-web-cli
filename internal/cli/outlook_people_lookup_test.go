package cli

import (
	"bytes"
	"encoding/json"
	"strings"
	"testing"
)

func TestOutlookExactEmailLookupDryRun(t *testing.T) {
	cmd := newOutlookPersonalContacts(&rootFlags{dryRun: true})
	var out bytes.Buffer
	cmd.SetOut(&out)
	cmd.SetErr(&bytes.Buffer{})
	cmd.SetArgs([]string{"lookup-email", "--email", "a@example.invalid"})
	if err := cmd.Execute(); err != nil {
		t.Fatal(err)
	}
	var data map[string]any
	if json.Unmarshal(out.Bytes(), &data) != nil || data["dry_run"] != true || !strings.Contains(data["scope"].(string), "directory") {
		t.Fatal("incorrect lookup preview")
	}
	cmd = newOutlookPersonalContacts(&rootFlags{dryRun: true})
	cmd.SetOut(&bytes.Buffer{})
	cmd.SetErr(&bytes.Buffer{})
	cmd.SetArgs([]string{"lookup-email", "--email", "not-an-email"})
	if cmd.Execute() == nil {
		t.Fatal("accepted invalid email")
	}
}
