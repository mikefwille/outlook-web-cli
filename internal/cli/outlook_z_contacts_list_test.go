package cli

import (
	"bytes"
	"encoding/json"
	"github.com/spf13/cobra"
	"testing"
)

func TestOutlookContactsListPreview(t *testing.T) {
	for _, tc := range []struct {
		args  []string
		valid bool
	}{{[]string{"list"}, true}, {[]string{"list", "--limit", "0"}, false}, {[]string{"list", "--offset", "1"}, false}} {
		group := &cobra.Command{Use: "contacts"}
		addOutlookContactsList(group, &rootFlags{dryRun: true})
		var out bytes.Buffer
		group.SetOut(&out)
		group.SetErr(&bytes.Buffer{})
		group.SetArgs(tc.args)
		err := group.Execute()
		if (err == nil) != tc.valid {
			t.Fatalf("unexpected result %v for %v", err, tc.args)
		}
		if tc.valid {
			var r map[string]any
			if json.Unmarshal(out.Bytes(), &r) != nil || r["dry_run"] != true {
				t.Fatal("invalid preview")
			}
		}
	}
	root := newRootCmd(&rootFlags{})
	group := findSubcommand(root, "contacts")
	if group == nil || findSubcommand(group, "list") == nil {
		t.Fatal("list not registered")
	}
}
