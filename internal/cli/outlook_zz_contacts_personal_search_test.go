package cli

import (
	"bytes"
	"encoding/json"
	"github.com/spf13/cobra"
	"testing"
)

func TestOutlookPersonalSearchPreview(t *testing.T) {
	for _, tc := range []struct {
		args  []string
		valid bool
	}{{[]string{"search-personal", "Ada"}, true}, {[]string{"search-personal"}, false}, {[]string{"search-personal", "Ada", "--offset", "1"}, false}, {[]string{"search-personal", "Ada", "--limit", "0"}, false}} {
		group := &cobra.Command{Use: "contacts"}
		addOutlookPersonalSearch(group, &rootFlags{dryRun: true})
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
	if group == nil || findSubcommand(group, "search-personal") == nil {
		t.Fatal("personal search not registered")
	}
}
