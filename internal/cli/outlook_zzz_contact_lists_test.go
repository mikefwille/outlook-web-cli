package cli

import (
	"bytes"
	"encoding/json"
	"testing"
)

func TestContactListPreviewRequiresGenericValidInputs(t *testing.T) {
	for _, tc := range []struct {
		args []string
		ok   bool
	}{
		{[]string{"create", "--name", "Friends", "--email", "one@example.invalid"}, true},
		{[]string{"create", "--name", "Friends", "--email", "bad"}, false},
		{[]string{"update", "list-id", "--notes", ""}, true},
		{[]string{"add-member", "list-id", "--email", "two@example.invalid"}, true},
		{[]string{"remove-member", "../bad", "--email", "two@example.invalid"}, false},
		{[]string{"delete", "list-id"}, true},
	} {
		cmd := newContactListsCommand(&rootFlags{dryRun: true})
		var out bytes.Buffer
		cmd.SetOut(&out)
		cmd.SetErr(&bytes.Buffer{})
		cmd.SetArgs(tc.args)
		err := cmd.Execute()
		if (err == nil) != tc.ok {
			t.Fatalf("%v: unexpected error: %v", tc.args, err)
		}
		if tc.ok {
			var x map[string]any
			if json.Unmarshal(out.Bytes(), &x) != nil || x["preview"] != true {
				t.Fatal("preview missing")
			}
		}
	}
}
