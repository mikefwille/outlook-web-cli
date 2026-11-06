package cli

import (
	"bytes"
	"encoding/json"
	"reflect"
	"testing"

	"github.com/spf13/cobra"
)

func runAlias(t *testing.T, cmd *cobra.Command, args ...string) map[string]any {
	t.Helper()
	var out bytes.Buffer
	cmd.SetOut(&out)
	cmd.SetErr(&bytes.Buffer{})
	cmd.SetArgs(args)
	if err := cmd.Execute(); err != nil {
		t.Fatal(err)
	}
	var value map[string]any
	if json.Unmarshal(out.Bytes(), &value) != nil {
		t.Fatal("invalid structured result")
	}
	return value
}
func TestSyncParentMatchesFolderPreview(t *testing.T) {
	root := runAlias(t, newMailSyncCmd(&rootFlags{}), "--folder-id", "synthetic", "--max-items", "20")
	child := runAlias(t, newMailSyncCmd(&rootFlags{}), "folder", "--folder-id", "synthetic", "--max-items", "20")
	if !reflect.DeepEqual(root, child) || root["preview"] != true || root["server_writes"] != false {
		t.Fatal("sync paths diverged")
	}
}
func TestMessageCategoryParentMatchesNamedActions(t *testing.T) {
	for _, kind := range []string{"get", "add", "remove"} {
		args := []string{"synthetic", "--action", kind}
		childArgs := []string{kind, "synthetic"}
		if kind != "get" {
			args = append(args, "--category", "Example")
			childArgs = append(childArgs, "--category", "Example")
		}
		root := runAlias(t, categoryTestRoot(), append([]string{"categories"}, args...)...)
		child := runAlias(t, categoryTestRoot(), append([]string{"categories"}, childArgs...)...)
		if !reflect.DeepEqual(root, child) {
			t.Fatalf("category %s paths diverged", kind)
		}
	}
	cmd := categoryTestRoot()
	cmd.SetOut(&bytes.Buffer{})
	cmd.SetErr(&bytes.Buffer{})
	cmd.SetArgs([]string{"categories", "synthetic", "--action", "get", "--execute"})
	if cmd.Execute() == nil {
		t.Fatal("get accepted an execution flag")
	}
}

func categoryTestRoot() *cobra.Command {
	root := &cobra.Command{Use: "outlook-test", SilenceErrors: true, SilenceUsage: true}
	root.AddCommand(newMailCategoriesCmd(&rootFlags{dryRun: true}))
	return root
}
