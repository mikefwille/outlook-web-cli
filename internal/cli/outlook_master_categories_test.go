package cli

import (
	"bytes"
	"encoding/json"
	"testing"
)

func TestMasterCategoryDryRunExposesItemWideEffects(t *testing.T) {
	for _, kind := range []string{"create", "update", "delete"} {
		cmd := newMasterCategoryAction(&rootFlags{dryRun: true}, kind)
		var out bytes.Buffer
		cmd.SetOut(&out)
		cmd.SetErr(&bytes.Buffer{})
		args := []string{}
		if kind != "create" {
			args = append(args, "sample-id")
		}
		if kind != "delete" {
			args = append(args, "--name", "Fixture", "--color", "20")
		}
		cmd.SetArgs(args)
		if err := cmd.Execute(); err != nil {
			t.Fatal(err)
		}
		var result map[string]any
		if json.Unmarshal(out.Bytes(), &result) != nil || result["preview"] != true {
			t.Fatal("invalid preview")
		}
		if kind != "create" && (result["requires_allow_affect_items"] != true || result["item_effects"] == nil) {
			t.Fatal("preview omitted category-wide effects")
		}
	}
}
