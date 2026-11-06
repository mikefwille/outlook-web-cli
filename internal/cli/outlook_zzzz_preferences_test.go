package cli

import (
	"bytes"
	"encoding/json"
	"testing"
)

func TestFocusedInboxPreview(t *testing.T) {
	cmd := newFocusedInboxUpdate(&rootFlags{dryRun: true})
	var out bytes.Buffer
	cmd.SetOut(&out)
	cmd.SetErr(&bytes.Buffer{})
	cmd.SetArgs([]string{"--focused-inbox=false"})
	if err := cmd.Execute(); err != nil {
		t.Fatal(err)
	}
	var v map[string]any
	if json.Unmarshal(out.Bytes(), &v) != nil || v["preview"] != true || v["focused_inbox"] != false {
		t.Fatal("invalid focused-inbox preview")
	}
	cmd = newFocusedInboxUpdate(&rootFlags{dryRun: true})
	cmd.SetOut(&bytes.Buffer{})
	cmd.SetErr(&bytes.Buffer{})
	cmd.SetArgs([]string{})
	if cmd.Execute() == nil {
		t.Fatal("missing desired setting accepted")
	}
}
