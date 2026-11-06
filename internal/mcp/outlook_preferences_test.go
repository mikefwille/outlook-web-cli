package mcp

import (
	"testing"

	"github.com/mark3labs/mcp-go/server"
)

func TestPreferencesMCPBoundary(t *testing.T) {
	for _, ro := range []bool{false, true} {
		if ro {
			t.Setenv("OUTLOOK_WEB_MCP_READ_ONLY", "1")
		} else {
			t.Setenv("OUTLOOK_WEB_MCP_READ_ONLY", "")
		}
		s := server.NewMCPServer("test", "test")
		RegisterTools(s)
		tools := s.ListTools()
		for _, name := range []string{"rules_list", "settings_auto_replies"} {
			v := tools[name]
			if v == nil || v.Tool.Annotations.ReadOnlyHint == nil || !*v.Tool.Annotations.ReadOnlyHint {
				t.Errorf("missing read-only tool %s", name)
			}
		}
		v := tools["settings_update"]
		if ro {
			if v != nil {
				t.Fatal("read-only server exposed settings update")
			}
			continue
		}
		if v == nil {
			t.Fatal("settings update missing")
		}
		if _, ok := v.Tool.InputSchema.Properties["execute"]; !ok {
			t.Fatal("execute flag missing")
		}
	}
}
