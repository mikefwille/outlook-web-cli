package mcp

import (
	"github.com/mark3labs/mcp-go/server"
	"testing"
)

func TestRuleMCPBoundary(t *testing.T) {
	for _, ro := range []bool{false, true} {
		if ro {
			t.Setenv("OUTLOOK_WEB_MCP_READ_ONLY", "1")
		} else {
			t.Setenv("OUTLOOK_WEB_MCP_READ_ONLY", "")
		}
		s := server.NewMCPServer("test", "test")
		RegisterTools(s)
		tools := s.ListTools()
		if tools["rules_get"] == nil {
			t.Fatal("rules get missing")
		}
		for _, name := range []string{"rules_create", "rules_enable", "rules_disable", "rules_update", "rules_reorder", "rules_delete"} {
			tool := tools[name]
			if ro {
				if tool != nil {
					t.Fatalf("read-only exposed %s", name)
				}
				continue
			}
			if tool == nil {
				t.Fatalf("tool missing %s", name)
			}
			if _, ok := tool.Tool.InputSchema.Properties["execute"]; !ok {
				t.Fatalf("missing execute on %s", name)
			}
		}
		if !ro {
			if _, ok := tools["rules_create"].Tool.InputSchema.Properties["allow-enabled"]; !ok {
				t.Fatal("missing enabled acknowledgement")
			}
			for _, field := range []string{"subject-contains", "category"} {
				if _, ok := tools["rules_update"].Tool.InputSchema.Properties[field]; !ok {
					t.Fatalf("rule update missing %s", field)
				}
			}
			for name, fields := range map[string][]string{"rules_update": {"allow-enabled", "restore-disabled"}, "rules_reorder": {"id", "allow-state-reapply"}} {
				for _, field := range fields {
					if _, ok := tools[name].Tool.InputSchema.Properties[field]; !ok {
						t.Fatalf("%s missing %s", name, field)
					}
				}
			}
		}
	}
}
