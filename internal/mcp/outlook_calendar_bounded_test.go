package mcp

import (
	"github.com/mark3labs/mcp-go/server"
	"testing"
)

func TestCalendarBoundedReadsMCPBoundary(t *testing.T) {
	for _, readOnly := range []bool{false, true} {
		if readOnly {
			t.Setenv("OUTLOOK_WEB_MCP_READ_ONLY", "1")
		} else {
			t.Setenv("OUTLOOK_WEB_MCP_READ_ONLY", "")
		}
		s := server.NewMCPServer("test", "test")
		RegisterTools(s)
		tools := s.ListTools()
		for _, name := range []string{"calendar_search", "calendar_find_times"} {
			x := tools[name]
			if x == nil {
				t.Fatalf("missing %s", name)
			}
			if x.Tool.Annotations.ReadOnlyHint == nil || !*x.Tool.Annotations.ReadOnlyHint {
				t.Fatalf("%s is not read-only", name)
			}
			for _, key := range []string{"from", "to"} {
				if _, ok := x.Tool.InputSchema.Properties[key]; !ok {
					t.Fatalf("%s missing %s", name, key)
				}
			}
		}
		if _, ok := tools["calendar_find_times"].Tool.InputSchema.Properties["principal"]; ok {
			t.Fatal("unverified other-principal availability exposed")
		}
	}
}
