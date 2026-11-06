package mcp

import (
	"github.com/mark3labs/mcp-go/server"
	"testing"
)

func TestBriefingMCPIsBoundedReadWithoutPaths(t *testing.T) {
	for _, ro := range []bool{false, true} {
		if ro {
			t.Setenv("OUTLOOK_WEB_MCP_READ_ONLY", "1")
		} else {
			t.Setenv("OUTLOOK_WEB_MCP_READ_ONLY", "")
		}
		s := server.NewMCPServer("test", "test")
		RegisterTools(s)
		x := s.ListTools()["briefing"]
		if x == nil || x.Tool.Annotations.ReadOnlyHint == nil || !*x.Tool.Annotations.ReadOnlyHint {
			t.Fatal("briefing read missing")
		}
		for _, key := range []string{"folder-id", "calendar-id", "from", "to", "limit"} {
			if _, ok := x.Tool.InputSchema.Properties[key]; !ok {
				t.Fatalf("missing %s", key)
			}
		}
		if _, ok := x.Tool.InputSchema.Properties["output"]; ok {
			t.Fatal("local output path exposed")
		}
	}
}
