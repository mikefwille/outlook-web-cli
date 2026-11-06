package mcp

import (
	"github.com/mark3labs/mcp-go/server"
	"testing"
)

func TestMailCopyMCPRequiresExplicitWrite(t *testing.T) {
	for _, ro := range []bool{false, true} {
		if ro {
			t.Setenv("OUTLOOK_WEB_MCP_READ_ONLY", "1")
		} else {
			t.Setenv("OUTLOOK_WEB_MCP_READ_ONLY", "")
		}
		s := server.NewMCPServer("test", "test")
		RegisterTools(s)
		tool := s.ListTools()["mail_copy"]
		if ro {
			if tool != nil {
				t.Fatal("read-only exposed copy")
			}
			continue
		}
		if tool == nil {
			t.Fatal("mail copy missing")
		}
		if tool.Tool.Annotations.ReadOnlyHint != nil && *tool.Tool.Annotations.ReadOnlyHint {
			t.Fatal("copy marked read-only")
		}
		for _, name := range []string{"id", "folder-id", "change-key", "execute"} {
			if _, ok := tool.Tool.InputSchema.Properties[name]; !ok {
				t.Fatalf("missing %s", name)
			}
		}
	}
}
