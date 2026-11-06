package mcp

import (
	"testing"

	"github.com/mark3labs/mcp-go/server"
)

func TestAccountsMCPBoundary(t *testing.T) {
	for _, readonly := range []bool{false, true} {
		if readonly {
			t.Setenv("OUTLOOK_WEB_MCP_READ_ONLY", "1")
		} else {
			t.Setenv("OUTLOOK_WEB_MCP_READ_ONLY", "")
		}
		s := server.NewMCPServer("test", "test")
		RegisterTools(s)
		tools := s.ListTools()
		if tools["accounts_list"] == nil {
			t.Fatal("accounts list missing")
		}
		for _, name := range []string{"accounts_use", "accounts_register_current"} {
			tool := tools[name]
			if readonly {
				if tool != nil {
					t.Fatalf("read-only exposed %s", name)
				}
				continue
			}
			if tool == nil {
				t.Fatalf("missing %s", name)
			}
			if _, ok := tool.Tool.InputSchema.Properties["execute"]; !ok {
				t.Fatalf("missing explicit execute flag on %s", name)
			}
			for field := range tool.Tool.InputSchema.Properties {
				if field == "path" || field == "session-file" || field == "operations-file" {
					t.Fatalf("private path exposed by %s", name)
				}
			}
		}
	}
}

func TestDirectoryProfileGetMCPIsReadOnly(t *testing.T) {
	t.Setenv("OUTLOOK_WEB_MCP_READ_ONLY", "1")
	s := server.NewMCPServer("test", "test")
	RegisterTools(s)
	tool := s.ListTools()["directory_get"]
	if tool == nil || tool.Tool.Annotations.ReadOnlyHint == nil || !*tool.Tool.Annotations.ReadOnlyHint {
		t.Fatal("directory profile get is missing or not read-only")
	}
	if _, ok := tool.Tool.InputSchema.Properties["people-id"]; !ok {
		t.Fatal("directory profile get does not accept its People ID")
	}
	for _, name := range []string{"directory_manager", "directory_reports"} {
		tool = s.ListTools()[name]
		if tool == nil || tool.Tool.Annotations.ReadOnlyHint == nil || !*tool.Tool.Annotations.ReadOnlyHint {
			t.Fatalf("%s is missing or not read-only", name)
		}
		for _, field := range []string{"object-id", "email", "upn"} {
			if _, ok := tool.Tool.InputSchema.Properties[field]; !ok {
				t.Fatalf("%s missing %s", name, field)
			}
		}
	}
}
