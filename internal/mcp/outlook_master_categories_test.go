package mcp

import (
	"github.com/mark3labs/mcp-go/server"
	"testing"
)

func TestMasterCategoryMCPBoundary(t *testing.T) {
	for _, readOnly := range []bool{false, true} {
		if readOnly {
			t.Setenv("OUTLOOK_WEB_MCP_READ_ONLY", "1")
		} else {
			t.Setenv("OUTLOOK_WEB_MCP_READ_ONLY", "")
		}
		s := server.NewMCPServer("test", "test")
		RegisterTools(s)
		tools := s.ListTools()
		for _, name := range []string{"categories_create", "categories_update", "categories_delete"} {
			x := tools[name]
			if readOnly {
				if x != nil {
					t.Fatalf("read-only exposed %s", name)
				}
				continue
			}
			if x == nil {
				t.Fatalf("missing %s", name)
			}
			for _, key := range []string{"execute", "if-version"} {
				if _, ok := x.Tool.InputSchema.Properties[key]; !ok {
					t.Fatalf("%s missing %s", name, key)
				}
			}
			if name != "categories_create" {
				if _, ok := x.Tool.InputSchema.Properties["allow-affect-items"]; !ok {
					t.Fatal("missing item-effect acknowledgement")
				}
			}
		}
	}
}
