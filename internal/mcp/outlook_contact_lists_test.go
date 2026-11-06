package mcp

import (
	"testing"

	"github.com/mark3labs/mcp-go/server"
)

func TestContactListMCPBoundary(t *testing.T) {
	for _, readOnly := range []bool{false, true} {
		if readOnly {
			t.Setenv("OUTLOOK_WEB_MCP_READ_ONLY", "1")
		} else {
			t.Setenv("OUTLOOK_WEB_MCP_READ_ONLY", "")
		}
		s := server.NewMCPServer("test", "test")
		RegisterTools(s)
		tools := s.ListTools()
		for _, name := range []string{"contacts_lists_list", "contacts_lists_get", "contact_lists_list", "contact_lists_get"} {
			entry := tools[name]
			if entry == nil || entry.Tool.Annotations.ReadOnlyHint == nil || !*entry.Tool.Annotations.ReadOnlyHint {
				t.Errorf("missing read-only tool %s", name)
			}
		}
		for _, name := range []string{"contacts_lists_create", "contacts_lists_update", "contacts_lists_add_member", "contacts_lists_remove_member", "contacts_lists_delete", "contact_lists_create", "contact_lists_update", "contact_lists_add_member", "contact_lists_remove_member", "contact_lists_delete"} {
			entry := tools[name]
			if readOnly {
				if entry != nil {
					t.Errorf("read-only server exposed %s", name)
				}
				continue
			}
			if entry == nil {
				t.Errorf("missing write tool %s", name)
				continue
			}
			if entry.Tool.Annotations.ReadOnlyHint != nil && *entry.Tool.Annotations.ReadOnlyHint {
				t.Errorf("write tool %s marked read-only", name)
			}
			if _, ok := entry.Tool.InputSchema.Properties["execute"]; !ok {
				t.Errorf("%s lacks explicit execute", name)
			}
		}
	}
}
