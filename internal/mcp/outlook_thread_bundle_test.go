package mcp

import (
	"github.com/mark3labs/mcp-go/server"
	"testing"
)

func TestThreadBundleMCPDoesNotExposePaths(t *testing.T) {
	for _, ro := range []bool{false, true} {
		if ro {
			t.Setenv("OUTLOOK_WEB_MCP_READ_ONLY", "1")
		} else {
			t.Setenv("OUTLOOK_WEB_MCP_READ_ONLY", "")
		}
		s := server.NewMCPServer("test", "test")
		RegisterTools(s)
		x := s.ListTools()["threads_bundle"]
		if x == nil || x.Tool.Annotations.ReadOnlyHint == nil || !*x.Tool.Annotations.ReadOnlyHint {
			t.Fatal("bounded thread read missing")
		}
		if _, ok := x.Tool.InputSchema.Properties["conversation-id"]; !ok {
			t.Fatal("thread ID missing")
		}
		if _, ok := x.Tool.InputSchema.Properties["output"]; ok {
			t.Fatal("local output path exposed")
		}
	}
}
