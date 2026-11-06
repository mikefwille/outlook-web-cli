package cobratree

import (
	"context"
	"testing"

	mcplib "github.com/mark3labs/mcp-go/mcp"
	"github.com/spf13/cobra"
)

func TestMCPExecutionFlagRequiresJSONBoolean(t *testing.T) {
	t.Setenv("OUTLOOK_WEB_MCP_READ_ONLY", "")
	cmd := &cobra.Command{Use: "drafts-create"}
	cmd.Flags().Bool("execute", false, "")
	for _, tc := range []struct {
		value any
		allow bool
	}{{"true", false}, {"false", false}, {1.0, false}, {nil, false}, {true, true}, {false, true}} {
		called := false
		handler := guardMirroredMCPCall(cmd, func(context.Context, mcplib.CallToolRequest) (*mcplib.CallToolResult, error) {
			called = true
			return mcplib.NewToolResultText("ok"), nil
		})
		result, err := handler(context.Background(), mcplib.CallToolRequest{Params: mcplib.CallToolParams{Arguments: map[string]any{"execute": tc.value}}})
		if err != nil || called != tc.allow || result == nil || result.IsError == tc.allow {
			t.Fatalf("boolean guard allowed=%v called=%v err=%v", tc.allow, called, err)
		}
	}
}
