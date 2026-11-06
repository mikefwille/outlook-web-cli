package cli

import (
	"fmt"
	"strings"

	"github.com/spf13/cobra"
)

func init() {
	registerNovelCommand(func(root *cobra.Command, flags *rootFlags) {
		var requested string
		cmd := &cobra.Command{Use: "capabilities", Short: "Inspect locally implemented Outlook command paths", Long: "Lists commands present in this build, or checks one exact command path. This is local discovery only; it does not establish account permission, live availability, or complete feature coverage.", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "local"}, RunE: func(cmd *cobra.Command, args []string) error {
			if len(args) != 0 {
				return usageErr(fmt.Errorf("capabilities takes --command, not positional arguments"))
			}
			entries := outlookWhichIndex(root)
			if requested == "" {
				return flags.printJSON(cmd, map[string]any{"source": "local", "commands": entries, "permission_checked": false, "live_checked": false})
			}
			requested = strings.Join(strings.Fields(requested), " ")
			for _, entry := range entries {
				if entry.Command == requested {
					return flags.printJSON(cmd, map[string]any{"source": "local", "command": requested, "implemented": true, "description": entry.Description, "permission_checked": false, "live_checked": false})
				}
			}
			return flags.printJSON(cmd, map[string]any{"source": "local", "command": requested, "implemented": false, "permission_checked": false, "live_checked": false})
		}}
		cmd.Flags().StringVar(&requested, "command", "", "Exact command path to inspect")
		root.AddCommand(cmd)
	})
}
