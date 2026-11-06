package cli

import (
	"strings"

	"github.com/spf13/cobra"
)

// Build discovery from the final command tree at invocation time so newly
// integrated operations and removed generic RPC handlers cannot drift apart.
func outlookWhichIndex(root *cobra.Command) []whichEntry {
	groups := map[string]bool{"account": true, "accounts": true, "attachments": true, "auth": true, "briefing": true, "calendar": true, "calendars": true, "capabilities": true, "categories": true, "contacts": true, "contact-lists": true, "deleted": true, "directory": true, "doctor": true, "drafts": true, "export": true, "folders": true, "mail": true, "search": true, "settings": true, "rules": true, "recipients": true, "sql": true, "sync": true, "threads": true}
	var result []whichEntry
	var walk func(*cobra.Command, []string, string)
	walk = func(cmd *cobra.Command, path []string, group string) {
		if cmd.Hidden {
			return
		}
		path = append(path, cmd.Name())
		if cmd.Runnable() {
			result = append(result, whichEntry{Command: strings.Join(path, " "), Description: cmd.Short, Group: group})
		}
		for _, child := range cmd.Commands() {
			walk(child, path, group)
		}
	}
	for _, group := range root.Commands() {
		if groups[group.Name()] {
			walk(group, nil, group.Name())
		}
	}
	return result
}

func init() {
	registerNovelCommand(func(root *cobra.Command, _ *rootFlags) {
		which := findSubcommand(root, "which")
		if which == nil {
			return
		}
		previous := which.PreRunE
		which.PreRunE = func(cmd *cobra.Command, args []string) error {
			whichIndex = outlookWhichIndex(root)
			if previous != nil {
				return previous(cmd, args)
			}
			return nil
		}
	})
}
