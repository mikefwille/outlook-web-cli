package cli

import "testing"

func TestOutlookConversationActionsAreRegistered(t *testing.T) {
	root := RootCmd()
	threads := findSubcommand(root, "threads")
	if threads == nil {
		t.Fatal("threads group missing")
	}
	for _, name := range []string{"mark-read", "mark-unread", "flag", "complete", "move"} {
		cmd := findSubcommand(threads, name)
		if cmd == nil || cmd.Flags().Lookup("folder-id") == nil || cmd.Flags().Lookup("execute") == nil {
			t.Fatalf("conversation action %s not registered", name)
		}
	}
}
