package cli

import (
	"bytes"
	"encoding/json"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/spf13/cobra"
)

func TestDraftTextInputFileAndStdin(t *testing.T) {
	file := filepath.Join(t.TempDir(), "body.txt")
	if err := os.WriteFile(file, []byte("first\nsecond"), 0600); err != nil {
		t.Fatal(err)
	}
	for _, tc := range []struct {
		args        []string
		stdin, want string
		ok          bool
	}{
		{[]string{"--text-file", file}, "", "first\nsecond", true},
		{[]string{"--text-file", "-"}, "from stdin\n", "from stdin\n", true},
		{[]string{"--text", "direct"}, "", "direct", true},
		{[]string{"--text", "direct", "--text-file", file}, "", "", false},
		{[]string{"--text-file", "-"}, strings.Repeat("x", (1<<20)+1), "", false},
		{[]string{"--text-file", "-"}, string([]byte{0xff}), "", false},
	} {
		cmd := &cobra.Command{}
		var text, path string
		cmd.Flags().StringVar(&text, "text", "", "")
		cmd.Flags().StringVar(&path, "text-file", "", "")
		if err := cmd.ParseFlags(tc.args); err != nil {
			t.Fatal(err)
		}
		cmd.SetIn(strings.NewReader(tc.stdin))
		got, err := draftTextInput(cmd, text, path)
		if (err == nil) != tc.ok || tc.ok && got != tc.want {
			t.Fatalf("input case failed: ok=%v err=%v", tc.ok, err)
		}
	}
}

func TestRichDraftPreviewHasMarkupAndPlainCompanion(t *testing.T) {
	group := newDraftsCmd(&rootFlags{})
	var out bytes.Buffer
	group.SetOut(&out)
	group.SetErr(&bytes.Buffer{})
	group.SetArgs([]string{"create", "--subject", "Synthetic", "--html", "<div>Hello <b>world</b></div>"})
	if err := group.Execute(); err != nil {
		t.Fatal(err)
	}
	var result map[string]any
	if json.Unmarshal(out.Bytes(), &result) != nil || result["html"] != "<div>Hello <b>world</b></div>" || result["text"] != "Hello world" || result["sends_mail"] != false {
		t.Fatal("rich preview missing markup or plain text")
	}
}
