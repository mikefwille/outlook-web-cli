package cli

import (
	"bytes"
	"os"
	"path/filepath"
	"testing"
)

func TestLogoutOnlyDeactivatesRecognizedPointer(t *testing.T) {
	dir := t.TempDir()
	if err := os.Chmod(dir, 0700); err != nil {
		t.Fatal(err)
	}
	pointer := filepath.Join(dir, "current-outlook.json")
	saved := filepath.Join(dir, "session-outlook-test.json")
	if err := os.WriteFile(pointer, []byte(`{"version":1,"service":"outlook","export_id":"abc-1234abcd"}`), 0600); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(saved, []byte("kept"), 0600); err != nil {
		t.Fatal(err)
	}
	t.Setenv("OUTLOOK_WEB_SESSION_POINTER", pointer)
	preview := newOutlookLogoutCmd(&rootFlags{asJSON: true})
	preview.SetOut(&bytes.Buffer{})
	preview.SetArgs(nil)
	if err := preview.Execute(); err != nil {
		t.Fatal(err)
	}
	if _, err := os.Stat(pointer); err != nil {
		t.Fatal("preview changed session")
	}
	cmd := newOutlookLogoutCmd(&rootFlags{asJSON: true})
	cmd.SetOut(&bytes.Buffer{})
	cmd.SetArgs([]string{"--execute"})
	if err := cmd.Execute(); err != nil {
		t.Fatal(err)
	}
	if _, err := os.Stat(pointer); !os.IsNotExist(err) {
		t.Fatal("pointer remains active")
	}
	if data, err := os.ReadFile(saved); err != nil || string(data) != "kept" {
		t.Fatal("saved export was modified")
	}
}
func TestLogoutRejectsUnsafeOrLockedPointer(t *testing.T) {
	dir := t.TempDir()
	if err := os.Chmod(dir, 0700); err != nil {
		t.Fatal(err)
	}
	pointer := filepath.Join(dir, "current-outlook.json")
	t.Setenv("OUTLOOK_WEB_SESSION_POINTER", pointer)
	if err := os.WriteFile(pointer, []byte(`{"version":1,"service":"teams","export_id":"abc-1234abcd"}`), 0600); err != nil {
		t.Fatal(err)
	}
	if _, err := safeOutlookPointer(pointer); err == nil {
		t.Fatal("foreign session pointer accepted")
	}
	if err := os.WriteFile(pointer, []byte(`{"version":1,"service":"outlook","export_id":"abc-1234abcd"}`), 0600); err != nil {
		t.Fatal(err)
	}
	if err := os.Mkdir(filepath.Join(dir, ".m365-session-outlook.lock"), 0700); err != nil {
		t.Fatal(err)
	}
	cmd := newOutlookLogoutCmd(&rootFlags{asJSON: true})
	cmd.SilenceErrors = true
	cmd.SilenceUsage = true
	cmd.SetArgs([]string{"--execute"})
	if err := cmd.Execute(); err == nil {
		t.Fatal("concurrently locked session was deactivated")
	}
	if _, err := os.Stat(pointer); err != nil {
		t.Fatal("locked pointer changed")
	}
}
