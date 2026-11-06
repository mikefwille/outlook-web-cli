package cli

import (
	"bytes"
	"encoding/json"
	"os"
	"path/filepath"
	"testing"
)

func TestOutlookAttachmentPreviewAndHashGate(t *testing.T) {
	file := filepath.Join(t.TempDir(), "fixture.txt")
	if err := os.WriteFile(file, []byte("fixture bytes"), 0600); err != nil {
		t.Fatal(err)
	}
	cmd := newDraftAttachCmd(&rootFlags{})
	var out bytes.Buffer
	cmd.SetOut(&out)
	cmd.SetErr(&bytes.Buffer{})
	cmd.SetArgs([]string{"id", "--change-key", "key", "--file", file})
	if err := cmd.Execute(); err != nil {
		t.Fatal(err)
	}
	var result struct {
		Preview bool   `json:"preview"`
		Size    int    `json:"size"`
		Hash    string `json:"sha256"`
		Sends   bool   `json:"sends_mail"`
	}
	if json.Unmarshal(out.Bytes(), &result) != nil || !result.Preview || result.Size != len("fixture bytes") || len(result.Hash) != 64 || result.Sends {
		t.Fatal("invalid attachment preview")
	}
	cmd = newDraftAttachCmd(&rootFlags{})
	cmd.SetOut(&bytes.Buffer{})
	cmd.SetErr(&bytes.Buffer{})
	cmd.SetArgs([]string{"id", "--change-key", "key", "--file", file, "--sha256", stringsOfZeroes(), "--execute"})
	if err := cmd.Execute(); err == nil {
		t.Fatal("accepted mismatched digest")
	}
}

func stringsOfZeroes() string {
	return "0000000000000000000000000000000000000000000000000000000000000000"
}

func TestAttachmentPreviewAcceptsTwoMiBAndRejectsOverLimit(t *testing.T) {
	file := filepath.Join(t.TempDir(), "ordinary.bin")
	if err := os.WriteFile(file, make([]byte, 2<<20), 0600); err != nil {
		t.Fatal(err)
	}
	cmd := newDraftAttachCmd(&rootFlags{})
	var output bytes.Buffer
	cmd.SetOut(&output)
	cmd.SetArgs([]string{"id", "--change-key", "key", "--file", file})
	if err := cmd.Execute(); err != nil {
		t.Fatalf("2 MiB preview rejected: %v", err)
	}
	var got struct {
		Size    int  `json:"size"`
		Preview bool `json:"preview"`
	}
	if json.Unmarshal(output.Bytes(), &got) != nil || !got.Preview || got.Size != 2<<20 {
		t.Fatal("invalid 2 MiB preview")
	}
	if err := os.WriteFile(file, make([]byte, (4<<20)+1), 0600); err != nil {
		t.Fatal(err)
	}
	if _, _, err := readAttachmentFile(file); err == nil {
		t.Fatal("oversized attachment accepted")
	}
}
