package cli

import (
	"os"
	"path/filepath"
	"testing"
)

func TestMailExportCreatesPrivateFileWithoutOverwrite(t *testing.T) {
	path := filepath.Join(t.TempDir(), "message.json")
	if err := writeNewMailExport(path, []byte(`{"synthetic":true}`)); err != nil {
		t.Fatal(err)
	}
	info, err := os.Stat(path)
	if err != nil || info.Mode().Perm()&0077 != 0 {
		t.Fatal("export permissions are not private")
	}
	if err := writeNewMailExport(path, []byte("replacement")); err == nil {
		t.Fatal("existing file overwritten")
	}
	data, err := os.ReadFile(path)
	if err != nil || string(data) != `{"synthetic":true}` {
		t.Fatal("existing file changed")
	}
}
