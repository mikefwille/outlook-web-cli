package capture

import (
	"encoding/json"
	"os"
	"path/filepath"
	"testing"
	"time"
)

func TestCurrentPointerFollowsOnlyOutlookExport(t *testing.T) {
	dir := t.TempDir()
	v := syntheticCapture(t, time.Now().Add(time.Hour))
	b, _ := json.Marshal(v)
	if err := os.WriteFile(filepath.Join(dir, "session-outlook-demo123-abcdef12.json"), b, 0o600); err != nil {
		t.Fatal(err)
	}
	p := filepath.Join(dir, "current-outlook.json")
	if err := os.WriteFile(p, []byte(`{"version":1,"service":"outlook","export_id":"demo123-abcdef12"}`), 0o600); err != nil {
		t.Fatal(err)
	}
	if _, _, err := LoadCurrentFindFolder(p); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(p, []byte(`{"version":1,"service":"teams","export_id":"demo123-abcdef12"}`), 0o600); err != nil {
		t.Fatal(err)
	}
	if _, _, err := LoadCurrentFindFolder(p); err == nil {
		t.Fatal("accepted other service")
	}
	if err := os.WriteFile(p, []byte(`{"version":1,"service":"outlook","export_id":"../../secret"}`), 0o600); err != nil {
		t.Fatal(err)
	}
	if _, _, err := LoadCurrentFindFolder(p); err == nil {
		t.Fatal("accepted path traversal")
	}
}
