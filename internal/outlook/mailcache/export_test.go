package mailcache

import (
	"context"
	"os"
	"strings"
	"testing"
)

func TestExportUsesOnlySelectedPublishedItemsAndSnapshotTime(t *testing.T) {
	ctx := context.Background()
	root := t.TempDir()
	if err := os.Chmod(root, 0700); err != nil {
		t.Fatal(err)
	}
	s, err := Open(root, "account-one", true)
	if err != nil {
		t.Fatal(err)
	}
	defer s.Close()
	r, err := s.Start(ctx, "folder")
	if err != nil {
		t.Fatal(err)
	}
	r, err = s.AppendPage(ctx, r, Page{Items: []Item{testItem("one", "v1", "folder", "first"), testItem("two", "v2", "folder", "second")}, Next: 2, Total: 2, Last: true}, 10)
	if err != nil {
		t.Fatal(err)
	}
	if err = s.Publish(ctx, r); err != nil {
		t.Fatal(err)
	}
	out, err := s.Export(ctx, "folder", []string{"two", "one"})
	if err != nil || len(out.Items) != 2 || out.Items[0].ID != "two" || out.Items[1].ChangeKey != "v1" || out.SyncedAt == "" || !strings.Contains(string(out.Items[0].Data), "second") {
		t.Fatalf("selected snapshot invalid: %#v, %v", out, err)
	}
	for _, ids := range [][]string{nil, {"one", "missing"}, {"one", "one"}, {""}, {"one\ninvalid"}, make([]string, 101)} {
		if _, err := s.Export(ctx, "folder", ids); err == nil {
			t.Fatalf("invalid or incomplete export accepted: %d IDs", len(ids))
		}
	}
	if _, err := s.Export(ctx, "unsynced", []string{"one"}); err == nil {
		t.Fatal("unsynced folder accepted")
	}
	other, err := Open(root, "account-two", true)
	if err != nil {
		t.Fatal(err)
	}
	defer other.Close()
	if _, err := other.Export(ctx, "folder", []string{"one"}); err == nil {
		t.Fatal("cross-account export accepted")
	}
}
