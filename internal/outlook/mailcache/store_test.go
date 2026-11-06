package mailcache

import (
	"context"
	"encoding/json"
	"os"
	"path/filepath"
	"testing"
)

func testItem(id, key, folder, subject string) Item {
	b, _ := json.Marshal(map[string]any{"ItemId": map[string]string{"Id": id, "ChangeKey": key}, "ParentFolderId": map[string]string{"Id": folder}, "Subject": subject})
	return Item{ID: id, ChangeKey: key, Data: b}
}
func TestReconciliationCommitsPagesAndPreservesLastCompleteSnapshot(t *testing.T) {
	ctx := context.Background()
	root := t.TempDir()
	if err := os.Chmod(root, 0700); err != nil {
		t.Fatal(err)
	}
	s, err := Open(root, "account-1", true)
	if err != nil {
		t.Fatal(err)
	}
	defer s.Close()
	r, err := s.Start(ctx, "folder-a")
	if err != nil {
		t.Fatal(err)
	}
	r, err = s.AppendPage(ctx, r, Page{Items: []Item{testItem("1", "v1", "folder-a", "first")}, Next: 1, Total: 2}, 10)
	if err != nil {
		t.Fatal(err)
	}
	resumed, err := s.Get(ctx, r.ID)
	if err != nil || resumed.Offset != 1 || resumed.Count != 1 || resumed.Status != "scan" {
		t.Fatal("page and checkpoint were not committed together")
	}
	if _, err = s.AppendPage(ctx, resumed, Page{Items: []Item{testItem("2", "v1", "folder-a", "second")}, Next: 2, Total: 3, Last: true}, 10); err == nil {
		t.Fatal("inconsistent folder total accepted")
	}
	if got, _ := s.Get(ctx, r.ID); got.Count != 1 || got.Offset != 1 {
		t.Fatal("failed page changed checkpoint")
	}
	r, err = s.AppendPage(ctx, resumed, Page{Items: []Item{testItem("2", "v1", "folder-a", "second")}, Next: 2, Total: 2, Last: true}, 10)
	if err != nil {
		t.Fatal(err)
	}
	versions, err := s.Versions(ctx, r.ID)
	if err != nil || len(versions) != 2 {
		t.Fatal("staged versions missing")
	}
	if err = s.Publish(ctx, r); err != nil {
		t.Fatal(err)
	}
	hits, err := s.Search(ctx, "first", "folder-a", 10)
	if err != nil || len(hits) != 1 || hits[0].SyncedAt == "" {
		t.Fatal("published snapshot unavailable")
	}
	incomplete, err := s.Start(ctx, "folder-a")
	if err != nil {
		t.Fatal(err)
	}
	_, err = s.AppendPage(ctx, incomplete, Page{Items: []Item{testItem("3", "v1", "folder-a", "new")}, Next: 1, Total: 2}, 10)
	if err != nil {
		t.Fatal(err)
	}
	hits, err = s.Search(ctx, "first", "folder-a", 10)
	if err != nil || len(hits) != 1 {
		t.Fatal("partial reconciliation overwrote completed snapshot")
	}
	other, err := Open(root, "account-2", true)
	if err != nil {
		t.Fatal(err)
	}
	defer other.Close()
	hits, err = other.Search(ctx, "first", "folder-a", 10)
	if err != nil || len(hits) != 0 {
		t.Fatal("account cache isolation failed")
	}
}
func TestRejectInsecureCacheDirectory(t *testing.T) {
	root := t.TempDir()
	if err := os.Chmod(root, 0755); err != nil {
		t.Fatal(err)
	}
	if s, err := Open(root, "account", true); err == nil {
		s.Close()
		t.Fatal("permissive root accepted")
	}
	if err := os.Chmod(root, 0700); err != nil {
		t.Fatal(err)
	}
	target := t.TempDir()
	if err := os.Symlink(target, filepath.Join(root, "mail-cache")); err != nil {
		t.Fatal(err)
	}
	if s, err := Open(root, "account", true); err == nil {
		s.Close()
		t.Fatal("symlink cache accepted")
	}
}
func TestParsePageRequiresBoundedCompleteMetadata(t *testing.T) {
	raw := []byte(`{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Success","ResponseCode":"NoError","RootFolder":{"Items":[{"ItemId":{"Id":"x","ChangeKey":"v"},"ParentFolderId":{"Id":"f"}}],"IndexedPagingOffset":1,"TotalItemsInView":1,"IncludesLastItemInRange":true}}]}}}`)
	p, err := ParsePage(raw, "f")
	if err != nil || len(p.Items) != 1 || !p.Last {
		t.Fatal("valid page rejected")
	}
	if _, err = ParsePage(raw, "other"); err == nil {
		t.Fatal("cross-folder item accepted")
	}
	if _, err = ParsePage([]byte(`{"Body":{}}`), "f"); err == nil {
		t.Fatal("missing metadata accepted")
	}
}

func TestReconcileVerifiesBeforePublishingAndCanResume(t *testing.T) {
	ctx := context.Background()
	root := t.TempDir()
	if err := os.Chmod(root, 0700); err != nil {
		t.Fatal(err)
	}
	s, err := Open(root, "account", true)
	if err != nil {
		t.Fatal(err)
	}
	defer s.Close()
	r, err := s.Start(ctx, "folder")
	if err != nil {
		t.Fatal(err)
	}
	calls := 0
	fetch := func(_ context.Context, offset, limit int) (Page, error) {
		calls++
		if offset != 0 || limit != 100 {
			t.Fatal("unexpected request")
		}
		return Page{Items: []Item{testItem("1", "v1", "folder", "hello")}, Next: 1, Total: 1, Last: true}, nil
	}
	r, err = Reconcile(ctx, s, r, fetch, 100)
	if err != nil || r.Status != "complete" || calls != 2 {
		t.Fatal("complete verified run failed")
	}
	r, err = s.Start(ctx, "folder")
	if err != nil {
		t.Fatal(err)
	}
	calls = 0
	changed := func(_ context.Context, offset, limit int) (Page, error) {
		calls++
		key := "v1"
		if calls == 2 {
			key = "v2"
		}
		return Page{Items: []Item{testItem("1", key, "folder", "changed")}, Next: 1, Total: 1, Last: true}, nil
	}
	if _, err = Reconcile(ctx, s, r, changed, 100); err == nil {
		t.Fatal("changed folder was published")
	}
	r, err = s.Get(ctx, r.ID)
	if err != nil || r.Status != "stale" {
		t.Fatal("changed run not marked stale")
	}
	hits, err := s.Search(ctx, "hello", "folder", 10)
	if err != nil || len(hits) != 1 {
		t.Fatal("changed scan replaced prior snapshot")
	}
	r, err = s.Start(ctx, "folder")
	if err != nil {
		t.Fatal(err)
	}
	r, err = s.AppendPage(ctx, r, Page{Items: []Item{testItem("1", "v1", "folder", "hello")}, Next: 1, Total: 1, Last: true}, 100)
	if err != nil {
		t.Fatal(err)
	}
	calls = 0
	r, err = Reconcile(ctx, s, r, fetch, 100)
	if err != nil || r.Status != "complete" || calls != 1 {
		t.Fatal("verification checkpoint did not resume")
	}
}

func TestReconcileResumesAnInterruptedPageScan(t *testing.T) {
	ctx := context.Background()
	root := t.TempDir()
	if err := os.Chmod(root, 0700); err != nil {
		t.Fatal(err)
	}
	s, err := Open(root, "account", true)
	if err != nil {
		t.Fatal(err)
	}
	defer s.Close()
	r, err := s.Start(ctx, "folder")
	if err != nil {
		t.Fatal(err)
	}
	r, err = s.AppendPage(ctx, r, Page{Items: []Item{testItem("1", "v1", "folder", "first")}, Next: 1, Total: 2}, 10)
	if err != nil {
		t.Fatal(err)
	}
	var offsets []int
	fetch := func(_ context.Context, offset, limit int) (Page, error) {
		offsets = append(offsets, offset)
		if offset == 0 {
			return Page{Items: []Item{testItem("1", "v1", "folder", "first")}, Next: 1, Total: 2}, nil
		}
		return Page{Items: []Item{testItem("2", "v1", "folder", "second")}, Next: 2, Total: 2, Last: true}, nil
	}
	r, err = Reconcile(ctx, s, r, fetch, 10)
	if err != nil || r.Status != "complete" || len(offsets) != 3 || offsets[0] != 1 || offsets[1] != 0 || offsets[2] != 1 {
		t.Fatal("resume did not continue from its checkpoint and verify from the start")
	}
}

func TestSQLIsBoundedAndAuthorizesOnlyPublishedReads(t *testing.T) {
	ctx := context.Background()
	root := t.TempDir()
	if err := os.Chmod(root, 0700); err != nil {
		t.Fatal(err)
	}
	s, err := Open(root, "account", true)
	if err != nil {
		t.Fatal(err)
	}
	defer s.Close()
	r, err := s.Start(ctx, "folder")
	if err != nil {
		t.Fatal(err)
	}
	r, err = s.AppendPage(ctx, r, Page{Items: []Item{testItem("1", "v1", "folder", "one"), testItem("2", "v1", "folder", "two")}, Next: 2, Total: 2, Last: true}, 10)
	if err != nil {
		t.Fatal(err)
	}
	if err = s.Publish(ctx, r); err != nil {
		t.Fatal(err)
	}
	result, err := s.Query(ctx, `SELECT json_extract(data, '$.Subject') AS subject FROM items ORDER BY subject`, 1)
	if err != nil || !result.Truncated || len(result.Rows) != 1 || result.Rows[0][0] != "one" {
		t.Fatalf("bounded SELECT failed: %v", err)
	}
	result, err = s.Query(ctx, `SELECT count(*) FROM items`, 10)
	if err != nil || len(result.Rows) != 1 || result.Rows[0][0] != int64(2) {
		t.Fatalf("aggregate failed: %v", err)
	}
	for _, q := range []string{`DELETE FROM items`, `SELECT * FROM items; DELETE FROM items`, `SELECT * FROM staged`, `SELECT * FROM sqlite_master`, `SELECT load_extension('bad')`, `SELECT readfile('/etc/passwd')`, `SELECT randomblob(1000000)`, `PRAGMA database_list`, `WITH RECURSIVE x(n) AS (SELECT 1) SELECT n FROM x`} {
		if _, err = s.Query(ctx, q, 10); err == nil {
			t.Fatalf("unsafe query accepted: %s", q)
		}
	}
	hits, err := s.Search(ctx, "one", "folder", 10)
	if err != nil || len(hits) != 1 {
		t.Fatal("rejected SQL altered the cache")
	}
}
