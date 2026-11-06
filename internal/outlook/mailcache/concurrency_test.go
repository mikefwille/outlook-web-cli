package mailcache

import (
	"context"
	"os"
	"sync"
	"testing"
)

// Two independent store handles model overlapping CLI processes attempting
// the same page. Exactly one may advance the transactional checkpoint.
func TestConcurrentPageAppendAdvancesCheckpointOnce(t *testing.T) {
	ctx := context.Background()
	root := t.TempDir()
	if err := os.Chmod(root, 0700); err != nil {
		t.Fatal(err)
	}
	first, err := Open(root, "account", true)
	if err != nil {
		t.Fatal(err)
	}
	defer first.Close()
	second, err := Open(root, "account", true)
	if err != nil {
		t.Fatal(err)
	}
	defer second.Close()
	run, err := first.Start(ctx, "folder")
	if err != nil {
		t.Fatal(err)
	}
	page := Page{Items: []Item{testItem("one", "v1", "folder", "fixture")}, Next: 1, Total: 1, Last: true}
	start := make(chan struct{})
	results := make(chan error, 2)
	var wait sync.WaitGroup
	for _, store := range []*Store{first, second} {
		wait.Add(1)
		go func(store *Store) {
			defer wait.Done()
			<-start
			_, err := store.AppendPage(ctx, run, page, 10)
			results <- err
		}(store)
	}
	close(start)
	wait.Wait()
	close(results)
	successes := 0
	for err := range results {
		if err == nil {
			successes++
		}
	}
	if successes != 1 {
		t.Fatalf("expected one successful page append, got %d", successes)
	}
	checkpoint, err := first.Get(ctx, run.ID)
	if err != nil || checkpoint.Count != 1 || checkpoint.Offset != 1 || checkpoint.Status != "verify" {
		t.Fatalf("unexpected checkpoint after concurrent appends: %#v, %v", checkpoint, err)
	}
	versions, err := first.Versions(ctx, run.ID)
	if err != nil || len(versions) != 1 || versions["one"] != "v1" {
		t.Fatal("staged page did not remain consistent")
	}
}
