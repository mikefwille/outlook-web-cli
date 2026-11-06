package cli

import (
	"fmt"
	"os"
	"path/filepath"
	"sync"
	"testing"
)

// A CLI invocation pins its first session even if another process changes the
// shared pointer; each concurrent caller must also receive independent headers.
func TestConcurrentInvocationKeepsPinnedSessionAndIndependentHeaders(t *testing.T) {
	dir := t.TempDir()
	if err := os.Chmod(dir, 0700); err != nil {
		t.Fatal(err)
	}
	pointer := filepath.Join(dir, "current-outlook.json")
	t.Setenv("OUTLOOK_WEB_SESSION_POINTER", pointer)
	t.Setenv("OUTLOOK_WEB_SESSION_FILE", "")
	accountFixture(t, dir, "exporta-aaaaaaaa", "alpha", "aud-a", "anchor-a")
	alphaPointer, err := os.ReadFile(pointer)
	if err != nil {
		t.Fatal(err)
	}
	accountFixture(t, dir, "exportb-bbbbbbbb", "beta", "aud-b", "anchor-b")
	betaPointer, err := os.ReadFile(pointer)
	if err != nil {
		t.Fatal(err)
	}
	pointAtAccount(t, dir, "exporta-aaaaaaaa")
	finish := beginOutlookInvocation()
	defer finish()
	pinned, _, err := currentSession()
	if err != nil {
		t.Fatal(err)
	}
	start := make(chan struct{})
	errors := make(chan error, 9)
	var wait sync.WaitGroup
	wait.Add(9)
	go func() {
		defer wait.Done()
		<-start
		for i := 0; i < 80; i++ {
			value := alphaPointer
			if i%2 != 0 {
				value = betaPointer
			}
			next := filepath.Join(dir, ".pointer-next")
			if err := os.WriteFile(next, value, 0600); err != nil {
				errors <- err
				return
			}
			if err := os.Rename(next, pointer); err != nil {
				errors <- err
				return
			}
		}
	}()
	for worker := 0; worker < 8; worker++ {
		go func(worker int) {
			defer wait.Done()
			<-start
			for i := 0; i < 100; i++ {
				current, _, err := currentSession()
				if err != nil || current.AccountKey != pinned.AccountKey || current.Headers.Get("Cookie") != "session=fixture-secret" {
					errors <- fmt.Errorf("pinned session or independent headers changed")
					return
				}
				current.Headers.Set("Cookie", fmt.Sprintf("local-%d-%d", worker, i))
			}
		}(worker)
	}
	close(start)
	wait.Wait()
	close(errors)
	for err := range errors {
		t.Error(err)
	}
}
