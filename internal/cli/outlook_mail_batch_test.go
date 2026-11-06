package cli

import (
	"context"
	"fmt"
	"reflect"
	"testing"
)

func TestMailBatchStopsWithoutRetryOnUncertainWrite(t *testing.T) {
	var calls []string
	runner := func(_ context.Context, id, key string, write bool) (map[string]any, error) {
		if !write {
			calls = append(calls, "preview:"+id)
			return map[string]any{"change_key": "v-" + id}, nil
		}
		calls = append(calls, "write:"+id)
		if id == "two" {
			return nil, fmt.Errorf("write outcome unknown")
		}
		return map[string]any{"accepted": true}, nil
	}
	got, err := runMailBatch(context.Background(), []string{"one", "two", "three"}, []string{"v-one", "v-two", "v-three"}, true, runner)
	if err == nil || got[0].Status != "accepted" || got[1].Status != "failed_or_unknown" || got[2].Status != "not_attempted" {
		t.Fatal("partial outcomes were not preserved")
	}
	want := []string{"preview:one", "preview:two", "preview:three", "write:one", "write:two"}
	if !reflect.DeepEqual(calls, want) {
		t.Fatalf("unexpected action order: %v", calls)
	}
}
func TestMailBatchConflictsPreventEveryWrite(t *testing.T) {
	writes := 0
	runner := func(_ context.Context, id, key string, write bool) (map[string]any, error) {
		if write {
			writes++
		}
		return map[string]any{"change_key": "current-" + id}, nil
	}
	got, err := runMailBatch(context.Background(), []string{"one", "two"}, []string{"current-one", "stale"}, true, runner)
	if err == nil || writes != 0 || got[1].Status != "version_conflict" {
		t.Fatal("stale batch wrote a message")
	}
}
func TestMailBatchRejectsDuplicateIDsBeforeIO(t *testing.T) {
	cmd := newMailBatchCmd(&rootFlags{})
	cmd.SilenceUsage = true
	cmd.SilenceErrors = true
	cmd.SetArgs([]string{"--operation", "flag", "--id", "same", "--id", "same"})
	if err := cmd.Execute(); err == nil {
		t.Fatal("duplicate message IDs accepted")
	}
}
