package mailcache

import (
	"context"
	"fmt"
)

type FetchPage func(context.Context, int, int) (Page, error)

// Reconcile resumes page checkpoints, then scans the folder a second time to
// compare every item ID and change key before atomically publishing. An
// incomplete or changed scan never replaces the last complete snapshot. This
// is bounded reconciliation, not a server-provided delta cursor or an atomic
// server snapshot; changes after verification can make the cache stale.
func Reconcile(ctx context.Context, s *Store, r Run, fetch FetchPage, maxItems int) (Run, error) {
	if maxItems < 1 || maxItems > 10000 {
		return r, fmt.Errorf("max items must be 1–10000")
	}
	if r.Status != "scan" && r.Status != "verify" {
		return r, fmt.Errorf("reconciliation cannot be resumed in its current state")
	}
	if r.Status == "scan" {
		for pages := 0; r.Status == "scan" && pages < 101; pages++ {
			p, err := fetch(ctx, r.Offset, 100)
			if err != nil {
				return r, err
			}
			next, err := s.AppendPage(ctx, r, p, maxItems)
			if err != nil {
				_ = s.MarkStale(ctx, r.ID)
				return r, err
			}
			r = next
		}
		if r.Status != "verify" {
			_ = s.MarkStale(ctx, r.ID)
			return r, fmt.Errorf("folder exceeds the page limit")
		}
	}
	versions, err := s.Versions(ctx, r.ID)
	if err != nil {
		return r, err
	}
	if len(versions) != r.Count {
		return r, fmt.Errorf("staged snapshot is incomplete")
	}
	seen := map[string]bool{}
	offset := 0
	finished := false
	for pages := 0; pages < 101; pages++ {
		p, err := fetch(ctx, offset, 100)
		if err != nil {
			return r, err
		}
		if p.Total != r.Total || len(seen)+len(p.Items) > r.Count {
			_ = s.MarkStale(ctx, r.ID)
			return r, fmt.Errorf("folder changed during verification; start a new reconciliation")
		}
		for _, item := range p.Items {
			if seen[item.ID] || versions[item.ID] == "" || versions[item.ID] != item.ChangeKey {
				_ = s.MarkStale(ctx, r.ID)
				return r, fmt.Errorf("folder changed during verification; start a new reconciliation")
			}
			seen[item.ID] = true
		}
		if p.Last {
			finished = true
			break
		}
		if p.Next <= offset {
			_ = s.MarkStale(ctx, r.ID)
			return r, fmt.Errorf("folder paging did not advance")
		}
		offset = p.Next
	}
	if !finished || len(seen) != r.Count {
		_ = s.MarkStale(ctx, r.ID)
		return r, fmt.Errorf("folder verification was incomplete")
	}
	if err = s.Publish(ctx, r); err != nil {
		return r, err
	}
	return s.Get(ctx, r.ID)
}
