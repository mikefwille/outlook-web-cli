package capture

import (
	"testing"
	"time"
)

func TestRebindRequiresMatchingAccountAudienceAndMailbox(t *testing.T) {
	v := syntheticCapture(t, time.Now().Add(time.Hour))
	p := writeCapture(t, v, 0o600)
	observed, _, err := LoadFindFolder(p)
	if err != nil {
		t.Fatal(err)
	}
	current := observed
	current.Headers = observed.Headers.Clone()
	current.Headers.Set("Cookie", "session=fresh")
	got, err := RebindRead(observed, current)
	if err != nil || got.Headers.Get("Cookie") != "session=fresh" || observed.Headers.Get("Cookie") == "session=fresh" {
		t.Fatal("safe refresh failed or mutated source")
	}
	bad := current
	bad.AccountKey = "other"
	if _, err := RebindRead(observed, bad); err == nil {
		t.Fatal("cross-account credentials accepted")
	}
	bad = current
	bad.Headers = current.Headers.Clone()
	bad.Headers.Set("X-AnchorMailbox", "other")
	if _, err := RebindRead(observed, bad); err == nil {
		t.Fatal("cross-mailbox credentials accepted")
	}
	bad = current
	bad.Headers = current.Headers.Clone()
	bad.Headers.Set("Authorization", "Bearer invalid")
	if _, err := RebindRead(observed, bad); err == nil {
		t.Fatal("unknown audience accepted")
	}
}
