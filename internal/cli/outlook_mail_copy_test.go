package cli

import (
	"context"
	"errors"
	"testing"
)

func copyState(id, folder, subject string) messageActionState {
	var s messageActionState
	s.ItemID.ID, s.ItemID.ChangeKey, s.Parent.ID, s.Subject = id, "version", folder, subject
	return s
}

func TestCopyReadbackRequiresBothOriginalAndDestination(t *testing.T) {
	before := copyState("source", "inbox", "Subject")
	for _, tc := range []struct {
		name                 string
		source, copied       messageActionState
		sourceErr, copyErr   error
		wantSource, wantCopy bool
	}{
		{"both", before, copyState("copy", "destination", "Subject"), nil, nil, true, true},
		{"source moved", copyState("source", "elsewhere", "Subject"), copyState("copy", "destination", "Subject"), nil, nil, false, true},
		{"wrong folder", before, copyState("copy", "elsewhere", "Subject"), nil, nil, true, false},
		{"wrong subject", before, copyState("copy", "destination", "Different"), nil, nil, true, false},
		{"source unavailable", before, copyState("copy", "destination", "Subject"), errors.New("offline"), nil, false, true},
		{"copy unavailable", before, copyState("copy", "destination", "Subject"), nil, errors.New("offline"), true, false},
	} {
		t.Run(tc.name, func(t *testing.T) {
			calls := 0
			read := func(_ context.Context, id string) (messageActionState, error) {
				calls++
				if id == "source" {
					return tc.source, tc.sourceErr
				}
				return tc.copied, tc.copyErr
			}
			a, b := verifyCopyReadback(context.Background(), "source", "copy", "destination", before, read)
			if calls != 2 || a != tc.wantSource || b != tc.wantCopy {
				t.Fatal("copy readback did not preserve both outcomes")
			}
		})
	}
}
