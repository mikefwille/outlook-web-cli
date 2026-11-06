package drafts

import (
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"testing"
)

func TestCreateResponseRequiresServerDraftIdentity(t *testing.T) {
	valid := []byte(`{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Success","ResponseCode":"NoError","Items":[{"ItemId":{"Id":"synthetic","ChangeKey":"key"},"IsDraft":true}]}]}}}`)
	v, state := ParseCreate(valid)
	if state != transport.EnvelopeOK || v.ID != "synthetic" || v.ChangeKey != "key" || !v.IsDraft {
		t.Fatal("valid draft response rejected")
	}
	for _, bad := range []string{
		`{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Error","ResponseCode":"ErrorAccessDenied"}]}}}`,
		`{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Success","ResponseCode":"NoError","Items":[{"ItemId":{"Id":"synthetic","ChangeKey":"key"},"IsDraft":false}]}]}}}`,
	} {
		if _, state := ParseCreate([]byte(bad)); state != transport.EnvelopeError {
			t.Fatal("accepted non-draft result")
		}
	}
}
