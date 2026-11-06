package folders

import (
	"testing"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func TestFolderMutationRejectsAmbiguousSuccess(t *testing.T) {
	for _, raw := range []string{`{}`, `{"Body":{"__type":"DeleteFolderResponse:#Exchange","ResponseMessages":{"Items":[{"__type":"SingleResponseMessage:#Exchange","ResponseClass":"Success"}]}}}`, `{"Body":{"__type":"DeleteFolderResponse:#Exchange","ResponseMessages":{"Items":[{"__type":"SingleResponseMessage:#Exchange","ResponseClass":"Error","ResponseCode":"AccessDenied"}]}}}`} {
		_, state := ParseResult("delete", []byte(raw))
		if state == transport.EnvelopeOK {
			t.Fatal("incomplete or rejected delete treated as success")
		}
	}
}
