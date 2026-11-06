package mailactions

import (
	"strings"
	"testing"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func TestMoveRequiresACompleteSingleResult(t *testing.T) {
	for _, body := range []string{`{}`, `{"Body":{"__type":"MoveItemResponse:#Exchange","ResponseMessages":{"Items":[{"__type":"ItemInfoResponseMessage:#Exchange","ResponseClass":"Success","ResponseCode":"NoError","Items":[]}]}}}`, `{"Body":{"__type":"MoveItemResponse:#Exchange","ResponseMessages":{"Items":[{"ResponseClass":"Error","ResponseCode":"ErrorAccessDenied"}]}}}`} {
		if _, state := ParseMove([]byte(body)); state == transport.EnvelopeOK {
			t.Fatal("incomplete move accepted")
		}
	}
	if _, state := ParseMove([]byte(strings.Repeat("x", 10))); state == transport.EnvelopeOK {
		t.Fatal("invalid JSON accepted")
	}
}
