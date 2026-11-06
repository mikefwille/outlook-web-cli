package mailactions

import (
	"encoding/json"

	"testing"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func TestActionRejectsIncompleteSuccess(t *testing.T) {
	for _, body := range []string{`{}`, `{"Body":{"__type":"UpdateItemResponse:#Exchange","ResponseMessages":{"Items":[{"__type":"UpdateItemResponseMessage:#Exchange","ResponseClass":"Success","ResponseCode":"NoError","Items":[]}]}}}`, `{"Body":{"__type":"DeleteItemResponse:#Exchange","ResponseMessages":{"Items":[{"__type":"DeleteItemResponseMessage:#Exchange","ResponseClass":"Error","ResponseCode":"ErrorAccessDenied"}]}}}`} {
		if !json.Valid([]byte(body)) {
			t.Fatal("invalid test")
		}
		_, state := ParseResult("mark-read", []byte(body))
		if state == transport.EnvelopeOK {
			t.Fatal("invalid action response accepted")
		}
	}
}
