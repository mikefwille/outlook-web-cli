package conversation

import (
	"testing"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func TestEnvelopeRequiresExplicitSuccess(t *testing.T) {
	for _, tc := range []struct {
		body string
		ok   bool
	}{
		{`{"Body":{"ResponseMessages":{"Items":[{"__type":"ApplyConversationActionResponseMessage:#Exchange","ResponseClass":"Success","ResponseCode":"NoError"}]}}}`, true},
		{`{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Success","ResponseCode":"NoError"}]}}}`, false},
		{`{"Body":{"ResponseMessages":{"Items":[{"__type":"ApplyConversationActionResponseMessage:#Exchange","ResponseClass":"Error","ResponseCode":"ErrorAccessDenied"}]}}}`, false},
	} {
		if (Validate([]byte(tc.body)) == transport.EnvelopeOK) != tc.ok {
			t.Fatal("unexpected action envelope result")
		}
	}
}
