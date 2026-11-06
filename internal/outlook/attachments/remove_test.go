package attachments

import (
	"testing"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func TestAttachmentRemovalRejectsIncompleteSuccess(t *testing.T) {
	for _, data := range []string{`{}`, `{"Body":{"__type":"DeleteAttachmentResponse:#Exchange","ResponseMessages":{"Items":[{"__type":"DeleteAttachmentResponseMessage:#Exchange","ResponseClass":"Success","ResponseCode":"NoError"}]}}}`, `{"Body":{"__type":"DeleteAttachmentResponse:#Exchange","ResponseMessages":{"Items":[{"ResponseClass":"Error","ResponseCode":"ErrorAccessDenied"}]}}}`} {
		if _, state := ParseRemoved([]byte(data)); state == transport.EnvelopeOK {
			t.Fatal("incomplete removal accepted")
		}
	}
}
