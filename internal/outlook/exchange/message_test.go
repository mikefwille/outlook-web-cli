package exchange

import (
	"strings"
	"testing"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func TestMessageGetPreservesTruncationAndRejectsErrors(t *testing.T) {
	const good = `{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Success","ResponseCode":"NoError","Items":[{"__type":"Message:#Exchange","ItemId":{"Id":"id"},"NormalizedBody":{"Value":"text","IsTruncated":true}}]}]}}}`
	id, item, state := ParseMessageGet([]byte(good))
	if state != transport.EnvelopeOK || id != "id" || !strings.Contains(string(item), `"IsTruncated":true`) {
		t.Fatal("lost item or truncation metadata")
	}
	for _, bad := range []string{strings.Replace(good, "NoError", "ErrorItemNotFound", 1), strings.Replace(good, "Message:#Exchange", "CalendarItem:#Exchange", 1)} {
		if ValidateMessageGet([]byte(bad)) == transport.EnvelopeOK {
			t.Fatal("accepted failed or non-message item")
		}
	}
}
