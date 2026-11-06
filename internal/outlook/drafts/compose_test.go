package drafts

import (
	"strings"
	"testing"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func TestComposeSnapshotRequiresCompleteDraftAndRecipients(t *testing.T) {
	const input = `{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Success","ResponseCode":"NoError","Items":[{"__type":"Message:#Exchange","ItemId":{"Id":"id","ChangeKey":"key"},"IsDraft":true,"Subject":"subject","Body":{"BodyType":"HTML","Value":"body","IsTruncated":false},"ToRecipients":[{"Name":"Alice","EmailAddress":"alice@example.invalid","RoutingType":"SMTP"}],"BccRecipients":[{"EmailAddress":"blind@example.invalid","RoutingType":"SMTP"}]}]}]}}}`
	s, state := ParseCompose([]byte(input))
	if state != transport.EnvelopeOK || len(s.To) != 1 || len(s.Bcc) != 1 || s.Body != "body" {
		t.Fatal("valid complete snapshot rejected")
	}
	for _, bad := range []string{strings.Replace(input, `"IsTruncated":false`, `"IsTruncated":true`, 1), strings.Replace(input, `"IsDraft":true`, `"IsDraft":false`, 1), strings.Replace(input, `"RoutingType":"SMTP"`, `"RoutingType":"EX"`, 1)} {
		if _, state := ParseCompose([]byte(bad)); state == transport.EnvelopeOK {
			t.Fatal("accepted unsafe snapshot")
		}
	}
}
