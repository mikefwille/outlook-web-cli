package drafts

import (
	"strings"
	"testing"
)

const replyTemplate = `{"__type":"CreateItemJsonRequest:#Exchange","Body":{"__type":"CreateItemRequest:#Exchange","MessageDisposition":"SaveOnly","SendOnNotFoundError":true,"SuppressMarkAsReadOnReplyOrForward":false,"Items":[{"__type":"ReplyToItem:#Exchange","MessageDisposition":"SaveOnly","ReferenceItemId":{"__type":"ItemId:#Exchange","Id":"old-id","ChangeKey":"old-key"},"ReferenceItemDocumentId":999,"Subject":"old-subject","NewBodyContent":{"BodyType":"HTML","Value":"old-body"},"ToRecipients":[{"EmailAddress":"fixture@example.invalid"}],"CcRecipients":[],"BccRecipients":[],"From":{},"ExtendedProperty":[{"Value":"old-value"}]}]}}`

func TestResponseDraftNeverCopiesFixtureAndCannotSend(t *testing.T) {
	r, err := ParseRecipients([]string{"self@example.invalid"}, nil, nil)
	if err != nil {
		t.Fatal(err)
	}
	b, err := PrepareResponseDraft([]byte(replyTemplate), "reply", "new-id", "new-key", "Re: subject", "<hello>", r)
	if err != nil {
		t.Fatal(err)
	}
	s := string(b)
	for _, forbidden := range []string{"fixture@example.invalid", "old-id", "old-key", "old-body", "old-value", "ReferenceItemDocumentId", "<hello>"} {
		if strings.Contains(s, forbidden) {
			t.Fatalf("fixture data or unescaped markup retained: %s", forbidden)
		}
	}
	for _, required := range []string{`"MessageDisposition":"SaveOnly"`, `"SendOnNotFoundError":false`, `"SuppressMarkAsReadOnReplyOrForward":true`, `"ShouldIgnoreChangeKey":false`, `self@example.invalid`} {
		if !strings.Contains(s, required) {
			t.Fatal("required safeguard missing")
		}
	}
	if _, err := PrepareResponseDraft([]byte(replyTemplate), "forward", "id", "key", "subject", "text", r); err == nil {
		t.Fatal("accepted wrong response type")
	}
}
