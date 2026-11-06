package drafts

import (
	"encoding/json"

	"strings"
	"testing"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

const createTemplate = `{"__type":"CreateItemJsonRequest:#Exchange","Body":{"__type":"CreateItemRequest:#Exchange","ComposeOperation":"newMail","MessageDisposition":"SaveOnly","SendOnNotFoundError":true,"Items":[{"__type":"Message:#Exchange","MessageDisposition":"SaveOnly","ToRecipients":[],"CcRecipients":[],"BccRecipients":[],"ReplyTo":[],"appendOnSend":[],"prependOnSend":[],"atMentionedRecipients":[],"recipientsAddedViaAtMention":[],"internetHeaders":[],"internetHeadersToRemove":[],"Body":{"BodyType":"HTML","Value":"old"},"Subject":"old","plainText":"old"}]}}`
const updateTemplate = `{"__type":"UpdateItemJsonRequest:#Exchange","Body":{"MessageDisposition":"SaveOnly","SendCalendarInvitationsOrCancellations":"SendToNone","SendOnNotFoundError":true,"ItemChanges":[{"ItemId":{"__type":"ItemId:#Exchange","Id":"old","ChangeKey":"old"},"Updates":[{"Path":{"FieldURI":"ToRecipients"},"Item":{"__type":"Message:#Exchange","ToRecipients":[]}},{"Path":{"FieldURI":"Subject"},"Item":{"__type":"Message:#Exchange","Subject":"old"}},{"Path":{"FieldURI":"Body"},"Item":{"__type":"Message:#Exchange","Body":{"BodyType":"HTML","Value":"old"}}}]}]}}`
const discardTemplate = `{"__type":"DeleteItemJsonRequest:#Exchange","Body":{"DeleteType":"MoveToDeletedItems","SendMeetingCancellations":"SendToNone","ItemIds":[{"__type":"ItemId:#Exchange","Id":"old","ChangeKey":"old"}]}}`
const getTemplate = `{"__type":"GetItemJsonRequest:#Exchange","Body":{"__type":"GetItemRequest:#Exchange","ItemIds":[{"__type":"ItemId:#Exchange","Id":"old"}]}}`
const sendTemplate = `{"__type":"UpdateItemJsonRequest:#Exchange","Body":{"__type":"UpdateItemRequest:#Exchange","MessageDisposition":"SendAndSaveCopy","SendCalendarInvitationsOrCancellations":"SendToNone","SendOnNotFoundError":true,"ConflictResolution":"AlwaysOverwrite","ItemChanges":[{"ItemId":{"__type":"ItemId:#Exchange","Id":"old","ChangeKey":"old"},"Updates":[{"__type":"SetItemField:#Exchange","Path":{"FieldURI":"ToRecipients"},"Item":{"__type":"Message:#Exchange","ToRecipients":[{"EmailAddress":"fixture@example.invalid"}]}},{"__type":"SetItemField:#Exchange","Path":{"FieldURI":"Subject"},"Item":{"__type":"Message:#Exchange","Subject":"fixture"}},{"__type":"SetItemField:#Exchange","Path":{"FieldURI":"Body"},"Item":{"__type":"Message:#Exchange","Body":{"Value":"fixture"}}}]}]}}`

func TestGetSelectsExactlyOneID(t *testing.T) {
	b, err := PrepareGet([]byte(getTemplate), "new-id")
	if err != nil || strings.Contains(string(b), "old") || !strings.Contains(string(b), "new-id") {
		t.Fatal("failed to replace fixture ID")
	}
	if _, err := PrepareGet([]byte(getTemplate), ""); err == nil {
		t.Fatal("accepted empty ID")
	}
}

func TestCreateIsSaveOnlyRecipientFreeAndEscapesHTML(t *testing.T) {
	b, err := PrepareCreate([]byte(createTemplate), "Synthetic", "<script> & text\nnext")
	if err != nil {
		t.Fatal(err)
	}
	var root map[string]any
	if json.Unmarshal(b, &root) != nil {
		t.Fatal("invalid JSON")
	}
	s := string(b)
	if strings.Contains(s, "<script>") || !strings.Contains(s, "SendOnNotFoundError\":false") || !strings.Contains(s, "SaveOnly") {
		t.Fatalf("unsafe body: %s", s)
	}
	bad := strings.Replace(createTemplate, `"ToRecipients":[]`, `"ToRecipients":[{"email":"outside@example.invalid"}]`, 1)
	if _, err := PrepareCreate([]byte(bad), "x", "y"); err == nil {
		t.Fatal("accepted recipient")
	}
}

func TestUpdatePreservesRecipientsByOmittingTheirFields(t *testing.T) {
	b, err := PrepareUpdate([]byte(updateTemplate), "new-id", "new-key", "new subject", "new body")
	if err != nil {
		t.Fatal(err)
	}
	if strings.Contains(string(b), "ToRecipients") || strings.Contains(string(b), `"Id":"old"`) || !strings.Contains(string(b), "new subject") || !strings.Contains(string(b), `"ConflictResolution":"NeverOverwrite"`) {
		t.Fatalf("wrong update: %s", b)
	}
}

func TestRecipientChangesAreExplicitAndValidated(t *testing.T) {
	r, err := ParseRecipients([]string{"Alice <alice@example.invalid>"}, []string{"copy@example.invalid"}, nil)
	if err != nil {
		t.Fatal(err)
	}
	b, err := PrepareCreateWithRecipients([]byte(createTemplate), "subject", "body", r)
	if err != nil || !strings.Contains(string(b), "alice@example.invalid") || !strings.Contains(string(b), "copy@example.invalid") || !strings.Contains(string(b), `"MessageDisposition":"SaveOnly"`) {
		t.Fatal("addressed draft not saved safely")
	}
	to := r.To
	u, err := PrepareUpdateWithRecipients([]byte(updateTemplate), "id", "key", "subject", "body", &RecipientChanges{To: &to})
	if err != nil || !strings.Contains(string(u), "alice@example.invalid") || strings.Contains(string(u), "copy@example.invalid") {
		t.Fatal("update did not limit recipient changes")
	}
	if _, err := ParseRecipients([]string{"alice@example.invalid\r\nBcc: evil@example.invalid"}, nil, nil); err == nil {
		t.Fatal("accepted header injection")
	}
	r.To[0].MailboxType = "InternalMailbox"
	if _, err := PrepareCreateWithRecipients([]byte(createTemplate), "subject", "body", r); err == nil {
		t.Fatal("accepted unvalidated recipient representation")
	}
}

func TestSendDoesNotCopyFixtureRecipientsOrBody(t *testing.T) {
	b, err := PrepareSend([]byte(sendTemplate), "id", "key", "reviewed subject")
	if err != nil {
		t.Fatal(err)
	}
	s := string(b)
	if strings.Contains(s, "fixture@example.invalid") || strings.Contains(s, `"Body":{"Value":"fixture"}`) || !strings.Contains(s, `"ConflictResolution":"NeverOverwrite"`) || !strings.Contains(s, `"SendOnNotFoundError":false`) || !strings.Contains(s, `"SendAndSaveCopy"`) {
		t.Fatal("send copied fixture content or lost protections")
	}
	bad := strings.Replace(sendTemplate, "SendAndSaveCopy", "SaveOnly", 1)
	if _, err := PrepareSend([]byte(bad), "id", "key", "subject"); err == nil {
		t.Fatal("accepted non-send template")
	}
	accepted := `{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Success","ResponseCode":"NoError","Items":null}]}}}`
	if ParseSend([]byte(accepted)) != transport.EnvelopeOK {
		t.Fatal("rejected server acceptance")
	}
	conflict := `{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Error","ResponseCode":"ErrorIrresolvableConflict"}]}}}`
	if ParseSend([]byte(conflict)) != transport.EnvelopeConflict {
		t.Fatal("failed to classify conflict")
	}
}

func TestDraftResponsesRequireDraftAndSuccess(t *testing.T) {
	get := `{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Success","ResponseCode":"NoError","Items":[{"__type":"Message:#Exchange","ItemId":{"Id":"id","ChangeKey":"key"},"ParentFolderId":{"Id":"folder"},"IsDraft":true,"Subject":"subject","NormalizedBody":{"Value":"partial","IsTruncated":true}}]}]}}}`
	d, state := ParseGet([]byte(get))
	if state != 1 || d.ID != "id" || d.ParentFolderID != "folder" || !strings.Contains(string(d.Body), `"IsTruncated":true`) {
		t.Fatal("failed valid draft")
	}
	for _, bad := range []string{strings.Replace(get, `"IsDraft":true`, `"IsDraft":false`, 1), strings.Replace(get, `"ResponseCode":"NoError"`, `"ResponseCode":"ErrorItemNotFound"`, 1)} {
		if _, state := ParseGet([]byte(bad)); state == 1 {
			t.Fatal("accepted non-draft or failed response")
		}
	}
	updated := `{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Success","ResponseCode":"NoError","Items":[{"__type":"Message:#Exchange","ItemId":{"Id":"id","ChangeKey":"new-key"}}]}]}}}`
	if _, state := ParseUpdate([]byte(updated)); state != 1 {
		t.Fatal("rejected update")
	}
	discarded := `{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Success","ResponseCode":"NoError","MovedItemId":{"Id":"moved"}}]}}}`
	if ParseDiscard([]byte(discarded)) != 1 || ParseDiscard([]byte(`{"Body":{}}`)) == 1 {
		t.Fatal("discard response validation")
	}
}

func TestDiscardCannotBecomePermanentOrNotify(t *testing.T) {
	if _, err := PrepareDiscard([]byte(discardTemplate), "id", "key"); err != nil {
		t.Fatal(err)
	}
	for _, bad := range []string{strings.Replace(discardTemplate, "MoveToDeletedItems", "HardDelete", 1), strings.Replace(discardTemplate, "SendToNone", "SendToAllAndSaveCopy", 1)} {
		if _, err := PrepareDiscard([]byte(bad), "id", "key"); err == nil {
			t.Fatal("accepted unsafe deletion")
		}
	}
}
