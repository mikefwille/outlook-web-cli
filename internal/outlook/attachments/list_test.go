package attachments

import (
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"testing"
)

func TestMessageAttachmentMetadataRequiresCompleteEntries(t *testing.T) {
	base := `{"__type":"Message:#Exchange","ItemId":{"Id":"message","ChangeKey":"version"},"HasAttachments":true,"Attachments":[{"__type":"FileAttachment:#Exchange","AttachmentId":{"Id":"attachment"},"Name":"note.txt","ContentType":"text/plain","Size":288}]}`
	v, state := ParseMessage([]byte(base))
	if state != transport.EnvelopeOK || len(v.Attachments) != 1 || v.Attachments[0].Name != "note.txt" || v.Attachments[0].ReportedSize == nil || *v.Attachments[0].ReportedSize != 288 {
		t.Fatal("valid metadata rejected")
	}
	for _, raw := range []string{
		`{"__type":"Message:#Exchange","ItemId":{"Id":"message","ChangeKey":"version"},"HasAttachments":true}`,
		`{"__type":"Message:#Exchange","ItemId":{"Id":"message","ChangeKey":"version"}}`,
		`{"__type":"Message:#Exchange","ItemId":{"Id":"message","ChangeKey":"version"},"HasAttachments":true,"Attachments":[{"__type":"FileAttachment:#Exchange","Name":"note.txt"}]}`,
	} {
		if _, s := ParseMessage([]byte(raw)); s == transport.EnvelopeOK {
			t.Fatal("incomplete metadata accepted")
		}
	}
}
