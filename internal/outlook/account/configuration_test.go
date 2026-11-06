package account

import (
	"testing"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func TestConfigurationReturnsOnlyRequestedSections(t *testing.T) {
	b := []byte(`{"MasterCategoryList":{"MasterList":[{"Id":"category-1","Name":"Example","Color":2,"KeyboardShortcut":0}]},"UserOptions":{"SignatureText":null,"SignatureHtml":"<b>Hi</b>","AutoAddSignature":true,"AutoAddSignatureOnReply":false,"IsFocusedInboxEnabled":true},"unrelated_secret":"do-not-return"}`)
	items, state := ParseCategories(b)
	if state != transport.EnvelopeOK || len(items) != 1 || items[0].ID != "category-1" || items[0].Name != "Example" {
		t.Fatal("category parse failed")
	}
	settings, state := ParseSettings(b)
	if state != transport.EnvelopeOK || settings.SignatureText != nil || settings.SignatureHTML == nil || !*settings.AutoAddSignature || *settings.AutoAddSignatureOnReply {
		t.Fatal("settings parse failed")
	}
	for _, bad := range []string{`{}`, `{"MasterCategoryList":{"MasterList":[{"Id":"a","Name":"x"}]}}`} {
		if _, s := ParseCategories([]byte(bad)); s == transport.EnvelopeOK {
			t.Fatal("incomplete category data accepted")
		}
	}
}
