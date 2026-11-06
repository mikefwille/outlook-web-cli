package account

import (
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"testing"
)

func TestAliasesRequireProxyListAndDeduplicate(t *testing.T) {
	b := []byte(`{"SessionSettings":{"UserEmailAddress":"person@example.invalid","UserPrincipalName":"person@example.invalid","UserProxyAddresses":["person@example.invalid","SMTP:alias@example.invalid","x500:/legacy"]}}`)
	a, s := ParseAliases(b)
	if s != transport.EnvelopeOK || len(a) != 2 {
		t.Fatal("valid aliases not preserved")
	}
	if _, s = ParseAliases([]byte(`{"SessionSettings":{"UserEmailAddress":"person@example.invalid"}}`)); s == transport.EnvelopeOK {
		t.Fatal("missing proxy list accepted")
	}
}
