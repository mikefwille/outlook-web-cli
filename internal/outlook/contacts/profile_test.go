package contacts

import (
	"encoding/json"
	"testing"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func TestParsePeopleProfileRequiresMatchingProfileSource(t *testing.T) {
	const profile = `{"value":[{"id":"profile-id","changeKey":"version","names":[],"personSources":[{"type":"Profile"}],"name":{"displayName":"Sample Person"}}]}`
	raw, state := ParseProfile([]byte(profile), "profile-id")
	var data map[string]any
	if state != transport.EnvelopeOK || json.Unmarshal(raw, &data) != nil || data["id"] != "profile-id" || data["name"].(map[string]any)["displayName"] != "Sample Person" {
		t.Fatal("full People profile not preserved")
	}
	if _, state = ParseProfile([]byte(profile), "different-id"); state == transport.EnvelopeOK {
		t.Fatal("accepted a different People identity")
	}
	personal := `{"value":[{"id":"profile-id","changeKey":"version","names":[],"personSources":[{"type":"UserProvided"}]}]}`
	if raw, state = ParseProfile([]byte(personal), "profile-id"); state != transport.EnvelopeOK || raw != nil {
		t.Fatal("personal-only contact misrepresented as a profile")
	}
	for _, invalid := range []string{`{}`, `{"value":[]}`, `{"value":[{"id":"profile-id","changeKey":"version"}]}`} {
		if _, state = ParseProfile([]byte(invalid), "profile-id"); state == transport.EnvelopeOK {
			t.Fatal("incomplete profile accepted")
		}
	}
}
