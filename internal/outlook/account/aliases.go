// Package account extracts only the signed-in mailbox's address identities.
package account

import (
	"encoding/json"
	"fmt"
	"net/mail"
	"strings"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/webcapture"
)

func PrepareAliases(detail string, current transport.Session) (transport.Session, transport.Operation, error) {
	fail := func() (transport.Session, transport.Operation, error) {
		return transport.Session{}, transport.Operation{}, fmt.Errorf("account configuration does not match the observed read contract")
	}
	d, err := webcapture.Load(detail, "", current)
	if err != nil {
		return fail()
	}
	app := d.URL.Query().Get("app")
	if d.Method != "POST" || d.URL.Path != "/owa/service.svc" || d.URL.Query().Get("action") != "GetOwaUserConfiguration" || (app != "Mail" && app != "People") || d.Session.Headers.Get("Action") != "GetOwaUserConfiguration" || !d.HeaderBody {
		return fail()
	}
	var body map[string]any
	if json.Unmarshal(d.Body, &body) != nil || len(body) != 0 {
		return fail()
	}
	return d.Session, d.Operation("POST", d.URL.Path, []byte(`{}`), transport.Read, validateAliases), nil
}

// Aliases contains only SMTP addresses, not the rest of the configuration
// response, which can include user settings and session material.
func ParseAliases(data []byte) ([]string, transport.EnvelopeResult) {
	var root struct {
		Session *struct {
			Email     string    `json:"UserEmailAddress"`
			Principal string    `json:"UserPrincipalName"`
			Logon     string    `json:"LogonEmailAddress"`
			Proxy     *[]string `json:"UserProxyAddresses"`
		} `json:"SessionSettings"`
	}
	if json.Unmarshal(data, &root) != nil || root.Session == nil || root.Session.Proxy == nil || len(*root.Session.Proxy) > 500 {
		return nil, transport.EnvelopeError
	}
	items := append([]string{}, (*root.Session.Proxy)...)
	items = append(items, root.Session.Email, root.Session.Principal, root.Session.Logon)
	seen := map[string]bool{}
	out := []string{}
	for _, s := range items {
		if strings.HasPrefix(strings.ToLower(s), "smtp:") {
			s = s[5:]
		}
		if s == "" {
			continue
		}
		if len(s) > 320 || strings.ContainsAny(s, "\r\n\x00") {
			return nil, transport.EnvelopeError
		}
		a, err := mail.ParseAddress(s)
		if err != nil || a.Name != "" || a.Address != s || strings.Count(s, "@") != 1 {
			continue
		}
		key := strings.ToLower(s)
		if !seen[key] {
			seen[key] = true
			out = append(out, s)
		}
	}
	primary := strings.TrimPrefix(strings.TrimPrefix(root.Session.Email, "SMTP:"), "smtp:")
	if len(out) == 0 || primary == "" || !seen[strings.ToLower(primary)] {
		return nil, transport.EnvelopeError
	}
	return out, transport.EnvelopeOK
}
func validateAliases(b []byte) transport.EnvelopeResult { _, s := ParseAliases(b); return s }
