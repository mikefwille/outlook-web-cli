package drafts

import (
	"net/mail"
	"strings"
)

// Recipient is a caller-supplied SMTP address. MailboxType OneOff uses the
// documented EWS representation for an address without a directory identity;
// the actual Outlook web endpoint must also pass a live fixture test.
type Recipient struct {
	Name         string `json:"Name"`
	EmailAddress string `json:"EmailAddress"`
	MailboxType  string `json:"MailboxType"`
	RoutingType  string `json:"RoutingType"`
}

type Recipients struct {
	To  []Recipient `json:"to"`
	Cc  []Recipient `json:"cc"`
	Bcc []Recipient `json:"bcc"`
}

// A nil field leaves the existing list alone; an empty slice clears it.
type RecipientChanges struct {
	To  *[]Recipient
	Cc  *[]Recipient
	Bcc *[]Recipient
}

func validRecipient(r Recipient) bool {
	if len(r.Name) > 512 || len(r.EmailAddress) > 320 || strings.ContainsAny(r.Name+r.EmailAddress, "\r\n\x00") || r.RoutingType != "SMTP" || r.MailboxType != "OneOff" || strings.Count(r.EmailAddress, "@") != 1 || strings.HasPrefix(r.EmailAddress, "@") || strings.HasSuffix(r.EmailAddress, "@") {
		return false
	}
	a, err := mail.ParseAddress(r.EmailAddress)
	return err == nil && a.Name == "" && a.Address == r.EmailAddress
}

func ParseRecipients(to, cc, bcc []string) (Recipients, error) {
	if len(to)+len(cc)+len(bcc) > 500 {
		return Recipients{}, ErrInvalid
	}
	parse := func(in []string) ([]Recipient, error) {
		out := make([]Recipient, 0, len(in))
		for _, raw := range in {
			if len(raw) > 1024 || strings.ContainsAny(raw, "\r\n\x00") {
				return nil, ErrInvalid
			}
			a, err := mail.ParseAddress(raw)
			if err != nil || len(a.Address) > 320 || len(a.Name) > 512 || strings.Count(a.Address, "@") != 1 || strings.HasPrefix(a.Address, "@") || strings.HasSuffix(a.Address, "@") {
				return nil, ErrInvalid
			}
			name := a.Name
			if name == "" {
				name = a.Address
			}
			r := Recipient{Name: name, EmailAddress: a.Address, MailboxType: "OneOff", RoutingType: "SMTP"}
			if !validRecipient(r) {
				return nil, ErrInvalid
			}
			out = append(out, r)
		}
		return out, nil
	}
	t, err := parse(to)
	if err != nil {
		return Recipients{}, err
	}
	c, err := parse(cc)
	if err != nil {
		return Recipients{}, err
	}
	b, err := parse(bcc)
	if err != nil {
		return Recipients{}, err
	}
	return Recipients{To: t, Cc: c, Bcc: b}, nil
}
