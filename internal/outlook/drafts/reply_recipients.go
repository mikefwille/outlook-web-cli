package drafts

import (
	"fmt"
	"net/mail"
	"strings"
)

// SourceMailbox describes a mailbox in the observed GetItem response.
type SourceMailbox struct {
	Name    string `json:"Name"`
	Address string `json:"EmailAddress"`
	Routing string `json:"RoutingType"`
}
type SourceParty struct {
	Mailbox SourceMailbox `json:"Mailbox"`
}
type SourceRecipientCounts struct {
	To *int `json:"ToRecipientsCount"`
	Cc *int `json:"CcRecipientsCount"`
}
type ReplySource struct {
	From    *SourceParty           `json:"From"`
	ReplyTo []SourceMailbox        `json:"ReplyTo"`
	To      []SourceMailbox        `json:"ToRecipients"`
	Cc      []SourceMailbox        `json:"CcRecipients"`
	Counts  *SourceRecipientCounts `json:"RecipientCounts"`
}

// ResolveReply uses explicitly returned ReplyTo or From and, for reply-all,
// the complete returned To/Cc lists. All account aliases are excluded except
// as a fallback for a conversation addressed only to the sender's own account.
// Bcc is never inferred from the source message.
func ResolveReply(s ReplySource, all bool, aliases []string) (Recipients, error) {
	fail := func() (Recipients, error) {
		return Recipients{}, fmt.Errorf("source recipient information is incomplete or not SMTP; provide explicit recipients")
	}
	target := s.ReplyTo
	if len(target) == 0 {
		if s.From == nil {
			return fail()
		}
		target = []SourceMailbox{s.From.Mailbox}
	}
	// The observed GetItem request returns at most 20 recipients. With no
	// ReplyTo count field, a full-sized list may be truncated.
	if len(target) >= 20 || len(s.To)+len(s.Cc) > 500 {
		return fail()
	}
	if all {
		if s.Counts == nil || len(aliases) == 0 {
			return fail()
		}
		count := func(v *int) int {
			if v == nil {
				return 0
			}
			return *v
		}
		if count(s.Counts.To) != len(s.To) || count(s.Counts.Cc) != len(s.Cc) {
			return fail()
		}
	}
	own := map[string]bool{}
	for _, a := range aliases {
		own[strings.ToLower(a)] = true
	}
	seen := map[string]bool{}
	to, cc, self := []string{}, []string{}, []string{}
	add := func(v SourceMailbox, dest *[]string) bool {
		if !strings.EqualFold(v.Routing, "SMTP") || v.Address == "" || strings.ContainsAny(v.Name+v.Address, "\r\n\x00") {
			return false
		}
		a, err := mail.ParseAddress(v.Address)
		if err != nil || a.Name != "" || a.Address != v.Address {
			return false
		}
		key := strings.ToLower(v.Address)
		if seen[key] {
			return true
		}
		seen[key] = true
		encoded := (&mail.Address{Name: v.Name, Address: v.Address}).String()
		if all && own[key] {
			self = append(self, encoded)
		} else {
			*dest = append(*dest, encoded)
		}
		return true
	}
	for _, v := range target {
		if !add(v, &to) {
			return fail()
		}
	}
	if all {
		for _, v := range s.To {
			if !add(v, &to) {
				return fail()
			}
		}
		for _, v := range s.Cc {
			if !add(v, &cc) {
				return fail()
			}
		}
	}
	if len(to)+len(cc) == 0 && len(self) > 0 {
		to = append(to, self[0])
	}
	if len(to)+len(cc) == 0 {
		return fail()
	}
	r, err := ParseRecipients(to, cc, nil)
	if err != nil {
		return fail()
	}
	return r, nil
}
