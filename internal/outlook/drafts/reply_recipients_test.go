package drafts

import "testing"

func TestResolveReplyAllExcludesAliasesDeduplicatesAndChecksCounts(t *testing.T) {
	n := func(v int) *int { return &v }
	m := func(a string) SourceMailbox { return SourceMailbox{Address: a, Routing: "SMTP"} }
	s := ReplySource{From: &SourceParty{Mailbox: m("author@example.invalid")}, To: []SourceMailbox{m("me@example.invalid"), m("other@example.invalid")}, Cc: []SourceMailbox{m("alias@example.invalid"), m("other@example.invalid")}, Counts: &SourceRecipientCounts{To: n(2), Cc: n(2)}}
	r, err := ResolveReply(s, true, []string{"me@example.invalid", "alias@example.invalid"})
	if err != nil || len(r.To) != 2 || len(r.Cc) != 0 || len(r.Bcc) != 0 {
		t.Fatal("reply-all resolution failed")
	}
	s.Counts.To = n(3)
	if _, err = ResolveReply(s, true, []string{"me@example.invalid"}); err == nil {
		t.Fatal("truncated recipient list accepted")
	}
	s.Counts.To = n(2)
	s.ReplyTo = []SourceMailbox{m("reply@example.invalid")}
	r, err = ResolveReply(s, false, nil)
	if err != nil || len(r.To) != 1 || r.To[0].EmailAddress != "reply@example.invalid" {
		t.Fatal("reply-to not preferred")
	}
	s = ReplySource{From: &SourceParty{Mailbox: m("me@example.invalid")}, To: []SourceMailbox{m("me@example.invalid")}, Counts: &SourceRecipientCounts{To: n(1)}}
	r, err = ResolveReply(s, true, []string{"me@example.invalid"})
	if err != nil || len(r.To) != 1 || r.To[0].EmailAddress != "me@example.invalid" {
		t.Fatal("self-only reply lost")
	}
}
