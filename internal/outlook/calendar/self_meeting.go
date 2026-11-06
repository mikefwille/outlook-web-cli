package calendar

import (
	"encoding/base64"
	"encoding/json"
	"fmt"
	"strings"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

type selfMeetingIdentity struct{ Email, ObjectID, Name string }

func selfIdentity(session transport.Session) (selfMeetingIdentity, error) {
	email, err := sessionUPN(session)
	if err != nil {
		return selfMeetingIdentity{}, err
	}
	auth := strings.SplitN(session.Headers.Get("Authorization"), " ", 2)
	parts := strings.Split(auth[1], ".")
	data, err := base64.RawURLEncoding.DecodeString(parts[1])
	if err != nil {
		return selfMeetingIdentity{}, fmt.Errorf("current session lacks self meeting identity")
	}
	var claims struct {
		ObjectID string `json:"oid"`
		Name     string `json:"name"`
	}
	if json.Unmarshal(data, &claims) != nil || claims.ObjectID == "" || len(claims.ObjectID) > 128 || claims.Name == "" || len(claims.Name) > 512 || strings.ContainsAny(claims.ObjectID+claims.Name, "\r\n\x00") {
		return selfMeetingIdentity{}, fmt.Errorf("current session lacks self meeting identity")
	}
	return selfMeetingIdentity{Email: email, ObjectID: claims.ObjectID, Name: claims.Name}, nil
}

func bindSelfMeetingAttendee(item map[string]any, session transport.Session, allowRuntimePlaceholder bool) error {
	self, err := selfIdentity(session)
	if err != nil {
		return err
	}
	attendees, ok := item["RequiredAttendees"].([]any)
	if !ok || len(attendees) != 1 || !arrayEmpty(item["OptionalAttendees"]) || !arrayEmpty(item["Resources"]) {
		return bad()
	}
	for _, key := range []string{"Organizer", "Sender", "From", "ReplyTo", "ToRecipients", "CcRecipients", "BccRecipients"} {
		if !arrayEmpty(item[key]) {
			return bad()
		}
	}
	attendee := obj(attendees[0])
	mailbox := obj(attendee["Mailbox"])
	if !exactKeys(attendee, "IsCoauthor", "Mailbox") || !exactKeys(mailbox, "EmailAddress", "Name", "MailboxType", "RoutingType", "AADObjectId") || attendee["IsCoauthor"] != false || mailbox["MailboxType"] != "InternalMailbox" || mailbox["RoutingType"] != "SMTP" {
		return bad()
	}
	matchesSelf := strings.EqualFold(fmt.Sprint(mailbox["EmailAddress"]), self.Email) && strings.EqualFold(fmt.Sprint(mailbox["AADObjectId"]), self.ObjectID)
	placeholder := allowRuntimePlaceholder && mailbox["EmailAddress"] == "template@example.invalid" && mailbox["AADObjectId"] == "template-object" && mailbox["Name"] == "template self"
	if !matchesSelf && !placeholder {
		return bad()
	}
	attendee["Mailbox"] = map[string]any{"EmailAddress": self.Email, "Name": self.Name, "MailboxType": "InternalMailbox", "RoutingType": "SMTP", "AADObjectId": self.ObjectID}
	return nil
}

// safeSelfOnlineEvent matches only the observed organizer-owned, self-only
// online event; it does not authorize attendee notifications or RSVP.
func (e Event) safeSelfOnlineEvent(session transport.Session) bool {
	self, err := selfIdentity(session)
	if err != nil || e.IsMeeting == nil || *e.IsMeeting || e.IsOnlineMeeting == nil || !*e.IsOnlineMeeting || e.IsRecurring == nil || *e.IsRecurring || e.IsOrganizer == nil || !*e.IsOrganizer || e.MeetingRequestWasSent == nil || !*e.MeetingRequestWasSent || e.IsAllDayEvent == nil || *e.IsAllDayEvent || e.CalendarItemType != "Single" || len(e.RequiredAttendees) != 1 || e.OptionalAttendees == nil || len(e.OptionalAttendees) != 0 || e.Resources == nil || len(e.Resources) != 0 {
		return false
	}
	var raw struct {
		ResponseType string
		StartZone    string `json:"StartTimeZoneId"`
		EndZone      string `json:"EndTimeZoneId"`
		Organizer    struct {
			Mailbox struct {
				Email string `json:"EmailAddress"`
				SMTP  string `json:"SmtpEmailAddress"`
			}
		} `json:"Organizer"`
	}
	var attendee struct {
		Mailbox struct {
			Email string `json:"EmailAddress"`
		} `json:"Mailbox"`
	}
	if json.Unmarshal(e.Raw, &raw) != nil || json.Unmarshal(e.RequiredAttendees[0], &attendee) != nil {
		return false
	}
	return raw.ResponseType == "Organizer" && raw.StartZone == "Mountain Standard Time" && raw.EndZone == "Mountain Standard Time" && (strings.EqualFold(raw.Organizer.Mailbox.Email, self.Email) || strings.EqualFold(raw.Organizer.Mailbox.SMTP, self.Email)) && strings.EqualFold(attendee.Mailbox.Email, self.Email)
}
