package calendar

import (
	"encoding/json"
	"fmt"
	"unicode/utf8"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/exchange"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

type BriefEvent struct {
	ID               string  `json:"id"`
	Subject          *string `json:"subject,omitempty"`
	SubjectTruncated bool    `json:"subject_truncated,omitempty"`
	Start            *string `json:"start,omitempty"`
	End              *string `json:"end,omitempty"`
	StartTimeZone    *string `json:"start_time_zone,omitempty"`
	EndTimeZone      *string `json:"end_time_zone,omitempty"`
	IsAllDay         *bool   `json:"is_all_day,omitempty"`
}
type BriefAgenda struct {
	ItemsSeen     int          `json:"items_seen"`
	MissingIDs    int          `json:"items_without_id"`
	MissingFields int          `json:"items_with_missing_subject_or_times"`
	Events        []BriefEvent `json:"events"`
	Truncated     bool         `json:"truncated"`
	Complete      bool         `json:"complete"`
}

// BriefAgendaView projects a single observed agenda response. The service may
// have omitted events; source order and missing fields remain explicit.
func BriefAgendaView(data []byte, limit int) (BriefAgenda, error) {
	if limit < 1 || limit > 50 || exchange.ValidateRead("GetCalendarView", data) != transport.EnvelopeOK {
		return BriefAgenda{}, fmt.Errorf("invalid bounded calendar response or limit")
	}
	var x struct {
		Body struct {
			Items []json.RawMessage `json:"Items"`
		} `json:"Body"`
	}
	if json.Unmarshal(data, &x) != nil {
		return BriefAgenda{}, fmt.Errorf("invalid calendar response")
	}
	out := BriefAgenda{ItemsSeen: len(x.Body.Items), Events: []BriefEvent{}}
	for _, raw := range x.Body.Items {
		var e struct {
			ItemID *struct {
				ID string `json:"Id"`
			} `json:"ItemId"`
			Subject       *string `json:"Subject"`
			Start         *string `json:"Start"`
			End           *string `json:"End"`
			StartTimeZone *string `json:"StartTimeZoneId"`
			EndTimeZone   *string `json:"EndTimeZoneId"`
			IsAllDay      *bool   `json:"IsAllDayEvent"`
		}
		if json.Unmarshal(raw, &e) != nil {
			return BriefAgenda{}, fmt.Errorf("invalid calendar event")
		}
		if e.ItemID == nil || e.ItemID.ID == "" || len(e.ItemID.ID) > 4096 {
			out.MissingIDs++
			continue
		}
		if e.Start != nil && len(*e.Start) > 64 {
			e.Start = nil
		}
		if e.End != nil && len(*e.End) > 64 {
			e.End = nil
		}
		if e.StartTimeZone != nil && len(*e.StartTimeZone) > 128 {
			e.StartTimeZone = nil
		}
		if e.EndTimeZone != nil && len(*e.EndTimeZone) > 128 {
			e.EndTimeZone = nil
		}
		if e.Subject == nil || e.Start == nil || e.End == nil {
			out.MissingFields++
		}
		if len(out.Events) == limit {
			out.Truncated = true
			continue
		}
		subject, truncated := briefSubject(e.Subject)
		out.Events = append(out.Events, BriefEvent{ID: e.ItemID.ID, Subject: subject, SubjectTruncated: truncated, Start: e.Start, End: e.End, StartTimeZone: e.StartTimeZone, EndTimeZone: e.EndTimeZone, IsAllDay: e.IsAllDay})
	}
	return out, nil
}

func briefSubject(s *string) (*string, bool) {
	if s == nil || len(*s) <= 512 {
		return s, false
	}
	v := (*s)[:512]
	for !utf8.ValidString(v) {
		v = v[:len(v)-1]
	}
	return &v, true
}
