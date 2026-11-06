package calendar

import (
	"encoding/json"
	"strings"
	"testing"
	"unicode/utf8"
)

func TestBriefAgendaReportsUnknownAndTruncation(t *testing.T) {
	raw := []byte(`{"Body":{"ResponseClass":"Success","ResponseCode":"NoError","Items":[{"ItemId":{"Id":"one"},"Subject":"A","Start":"start","End":"end","StartTimeZoneId":"UTC","EndTimeZoneId":"UTC","IsAllDayEvent":false,"Body":"private"},{"ItemId":{"Id":"two"},"Start":"start"},{"Subject":"no id"}]}}`)
	got, err := BriefAgendaView(raw, 1)
	if err != nil || got.ItemsSeen != 3 || got.MissingIDs != 1 || got.MissingFields != 1 || len(got.Events) != 1 || got.Events[0].ID != "one" || !got.Truncated || got.Complete || got.Events[0].StartTimeZone == nil || *got.Events[0].StartTimeZone != "UTC" || got.Events[0].IsAllDay == nil || *got.Events[0].IsAllDay {
		t.Fatalf("invalid bounded projection: %+v %v", got, err)
	}
	if _, err = BriefAgendaView([]byte(`{"Body":{"ResponseClass":"Error","ResponseCode":"ErrorAccessDenied","Items":[]}}`), 1); err == nil {
		t.Fatal("error envelope accepted")
	}
	if _, err = BriefAgendaView(raw, 51); err == nil {
		t.Fatal("oversized limit accepted")
	}
}

func TestBriefAgendaSubjectIsBoundedWithoutSplittingUTF8(t *testing.T) {
	subject := strings.Repeat("é", 300)
	raw, _ := json.Marshal(map[string]any{"Body": map[string]any{"ResponseClass": "Success", "ResponseCode": "NoError", "Items": []any{map[string]any{"ItemId": map[string]string{"Id": "item"}, "Subject": subject, "Start": "start", "End": "end"}}}})
	result, err := BriefAgendaView(raw, 1)
	if err != nil || len(result.Events) != 1 || !result.Events[0].SubjectTruncated || !utf8.ValidString(*result.Events[0].Subject) || len(*result.Events[0].Subject) > 512 {
		t.Fatal("calendar subject truncation was not safe and visible")
	}
}
