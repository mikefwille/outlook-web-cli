package exchange

import (
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"testing"
)

func TestObservedReadEnvelopes(t *testing.T) {
	cases := []struct {
		action, body string
		want         transport.EnvelopeResult
	}{
		{"FindConversation", `{"Body":{"Conversations":[],"ResponseClass":"Success","ResponseCode":"NoError"}}`, transport.EnvelopeOK},
		{"FindConversation", `{"Body":{"Conversations":[],"ResponseClass":"Error","ResponseCode":"ErrorAccessDenied"}}`, transport.EnvelopeError},
		{"GetCalendarView", `{"Body":{"Items":[],"ResponseClass":"Success","ResponseCode":"NoError"}}`, transport.EnvelopeOK},
		{"GetCalendarView", `{"Body":{"ResponseClass":"Success","ResponseCode":"NoError"}}`, transport.EnvelopeError},
		{"GetCalendarFoldersV2", `{"CalendarGroups":[],"CalendarFolders":[],"WasSuccessful":true,"ErrorCode":0}`, transport.EnvelopeOK},
		{"GetAccountInformation", `{"AccountInfo":{},"WasSuccessful":true,"ErrorCode":0}`, transport.EnvelopeOK},
		{"GetCalendarFoldersV2", `{"CalendarGroups":[],"CalendarFolders":[],"WasSuccessful":false,"ErrorCode":1}`, transport.EnvelopeError},
		{"FindItem", `{"Body":{"ResponseMessages":{"Items":[{"RootFolder":{"Items":[]},"ResponseClass":"Success","ResponseCode":"NoError"}]}}}`, transport.EnvelopeOK},
		{"GetConversationItems", `{"Body":{"ResponseMessages":{"Items":[{"Conversation":{},"ResponseClass":"Success","ResponseCode":"NoError"},{"ResponseClass":"Error","ResponseCode":"ErrorAccessDenied"}]}}}`, transport.EnvelopePartial},
		{"UnknownAction", `{}`, transport.EnvelopeError},
	}
	for _, tc := range cases {
		if got := ValidateRead(tc.action, []byte(tc.body)); got != tc.want {
			t.Errorf("%s: got %v want %v", tc.action, got, tc.want)
		}
	}
}
