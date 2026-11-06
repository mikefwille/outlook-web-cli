package drafts

import (
	"encoding/json"
	"testing"
)

func TestSelectUpdateFieldsPreservesOmittedValues(t *testing.T) {
	body := []byte(`{"__type":"UpdateItemJsonRequest:#Exchange","Body":{"MessageDisposition":"SaveOnly","ConflictResolution":"NeverOverwrite","SendCalendarInvitationsOrCancellations":"SendToNone","ItemChanges":[{"Updates":[{"Path":{"FieldURI":"Subject"}},{"Path":{"FieldURI":"Body"}},{"Path":{"FieldURI":"ToRecipients"}}]}]}}`)
	for _, tc := range []struct {
		subject, text bool
		want          []string
	}{{true, false, []string{"Subject", "ToRecipients"}}, {false, true, []string{"Body", "ToRecipients"}}, {false, false, []string{"ToRecipients"}}} {
		result, err := SelectUpdateFields(body, tc.subject, tc.text)
		if err != nil {
			t.Fatal(err)
		}
		var root struct {
			Body struct {
				Changes []struct {
					Updates []struct {
						Path struct {
							Field string `json:"FieldURI"`
						} `json:"Path"`
					} `json:"Updates"`
				} `json:"ItemChanges"`
			} `json:"Body"`
		}
		if json.Unmarshal(result, &root) != nil || len(root.Body.Changes) != 1 || len(root.Body.Changes[0].Updates) != len(tc.want) {
			t.Fatal("unexpected update shape")
		}
		for i, field := range tc.want {
			if root.Body.Changes[0].Updates[i].Path.Field != field {
				t.Fatal("unrequested field would be changed")
			}
		}
	}
	if _, err := SelectUpdateFields([]byte(`{}`), true, false); err == nil {
		t.Fatal("invalid update accepted")
	}
}
