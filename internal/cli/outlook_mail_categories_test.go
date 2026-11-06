package cli

import (
	"encoding/json"
	"testing"
)

func TestMessageCategoryReadDistinguishesMissingAndEmpty(t *testing.T) {
	for _, tc := range []struct {
		raw      string
		complete bool
		length   int
	}{
		{`{"ItemId":{"Id":"item","ChangeKey":"version"},"Categories":[]}`, true, 0},
		{`{"ItemId":{"Id":"item","ChangeKey":"version"},"Categories":["Keep"]}`, true, 1},
		{`{"ItemId":{"Id":"item","ChangeKey":"version"}}`, false, 0},
		{`{"ItemId":{"Id":"item","ChangeKey":"version"},"Categories":null}`, false, 0},
	} {
		var item messageCategoryState
		if json.Unmarshal([]byte(tc.raw), &item) != nil || item.ItemID.ID != "item" || item.ItemID.ChangeKey != "version" || (item.Categories != nil) != tc.complete {
			t.Fatal("category completeness was not preserved")
		}
		if tc.complete && len(*item.Categories) != tc.length {
			t.Fatal("category list length incorrect")
		}
	}
}
