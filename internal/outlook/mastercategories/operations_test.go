package mastercategories

import (
	"encoding/json"
	"strings"
	"testing"
)

func TestMasterSnapshotsAndDeltaPreserveOtherCategories(t *testing.T) {
	beforeJSON := []byte(`{"MasterCategoryList":{"MasterList":[{"Name":"Keep","Color":1,"Id":"keep-id","LastTimeUsed":"2026-01-01T00:00:00Z","KeyboardShortcut":0,"Unknown":"preserve"},{"Name":"Old","Color":20,"Id":"old-id","LastTimeUsed":"2026-01-01T00:00:00Z","KeyboardShortcut":0}]}}`)
	before, err := ParseConfiguration(beforeJSON)
	if err != nil || len(before.Categories) != 2 || len(before.Version) != 64 || before.Categories[0].KnownFieldsOnly || !before.Categories[1].KnownFieldsOnly {
		t.Fatal("configuration or unknown fields not recognized")
	}
	afterJSON := []byte(`{"MasterList":[{"Name":"Keep","Color":1,"Id":"keep-id","LastTimeUsed":"2026-01-02T00:00:00Z","KeyboardShortcut":0,"Unknown":"preserve"},{"Name":"New","Color":23,"Id":"new-id","LastTimeUsed":"2026-01-01T00:00:00Z","KeyboardShortcut":0}],"WasSuccessful":true,"ErrorCode":0,"ResponseCode":"NoError","ResponseClass":"Success"}`)
	after, err := ParseResponse(afterJSON)
	if err != nil {
		t.Fatal(err)
	}
	p := Params{OldName: "Old", New: Category{Name: "New", Color: 23, KeyboardShortcut: 0}}
	if !VerifyDelta(before, after, "update", p) {
		t.Fatal("expected rename delta rejected")
	}
	changedUnknown := []byte(strings.Replace(string(afterJSON), `"Unknown":"preserve"`, `"Unknown":"changed"`, 1))
	changed, err := ParseResponse(changedUnknown)
	if err != nil || VerifyDelta(before, changed, "update", p) {
		t.Fatal("unrelated unknown-field change accepted")
	}
	after.Categories[0].Color = 2
	if VerifyDelta(before, after, "update", p) {
		t.Fatal("unrelated category change accepted")
	}
	if _, err = ParseConfiguration([]byte(`{"MasterCategoryList":{"MasterList":[{"Id":"a","Name":"X","Color":1}]}}`)); err == nil {
		t.Fatal("incomplete category accepted")
	}
	var doc map[string]any
	if json.Unmarshal(beforeJSON, &doc) != nil {
		t.Fatal("test input malformed")
	}
	if _, err = ParseResponse([]byte(`{"MasterList":[],"WasSuccessful":false,"ErrorCode":0,"ResponseCode":"NoError","ResponseClass":"Success"}`)); err == nil {
		t.Fatal("failed response accepted")
	}
}
