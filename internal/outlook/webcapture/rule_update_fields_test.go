package webcapture_test

import (
	"testing"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/rules"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

// This tests composition into the observed SetInboxRule shape. A successful
// service readback is still needed to establish live update behavior.
func TestRuleUpdateRewritesObservedFieldsWithoutFixtureData(t *testing.T) {
	body := map[string]any{
		"__type": "SetInboxRuleRequest:#Exchange",
		"Header": map[string]any{
			"__type": "JsonRequestHeaders:#Exchange", "RequestServerVersion": "V2018_01_08",
			"TimeZoneContext": map[string]any{"__type": "TimeZoneContext:#Exchange", "TimeZoneDefinition": map[string]any{"__type": "TimeZoneDefinitionType:#Exchange", "Id": "Example Time"}},
		},
		"InboxRule": map[string]any{"Name": "fixture name", "SubjectContainsWords": []string{"fixture subject"}, "ApplyCategory": []string{"fixture category"}, "StopProcessingRules": false, "Identity": map[string]any{"DisplayName": "fixture name", "RawIdentity": "fixture identity"}},
		"Force":     false,
	}
	detail, _ := captureFile(t, "POST", "/owa/service.svc?action=SetInboxRule&app=Mail&n=1", "SetInboxRule", body, true)
	p := rules.Params{ID: "new-identity", Name: "New name", Subject: "new subject", Category: "Green"}
	_, op, err := rules.Prepare("update", detail, accountSession(), p)
	if err != nil {
		t.Fatal(err)
	}
	if op.Class != transport.Write {
		t.Fatal("rule update not classified as write")
	}
	result := unwrapBody(t, op)
	rule := result["InboxRule"].(map[string]any)
	identity := rule["Identity"].(map[string]any)
	if rule["Name"] != "New name" || rule["SubjectContainsWords"].([]any)[0] != "new subject" || rule["ApplyCategory"].([]any)[0] != "Green" || rule["StopProcessingRules"] != false || identity["RawIdentity"] != "new-identity" || identity["DisplayName"] != "New name" || result["Force"] != false {
		t.Fatal("updated rule request did not preserve the narrow captured shape")
	}
}
