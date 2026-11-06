package exchange

import (
	"encoding/json"
	"net/url"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"testing"
)

func TestReplyToPropertyIsExplicitlyRequestedWithoutChangingIdentity(t *testing.T) {
	op := transport.Operation{Method: "POST", Path: "/owa/service.svc", Query: url.Values{"action": []string{"GetItem"}}, Class: transport.Read, Body: []byte(`{"Body":{"ItemShape":{"AdditionalProperties":[{"__type":"PropertyUri:#Exchange","FieldURI":"CanDelete"}]},"ItemIds":[{"Id":"source"}]}}`)}
	got, err := IncludeReplyTo(op)
	if err != nil {
		t.Fatal(err)
	}
	var root struct {
		Body struct {
			ItemShape struct {
				Additional []struct {
					Field string `json:"FieldURI"`
				} `json:"AdditionalProperties"`
			} `json:"ItemShape"`
			ItemIDs []struct {
				ID string `json:"Id"`
			} `json:"ItemIds"`
		} `json:"Body"`
	}
	if json.Unmarshal(got.Body, &root) != nil || len(root.Body.ItemShape.Additional) != 2 || root.Body.ItemShape.Additional[1].Field != "ReplyTo" || root.Body.ItemIDs[0].ID != "source" {
		t.Fatal("reply shape or item identity changed")
	}
	op.Class = transport.Write
	if _, err = IncludeReplyTo(op); err == nil {
		t.Fatal("accepted write operation")
	}
}

func TestIncludeCategoriesAddsOnlyRequestedProperty(t *testing.T) {
	op := transport.Operation{Method: "POST", Path: "/owa/service.svc", Query: url.Values{"action": []string{"GetItem"}}, Class: transport.Read, Body: []byte(`{"Body":{"ItemShape":{"AdditionalProperties":[{"__type":"PropertyUri:#Exchange","FieldURI":"CanDelete"}]},"ItemIds":[{"Id":"source"}]}}`)}
	result, err := IncludeCategories(op)
	if err != nil {
		t.Fatal(err)
	}
	var root struct {
		Body struct {
			Shape struct {
				Properties []struct {
					Field string `json:"FieldURI"`
				} `json:"AdditionalProperties"`
			} `json:"ItemShape"`
		} `json:"Body"`
	}
	if json.Unmarshal(result.Body, &root) != nil || len(root.Body.Shape.Properties) != 2 || root.Body.Shape.Properties[1].Field != "Categories" {
		t.Fatal("categories property missing")
	}
	again, err := IncludeCategories(result)
	if err != nil {
		t.Fatal(err)
	}
	if json.Unmarshal(again.Body, &root) != nil || len(root.Body.Shape.Properties) != 2 {
		t.Fatal("duplicate category property")
	}
}
