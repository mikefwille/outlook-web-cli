package drafts

import (
	"encoding/json"
	"strings"
	"testing"
)

func TestRichDraftKeepsHTMLAndCompanionText(t *testing.T) {
	source := `<div>Hello <strong>there</strong><br><a href="https://example.invalid">link</a></div>`
	plain, err := HTMLText(source)
	if err != nil || plain != "Hello there\nlink" {
		t.Fatalf("unexpected plain text: %q", plain)
	}
	create := []byte(`{"__type":"CreateItemJsonRequest:#Exchange","Body":{"MessageDisposition":"SaveOnly","Items":[{"__type":"Message:#Exchange","Subject":"keep","plainText":"Hello there\nlink","Body":{"BodyType":"HTML","Value":"placeholder"}}]}}`)
	result, err := SetHTMLBody(create, "CreateItem", source)
	if err != nil {
		t.Fatal(err)
	}
	var decoded struct {
		Body struct {
			Items []struct {
				Subject   string
				PlainText string `json:"plainText"`
				Body      struct{ Value string }
			}
		}
	}
	if json.Unmarshal(result, &decoded) != nil || len(decoded.Body.Items) != 1 || decoded.Body.Items[0].Subject != "keep" || decoded.Body.Items[0].PlainText != plain || decoded.Body.Items[0].Body.Value != source {
		t.Fatal("HTML create did not preserve subject, plain text and exact markup")
	}
	update := []byte(`{"__type":"UpdateItemJsonRequest:#Exchange","Body":{"MessageDisposition":"SaveOnly","ItemChanges":[{"Updates":[{"Path":{"FieldURI":"Subject"},"Item":{"__type":"Message:#Exchange","Subject":"keep"}},{"Path":{"FieldURI":"Body"},"Item":{"__type":"Message:#Exchange","Body":{"BodyType":"HTML","Value":"old"}}}]}]}}`)
	result, err = SetHTMLBody(update, "UpdateItem", source)
	var updated struct {
		Body struct {
			Changes []struct {
				Updates []struct {
					Item struct {
						Subject string
						Body    struct{ Value string }
					}
				}
			} `json:"ItemChanges"`
		}
	}
	if err != nil || json.Unmarshal(result, &updated) != nil || len(updated.Body.Changes) != 1 || len(updated.Body.Changes[0].Updates) != 2 || updated.Body.Changes[0].Updates[0].Item.Subject != "keep" || updated.Body.Changes[0].Updates[1].Item.Body.Value != source {
		t.Fatal("HTML update did not preserve other fields")
	}
	if _, err = SetHTMLBody([]byte(strings.Replace(string(update), "SaveOnly", "SendAndSaveCopy", 1)), "UpdateItem", source); err == nil {
		t.Fatal("HTML rewrite accepted a send")
	}
	if _, err = HTMLText("bad\x00input"); err == nil {
		t.Fatal("NUL accepted")
	}
}
