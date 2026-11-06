package exchange

import (
	"encoding/json"
	"testing"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func TestObservedEmptyFolderShapeIsAnEmptyArray(t *testing.T) {
	data := []byte(`{"Header":{"ServerVersionInfo":{}},"Body":{"__type":"FindFolderResponse:#Exchange","ResponseMessages":{"Items":[{"ResponseClass":"Success","ResponseCode":"NoError","RootFolder":{"Folders":[],"IndexedPagingOffset":0,"TotalItemsInView":0,"IncludesLastItemInRange":true}}]}}}`)
	items, state := ParseFindFolder(data)
	if state != transport.EnvelopeOK || len(items) != 1 || items[0].Page == nil || items[0].Page.Folders == nil || len(items[0].Page.Folders) != 0 {
		t.Fatalf("unexpected empty page: %#v state=%v", items, state)
	}
	encoded, _ := json.Marshal(items[0].Page.Folders)
	if string(encoded) != "[]" {
		t.Fatalf("empty folders encoded as %s", encoded)
	}
}

func TestFolderPagePreservesIDsAndContinuationMetadata(t *testing.T) {
	data := []byte(`{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Success","ResponseCode":"NoError","RootFolder":{"Folders":[{"FolderId":{"Id":"synthetic-id","ChangeKey":"synthetic-change"},"DisplayName":"Fixture","ChildFolderCount":2}],"IndexedPagingOffset":10,"TotalItemsInView":20,"IncludesLastItemInRange":false}}]}}}`)
	items, state := ParseFindFolder(data)
	if state != transport.EnvelopeOK || items[0].Page.Folders[0].ID.ID != "synthetic-id" || items[0].Page.Folders[0].ID.ChangeKey != "synthetic-change" || items[0].Page.IndexedPagingOffset != 10 || items[0].Page.IncludesLastItemInRange {
		t.Fatalf("lost fields: %#v state=%v", items, state)
	}
}

func TestExchangeErrorsAndMalformedBodiesAreNotEmptySuccess(t *testing.T) {
	cases := [][]byte{
		[]byte(`{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Error","ResponseCode":"ErrorAccessDenied"}]}}}`),
		[]byte(`{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Success","ResponseCode":"NoError","RootFolder":{}}]}}}`),
		[]byte(`{"Body":{"ResponseMessages":{"Items":[]}}}`),
		[]byte(`not-json`),
	}
	for _, data := range cases {
		if _, result := ParseFindFolder(data); result != transport.EnvelopeError {
			t.Fatalf("accepted %s", data)
		}
	}
}

func TestMixedResponsePreservesPerItemOutcome(t *testing.T) {
	data := []byte(`{"Body":{"ResponseMessages":{"Items":[{"ResponseClass":"Success","ResponseCode":"NoError","RootFolder":{"Folders":[],"IncludesLastItemInRange":true}},{"ResponseClass":"Error","ResponseCode":"ErrorAccessDenied"}]}}}`)
	items, state := ParseFindFolder(data)
	if state != transport.EnvelopePartial || len(items) != 2 || !items[0].Success || items[1].Success || items[1].Page != nil {
		t.Fatalf("partial outcome: %#v state=%v", items, state)
	}
}
