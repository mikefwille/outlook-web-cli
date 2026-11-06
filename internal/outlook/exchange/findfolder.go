// Package exchange contains parsers for observed Outlook web RPC envelopes.
package exchange

import (
	"encoding/json"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

type FolderID struct {
	ID        string `json:"Id"`
	ChangeKey string `json:"ChangeKey"`
}

type Folder struct {
	ID              FolderID `json:"FolderId"`
	ParentID        FolderID `json:"ParentFolderId"`
	DisplayName     string   `json:"DisplayName"`
	FolderClass     string   `json:"FolderClass"`
	DistinguishedID string   `json:"DistinguishedFolderId"`
	UnreadCount     *int     `json:"UnreadCount"`
	TotalCount      *int     `json:"TotalCount"`
	ChildCount      *int     `json:"ChildFolderCount"`
}

type FolderPage struct {
	Folders                 []Folder
	Parent                  *Folder
	IndexedPagingOffset     int
	TotalItemsInView        int
	IncludesLastItemInRange bool
}

type ItemResult struct {
	Index   int
	Success bool
	Page    *FolderPage
}

type findFolderWire struct {
	Body *struct {
		ResponseMessages *struct {
			Items []struct {
				ResponseClass string `json:"ResponseClass"`
				ResponseCode  string `json:"ResponseCode"`
				RootFolder    *struct {
					Folders                 *[]Folder `json:"Folders"`
					ParentFolder            *Folder   `json:"ParentFolder"`
					IndexedPagingOffset     int       `json:"IndexedPagingOffset"`
					TotalItemsInView        int       `json:"TotalItemsInView"`
					IncludesLastItemInRange *bool     `json:"IncludesLastItemInRange"`
				} `json:"RootFolder"`
			} `json:"Items"`
		} `json:"ResponseMessages"`
	} `json:"Body"`
}

// ParseFindFolder follows the observed Header/Body/ResponseMessages envelope.
// It fails closed on absent or malformed response metadata and preserves
// per-item outcomes rather than turning a partial batch into success.
func ParseFindFolder(data []byte) ([]ItemResult, transport.EnvelopeResult) {
	var wire findFolderWire
	if json.Unmarshal(data, &wire) != nil || wire.Body == nil || wire.Body.ResponseMessages == nil || len(wire.Body.ResponseMessages.Items) == 0 {
		return nil, transport.EnvelopeError
	}
	results := make([]ItemResult, 0, len(wire.Body.ResponseMessages.Items))
	successes := 0
	for i, item := range wire.Body.ResponseMessages.Items {
		result := ItemResult{Index: i}
		if item.ResponseClass == "Success" && item.ResponseCode == "NoError" && item.RootFolder != nil && item.RootFolder.Folders != nil && item.RootFolder.IncludesLastItemInRange != nil {
			folders := append([]Folder{}, (*item.RootFolder.Folders)...)
			result.Success = true
			result.Page = &FolderPage{Folders: folders, Parent: item.RootFolder.ParentFolder,
				IndexedPagingOffset: item.RootFolder.IndexedPagingOffset, TotalItemsInView: item.RootFolder.TotalItemsInView,
				IncludesLastItemInRange: *item.RootFolder.IncludesLastItemInRange}
			successes++
		}
		results = append(results, result)
	}
	if successes == len(results) {
		return results, transport.EnvelopeOK
	}
	if successes > 0 {
		return results, transport.EnvelopePartial
	}
	return results, transport.EnvelopeError
}

func ValidateFindFolder(data []byte) transport.EnvelopeResult {
	_, result := ParseFindFolder(data)
	return result
}
