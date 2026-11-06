// Package mailcache holds a private, account-scoped reconciliation of one
// observed mail-list query. It never writes to Outlook or changes drafts.
package mailcache

import (
	"encoding/json"
	"fmt"
)

type Item struct {
	ID, ChangeKey string
	Data          json.RawMessage
}
type Page struct {
	Items       []Item
	Next, Total int
	Last        bool
}

// ParsePage requires explicit paging and version metadata. A missing field is
// not interpreted as an empty page or proof that a folder has been exhausted.
func ParsePage(data []byte, folderID string) (Page, error) {
	var x struct {
		Body struct {
			Responses struct {
				Items []struct {
					Class string `json:"ResponseClass"`
					Code  string `json:"ResponseCode"`
					Root  *struct {
						Items *[]json.RawMessage `json:"Items"`
						Next  *int               `json:"IndexedPagingOffset"`
						Total *int               `json:"TotalItemsInView"`
						Last  *bool              `json:"IncludesLastItemInRange"`
					} `json:"RootFolder"`
				} `json:"Items"`
			} `json:"ResponseMessages"`
		} `json:"Body"`
	}
	bad := func() (Page, error) { return Page{}, fmt.Errorf("mail page lacks complete paging or item metadata") }
	if json.Unmarshal(data, &x) != nil || len(x.Body.Responses.Items) != 1 {
		return bad()
	}
	r := x.Body.Responses.Items[0]
	if r.Class != "Success" || r.Code != "NoError" || r.Root == nil || r.Root.Items == nil || r.Root.Next == nil || r.Root.Total == nil || r.Root.Last == nil || *r.Root.Total < 0 || len(*r.Root.Items) > 1000 {
		return bad()
	}
	p := Page{Next: *r.Root.Next, Total: *r.Root.Total, Last: *r.Root.Last, Items: make([]Item, 0, len(*r.Root.Items))}
	seen := map[string]bool{}
	for _, raw := range *r.Root.Items {
		var v struct {
			ItemID struct {
				ID        string `json:"Id"`
				ChangeKey string `json:"ChangeKey"`
			} `json:"ItemId"`
			Parent struct {
				ID string `json:"Id"`
			} `json:"ParentFolderId"`
		}
		if json.Unmarshal(raw, &v) != nil || v.ItemID.ID == "" || v.ItemID.ChangeKey == "" || len(v.ItemID.ID) > 4096 || len(v.ItemID.ChangeKey) > 4096 || v.Parent.ID != folderID || seen[v.ItemID.ID] {
			return bad()
		}
		seen[v.ItemID.ID] = true
		p.Items = append(p.Items, Item{ID: v.ItemID.ID, ChangeKey: v.ItemID.ChangeKey, Data: append(json.RawMessage(nil), raw...)})
	}
	return p, nil
}
