package account

import (
	"encoding/json"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

type Category struct {
	ID               string `json:"id"`
	Name             string `json:"name"`
	Color            int    `json:"color"`
	KeyboardShortcut int    `json:"keyboard_shortcut"`
	LastTimeUsed     string `json:"last_time_used,omitempty"`
}

// ParseCategories exposes only the user's master categories from the observed
// configuration; unrelated session and preference fields never leave parsing.
func ParseCategories(data []byte) ([]Category, transport.EnvelopeResult) {
	var root struct {
		Master *struct {
			Items *[]struct {
				ID, Name, LastTimeUsed  string
				Color, KeyboardShortcut *int
			} `json:"MasterList"`
		} `json:"MasterCategoryList"`
	}
	if json.Unmarshal(data, &root) != nil || root.Master == nil || root.Master.Items == nil || len(*root.Master.Items) > 5000 {
		return nil, transport.EnvelopeError
	}
	items := make([]Category, 0, len(*root.Master.Items))
	seen := map[string]bool{}
	for _, x := range *root.Master.Items {
		if x.ID == "" || x.Name == "" || x.Color == nil || x.KeyboardShortcut == nil || seen[x.ID] {
			return nil, transport.EnvelopeError
		}
		seen[x.ID] = true
		items = append(items, Category{ID: x.ID, Name: x.Name, Color: *x.Color, KeyboardShortcut: *x.KeyboardShortcut, LastTimeUsed: x.LastTimeUsed})
	}
	return items, transport.EnvelopeOK
}

type Settings struct {
	SignatureText           *string `json:"signature_text"`
	SignatureHTML           *string `json:"signature_html"`
	AutoAddSignature        *bool   `json:"auto_add_signature"`
	AutoAddSignatureOnReply *bool   `json:"auto_add_signature_on_reply"`
	FocusedInboxEnabled     *bool   `json:"focused_inbox_enabled"`
}

// ParseSettings intentionally returns only independently observed settings.
// Automatic replies are not present in this response and are not inferred.
func ParseSettings(data []byte) (Settings, transport.EnvelopeResult) {
	var root struct {
		Options *struct {
			SignatureText           *string `json:"SignatureText"`
			SignatureHTML           *string `json:"SignatureHtml"`
			AutoAddSignature        *bool   `json:"AutoAddSignature"`
			AutoAddSignatureOnReply *bool   `json:"AutoAddSignatureOnReply"`
			FocusedInboxEnabled     *bool   `json:"IsFocusedInboxEnabled"`
		} `json:"UserOptions"`
	}
	if json.Unmarshal(data, &root) != nil || root.Options == nil || root.Options.AutoAddSignature == nil || root.Options.AutoAddSignatureOnReply == nil || root.Options.FocusedInboxEnabled == nil {
		return Settings{}, transport.EnvelopeError
	}
	x := root.Options
	return Settings{SignatureText: x.SignatureText, SignatureHTML: x.SignatureHTML, AutoAddSignature: x.AutoAddSignature, AutoAddSignatureOnReply: x.AutoAddSignatureOnReply, FocusedInboxEnabled: x.FocusedInboxEnabled}, transport.EnvelopeOK
}
