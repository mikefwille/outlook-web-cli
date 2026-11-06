package exchange

import (
	"encoding/json"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

// ValidateRead checks the operation-specific success envelope before a 2xx
// response can be presented as data. It deliberately retains raw item data
// for the caller; request-specific fields and bodies are not logged.
func ValidateRead(action string, data []byte) transport.EnvelopeResult {
	var root struct {
		WasSuccessful   *bool           `json:"WasSuccessful"`
		ErrorCode       *int            `json:"ErrorCode"`
		CalendarGroups  json.RawMessage `json:"CalendarGroups"`
		CalendarFolders json.RawMessage `json:"CalendarFolders"`
		AccountInfo     json.RawMessage `json:"AccountInfo"`
		Body            *struct {
			ResponseClass    string          `json:"ResponseClass"`
			ResponseCode     string          `json:"ResponseCode"`
			Conversations    json.RawMessage `json:"Conversations"`
			Items            json.RawMessage `json:"Items"`
			ResponseMessages *struct {
				Items []struct {
					ResponseClass string          `json:"ResponseClass"`
					ResponseCode  string          `json:"ResponseCode"`
					RootFolder    json.RawMessage `json:"RootFolder"`
					Conversation  json.RawMessage `json:"Conversation"`
				} `json:"Items"`
			} `json:"ResponseMessages"`
		} `json:"Body"`
	}
	if json.Unmarshal(data, &root) != nil {
		return transport.EnvelopeError
	}
	switch action {
	case "FindFolder":
		return ValidateFindFolder(data)
	case "GetCalendarFoldersV2":
		if root.WasSuccessful != nil && *root.WasSuccessful && root.ErrorCode != nil && *root.ErrorCode == 0 && isArray(root.CalendarGroups) && isArray(root.CalendarFolders) {
			return transport.EnvelopeOK
		}
	case "GetAccountInformation":
		if root.WasSuccessful != nil && *root.WasSuccessful && root.ErrorCode != nil && *root.ErrorCode == 0 && isObject(root.AccountInfo) {
			return transport.EnvelopeOK
		}
	case "FindConversation":
		if root.Body != nil && root.Body.ResponseClass == "Success" && root.Body.ResponseCode == "NoError" && isArray(root.Body.Conversations) {
			return transport.EnvelopeOK
		}
	case "GetCalendarView":
		if root.Body != nil && root.Body.ResponseClass == "Success" && root.Body.ResponseCode == "NoError" && isArray(root.Body.Items) {
			return transport.EnvelopeOK
		}
	case "FindItem", "GetConversationItems":
		if root.Body == nil || root.Body.ResponseMessages == nil || len(root.Body.ResponseMessages.Items) == 0 {
			return transport.EnvelopeError
		}
		good := 0
		for _, item := range root.Body.ResponseMessages.Items {
			if item.ResponseClass != "Success" || item.ResponseCode != "NoError" {
				continue
			}
			if action == "FindItem" && isObject(item.RootFolder) || action == "GetConversationItems" && isObject(item.Conversation) {
				good++
			}
		}
		if good == len(root.Body.ResponseMessages.Items) {
			return transport.EnvelopeOK
		}
		if good > 0 {
			return transport.EnvelopePartial
		}
	}
	return transport.EnvelopeError
}

func isArray(data json.RawMessage) bool {
	var v []json.RawMessage
	return len(data) > 0 && string(data) != "null" && json.Unmarshal(data, &v) == nil
}

func isObject(data json.RawMessage) bool {
	var v map[string]json.RawMessage
	return len(data) > 0 && string(data) != "null" && json.Unmarshal(data, &v) == nil
}
