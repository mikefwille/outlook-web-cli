package drafts

import "encoding/json"

// SelectUpdateFields filters a previously validated SaveOnly draft update.
// Omitting a field leaves the server's existing value intact, so a recipient
// or subject-only edit does not rewrite a potentially rich or truncated body.
func SelectUpdateFields(body []byte, subject, text bool) ([]byte, error) {
	var root map[string]any
	if json.Unmarshal(body, &root) != nil || root["__type"] != "UpdateItemJsonRequest:#Exchange" {
		return nil, ErrInvalid
	}
	b, ok := root["Body"].(map[string]any)
	if !ok || b["MessageDisposition"] != "SaveOnly" || b["ConflictResolution"] != "NeverOverwrite" || b["SendCalendarInvitationsOrCancellations"] != "SendToNone" {
		return nil, ErrInvalid
	}
	changes, ok := b["ItemChanges"].([]any)
	if !ok || len(changes) != 1 {
		return nil, ErrInvalid
	}
	change, ok := changes[0].(map[string]any)
	if !ok {
		return nil, ErrInvalid
	}
	updates, ok := change["Updates"].([]any)
	if !ok {
		return nil, ErrInvalid
	}
	filtered := make([]any, 0, len(updates))
	seen := map[string]bool{}
	for _, entry := range updates {
		u, ok := entry.(map[string]any)
		if !ok {
			return nil, ErrInvalid
		}
		path, ok := u["Path"].(map[string]any)
		if !ok {
			return nil, ErrInvalid
		}
		field, ok := path["FieldURI"].(string)
		if !ok || seen[field] {
			return nil, ErrInvalid
		}
		seen[field] = true
		switch field {
		case "Subject":
			if subject {
				filtered = append(filtered, entry)
			}
		case "Body":
			if text {
				filtered = append(filtered, entry)
			}
		case "ToRecipients", "CcRecipients", "BccRecipients":
			filtered = append(filtered, entry)
		default:
			return nil, ErrInvalid
		}
	}
	if subject && !seen["Subject"] || text && !seen["Body"] || len(filtered) == 0 {
		return nil, ErrInvalid
	}
	change["Updates"] = filtered
	return json.Marshal(root)
}
