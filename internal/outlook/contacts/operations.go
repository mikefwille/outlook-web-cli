// Package contacts implements personal PeopleGraph requests observed in the UI.
// It never targets the organization directory.
package contacts

import (
	"encoding/json"
	"fmt"
	"net/mail"
	"regexp"
	"strings"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/webcapture"
)

var idPattern = regexp.MustCompile(`^[A-Za-z0-9_-]+$`)
var itemPath = regexp.MustCompile(`^/PeopleGraphVx/v1\.0/contacts\('[A-Za-z0-9_-]+'\)$`)
var updatePath = regexp.MustCompile(`^/PeopleGraphVx/v1\.0/contacts\('[A-Za-z0-9_-]+'\)/UpdateCollections$`)

func validID(s string) bool { return len(s) <= 4096 && idPattern.MatchString(s) }
func ValidEmail(s string) bool {
	if len(s) == 0 || len(s) > 320 || s != strings.TrimSpace(s) {
		return false
	}
	a, err := mail.ParseAddress(s)
	return err == nil && a.Address == s && a.Name == ""
}
func bad() error {
	return fmt.Errorf("contact request does not match the captured personal-contact contract")
}
func object(v any) map[string]any { x, _ := v.(map[string]any); return x }
func empty(v any) bool {
	if v == nil {
		return true
	}
	a, ok := v.([]any)
	return ok && len(a) == 0
}

type Params struct {
	ID, First, Last                                  string
	Current                                          *Contact
	NameChange                                       bool
	Email, Phone, Company, Notes                     *string
	ClearEmail, ClearPhone, ClearCompany, ClearNotes bool
}

func userSource() map[string]any { return map[string]any{"type": "UserProvided"} }

// updateEntry selects one user-provided element and uses the observed POST,
// PATCH or DELETE action. Ambiguous collections fail rather than guessing.
func updateEntry(c *Contact, collection, requiredType string, value string, clear bool) (map[string]any, error) {
	if !clear && strings.TrimSpace(value) == "" {
		return nil, fmt.Errorf("%s value cannot be empty; use the clear flag", collection)
	}
	values, _ := c.body[collection].([]any)
	var chosen map[string]any
	for _, v := range values {
		x := object(v)
		if object(x["source"])["type"] != "UserProvided" {
			continue
		}
		if requiredType != "" && x["type"] != requiredType {
			continue
		}
		if chosen != nil {
			return nil, fmt.Errorf("%s has multiple matching values; choose an explicit value once supported", collection)
		}
		chosen = x
	}
	if chosen == nil {
		if clear {
			return nil, fmt.Errorf("%s has no matching editable value to clear", collection)
		}
		body := map[string]any{"source": userSource()}
		switch collection {
		case "emails":
			body["address"] = value
			body["type"] = "Work"
		case "phones":
			body["number"] = value
			body["type"] = "Business"
		case "positions":
			body["detail"] = map[string]any{"jobTitle": "", "company": map[string]any{"department": "", "displayName": value, "officeLocation": "", "pronunciation": "", "address": map[string]any{"type": "Unknown"}}}
			body["isCurrent"] = false
			body["categories"] = []any{}
		case "notes":
			body["detail"] = map[string]any{"content": value, "contentType": "text"}
		default:
			return nil, bad()
		}
		return map[string]any{"body": body, "action": "POST"}, nil
	}
	id, ok := chosen["id"].(string)
	if !ok || id == "" {
		return nil, bad()
	}
	if clear {
		if collection == "positions" && !companyOnly(chosen) {
			return nil, fmt.Errorf("company position has other information; refusing to delete it")
		}
		if collection == "notes" && object(chosen["detail"])["contentType"] != "text" {
			return nil, fmt.Errorf("only an existing plain-text note can be cleared")
		}
		return map[string]any{"id": id, "action": "DELETE"}, nil
	}
	body := map[string]any{"source": userSource()}
	switch collection {
	case "emails":
		body["address"] = value
		body["type"] = "Work"
	case "phones":
		body["number"] = value
		body["type"] = "Business"
	case "notes":
		detail := object(chosen["detail"])
		if detail["contentType"] != "text" {
			return nil, fmt.Errorf("only an existing plain-text note can be updated")
		}
		body["detail"] = map[string]any{"content": value, "contentType": "text"}
	case "positions":
		detail := object(chosen["detail"])
		company := object(detail["company"])
		address := object(company["address"])
		if detail == nil || company == nil || address == nil {
			return nil, bad()
		}
		if _, ok := detail["jobTitle"]; !ok {
			return nil, bad()
		}
		if _, ok := chosen["isCurrent"].(bool); !ok {
			return nil, bad()
		}
		if _, ok := chosen["categories"].([]any); !ok {
			return nil, bad()
		}
		for k := range detail {
			if k != "jobTitle" && k != "company" {
				return nil, bad()
			}
		}
		for k := range company {
			if k != "department" && k != "displayName" && k != "officeLocation" && k != "pronunciation" && k != "address" {
				return nil, bad()
			}
		}
		// Preserve the complete captured address object, title, category and status.
		copyCompany := map[string]any{}
		for k, v := range company {
			copyCompany[k] = v
		}
		copyCompany["displayName"] = value
		body["detail"] = map[string]any{"jobTitle": detail["jobTitle"], "company": copyCompany}
		body["isCurrent"] = chosen["isCurrent"]
		body["categories"] = chosen["categories"]
	default:
		return nil, bad()
	}
	return map[string]any{"id": id, "body": body, "action": "PATCH"}, nil
}

// The captured clear-company operation deletes a whole position. Only allow
// that when doing so cannot discard a job title, address, manager or metadata.
func companyOnly(position map[string]any) bool {
	detail := object(position["detail"])
	company := object(detail["company"])
	address := object(company["address"])
	if detail == nil || company == nil || address == nil {
		return false
	}
	if v := detail["jobTitle"]; v != nil && v != "" {
		return false
	}
	for k, v := range detail {
		if k != "jobTitle" && k != "company" && v != nil && v != "" {
			return false
		}
	}
	for k, v := range company {
		if k == "displayName" || k == "address" {
			continue
		}
		if v != nil && v != "" {
			return false
		}
	}
	for k, v := range address {
		if k == "type" && v == "Unknown" {
			continue
		}
		if v != nil && v != "" {
			return false
		}
	}
	current, ok := position["isCurrent"].(bool)
	if !ok || current {
		return false
	}
	if !empty(position["categories"]) || !empty(position["colleagues"]) || position["manager"] != nil {
		return false
	}
	return true
}

func Prepare(action, detail, request string, current transport.Session, p Params) (transport.Session, transport.Operation, error) {
	fail := func() (transport.Session, transport.Operation, error) {
		return transport.Session{}, transport.Operation{}, bad()
	}
	switch action {
	case "get", "lookup-email", "create", "update", "delete":
	default:
		return fail()
	}
	d, err := webcapture.Load(detail, request, current)
	if err != nil {
		return transport.Session{}, transport.Operation{}, err
	}
	var root map[string]any
	if action != "delete" && (len(d.Body) == 0 || json.Unmarshal(d.Body, &root) != nil) {
		return fail()
	}
	path := d.URL.Path
	class := transport.Read
	method := "POST"
	validate := validateObject
	switch action {
	case "lookup-email":
		if d.Method != "POST" || path != "/PeopleGraphVx/v1.0/peopleLookup" || p.Email == nil || !ValidEmail(*p.Email) || len(root) != 1 {
			return fail()
		}
		emails, ok := root["Emails"].([]any)
		if !ok || len(emails) != 1 {
			return fail()
		}
		root["Emails"] = []any{*p.Email}
		validate = validateEmailLookup
	case "get":
		if d.Method != "POST" || path != "/PeopleGraphVx/v1.0/peopleLookup" || !validID(p.ID) {
			return fail()
		}
		ids, ok := root["ContactIds"].([]any)
		if !ok || len(ids) != 1 {
			return fail()
		}
		root["ContactIds"] = []any{p.ID}
		validate = validateLookup
	case "create":
		if d.Method != "POST" || path != "/PeopleGraphVx/v1.0/contacts" || strings.TrimSpace(p.First+" "+p.Last) == "" || p.ClearEmail || p.ClearPhone || p.ClearCompany || p.ClearNotes {
			return fail()
		}
		names, ok := root["names"].([]any)
		if !ok || len(names) != 1 {
			return fail()
		}
		name := object(names[0])
		source := object(name["source"])
		if name == nil || source["type"] != "UserProvided" || len(source) != 1 {
			return fail()
		}
		allowedName := map[string]bool{"first": true, "last": true, "middle": true, "initials": true, "suffix": true, "title": true, "nickname": true, "maiden": true, "pronunciation": true, "source": true}
		for k := range name {
			if !allowedName[k] {
				return fail()
			}
		}
		for _, k := range []string{"emails", "phones", "positions", "addresses", "anniversaries", "notes", "relationships", "tags", "webAccounts", "websites", "extensions"} {
			if !empty(root[k]) {
				return fail()
			}
		}
		// Refuse to copy any additional name values from a fixture.
		for _, k := range []string{"middle", "initials", "suffix", "title", "nickname", "maiden"} {
			if v := name[k]; v != nil && v != "" {
				return fail()
			}
		}
		pronunciation := object(name["pronunciation"])
		for _, k := range []string{"first", "last"} {
			if v := pronunciation[k]; v != nil && v != "" {
				return fail()
			}
		}
		name["first"] = p.First
		name["last"] = p.Last
		for _, f := range []*string{p.Email, p.Phone, p.Company, p.Notes} {
			if f != nil && strings.TrimSpace(*f) == "" {
				return fail()
			}
		}
		if p.Email != nil {
			root["emails"] = []any{map[string]any{"address": *p.Email, "type": "Work", "source": userSource()}}
		}
		if p.Phone != nil {
			root["phones"] = []any{map[string]any{"number": *p.Phone, "type": "Business", "source": userSource()}}
		}
		if p.Company != nil {
			root["positions"] = []any{map[string]any{"detail": map[string]any{"jobTitle": "", "company": map[string]any{"department": "", "displayName": *p.Company, "officeLocation": "", "pronunciation": "", "address": map[string]any{"type": "Unknown"}}}, "isCurrent": false, "categories": []any{}, "source": userSource()}}
		}
		if p.Notes != nil {
			root["notes"] = []any{map[string]any{"detail": map[string]any{"content": *p.Notes, "contentType": "text"}, "source": userSource()}}
		}
		class = transport.Write
	case "update":
		if d.Method != "POST" || !updatePath.MatchString(path) || !validID(p.ID) || p.Current == nil || !p.Current.SafeEditable() || p.Current.ID != p.ID || p.NameChange && strings.TrimSpace(p.First+" "+p.Last) == "" {
			return fail()
		}
		names, ok := root["names"].([]any)
		if !ok || len(names) != 1 {
			return fail()
		}
		entry := object(names[0])
		if entry["action"] != "PATCH" {
			return fail()
		}
		for _, k := range []string{"emails", "phones", "positions", "addresses", "anniversaries", "notes", "relationships", "tags", "webAccounts", "websites"} {
			if !empty(root[k]) {
				return fail()
			}
		}
		if p.NameChange {
			entry["id"] = p.Current.NameID
			body := object(entry["body"])
			if body == nil {
				return fail()
			}
			for _, k := range []string{"first", "last", "middle", "initials", "suffix", "title", "nickname", "maiden", "pronunciation"} {
				if v, ok := p.Current.name[k]; ok {
					body[k] = v
				}
			}
			body["first"] = p.First
			body["last"] = p.Last
			body["source"] = userSource()
		} else {
			root["names"] = []any{}
		}
		for _, field := range []struct {
			collection, typ string
			value           *string
			clear           bool
		}{
			{"emails", "Work", p.Email, p.ClearEmail}, {"phones", "Business", p.Phone, p.ClearPhone}, {"positions", "", p.Company, p.ClearCompany}, {"notes", "", p.Notes, p.ClearNotes},
		} {
			if field.value != nil && field.clear {
				return fail()
			}
			if field.value == nil && !field.clear {
				continue
			}
			value := ""
			if field.value != nil {
				value = *field.value
			}
			e, err := updateEntry(p.Current, field.collection, field.typ, value, field.clear)
			if err != nil {
				return transport.Session{}, transport.Operation{}, err
			}
			root[field.collection] = []any{e}
		}
		if !p.NameChange && p.Email == nil && p.Phone == nil && p.Company == nil && p.Notes == nil && !p.ClearEmail && !p.ClearPhone && !p.ClearCompany && !p.ClearNotes {
			return fail()
		}
		path = "/PeopleGraphVx/v1.0/contacts('" + p.ID + "')/UpdateCollections"
		class = transport.Write
	case "delete":
		if d.Method != "DELETE" || !itemPath.MatchString(path) || !validID(p.ID) || len(d.Body) != 0 {
			return fail()
		}
		path = "/PeopleGraphVx/v1.0/contacts('" + p.ID + "')"
		method = "DELETE"
		class = transport.Write
		validate = validateDelete
	}
	var body []byte
	if action != "delete" {
		body, err = json.Marshal(root)
		if err != nil {
			return fail()
		}
	}
	return d.Session, d.Operation(method, path, body, class, validate), nil
}

type Contact struct {
	ID                  string          `json:"id"`
	ChangeKey           string          `json:"change_key"`
	First               string          `json:"first"`
	Last                string          `json:"last"`
	NameID              string          `json:"name_id"`
	Raw                 json.RawMessage `json:"raw"`
	name                map[string]any
	body                map[string]any
	editable, deletable bool
}

func ParseContact(data []byte, lookup bool) (Contact, transport.EnvelopeResult) {
	var raw json.RawMessage = data
	if lookup {
		var wrapper struct {
			Value []json.RawMessage `json:"value"`
		}
		if json.Unmarshal(data, &wrapper) != nil || len(wrapper.Value) != 1 {
			return Contact{}, transport.EnvelopeError
		}
		raw = wrapper.Value[0]
	}
	var v struct {
		ID        string           `json:"id"`
		ChangeKey string           `json:"changeKey"`
		Names     []map[string]any `json:"names"`
	}
	// Decode source booleans separately because their JSON names differ.
	var body map[string]any
	if json.Unmarshal(raw, &v) != nil || json.Unmarshal(raw, &body) != nil || !validID(v.ID) || v.ChangeKey == "" {
		return Contact{}, transport.EnvelopeError
	}
	c := Contact{ID: v.ID, ChangeKey: v.ChangeKey, Raw: append(json.RawMessage(nil), raw...), body: body}
	for _, n := range v.Names {
		if object(n["source"])["type"] != "UserProvided" {
			continue
		}
		if c.name != nil {
			return Contact{}, transport.EnvelopeError
		}
		c.name = n
		c.NameID, _ = n["id"].(string)
		c.First, _ = n["first"].(string)
		c.Last, _ = n["last"].(string)
	}
	sources, _ := body["personSources"].([]any)
	// A unified record linked to directory or unknown sources is not proven
	// safe for mutation by the captured personal-contact lifecycle.
	if len(sources) != 1 {
		return c, transport.EnvelopeOK
	}
	for _, v := range sources {
		s := object(v)
		if s["type"] == "UserProvided" {
			e, _ := s["isEditable"].(bool)
			d, _ := s["isDeletable"].(bool)
			c.editable = c.editable || e
			c.deletable = c.deletable || d
		}
	}
	return c, transport.EnvelopeOK
}
func (c Contact) SafeEditable() bool                   { return c.editable && c.name != nil && c.NameID != "" }
func (c Contact) SafeDeletable() bool                  { return c.deletable }
func validateObject(b []byte) transport.EnvelopeResult { _, s := ParseContact(b, false); return s }
func validateLookup(b []byte) transport.EnvelopeResult { _, s := ParseContact(b, true); return s }

// ParseProfile accepts the full observed People lookup only when the returned
// identity matches the requested People ID. A personal-only record is not
// presented as a directory/profile record; nil with EnvelopeOK means that the
// service returned a valid record with no observed Profile source.
func ParseProfile(data []byte, id string) (json.RawMessage, transport.EnvelopeResult) {
	c, state := ParseContact(data, true)
	if state != transport.EnvelopeOK || c.ID != id {
		return nil, transport.EnvelopeError
	}
	var record struct {
		Sources []struct {
			Type string `json:"type"`
		} `json:"personSources"`
	}
	if json.Unmarshal(c.Raw, &record) != nil || record.Sources == nil {
		return nil, transport.EnvelopeError
	}
	for _, source := range record.Sources {
		if source.Type == "Profile" {
			return append(json.RawMessage(nil), c.Raw...), transport.EnvelopeOK
		}
	}
	return nil, transport.EnvelopeOK
}

// ParseEmailLookup returns the exact People results, including any directory
// profiles. An empty value array is an ordinary successful result.
func ParseEmailLookup(data []byte) ([]json.RawMessage, transport.EnvelopeResult) {
	var root struct {
		Value json.RawMessage `json:"value"`
	}
	if json.Unmarshal(data, &root) != nil || len(root.Value) == 0 || string(root.Value) == "null" {
		return nil, transport.EnvelopeError
	}
	var values []json.RawMessage
	if json.Unmarshal(root.Value, &values) != nil {
		return nil, transport.EnvelopeError
	}
	for _, raw := range values {
		var item map[string]json.RawMessage
		if json.Unmarshal(raw, &item) != nil || item == nil {
			return nil, transport.EnvelopeError
		}
	}
	if values == nil {
		values = []json.RawMessage{}
	}
	return values, transport.EnvelopeOK
}
func validateEmailLookup(b []byte) transport.EnvelopeResult { _, s := ParseEmailLookup(b); return s }
func validateDelete(b []byte) transport.EnvelopeResult {
	if len(strings.TrimSpace(string(b))) == 0 {
		return transport.EnvelopeOK
	}
	return transport.EnvelopeError
}
