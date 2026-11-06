// Package contactlists implements only the personal distribution-list requests
// observed in Outlook on the web. No directory list or mail is modified.
package contactlists

import (
	"crypto/rand"
	"encoding/base64"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"net/url"
	"regexp"
	"strings"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/contacts"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/webcapture"
)

const Path = "/api/v2.0/me/personaldistributionlists"
const Select = "id,createdDateTime,lastModifiedDateTime,personId,displayName,notes,members"
const MembersSelect = "id,createdDateTime,lastModifiedDateTime,personId,displayName,notes"
const ServerLimit = 500

var idRE = regexp.MustCompile(`^[A-Za-z0-9_-]+={0,2}$`)
var numberRE = regexp.MustCompile(`^[0-9]{1,12}$`)

func ValidID(s string) bool { return len(s) > 0 && len(s) <= 4096 && idRE.MatchString(s) }

func ValidName(s string) bool {
	return len(s) > 0 && len(s) <= 255 && strings.TrimSpace(s) == s && !strings.ContainsAny(s, "\x00\r\n")
}

func ValidNotes(s string) bool { return len(s) <= 32768 && !strings.ContainsRune(s, '\x00') }

type Params struct {
	ID, Name, Notes, Email, MemberName string
}

func fail() error {
	return fmt.Errorf("contact list request does not match the observed personal-list contract")
}

func checkQuery(q url.Values, action string) bool {
	for key, values := range q {
		if len(values) != 1 || len(values[0]) > 8192 || strings.ContainsAny(values[0], "\x00\r\n") {
			return false
		}
		switch key {
		case "$select", "$expand", "$orderBy", "$top", "n", "cri", "cv":
		default:
			return false
		}
	}
	if !numberRE.MatchString(q.Get("n")) || q.Get("cri") == "" || q.Get("cv") == "" {
		return false
	}
	switch action {
	case "list":
		return len(q) == 6 && q.Get("$select") == Select && q.Get("$top") == "500" && q.Get("$orderBy") == "DisplayName Asc"
	case "create", "update":
		return len(q) == 4 && q.Get("$select") == Select
	case "add-member", "remove-member":
		return len(q) == 5 && q.Get("$select") == MembersSelect && q.Get("$expand") == "members"
	case "delete":
		return len(q) == 3
	}
	return false
}

// Fresh correlation values ensure private capture IDs are not copied into a
// portable runtime bundle. They are not credentials or account identifiers.
func freshQuery(q url.Values) (url.Values, error) {
	var b [32]byte
	if _, err := rand.Read(b[:]); err != nil {
		return nil, fail()
	}
	b[6] = b[6]&0x0f | 0x40
	b[8] = b[8]&0x3f | 0x80
	id := hex.EncodeToString(b[:16])
	q.Set("n", "1")
	q.Set("cri", id[:8]+"-"+id[8:12]+"-"+id[12:16]+"-"+id[16:20]+"-"+id[20:])
	q.Set("cv", base64.RawStdEncoding.EncodeToString(b[16:])+".0")
	return q, nil
}

func object(b []byte) (map[string]json.RawMessage, bool) {
	var m map[string]json.RawMessage
	err := json.Unmarshal(b, &m)
	return m, err == nil && m != nil
}

func exactKeys(m map[string]json.RawMessage, keys ...string) bool {
	if len(m) != len(keys) {
		return false
	}
	for _, k := range keys {
		if _, ok := m[k]; !ok {
			return false
		}
	}
	return true
}

func validCapturedBody(action string, b []byte) bool {
	m, ok := object(b)
	if !ok {
		return false
	}
	switch action {
	case "update":
		if !exactKeys(m, "DisplayName", "Notes") {
			return false
		}
		var name, notes string
		return json.Unmarshal(m["DisplayName"], &name) == nil && json.Unmarshal(m["Notes"], &notes) == nil && ValidName(name) && ValidNotes(notes)
	case "create", "add-member", "remove-member":
		if action == "create" {
			if !exactKeys(m, "DisplayName", "Members") {
				return false
			}
			var n string
			if json.Unmarshal(m["DisplayName"], &n) != nil || !ValidName(n) {
				return false
			}
		} else if !exactKeys(m, "Members") {
			return false
		}
		var members []json.RawMessage
		if json.Unmarshal(m["Members"], &members) != nil || len(members) != 1 {
			return false
		}
		x, ok := object(members[0])
		if !ok {
			return false
		}
		if action == "remove-member" {
			if !exactKeys(x, "Key", "RoutingType", "Type") {
				return false
			}
		} else if !exactKeys(x, "DisplayName", "Key", "RoutingType", "Type") {
			return false
		}
		var email, routing, typ string
		if json.Unmarshal(x["Key"], &email) != nil || json.Unmarshal(x["RoutingType"], &routing) != nil || json.Unmarshal(x["Type"], &typ) != nil || !contacts.ValidEmail(email) || routing != "SMTP" || typ != "OneOff" {
			return false
		}
		if action != "remove-member" {
			var n string
			return json.Unmarshal(x["DisplayName"], &n) == nil && ValidName(n)
		}
		return true
	}
	return false
}

// Prepare rewrites all fixture-specific fields and checks origin, account and
// token audience via webcapture before constructing the observed operation.
func Prepare(action, detail, request string, current transport.Session, p Params) (transport.Session, transport.Operation, error) {
	bad := func() (transport.Session, transport.Operation, error) {
		return transport.Session{}, transport.Operation{}, fail()
	}
	d, err := webcapture.Load(detail, request, current)
	if err != nil {
		return transport.Session{}, transport.Operation{}, err
	}
	q := d.URL.Query()
	if d.HeaderBody || !checkQuery(q, action) {
		return bad()
	}
	method, suffix := "", ""
	switch action {
	case "list":
		method = "GET"
	case "create":
		method = "POST"
	case "update":
		method = "PATCH"
	case "add-member":
		method = "POST"
		suffix = "/AddMembers"
	case "remove-member":
		method = "POST"
		suffix = "/DeleteMembers"
	case "delete":
		method = "DELETE"
	default:
		return bad()
	}
	if d.Method != method {
		return bad()
	}
	path := d.URL.Path
	if action == "list" || action == "create" {
		if path != Path {
			return bad()
		}
	} else {
		if !ValidID(p.ID) || !strings.HasPrefix(path, Path+"/") || !strings.HasSuffix(path, suffix) {
			return bad()
		}
		capturedID := strings.TrimSuffix(strings.TrimPrefix(path, Path+"/"), suffix)
		if !ValidID(capturedID) {
			return bad()
		}
		path = Path + "/" + p.ID + suffix
	}
	if action == "list" || action == "delete" {
		if len(d.Body) != 0 {
			return bad()
		}
	} else if !validCapturedBody(action, d.Body) {
		return bad()
	}
	var body []byte
	if action != "list" && action != "delete" {
		var value any
		switch action {
		case "create", "add-member":
			if !contacts.ValidEmail(p.Email) || !ValidName(p.MemberName) {
				return bad()
			}
			member := map[string]any{"DisplayName": p.MemberName, "Key": p.Email, "RoutingType": "SMTP", "Type": "OneOff"}
			if action == "create" {
				if !ValidName(p.Name) {
					return bad()
				}
				value = map[string]any{"DisplayName": p.Name, "Members": []any{member}}
			} else {
				value = map[string]any{"Members": []any{member}}
			}
		case "remove-member":
			if !contacts.ValidEmail(p.Email) {
				return bad()
			}
			value = map[string]any{"Members": []any{map[string]any{"Key": p.Email, "RoutingType": "SMTP", "Type": "OneOff"}}}
		case "update":
			if !ValidName(p.Name) || !ValidNotes(p.Notes) {
				return bad()
			}
			value = map[string]any{"DisplayName": p.Name, "Notes": p.Notes}
		}
		body, err = json.Marshal(value)
		if err != nil {
			return bad()
		}
	}
	q, err = freshQuery(q)
	if err != nil {
		return bad()
	}
	class := transport.Write
	validate := func(b []byte) transport.EnvelopeResult {
		if action == "delete" {
			if len(strings.TrimSpace(string(b))) == 0 {
				return transport.EnvelopeOK
			}
			return transport.EnvelopeError
		}
		item, state := ParseItem(b)
		if state != transport.EnvelopeOK || (p.ID != "" && item.ID != p.ID) || (action == "create" && item.Name != p.Name) {
			return transport.EnvelopeError
		}
		return transport.EnvelopeOK
	}
	if action == "list" {
		class = transport.Read
		validate = func(b []byte) transport.EnvelopeResult { _, s := ParsePage(b); return s }
	}
	op := d.Operation(method, path, body, class, validate)
	op.Query = q
	return d.Session, op, nil
}

type Item struct {
	ID      string            `json:"id"`
	ETag    string            `json:"etag"`
	Name    string            `json:"name"`
	Notes   string            `json:"notes"`
	Members []json.RawMessage `json:"members"`
}

func ParseItem(data []byte) (Item, transport.EnvelopeResult) {
	var x struct {
		ID      string          `json:"Id"`
		ETag    string          `json:"@odata.etag"`
		Name    *string         `json:"DisplayName"`
		Notes   *string         `json:"Notes"`
		Members json.RawMessage `json:"Members"`
	}
	if json.Unmarshal(data, &x) != nil || !ValidID(x.ID) || x.ETag == "" || len(x.ETag) > 1024 || strings.ContainsAny(x.ETag, "\x00\r\n") || x.Name == nil || !ValidName(*x.Name) || x.Notes == nil || !ValidNotes(*x.Notes) || len(x.Members) == 0 || string(x.Members) == "null" {
		return Item{}, transport.EnvelopeError
	}
	var members []json.RawMessage
	if json.Unmarshal(x.Members, &members) != nil || len(members) > 10000 {
		return Item{}, transport.EnvelopeError
	}
	for _, m := range members {
		if _, ok := object(m); !ok {
			return Item{}, transport.EnvelopeError
		}
	}
	if members == nil {
		members = []json.RawMessage{}
	}
	return Item{ID: x.ID, ETag: x.ETag, Name: *x.Name, Notes: *x.Notes, Members: members}, transport.EnvelopeOK
}

type Page struct {
	Items               []Item `json:"items"`
	Complete            bool   `json:"complete"`
	ContinuationPresent bool   `json:"continuation_present"`
}

func ParsePage(data []byte) (Page, transport.EnvelopeResult) {
	var root struct {
		Value json.RawMessage `json:"value"`
		Next  string          `json:"@odata.nextLink"`
	}
	if json.Unmarshal(data, &root) != nil || len(root.Value) == 0 || string(root.Value) == "null" {
		return Page{}, transport.EnvelopeError
	}
	var rows []json.RawMessage
	if json.Unmarshal(root.Value, &rows) != nil || len(rows) > ServerLimit {
		return Page{}, transport.EnvelopeError
	}
	items := make([]Item, 0, len(rows))
	seen := map[string]bool{}
	for _, row := range rows {
		item, state := ParseItem(row)
		if state != transport.EnvelopeOK || seen[item.ID] {
			return Page{}, transport.EnvelopeError
		}
		seen[item.ID] = true
		items = append(items, item)
	}
	return Page{Items: items, Complete: len(items) < ServerLimit && root.Next == "", ContinuationPresent: root.Next != ""}, transport.EnvelopeOK
}

// CheckMember only permits observed OneOff SMTP membership writes and refuses
// incomplete member representations or ambiguous duplicate matches.
func CheckMember(item Item, email string, adding bool) error {
	count := 0
	for _, raw := range item.Members {
		var m struct {
			Key     string `json:"Key"`
			Routing string `json:"RoutingType"`
			Type    string `json:"Type"`
		}
		if json.Unmarshal(raw, &m) != nil || m.Key == "" || m.Type == "" || m.Routing == "" {
			return fmt.Errorf("list has a member without a complete key and routing type; refusing membership change")
		}
		if strings.EqualFold(m.Key, email) {
			count++
			if m.Type != "OneOff" || m.Routing != "SMTP" {
				return fmt.Errorf("matching member is not an observed OneOff SMTP address")
			}
		}
	}
	if adding && count != 0 {
		return fmt.Errorf("member already exists")
	}
	if !adding && count != 1 {
		return fmt.Errorf("member was not found uniquely")
	}
	return nil
}
