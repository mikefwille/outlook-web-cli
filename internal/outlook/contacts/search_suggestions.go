package contacts

import (
	"crypto/rand"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/webcapture"
	"strings"
)

func randomID() (string, error) {
	var b [16]byte
	if _, err := rand.Read(b[:]); err != nil {
		return "", err
	}
	b[6] = (b[6] & 0x0f) | 0x40
	b[8] = (b[8] & 0x3f) | 0x80
	v := hex.EncodeToString(b[:])
	return v[:8] + "-" + v[8:12] + "-" + v[12:16] + "-" + v[16:20] + "-" + v[20:], nil
}

// PrepareSearch produces People suggestions from the captured People UI
// scenario. The observed response is bounded to six suggestions and combines
// mailbox and directory provenance; this is not an exhaustive search.
func PrepareSearch(detail, request string, current transport.Session, query string) (transport.Session, transport.Operation, error) {
	if strings.TrimSpace(query) == "" || len(query) > 512 || strings.ContainsRune(query, '\x00') {
		return transport.Session{}, transport.Operation{}, fmt.Errorf("search query must be nonempty and at most 512 bytes")
	}
	d, err := webcapture.Load(detail, request, current)
	if err != nil {
		return transport.Session{}, transport.Operation{}, err
	}
	if d.Method != "POST" || d.URL.Path != "/search/api/v1/suggestions" || d.URL.Query().Get("domain") != "People" || d.HeaderBody {
		return transport.Session{}, transport.Operation{}, bad()
	}
	var root map[string]any
	if json.Unmarshal(d.Body, &root) != nil || object(root["Scenario"])["Name"] != "owa.react.people" {
		return transport.Session{}, transport.Operation{}, bad()
	}
	entities, ok := root["EntityRequests"].([]any)
	if !ok || len(entities) != 1 {
		return transport.Session{}, transport.Operation{}, bad()
	}
	entity := object(entities[0])
	q := object(entity["Query"])
	provenances, ok := entity["Provenances"].([]any)
	if entity["EntityType"] != "People" || entity["Size"] != float64(6) || q == nil || !ok || len(provenances) != 2 || provenances[0] != "Mailbox" || provenances[1] != "Directory" {
		return transport.Session{}, transport.Operation{}, bad()
	}
	q["QueryString"] = query
	for _, key := range []string{"Cvid", "LogicalId"} {
		if _, ok := root[key]; !ok {
			return transport.Session{}, transport.Operation{}, bad()
		}
		id, e := randomID()
		if e != nil {
			return transport.Session{}, transport.Operation{}, e
		}
		root[key] = id
	}
	body, err := json.Marshal(root)
	if err != nil {
		return transport.Session{}, transport.Operation{}, bad()
	}
	return d.Session, d.Operation("POST", d.URL.Path, body, transport.Read, validateSearch), nil
}

type Suggestion struct {
	ID             string            `json:"id"`
	DisplayName    string            `json:"display_name"`
	GivenName      string            `json:"given_name,omitempty"`
	Surname        string            `json:"surname,omitempty"`
	EmailAddresses []string          `json:"email_addresses"`
	Phones         []json.RawMessage `json:"phones,omitempty"`
	PeopleType     string            `json:"people_type,omitempty"`
	PeopleSubtype  string            `json:"people_subtype,omitempty"`
	PersonID       string            `json:"person_id,omitempty"`
	PersonaID      string            `json:"persona_id,omitempty"`
	ADObjectID     string            `json:"directory_object_id,omitempty"`
	CompanyName    string            `json:"company,omitempty"`
	Department     string            `json:"department,omitempty"`
	JobTitle       string            `json:"job_title,omitempty"`
	GroupType      string            `json:"group_type,omitempty"`
}

func ParseSearch(data []byte) ([]Suggestion, transport.EnvelopeResult) {
	var root struct {
		Groups json.RawMessage `json:"Groups"`
	}
	if json.Unmarshal(data, &root) != nil || len(root.Groups) == 0 || string(root.Groups) == "null" {
		return nil, transport.EnvelopeError
	}
	var groups []struct {
		Type        string          `json:"Type"`
		Suggestions json.RawMessage `json:"Suggestions"`
	}
	if json.Unmarshal(root.Groups, &groups) != nil {
		return nil, transport.EnvelopeError
	}
	out := []Suggestion{}
	for _, g := range groups {
		if len(g.Suggestions) == 0 || string(g.Suggestions) == "null" {
			return nil, transport.EnvelopeError
		}
		var items []struct {
			ID, DisplayName, GivenName, Surname, PeopleType, PeopleSubtype, PersonID, PersonaID, ADObjectID, CompanyName, Department, JobTitle string
			EmailAddresses                                                                                                                     []string
			Phones                                                                                                                             []json.RawMessage
		}
		if json.Unmarshal(g.Suggestions, &items) != nil {
			return nil, transport.EnvelopeError
		}
		for _, x := range items {
			if x.ID == "" || x.DisplayName == "" {
				return nil, transport.EnvelopeError
			}
			emails := x.EmailAddresses
			if emails == nil {
				emails = []string{}
			}
			out = append(out, Suggestion{ID: x.ID, DisplayName: x.DisplayName, GivenName: x.GivenName, Surname: x.Surname, EmailAddresses: emails, Phones: x.Phones, PeopleType: x.PeopleType, PeopleSubtype: x.PeopleSubtype, PersonID: x.PersonID, PersonaID: x.PersonaID, ADObjectID: x.ADObjectID, CompanyName: x.CompanyName, Department: x.Department, JobTitle: x.JobTitle, GroupType: g.Type})
		}
	}
	return out, transport.EnvelopeOK
}
func validateSearch(b []byte) transport.EnvelopeResult { _, s := ParseSearch(b); return s }
