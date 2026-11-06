// Package search prepares the observed Outlook conversation-search request.
package search

import (
	"crypto/rand"
	"encoding/hex"
	"encoding/json"
	"errors"
	"regexp"
	"strings"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/webcapture"
)

var digits = regexp.MustCompile(`^[0-9]{1,6}$`)
var ErrInput = errors.New("invalid or unsupported Outlook search")

type Params struct {
	Query         string
	FolderID      string
	Offset, Limit int
}

func Prepare(detail, request string, current transport.Session, p Params) (transport.Session, transport.Operation, error) {
	fail := func() (transport.Session, transport.Operation, error) {
		return transport.Session{}, transport.Operation{}, ErrInput
	}
	if strings.TrimSpace(p.Query) == "" || len(p.Query) > 2048 || strings.ContainsAny(p.Query, "\x00\r\n") || len(p.FolderID) > 4096 || strings.ContainsAny(p.FolderID, "\x00\r\n") || p.Offset < 0 || p.Offset > 1_000_000 || p.Limit < 1 || p.Limit > 100 {
		return fail()
	}
	d, err := webcapture.Load(detail, request, current)
	if err != nil {
		return fail()
	}
	q := d.URL.Query()
	if d.Method != "POST" || d.URL.Path != "/searchservice/api/v2/query" || d.HeaderBody || len(q) != 2 || len(q["n"]) != 1 || !digits.MatchString(q.Get("n")) || len(q["cv"]) != 1 || q.Get("cv") == "" || len(q.Get("cv")) > 256 {
		return fail()
	}
	var root map[string]any
	if json.Unmarshal(d.Body, &root) != nil {
		return fail()
	}
	scenario, ok := root["Scenario"].(map[string]any)
	if !ok || scenario["Name"] != "owa.react" {
		return fail()
	}
	requests, ok := root["EntityRequests"].([]any)
	if !ok || len(requests) != 1 {
		return fail()
	}
	er, ok := requests[0].(map[string]any)
	if !ok || er["EntityType"] != "Conversation" {
		return fail()
	}
	sources, ok := er["ContentSources"].([]any)
	if !ok || len(sources) != 1 || sources[0] != "Exchange" {
		return fail()
	}
	query, ok := er["Query"].(map[string]any)
	if !ok {
		return fail()
	}
	query["QueryString"] = p.Query
	filter, ok := er["Filter"].(map[string]any)
	if !ok {
		return fail()
	}
	or, ok := filter["Or"].([]any)
	if !ok || len(or) != 2 {
		return fail()
	}
	first, ok := or[0].(map[string]any)
	if !ok {
		return fail()
	}
	term, ok := first["Term"].(map[string]any)
	if !ok || len(term) != 1 {
		return fail()
	}
	second, ok := or[1].(map[string]any)
	if !ok {
		return fail()
	}
	deleted, ok := second["Term"].(map[string]any)
	if !ok || len(deleted) != 1 || deleted["DistinguishedFolderName"] != "DeletedItems" {
		return fail()
	}
	if p.FolderID == "" {
		if term["DistinguishedFolderName"] != "msgfolderroot" {
			return fail()
		}
	} else {
		if _, ok := term["FolderId"].(string); !ok {
			return fail()
		}
		term["FolderId"] = p.FolderID
	}
	er["From"], er["Size"] = p.Offset, p.Limit
	// Avoid reusing the fixture's search correlation identifiers.
	for _, key := range []string{"Cvid", "LogicalId"} {
		id, e := uuid()
		if e != nil {
			return fail()
		}
		root[key] = id
	}
	body, err := json.Marshal(root)
	if err != nil {
		return fail()
	}
	// Do not replay the captured canary as a durable credential. The parent
	// fixture checks whether this read works using the current session alone.
	d.Session.Headers.Del("X-Owa-Canary")
	return d.Session, d.Operation("POST", d.URL.Path, body, transport.Read, Validate), nil
}

func uuid() (string, error) {
	var b [16]byte
	if _, err := rand.Read(b[:]); err != nil {
		return "", err
	}
	b[6] = (b[6] & 0x0f) | 0x40
	b[8] = (b[8] & 0x3f) | 0x80
	h := hex.EncodeToString(b[:])
	return h[:8] + "-" + h[8:12] + "-" + h[12:16] + "-" + h[16:20] + "-" + h[20:], nil
}

func Validate(data []byte) transport.EnvelopeResult {
	var root struct {
		EntitySets []struct {
			IsPartial  *bool `json:"IsPartial"`
			Properties struct {
				HasParseException *bool `json:"HasParseException"`
			} `json:"Properties"`
			ResultSets []struct {
				Total   *int            `json:"Total"`
				More    *bool           `json:"MoreResultsAvailable"`
				Results json.RawMessage `json:"Results"`
			} `json:"ResultSets"`
		} `json:"EntitySets"`
	}
	if json.Unmarshal(data, &root) != nil || len(root.EntitySets) == 0 {
		return transport.EnvelopeError
	}
	partial := false
	for _, set := range root.EntitySets {
		if set.IsPartial == nil || set.Properties.HasParseException == nil || *set.Properties.HasParseException || len(set.ResultSets) == 0 {
			return transport.EnvelopeError
		}
		if *set.IsPartial {
			partial = true
		}
		for _, r := range set.ResultSets {
			if r.Total == nil || *r.Total < 0 || r.More == nil {
				return transport.EnvelopeError
			}
			if len(r.Results) > 0 && string(r.Results) != "null" {
				var rows []json.RawMessage
				if json.Unmarshal(r.Results, &rows) != nil {
					return transport.EnvelopeError
				}
			}
		}
	}
	if partial {
		return transport.EnvelopePartial
	}
	return transport.EnvelopeOK
}
