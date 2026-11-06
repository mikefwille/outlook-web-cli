// Package directory replays the profile card's observed organization query.
// Its credential is a separate, account-bound Loki audience; Outlook session
// cookies and Outlook bearer tokens are never sent to that origin.
package directory

import (
	"bufio"
	"bytes"
	"context"
	"crypto/rand"
	"crypto/sha256"
	"encoding/base64"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"io"
	"mime/multipart"
	"net/http"
	"net/url"
	"path/filepath"
	"regexp"
	"strings"
	"time"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/capture"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/contacts"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

const origin = "https://nam.loki.delve.office.com"
const queryName = "useOrgExplorerQueryRef_OrganizationQuery"
const queryHash = "d700f004c154489effde5101064fee83a84ed5da3958ab72cd557722cb318bce"

var uuid = regexp.MustCompile(`(?i)^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$`)
var requestLine = regexp.MustCompile(`(?m)^## Request (https://\S+)\r?\nStatus: 200\r?$`)
var safeBoundary = regexp.MustCompile(`^[A-Za-z0-9'()+_,./:=?-]{1,70}$`)
var outerMethod = regexp.MustCompile(`(?m)^- :method:\s*POST\s*$`)
var outerType = regexp.MustCompile(`(?mi)^- content-type:\s*text/plain;charset=UTF-8\s*$`)
var outerOrigin = regexp.MustCompile(`(?mi)^- origin:\s*https://outlook\.cloud\.microsoft\s*$`)
var headerNames = map[string]bool{"accept": true, "content-type": true, "scenariotag": true, "x-client-language": true, "x-clientarchitectureversion": true, "x-clientcorrelationid": true, "x-clientfeature": true, "x-clientscenario": true, "x-clienttype": true, "x-correlationid": true, "x-hostappapp": true, "x-hostappcapabilities": true, "x-lpcversion": true, "x-rootcorrelationid": true}

type Identity struct {
	ObjectID string
	Email    string
	UPN      string
}

type Session struct {
	Version             int               `json:"version"`
	Origin              string            `json:"origin"`
	AccountKey          string            `json:"account_key"`
	PrimaryAudienceHash string            `json:"primary_audience_hash"`
	AnchorHash          string            `json:"anchor_hash"`
	DirectoryAudience   string            `json:"directory_audience_hash"`
	Authorization       string            `json:"authorization"`
	Query               string            `json:"query"`
	Headers             map[string]string `json:"headers"`
}

type Result struct {
	Manager     []json.RawMessage `json:"manager"`
	Reports     []json.RawMessage `json:"reports"`
	HasNextPage bool              `json:"has_next_page"`
	EndCursor   string            `json:"end_cursor,omitempty"`
}

func hash(s string) string { h := sha256.Sum256([]byte(s)); return hex.EncodeToString(h[:]) }

func tokenBinding(auth string) (string, string, time.Time, bool) {
	p := strings.SplitN(auth, " ", 2)
	if len(p) != 2 || !strings.EqualFold(p[0], "Bearer") || strings.ContainsAny(auth, "\r\n\x00") {
		return "", "", time.Time{}, false
	}
	pieces := strings.Split(p[1], ".")
	if len(pieces) != 3 {
		return "", "", time.Time{}, false
	}
	raw, err := base64.RawURLEncoding.DecodeString(pieces[1])
	if err != nil {
		return "", "", time.Time{}, false
	}
	var c struct {
		Tenant string `json:"tid"`
		Object string `json:"oid"`
		Aud    string `json:"aud"`
		Exp    int64  `json:"exp"`
	}
	if json.Unmarshal(raw, &c) != nil || c.Tenant == "" || c.Object == "" || c.Aud == "" || c.Exp <= 0 {
		return "", "", time.Time{}, false
	}
	return hash(c.Tenant + "\x00" + c.Object), hash(c.Aud), time.Unix(c.Exp, 0), true
}

func ValidIdentity(i Identity) bool {
	return uuid.MatchString(i.ObjectID) && contacts.ValidEmail(i.Email) && contacts.ValidEmail(i.UPN)
}

func validSession(s Session, current transport.Session) bool {
	b, err := capture.NewRuntimeBundle(current)
	if err != nil || current.Headers.Get("Cookie") == "" || s.Version != 1 || s.Origin != origin || s.AccountKey != b.AccountKey || s.PrimaryAudienceHash != b.AudienceHash || s.AnchorHash != b.AnchorHash || hash(s.Query) != queryHash {
		return false
	}
	account, aud, expires, ok := tokenBinding(s.Authorization)
	if !ok || account != s.AccountKey || aud != s.DirectoryAudience || aud == s.PrimaryAudienceHash || !time.Now().Before(expires) {
		return false
	}
	if s.Headers["accept"] != "multipart/mixed;deferSpec=20220824,application/json" || s.Headers["content-type"] != "application/json" {
		return false
	}
	for key, value := range s.Headers {
		if !headerNames[key] || len(value) == 0 || len(value) > 4096 || strings.ContainsAny(value, "\r\n\x00") {
			return false
		}
	}
	return true
}

// ExtractObserved checks a successful, exact profile-card request/response
// pair. The returned session is secret and may only be saved in private state.
func ExtractObserved(current transport.Session, detail, request, response []byte) (Session, error) {
	fail := func() (Session, error) {
		return Session{}, fmt.Errorf("directory capture is invalid or incompatible with the active account")
	}
	match := requestLine.FindSubmatch(detail)
	if len(match) != 2 {
		return fail()
	}
	u, err := url.Parse(string(match[1]))
	if err != nil || u.Scheme != "https" || u.Host != "nam.loki.delve.office.com" || u.User != nil || u.Path != "/api/v2/graphql" || u.Fragment != "" || len(u.Query()) != 2 || len(u.Query()["operationName"]) != 1 || len(u.Query()["ConvertSimpleRequest"]) != 1 || u.Query().Get("operationName") != queryName || u.Query().Get("ConvertSimpleRequest") != "POST" {
		return fail()
	}
	blocks := strings.SplitN(string(detail), "### Request Headers\n", 2)
	if len(blocks) != 2 {
		return fail()
	}
	requestHeaders := strings.SplitN(blocks[1], "\n### ", 2)[0]
	if !outerMethod.MatchString(requestHeaders) || !outerType.MatchString(requestHeaders) || !outerOrigin.MatchString(requestHeaders) {
		return fail()
	}
	var r struct {
		Operation string `json:"operationName"`
		Query     string `json:"query"`
		Variables struct {
			PersonIDs struct {
				ObjectID string `json:"aadObjectId"`
				Email    string `json:"smtp"`
				UPN      string `json:"userPrincipalName"`
			} `json:"personIds"`
		} `json:"variables"`
		Headers map[string]string `json:"__request_headers"`
	}
	if json.Unmarshal(request, &r) != nil || r.Operation != queryName || hash(r.Query) != queryHash {
		return fail()
	}
	id := Identity{ObjectID: r.Variables.PersonIDs.ObjectID, Email: r.Variables.PersonIDs.Email, UPN: r.Variables.PersonIDs.UPN}
	if !ValidIdentity(id) {
		return fail()
	}
	var observed struct {
		Variables map[string]any `json:"variables"`
	}
	if json.Unmarshal(request, &observed) != nil {
		return fail()
	}
	actualVars, e1 := json.Marshal(observed.Variables)
	expectedVars, e2 := json.Marshal(orgVariables(id))
	if e1 != nil || e2 != nil || !bytes.Equal(actualVars, expectedVars) {
		return fail()
	}
	if _, err = ParseOrg(response, id); err != nil {
		return fail()
	}
	b, err := capture.NewRuntimeBundle(current)
	if err != nil {
		return fail()
	}
	s := Session{Version: 1, Origin: origin, AccountKey: b.AccountKey, PrimaryAudienceHash: b.AudienceHash, AnchorHash: b.AnchorHash, Query: r.Query, Headers: map[string]string{}}
	for k, v := range r.Headers {
		key := strings.ToLower(k)
		if key == "authorization" {
			s.Authorization = v
		} else if headerNames[key] {
			s.Headers[key] = v
		}
	}
	_, s.DirectoryAudience, _, _ = tokenBinding(s.Authorization)
	if !validSession(s, current) {
		return fail()
	}
	return s, nil
}

// SessionPath derives a fixed filename from the active account binding, never
// from a caller-supplied private path or person identifier.
func SessionPath(pointer string, current transport.Session) (string, error) {
	p, err := capture.RuntimeBundlePathForAccount(pointer, current)
	if err != nil {
		return "", err
	}
	name := strings.TrimPrefix(filepath.Base(p), "outlook-operations-")
	return filepath.Join(filepath.Dir(p), "outlook-directory-"+name), nil
}

func Load(pointer string, current transport.Session) (Session, error) {
	p, err := SessionPath(pointer, current)
	if err != nil {
		return Session{}, &transport.Error{Kind: transport.KindAuth}
	}
	data, err := capture.ReadPrivateEvidence(p)
	var s Session
	if err != nil || json.Unmarshal(data, &s) != nil || !validSession(s, current) {
		return Session{}, &transport.Error{Kind: transport.KindAuth}
	}
	return s, nil
}

func orgVariables(id Identity) map[string]any {
	return map[string]any{"personIds": map[string]string{"aadObjectId": id.ObjectID, "smtp": id.Email, "userPrincipalName": id.UPN}, "jobTitleFilter": []any{}, "vivaSkillsFilter": []any{}, "hybridLocationFilter": "All", "directsToLoad": 20, "worksWithToLoad": 25, "peersToLoad": 20, "withHighlights": false, "withFiltering": false, "withVivaSkillsFiltering": false, "withPeers": false, "withNextLevelReports": false, "isMto": false}
}

func correlationID() (string, error) {
	var b [16]byte
	if _, err := rand.Read(b[:]); err != nil {
		return "", err
	}
	b[6] = (b[6] & 0x0f) | 0x40
	b[8] = (b[8] & 0x3f) | 0x80
	v := hex.EncodeToString(b[:])
	return v[:8] + "-" + v[8:12] + "-" + v[12:16] + "-" + v[16:20] + "-" + v[20:], nil
}

func Prepare(s Session, current transport.Session, id Identity) (transport.Session, transport.Operation, error) {
	if !ValidIdentity(id) {
		return transport.Session{}, transport.Operation{}, fmt.Errorf("supply a directory object UUID, exact email and user principal name")
	}
	if !validSession(s, current) {
		return transport.Session{}, transport.Operation{}, &transport.Error{Kind: transport.KindAuth}
	}
	// The captured ConvertSimpleRequest puts the Loki authorization and
	// backend headers in the JSON body, not in HTTP headers. Reproduce that
	// exact carrier only on the fixed HTTPS origin, with fresh correlation IDs.
	embedded := make(map[string]string, len(s.Headers)+1)
	for key, value := range s.Headers {
		embedded[key] = value
	}
	embedded["authorization"] = s.Authorization
	for _, key := range []string{"x-clientcorrelationid", "x-correlationid", "x-rootcorrelationid"} {
		if _, ok := embedded[key]; ok {
			value, err := correlationID()
			if err != nil {
				return transport.Session{}, transport.Operation{}, &transport.Error{Kind: transport.KindInput}
			}
			embedded[key] = value
		}
	}
	body, err := json.Marshal(map[string]any{"operationName": queryName, "variables": orgVariables(id), "query": s.Query, "__request_headers": embedded})
	if err != nil {
		return transport.Session{}, transport.Operation{}, &transport.Error{Kind: transport.KindInput}
	}
	_, _, exp, _ := tokenBinding(s.Authorization)
	session := transport.Session{AccountKey: current.AccountKey, Origin: origin, ExpiresAt: exp, Headers: http.Header{"Accept": {"*/*"}, "Content-Type": {"text/plain;charset=UTF-8"}, "Origin": {"https://outlook.cloud.microsoft"}}}
	op := transport.Operation{Method: "POST", Path: "/api/v2/graphql", Query: url.Values{"operationName": {queryName}, "ConvertSimpleRequest": {"POST"}}, Body: body, Class: transport.Read, Validate: func(data []byte) transport.EnvelopeResult {
		if _, err := ParseOrg(data, id); err != nil {
			return transport.EnvelopeError
		}
		return transport.EnvelopeOK
	}}
	return session, op, nil
}

func Fetch(ctx context.Context, current transport.Session, s Session, id Identity, rt http.RoundTripper) (Result, error) {
	session, op, err := Prepare(s, current, id)
	if err != nil {
		return Result{}, err
	}
	client, err := transport.New(session.AccountKey, session, rt)
	if err != nil {
		return Result{}, err
	}
	response, err := client.Do(ctx, session.AccountKey, op)
	if err != nil {
		return Result{}, err
	}
	result, err := ParseOrg(response.Body, id)
	if err != nil {
		return Result{}, &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
	}
	return result, nil
}

func errorsPresent(raw json.RawMessage) bool {
	if len(raw) == 0 || string(bytes.TrimSpace(raw)) == "null" {
		return false
	}
	var errors []json.RawMessage
	return json.Unmarshal(raw, &errors) != nil || len(errors) != 0
}

func jsonParts(data []byte) ([][]byte, bool) {
	trimmed := bytes.TrimSpace(data)
	if len(trimmed) == 0 || len(trimmed) > 8<<20 {
		return nil, false
	}
	if trimmed[0] == '{' {
		return [][]byte{trimmed}, false
	}
	reader := bufio.NewReader(bytes.NewReader(trimmed))
	line, err := reader.ReadString('\n')
	if err != nil {
		return nil, false
	}
	boundary := strings.TrimSpace(line)
	if !strings.HasPrefix(boundary, "--") || !safeBoundary.MatchString(strings.TrimPrefix(boundary, "--")) {
		return nil, false
	}
	parts := [][]byte{}
	mr := multipart.NewReader(bytes.NewReader(trimmed), strings.TrimPrefix(boundary, "--"))
	for {
		part, err := mr.NextPart()
		if err == io.EOF {
			break
		}
		if err != nil || len(parts) >= 32 {
			return nil, false
		}
		b, err := io.ReadAll(io.LimitReader(part, 2<<20+1))
		if err != nil || len(b) > 2<<20 {
			return nil, false
		}
		parts = append(parts, b)
	}
	return parts, true
}

// ParseOrg rejects incomplete multipart streams, GraphQL errors, an identity
// mismatch and missing relationship fields. A complete empty array is data;
// an omitted field or failed operation is not treated as an empty result.
func ParseOrg(data []byte, id Identity) (Result, error) {
	fail := func() (Result, error) {
		return Result{}, fmt.Errorf("directory organization response is incomplete or invalid")
	}
	parts, multipartResponse := jsonParts(data)
	if len(parts) == 0 {
		return fail()
	}
	var focus map[string]json.RawMessage
	finished := !multipartResponse
	for index, part := range parts {
		var doc struct {
			Data struct {
				Org struct {
					Focus map[string]json.RawMessage `json:"personInFocusV2"`
				} `json:"orgexplorer"`
			} `json:"data"`
			Errors      json.RawMessage `json:"errors"`
			HasNext     *bool           `json:"hasNext"`
			Incremental []struct {
				Path   []any                      `json:"path"`
				Data   map[string]json.RawMessage `json:"data"`
				Errors json.RawMessage            `json:"errors"`
			} `json:"incremental"`
		}
		if json.Unmarshal(part, &doc) != nil || errorsPresent(doc.Errors) || finished && multipartResponse {
			return fail()
		}
		if index == 0 {
			focus = doc.Data.Org.Focus
			if focus == nil {
				return fail()
			}
		} else if doc.Data.Org.Focus != nil {
			return fail()
		}
		for _, inc := range doc.Incremental {
			if errorsPresent(inc.Errors) || len(inc.Path) != 2 || inc.Path[0] != "orgexplorer" || inc.Path[1] != "personInFocusV2" || inc.Data == nil {
				return fail()
			}
			for key, value := range inc.Data {
				if old, exists := focus[key]; exists && !bytes.Equal(old, value) && (key == "aadObjectId" || key == "email" || key == "manager" || key == "directs") {
					return fail()
				}
				if _, exists := focus[key]; !exists {
					focus[key] = value
				}
			}
		}
		if multipartResponse {
			if doc.HasNext == nil {
				return fail()
			}
			finished = !*doc.HasNext
		} else if doc.HasNext != nil && *doc.HasNext {
			return fail()
		}
	}
	if !finished {
		return fail()
	}
	var oid, email string
	if json.Unmarshal(focus["aadObjectId"], &oid) != nil || !strings.EqualFold(oid, id.ObjectID) || json.Unmarshal(focus["email"], &email) != nil || !strings.EqualFold(email, id.Email) {
		return fail()
	}
	var manager struct {
		Nodes *[]json.RawMessage `json:"nodes"`
	}
	var directs struct {
		Edges *[]struct {
			Node json.RawMessage `json:"node"`
		} `json:"edges"`
		PageInfo *struct {
			HasNext   *bool   `json:"hasNextPage"`
			EndCursor *string `json:"endCursor"`
		} `json:"pageInfo"`
	}
	if json.Unmarshal(focus["manager"], &manager) != nil || manager.Nodes == nil || len(*manager.Nodes) > 1 || json.Unmarshal(focus["directs"], &directs) != nil || directs.Edges == nil || directs.PageInfo == nil || directs.PageInfo.HasNext == nil || len(*directs.Edges) > 20 {
		return fail()
	}
	out := Result{Manager: *manager.Nodes, Reports: make([]json.RawMessage, 0, len(*directs.Edges)), HasNextPage: *directs.PageInfo.HasNext}
	for _, node := range out.Manager {
		if !validPerson(node) {
			return fail()
		}
	}
	for _, edge := range *directs.Edges {
		if !validPerson(edge.Node) {
			return fail()
		}
		out.Reports = append(out.Reports, edge.Node)
	}
	if directs.PageInfo.EndCursor != nil {
		out.EndCursor = *directs.PageInfo.EndCursor
	}
	return out, nil
}

func validPerson(raw json.RawMessage) bool {
	var p struct {
		ID       string `json:"id"`
		ObjectID string `json:"aadObjectId"`
	}
	return json.Unmarshal(raw, &p) == nil && p.ID != "" && p.ObjectID != ""
}
