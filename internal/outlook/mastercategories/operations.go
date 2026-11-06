// Package mastercategories binds only the observed master-category delta and
// category-wide rename/delete operations. Calls are not atomic.
package mastercategories

import (
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"io"
	"regexp"
	"strings"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/webcapture"
)

type Category struct {
	Name             string `json:"Name"`
	Color            int    `json:"Color"`
	ID               string `json:"Id"`
	LastTimeUsed     string `json:"LastTimeUsed"`
	KeyboardShortcut int    `json:"KeyboardShortcut"`
	KnownFieldsOnly  bool   `json:"-"`
	PreservationHash string `json:"-"`
}

type Snapshot struct {
	Categories []Category
	Version    string
}

func ParseConfiguration(data []byte) (Snapshot, error) {
	var x struct {
		Master *struct {
			Items *[]json.RawMessage `json:"MasterList"`
		} `json:"MasterCategoryList"`
	}
	if json.Unmarshal(data, &x) != nil || x.Master == nil || x.Master.Items == nil {
		return Snapshot{}, fmt.Errorf("master category list missing")
	}
	return parseList(*x.Master.Items)
}
func ParseResponse(data []byte) (Snapshot, error) {
	var x struct {
		Items         *[]json.RawMessage `json:"MasterList"`
		Success       *bool              `json:"WasSuccessful"`
		ErrorCode     *int               `json:"ErrorCode"`
		ResponseCode  string             `json:"ResponseCode"`
		ResponseClass string             `json:"ResponseClass"`
	}
	if json.Unmarshal(data, &x) != nil || x.Items == nil || x.Success == nil || !*x.Success || x.ErrorCode == nil || *x.ErrorCode != 0 || x.ResponseCode != "NoError" || x.ResponseClass != "Success" {
		return Snapshot{}, fmt.Errorf("master category response was not confirmed successful")
	}
	return parseList(*x.Items)
}
func parseList(items []json.RawMessage) (Snapshot, error) {
	if len(items) > 5000 {
		return Snapshot{}, fmt.Errorf("master category list exceeds limit")
	}
	out := Snapshot{Categories: make([]Category, 0, len(items))}
	canonical := make([]json.RawMessage, 0, len(items))
	ids := map[string]bool{}
	names := map[string]bool{}
	for _, raw := range items {
		var c Category
		var fields map[string]json.RawMessage
		if json.Unmarshal(raw, &fields) != nil || fields == nil || json.Unmarshal(raw, &c) != nil || c.ID == "" || !validName(c.Name) || ids[c.ID] || names[strings.ToLower(c.Name)] {
			return Snapshot{}, fmt.Errorf("master category list contains incomplete or ambiguous entries")
		}
		for _, key := range []string{"Id", "Name", "Color", "LastTimeUsed", "KeyboardShortcut"} {
			if len(fields[key]) == 0 || string(fields[key]) == "null" {
				return Snapshot{}, fmt.Errorf("master category list contains incomplete entries")
			}
		}
		c.KnownFieldsOnly = len(fields) == 5
		preserved := make(map[string]json.RawMessage, len(fields))
		for key, value := range fields {
			if key != "LastTimeUsed" {
				preserved[key] = value
			}
		}
		stable, err := json.Marshal(preserved)
		if err != nil {
			return Snapshot{}, err
		}
		stableHash := sha256.Sum256(stable)
		c.PreservationHash = hex.EncodeToString(stableHash[:])
		ids[c.ID], names[strings.ToLower(c.Name)] = true, true
		// Canonicalize all fields, including unknown ones, for the precheck.
		clean, err := json.Marshal(fields)
		if err != nil {
			return Snapshot{}, err
		}
		canonical = append(canonical, clean)
		out.Categories = append(out.Categories, c)
	}
	encoded, err := json.Marshal(canonical)
	if err != nil {
		return Snapshot{}, err
	}
	hash := sha256.Sum256(encoded)
	out.Version = hex.EncodeToString(hash[:])
	return out, nil
}
func validName(s string) bool {
	return len(s) > 0 && len(s) <= 255 && strings.TrimSpace(s) == s && !strings.ContainsAny(s, "\x00\r\n")
}

var guid = regexp.MustCompile(`^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$`)
var digits = regexp.MustCompile(`^[0-9]{1,12}$`)

type Params struct {
	OldName string
	New     Category
}
type masterBody struct {
	Request struct {
		Type      string            `json:"__type"`
		Add       []Category        `json:"AddCategoryList"`
		Remove    []string          `json:"RemoveCategoryList"`
		Colors    []json.RawMessage `json:"ChangeCategoryColorList"`
		Times     []json.RawMessage `json:"UpdateCategoryLastTimeUsedList"`
		Shortcuts []json.RawMessage `json:"ChangeCategoryKeyboardShortcutList"`
	} `json:"request"`
}
type applyBody struct {
	Request struct {
		Type      string `json:"__type"`
		Operation struct {
			Action *int   `json:"Action"`
			Name   string `json:"CategoryName"`
			Rename string `json:"RenameCategoryName,omitempty"`
		} `json:"categoryOperation"`
	} `json:"request"`
}

func strict(b []byte, v any) bool {
	d := json.NewDecoder(bytes.NewReader(b))
	d.DisallowUnknownFields()
	if d.Decode(v) != nil {
		return false
	}
	var extra any
	return d.Decode(&extra) == io.EOF
}
func validNew(c Category) bool {
	return validName(c.Name) && guid.MatchString(c.ID) && c.Color >= 0 && c.Color <= 24 && c.LastTimeUsed != "" && len(c.LastTimeUsed) <= 64 && !strings.ContainsAny(c.LastTimeUsed, "\x00\r\n") && c.KeyboardShortcut >= 0 && c.KeyboardShortcut <= 99
}

// Prepare validates the original captured request shape before replacing all
// category-specific values. The runtime template contains only placeholders.
func Prepare(kind, detail string, current transport.Session, p Params) (transport.Session, transport.Operation, error) {
	fail := func() (transport.Session, transport.Operation, error) {
		return transport.Session{}, transport.Operation{}, fmt.Errorf("category operation does not match the observed request contract")
	}
	action := "UpdateMasterCategoryList"
	if kind == "rename-items" || kind == "delete-items" {
		action = "ApplyCategoryOperation"
	}
	if kind != "create" && kind != "update" && kind != "delete" && kind != "rename-items" && kind != "delete-items" {
		return fail()
	}
	d, err := webcapture.Load(detail, "", current)
	if err != nil {
		return fail()
	}
	q := d.URL.Query()
	if d.Method != "POST" || d.URL.Path != "/owa/service.svc" || len(q) != 3 || q.Get("action") != action || q.Get("app") != "Mail" || !digits.MatchString(q.Get("n")) || d.Session.Headers.Get("Action") != action || !d.HeaderBody {
		return fail()
	}
	for _, v := range q {
		if len(v) != 1 {
			return fail()
		}
	}
	var body []byte
	validator := validateMaster
	if action == "UpdateMasterCategoryList" {
		var x masterBody
		if !strict(d.Body, &x) || x.Request.Type != "UpdateMasterCategoryListRequest:#Exchange" || x.Request.Add == nil || x.Request.Remove == nil || x.Request.Colors == nil || x.Request.Times == nil || x.Request.Shortcuts == nil || len(x.Request.Colors) != 0 || len(x.Request.Times) != 0 || len(x.Request.Shortcuts) != 0 {
			return fail()
		}
		wantAdd, wantRemove := 0, 0
		if kind == "create" || kind == "update" {
			wantAdd = 1
		}
		if kind == "update" || kind == "delete" {
			wantRemove = 1
		}
		if len(x.Request.Add) != wantAdd || len(x.Request.Remove) != wantRemove {
			return fail()
		}
		if wantAdd == 1 {
			if !validNew(x.Request.Add[0]) || !validNew(p.New) {
				return fail()
			}
			x.Request.Add[0] = p.New
		}
		if wantRemove == 1 {
			if !validName(x.Request.Remove[0]) || !validName(p.OldName) {
				return fail()
			}
			x.Request.Remove[0] = p.OldName
		}
		body, err = json.Marshal(x)
	} else {
		var x applyBody
		if !strict(d.Body, &x) || x.Request.Type != "ApplyCategoryOperationRequest:#Exchange" || x.Request.Operation.Action == nil || !validName(x.Request.Operation.Name) || !validName(p.OldName) {
			return fail()
		}
		if kind == "rename-items" {
			if *x.Request.Operation.Action != 0 || !validName(x.Request.Operation.Rename) || !validName(p.New.Name) {
				return fail()
			}
			x.Request.Operation.Rename = p.New.Name
		} else if *x.Request.Operation.Action != 1 || x.Request.Operation.Rename != "" {
			return fail()
		}
		x.Request.Operation.Name = p.OldName
		body, err = json.Marshal(x)
		validator = validateApply
	}
	if err != nil {
		return fail()
	}
	return d.Session, d.Operation("POST", d.URL.Path, body, transport.Write, validator), nil
}
func validateMaster(b []byte) transport.EnvelopeResult {
	if _, err := ParseResponse(b); err != nil {
		return transport.EnvelopeError
	}
	return transport.EnvelopeOK
}
func validateApply(b []byte) transport.EnvelopeResult {
	var x struct {
		Ready *bool `json:"IsSearchFolderReady"`
	}
	if json.Unmarshal(b, &x) != nil || x.Ready == nil || !*x.Ready {
		return transport.EnvelopeError
	}
	return transport.EnvelopeOK
}

// VerifyDelta checks the successful master-list response before a dependent
// category-wide rename/delete. LastTimeUsed may change independently on use.
func VerifyDelta(before, after Snapshot, kind string, p Params) bool {
	current := map[string]Category{}
	for _, c := range after.Categories {
		current[c.Name] = c
	}
	for _, old := range before.Categories {
		if (kind == "update" || kind == "delete") && old.Name == p.OldName {
			continue
		}
		now, ok := current[old.Name]
		if !ok || now.ID != old.ID || now.Color != old.Color || now.KeyboardShortcut != old.KeyboardShortcut || now.PreservationHash != old.PreservationHash {
			return false
		}
		delete(current, old.Name)
	}
	if kind == "create" || kind == "update" {
		newCat, ok := current[p.New.Name]
		if !ok || newCat.Color != p.New.Color || newCat.KeyboardShortcut != p.New.KeyboardShortcut {
			return false
		}
		delete(current, p.New.Name)
	}
	return len(current) == 0
}
