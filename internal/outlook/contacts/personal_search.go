package contacts

import (
	"encoding/json"
	"fmt"
	"strings"
)

type searchableName struct {
	DisplayName string `json:"displayName"`
	First       string `json:"first"`
	Middle      string `json:"middle"`
	Last        string `json:"last"`
	Nickname    string `json:"nickname"`
}
type searchableContact struct {
	Name   searchableName   `json:"name"`
	Names  []searchableName `json:"names"`
	Emails []struct {
		Address string `json:"address"`
	} `json:"emails"`
	Phones []struct {
		Number string `json:"number"`
	} `json:"phones"`
	Positions []struct {
		Detail struct {
			JobTitle string `json:"jobTitle"`
			Company  struct {
				DisplayName string `json:"displayName"`
			} `json:"company"`
		} `json:"detail"`
	} `json:"positions"`
}

// FilterPersonalContacts performs a case-insensitive substring search only
// over the captured personal contact page. It does not query the directory.
func FilterPersonalContacts(page ListPage, query string) (ListPage, error) {
	q := strings.ToLower(strings.TrimSpace(query))
	if q == "" || len(q) > 512 || strings.ContainsRune(q, '\x00') {
		return ListPage{}, fmt.Errorf("search query must be nonempty and at most 512 bytes")
	}
	out := ListPage{Results: []json.RawMessage{}, NextCursor: page.NextCursor}
	for _, raw := range page.Results {
		var c searchableContact
		if json.Unmarshal(raw, &c) != nil {
			return ListPage{}, fmt.Errorf("invalid contact in search results")
		}
		values := []string{}
		addName := func(n searchableName) {
			values = append(values, n.DisplayName, n.First, n.Middle, n.Last, n.Nickname, strings.TrimSpace(n.First+" "+n.Last))
		}
		addName(c.Name)
		for _, n := range c.Names {
			addName(n)
		}
		for _, v := range c.Emails {
			values = append(values, v.Address)
		}
		for _, v := range c.Phones {
			values = append(values, v.Number)
		}
		for _, v := range c.Positions {
			values = append(values, v.Detail.JobTitle, v.Detail.Company.DisplayName)
		}
		for _, v := range values {
			if strings.Contains(strings.ToLower(v), q) {
				out.Results = append(out.Results, raw)
				break
			}
		}
	}
	return out, nil
}

// PersonalSearchWindow carries list continuation and snapshot behavior over
// to a filtered page while preserving uncertainty about an upstream cap.
func PersonalSearchWindow(source ListPage, query string, offset, limit int, cursor, snapshot string, serverLimit int) (ListWindow, error) {
	if serverLimit < 0 || serverLimit > 30000 {
		return ListWindow{}, fmt.Errorf("invalid server limit")
	}
	filtered, err := FilterPersonalContacts(source, query)
	if err != nil {
		return ListWindow{}, err
	}
	scope := "personal-search:" + strings.ToLower(strings.TrimSpace(query)) + "\x00" + cursor
	// The list window hashes the scope together with the matched IDs/versions, so
	// a token from another query or another server page cannot be reused.
	w, err := WindowList(filtered, offset, limit, scope, snapshot, 30000)
	if err != nil {
		return ListWindow{}, err
	}
	if w.NextSnapshot != "" {
		w.NextCursor = cursor
	}
	sourceCap := source.NextCursor == "" && (serverLimit == 0 || len(source.Results) >= serverLimit)
	w.ServerLimit = serverLimit
	w.ServerCapReached = source.NextCursor == "" && serverLimit > 0 && len(source.Results) >= serverLimit
	if sourceCap && w.NextSnapshot == "" {
		w.HasMore = nil
		w.Complete = false
		w.SyncCursor = ""
	}
	if !sourceCap && source.NextCursor == "" && w.HasMore != nil && !*w.HasMore && offset == 0 && cursor == "" {
		w.Complete = true
	}
	return w, nil
}
