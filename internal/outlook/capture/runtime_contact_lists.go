package capture

import (
	"encoding/json"
	"strings"
)

// Contact-list bundles accept only canonicalized synthetic paths and the
// observed query shape. The operation builder validates bodies again on use.
func validContactListTemplate(name string, t RuntimeWebTemplate) (bool, bool) {
	const path = "/api/v2.0/me/personaldistributionlists"
	const selectAll = "id,createdDateTime,lastModifiedDateTime,personId,displayName,notes,members"
	const selectExpanded = "id,createdDateTime,lastModifiedDateTime,personId,displayName,notes"
	var method, wantPath, selectValue, expand, top, order string
	switch name {
	case "contact_lists_list":
		method, wantPath, selectValue, top, order = "GET", path, selectAll, "500", "DisplayName Asc"
	case "contact_lists_create":
		method, wantPath, selectValue = "POST", path, selectAll
	case "contact_lists_update":
		method, wantPath, selectValue = "PATCH", path+"/template-id", selectAll
	case "contact_lists_add_member":
		method, wantPath, selectValue, expand = "POST", path+"/template-id/AddMembers", selectExpanded, "members"
	case "contact_lists_remove_member":
		method, wantPath, selectValue, expand = "POST", path+"/template-id/DeleteMembers", selectExpanded, "members"
	case "contact_lists_delete":
		method, wantPath = "DELETE", path+"/template-id"
	default:
		return false, false
	}
	if t.Method != method || t.Path != wantPath || t.HeaderCarrier || !digits.MatchString(t.Query.Get("n")) || t.Query.Get("cri") == "" || t.Query.Get("cv") == "" {
		return true, false
	}
	for k, values := range t.Query {
		if len(values) != 1 || len(values[0]) > 8192 || strings.ContainsAny(values[0], "\x00\r\n") {
			return true, false
		}
		switch k {
		case "$select", "$expand", "$top", "$orderBy", "n", "cri", "cv":
		default:
			return true, false
		}
	}
	want := 3
	for _, s := range []string{selectValue, expand, top, order} {
		if s != "" {
			want++
		}
	}
	if len(t.Query) != want || t.Query.Get("$select") != selectValue || t.Query.Get("$expand") != expand || t.Query.Get("$top") != top || t.Query.Get("$orderBy") != order {
		return true, false
	}
	if method == "GET" || method == "DELETE" {
		return true, len(t.Body) == 0 || string(t.Body) == "null"
	}
	return true, json.Valid(t.Body)
}
