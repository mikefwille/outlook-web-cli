package capture

import (
	"net/url"
	"testing"
)

func TestContactListRuntimeTemplatesRejectForeignPathsAndQueries(t *testing.T) {
	const path = "/api/v2.0/me/personaldistributionlists"
	const fields = "id,createdDateTime,lastModifiedDateTime,personId,displayName,notes,members"
	base := RuntimeWebTemplate{Method: "GET", Path: path, Query: url.Values{"n": {"1"}, "cri": {"correlation"}, "cv": {"cv"}, "$select": {fields}, "$top": {"500"}, "$orderBy": {"DisplayName Asc"}}}
	if handled, ok := validContactListTemplate("contact_lists_list", base); !handled || !ok {
		t.Fatal("observed list template rejected")
	}
	foreign := base
	foreign.Path = "/other"
	if _, ok := validContactListTemplate("contact_lists_list", foreign); ok {
		t.Fatal("foreign route accepted")
	}
	foreign = base
	foreign.Query = url.Values{"n": {"1"}, "cri": {"x"}, "cv": {"x"}, "$select": {fields}, "$top": {"500"}, "$orderBy": {"DisplayName Asc"}, "extra": {"x"}}
	if _, ok := validContactListTemplate("contact_lists_list", foreign); ok {
		t.Fatal("extra query accepted")
	}
	foreign = base
	foreign.Method = "DELETE"
	if _, ok := validContactListTemplate("contact_lists_list", foreign); ok {
		t.Fatal("wrong method accepted")
	}
}
