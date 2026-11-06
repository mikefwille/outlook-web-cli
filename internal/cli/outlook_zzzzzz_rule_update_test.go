package cli

import (
	"context"
	"encoding/json"
	"errors"
	"testing"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/rules"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func TestRuleUpdateReadbackAndOptionalRestore(t *testing.T) {
	base := rules.Rule{ID: "id", IdentityDisplay: "display", Name: "name", Subject: []string{"subject"}, Category: []string{"category"}, Version: "v1"}
	before := rules.Rule{Raw: json.RawMessage(`{"Priority":1,"Unchanged":"keep"}`)}
	desired := rules.Params{Name: "name", Subject: "subject", Category: "category"}
	tests := []struct {
		name                    string
		reads                   []rules.Rule
		readErrAt               int
		writeErr                bool
		restore                 bool
		wantStatus, wantRestore string
		wantWrites              int
	}{
		{"already disabled", []rules.Rule{base}, -1, false, true, "verified", "not_needed", 0},
		{"enabled no restore", []rules.Rule{{ID: "id", Name: "name", Subject: []string{"subject"}, Category: []string{"category"}, Enabled: true, Version: "v1"}}, -1, false, false, "verified", "not_requested", 0},
		{"restored", []rules.Rule{{ID: "id", IdentityDisplay: "display", Name: "name", Subject: []string{"subject"}, Category: []string{"category"}, Enabled: true, Version: "v1"}, {ID: "id", IdentityDisplay: "display", Name: "name", Subject: []string{"subject"}, Category: []string{"category"}, Enabled: true, Version: "v1"}, base}, -1, false, true, "verified", "verified", 1},
		{"state changed", []rules.Rule{{ID: "id", Name: "name", Subject: []string{"subject"}, Category: []string{"category"}, Enabled: true, Version: "v1"}, {ID: "id", Name: "name", Subject: []string{"subject"}, Category: []string{"category"}, Enabled: true, Version: "v2"}}, -1, false, true, "verified", "not_attempted_state_changed", 0},
		{"restore uncertain", []rules.Rule{{ID: "id", IdentityDisplay: "display", Name: "name", Subject: []string{"subject"}, Category: []string{"category"}, Enabled: true, Version: "v1"}, {ID: "id", IdentityDisplay: "display", Name: "name", Subject: []string{"subject"}, Category: []string{"category"}, Enabled: true, Version: "v1"}}, -1, true, true, "verified", "unconfirmed", 1},
		{"read unavailable", nil, 0, false, true, "unavailable", "not_attempted", 0},
		{"unexpected content", []rules.Rule{{ID: "id", Name: "wrong", Subject: []string{"subject"}, Category: []string{"category"}, Enabled: true}}, -1, false, true, "mismatch", "not_attempted", 0},
		{"other field changed", []rules.Rule{{ID: "id", Name: "name", Subject: []string{"subject"}, Category: []string{"category"}, Raw: json.RawMessage(`{"Priority":1,"Unchanged":"changed"}`)}}, -1, false, true, "mismatch", "not_attempted", 0},
	}
	for _, tc := range tests {
		t.Run(tc.name, func(t *testing.T) {
			reads, writes := 0, 0
			read := func(context.Context, string) (rules.Rule, error) {
				i := reads
				reads++
				if i == tc.readErrAt {
					return rules.Rule{}, errors.New("synthetic")
				}
				if i >= len(tc.reads) {
					t.Fatal("unexpected read")
				}
				item := tc.reads[i]
				if len(item.Raw) == 0 {
					item.Raw = before.Raw
				}
				return item, nil
			}
			write := func(_ context.Context, kind string, p rules.Params) (transport.Response, error) {
				writes++
				if kind != "disable" || p.ID != "id" || p.IdentityDisplay != "display" {
					t.Fatal("unsafe restore request")
				}
				if tc.writeErr {
					return transport.Response{}, errors.New("synthetic")
				}
				return transport.Response{}, nil
			}
			got := ruleUpdateOutcome(context.Background(), "id", desired, before, tc.restore, read, write)
			if got["accepted"] != true || got["readback_status"] != tc.wantStatus || got["restore_status"] != tc.wantRestore || writes != tc.wantWrites {
				t.Fatal("unexpected update/readback outcome")
			}
			if tc.name == "restored" && (got["enabled_after_update"] != true || got["final_enabled"] != false) {
				t.Fatal("restore hid the immediate enabled state")
			}
		})
	}
}
