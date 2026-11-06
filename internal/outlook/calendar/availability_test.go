package calendar

import (
	"crypto/sha256"
	"encoding/base64"
	"encoding/hex"
	"encoding/json"
	"net/http"
	"strings"
	"testing"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func TestAvailabilityRangeAndIdentity(t *testing.T) {
	for _, tc := range []struct {
		from, to string
		slots    int
		valid    bool
	}{
		{"2026-11-06T04:00:00-07:00", "2026-11-06T05:00:00-07:00", 4, true},
		{"2026-11-01T01:30:00-06:00", "2026-11-01T01:30:00-07:00", 4, true},
		{"2026-11-06T04:01:00-07:00", "2026-11-06T05:00:00-07:00", 0, false},
		{"2026-11-06T04:00:00-07:00", "2026-12-06T05:00:00-07:00", 0, false},
		{"2026-11-06T04:00:00-07:00", "2026-11-06T04:00:00-07:00", 0, false},
	} {
		_, _, slots, err := AvailabilityRange(tc.from, tc.to)
		if (err == nil) != tc.valid || slots != tc.slots {
			t.Fatalf("range validation failed: %v", err)
		}
	}
	claims, _ := json.Marshal(map[string]any{"upn": "viewer@example.invalid"})
	s := transport.Session{Headers: http.Header{"Authorization": []string{"Bearer e30." + base64.RawURLEncoding.EncodeToString(claims) + ".signature"}}}
	if got, err := sessionUPN(s); err != nil || got != "viewer@example.invalid" {
		t.Fatal("session identity rejected")
	}
	if ValidAvailabilityPrincipal("Viewer <viewer@example.invalid>") || ValidAvailabilityPrincipal("viewer@example.invalid\nother") {
		t.Fatal("invalid principal accepted")
	}
}

func TestAvailabilityRewriteBindsViewerPrincipalAndRange(t *testing.T) {
	query := "query GetSchedule { synthetic }"
	hash := sha256.Sum256([]byte(query))
	body, _ := json.Marshal([]any{map[string]any{"operationName": "GetSchedule", "query": query, "variables": map[string]any{"input": map[string]any{"userId": "old@example.invalid", "schedules": []any{"old@example.invalid"}, "availabilityViewInterval": 15, "startTime": map[string]any{"dateTime": "fixture-start", "timeZone": map[string]any{"name": "UTC"}}, "endTime": map[string]any{"dateTime": "fixture-end", "timeZone": map[string]any{"name": "UTC"}}}}}})
	encoded, slots, err := rewriteAvailability(body, "viewer@example.invalid", "target@example.invalid", "2026-11-06T04:00:00-07:00", "2026-11-06T05:00:00-07:00", hex.EncodeToString(hash[:]))
	if err != nil || slots != 4 || strings.Contains(string(encoded), "old@example.invalid") || strings.Contains(string(encoded), "fixture-") {
		t.Fatal("availability request retained fixture values")
	}
	var calls []map[string]any
	if json.Unmarshal(encoded, &calls) != nil {
		t.Fatal("invalid rewritten query")
	}
	i := obj(obj(calls[0]["variables"])["input"])
	if i["userId"] != "viewer@example.invalid" || i["schedules"].([]any)[0] != "target@example.invalid" || obj(i["startTime"])["dateTime"] != "2026-11-06T11:00:00.000" || obj(i["endTime"])["dateTime"] != "2026-11-06T12:00:00.000" {
		t.Fatal("availability binding incorrect")
	}
	if _, _, err = rewriteAvailability(body, "viewer@example.invalid", "target@example.invalid", "2026-11-06T04:00:00-07:00", "2026-11-06T05:00:00-07:00", strings.Repeat("0", 64)); err == nil {
		t.Fatal("unrecognized query accepted")
	}
}

func TestAvailabilityResponseKeepsErrorsAndUnknownCodes(t *testing.T) {
	good := `[{"data":{"getSchedule":{"schedules":[{"scheduleId":"viewer@example.invalid","availabilityView":"0022","scheduleItems":[],"workingHours":null,"error":null}]}}}]`
	v, state := ParseAvailability([]byte(good), "viewer@example.invalid", 4)
	if state != transport.EnvelopeOK || v.Status != "ok" || v.AvailabilityView == nil || *v.AvailabilityView != "0022" {
		t.Fatal("availability response rejected")
	}
	denied := `[{"data":{"getSchedule":{"schedules":[{"scheduleId":"viewer@example.invalid","error":{"code":"synthetic-denied"}}]}}}]`
	v, state = ParseAvailability([]byte(denied), "viewer@example.invalid", 4)
	if state != transport.EnvelopeOK || v.Status != "error" || v.AvailabilityView != nil || len(v.Error) == 0 {
		t.Fatal("per-principal error treated as free")
	}
	permission := strings.Replace(denied, "synthetic-denied", "ErrorAccessDenied", 1)
	v, state = ParseAvailability([]byte(permission), "viewer@example.invalid", 4)
	if state != transport.EnvelopeOK || v.Status != "error" || v.ErrorKind != "permission_denied" || v.AvailabilityView != nil {
		t.Fatal("permission denial was not distinguished")
	}
	for _, invalid := range []string{strings.Replace(good, `"0022"`, `"002"`, 1), strings.Replace(good, "viewer@example.invalid", "other@example.invalid", 1), `[{"errors":[{"message":"denied"}]}]`} {
		if _, state = ParseAvailability([]byte(invalid), "viewer@example.invalid", 4); state == transport.EnvelopeOK {
			t.Fatal("invalid availability accepted")
		}
	}
}
