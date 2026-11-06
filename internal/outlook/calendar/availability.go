package calendar

import (
	"crypto/sha256"
	"encoding/base64"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"net/mail"
	"strings"
	"time"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/webcapture"
)

const observedScheduleQueryHash = "9ec5dbedf881d402d59f210926bd5235c68e3db0205032de56a70c280189d08e"

type Availability struct {
	Principal        string            `json:"principal"`
	Status           string            `json:"status"`
	AvailabilityView *string           `json:"availability_view,omitempty"`
	ScheduleItems    []json.RawMessage `json:"schedule_items"`
	WorkingHours     json.RawMessage   `json:"working_hours,omitempty"`
	ErrorKind        string            `json:"error_kind,omitempty"`
	Error            json.RawMessage   `json:"error,omitempty"`
}

func scheduleAddress(s string) bool {
	if s == "" || len(s) > 320 || strings.TrimSpace(s) != s || strings.ContainsAny(s, "\r\n\x00") {
		return false
	}
	a, err := mail.ParseAddress(s)
	return err == nil && a.Name == "" && a.Address == s && strings.Contains(s, "@")
}

func ValidAvailabilityPrincipal(s string) bool { return scheduleAddress(s) }

func exactKeys(m map[string]any, keys ...string) bool {
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

func sessionUPN(s transport.Session) (string, error) {
	auth := strings.SplitN(s.Headers.Get("Authorization"), " ", 2)
	if len(auth) != 2 || !strings.EqualFold(auth[0], "Bearer") {
		return "", fmt.Errorf("current session does not include a usable account identity")
	}
	parts := strings.Split(auth[1], ".")
	if len(parts) != 3 {
		return "", fmt.Errorf("current session does not include a usable account identity")
	}
	data, err := base64.RawURLEncoding.DecodeString(parts[1])
	if err != nil {
		return "", fmt.Errorf("current session does not include a usable account identity")
	}
	var c struct {
		UPN        string `json:"upn"`
		UniqueName string `json:"unique_name"`
	}
	if json.Unmarshal(data, &c) != nil {
		return "", fmt.Errorf("current session does not include a usable account identity")
	}
	viewer := c.UPN
	if viewer == "" {
		viewer = c.UniqueName
	}
	if !scheduleAddress(viewer) {
		return "", fmt.Errorf("current session does not include a usable account identity")
	}
	return viewer, nil
}

// AvailabilityRange uses the observed UTC 15-minute query with a bounded
// window. The interval endpoints are inclusive start and exclusive end.
func AvailabilityRange(from, to string) (time.Time, time.Time, int, error) {
	a, e1 := time.Parse(time.RFC3339, from)
	b, e2 := time.Parse(time.RFC3339, to)
	if e1 != nil || e2 != nil || !a.Before(b) || b.Sub(a) > 28*24*time.Hour || !a.Equal(a.Truncate(15*time.Minute)) || !b.Equal(b.Truncate(15*time.Minute)) {
		return time.Time{}, time.Time{}, 0, fmt.Errorf("availability requires RFC3339 endpoints aligned to 15 minutes, with a range of at most 28 days")
	}
	return a.UTC(), b.UTC(), int(b.Sub(a) / (15 * time.Minute)), nil
}

func rewriteAvailability(body []byte, viewer, principal, from, to string, expectedHash string) ([]byte, int, error) {
	fail := func() ([]byte, int, error) {
		return nil, 0, fmt.Errorf("availability capture does not match the observed query")
	}
	if !scheduleAddress(viewer) || !scheduleAddress(principal) {
		return nil, 0, fmt.Errorf("availability principal must be one complete email address")
	}
	a, b, slots, err := AvailabilityRange(from, to)
	if err != nil {
		return nil, 0, err
	}
	var calls []map[string]any
	if json.Unmarshal(body, &calls) != nil || len(calls) != 1 || !exactKeys(calls[0], "operationName", "variables", "query") || calls[0]["operationName"] != "GetSchedule" {
		return fail()
	}
	query, ok := calls[0]["query"].(string)
	hash := sha256.Sum256([]byte(query))
	if !ok || hex.EncodeToString(hash[:]) != expectedHash {
		return fail()
	}
	variables := obj(calls[0]["variables"])
	input := obj(variables["input"])
	if !exactKeys(variables, "input") || !exactKeys(input, "userId", "schedules", "availabilityViewInterval", "startTime", "endTime") {
		return fail()
	}
	oldViewer, _ := input["userId"].(string)
	oldSchedules, ok := input["schedules"].([]any)
	if !scheduleAddress(oldViewer) || !ok || len(oldSchedules) != 1 || oldSchedules[0] != oldViewer || input["availabilityViewInterval"] != float64(15) || strings.Contains(query, oldViewer) {
		return fail()
	}
	for _, name := range []string{"startTime", "endTime"} {
		v := obj(input[name])
		zone := obj(v["timeZone"])
		if !exactKeys(v, "dateTime", "timeZone") || !exactKeys(zone, "name") || zone["name"] != "UTC" {
			return fail()
		}
	}
	input["userId"] = viewer
	input["schedules"] = []any{principal}
	input["availabilityViewInterval"] = 15
	obj(input["startTime"])["dateTime"] = a.Format("2006-01-02T15:04:05.000")
	obj(input["endTime"])["dateTime"] = b.Format("2006-01-02T15:04:05.000")
	encoded, err := json.Marshal(calls)
	if err != nil {
		return fail()
	}
	return encoded, slots, nil
}

func PrepareAvailability(detail, request string, current transport.Session, principal, from, to string) (transport.Session, transport.Operation, string, error) {
	viewer, err := sessionUPN(current)
	if err != nil {
		return transport.Session{}, transport.Operation{}, "", err
	}
	if principal == "" {
		principal = viewer
	}
	d, err := webcapture.Load(detail, request, current)
	if err != nil {
		return transport.Session{}, transport.Operation{}, "", err
	}
	if d.Method != "POST" || d.URL.Path != "/outlookgatewayb2/graphql" || d.HeaderBody {
		return transport.Session{}, transport.Operation{}, "", fmt.Errorf("availability capture does not match the observed query")
	}
	encoded, slots, err := rewriteAvailability(d.Body, viewer, principal, from, to, observedScheduleQueryHash)
	if err != nil {
		return transport.Session{}, transport.Operation{}, "", err
	}
	validate := func(data []byte) transport.EnvelopeResult {
		_, state := ParseAvailability(data, principal, slots)
		return state
	}
	return d.Session, d.Operation("POST", d.URL.Path, encoded, transport.Read, validate), principal, nil
}

func ParseAvailability(data []byte, principal string, slots int) (Availability, transport.EnvelopeResult) {
	var calls []struct {
		Errors []json.RawMessage `json:"errors"`
		Data   struct {
			GetSchedule *struct {
				Schedules []struct {
					ScheduleID   string          `json:"scheduleId"`
					View         *string         `json:"availabilityView"`
					Items        json.RawMessage `json:"scheduleItems"`
					WorkingHours json.RawMessage `json:"workingHours"`
					Error        json.RawMessage `json:"error"`
				} `json:"schedules"`
			} `json:"getSchedule"`
		} `json:"data"`
	}
	if json.Unmarshal(data, &calls) != nil || len(calls) != 1 || len(calls[0].Errors) != 0 || calls[0].Data.GetSchedule == nil || len(calls[0].Data.GetSchedule.Schedules) != 1 || slots < 1 {
		return Availability{}, transport.EnvelopeError
	}
	s := calls[0].Data.GetSchedule.Schedules[0]
	if !strings.EqualFold(s.ScheduleID, principal) {
		return Availability{}, transport.EnvelopeError
	}
	if len(s.Error) != 0 && string(s.Error) != "null" {
		var detail struct {
			Code         string `json:"code"`
			ResponseCode string `json:"responseCode"`
		}
		kind := "upstream_error"
		if json.Unmarshal(s.Error, &detail) == nil {
			code := detail.Code
			if code == "" {
				code = detail.ResponseCode
			}
			switch strings.ToLower(code) {
			case "erroraccessdenied", "accessdenied", "permissiondenied", "forbidden":
				kind = "permission_denied"
			}
		}
		return Availability{Principal: s.ScheduleID, Status: "error", ErrorKind: kind, Error: s.Error}, transport.EnvelopeOK
	}
	var items []json.RawMessage
	if s.View == nil || len(*s.View) != slots || len(s.Items) == 0 || string(s.Items) == "null" || json.Unmarshal(s.Items, &items) != nil {
		return Availability{}, transport.EnvelopeError
	}
	for _, c := range *s.View {
		if c < '0' || c > '9' {
			return Availability{}, transport.EnvelopeError
		}
	}
	if items == nil {
		items = []json.RawMessage{}
	}
	return Availability{Principal: s.ScheduleID, Status: "ok", AvailabilityView: s.View, ScheduleItems: items, WorkingHours: s.WorkingHours}, transport.EnvelopeOK
}
