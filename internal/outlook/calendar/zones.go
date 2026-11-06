package calendar

import (
	"fmt"
	"time"
)

// CanonicalTimeZone resolves the two event zones observed in Outlook captures.
func CanonicalTimeZone(name string) (string, error) {
	switch name {
	case "", "America/Denver", "Mountain Standard Time":
		return "Mountain Standard Time", nil
	case "America/New_York", "Eastern Standard Time":
		return "Eastern Standard Time", nil
	default:
		return "", fmt.Errorf("supported captured event time zones are America/Denver and America/New_York")
	}
}

func zoneLocation(name string) (*time.Location, error) {
	zone, err := CanonicalTimeZone(name)
	if err != nil {
		return nil, err
	}
	if zone == "Eastern Standard Time" {
		return time.LoadLocation("America/New_York")
	}
	return time.LoadLocation("America/Denver")
}

func localTimeInZone(value, name string) (string, error) {
	t, err := time.Parse(time.RFC3339, value)
	if err != nil {
		return "", fmt.Errorf("start and end must include an RFC3339 offset")
	}
	zone, err := zoneLocation(name)
	if err != nil {
		return "", err
	}
	const layout = "2006-01-02T15:04:05.000"
	v := t.In(zone).Format(layout)
	if t.Add(time.Hour).In(zone).Format(layout) == v || t.Add(-time.Hour).In(zone).Format(layout) == v {
		return "", fmt.Errorf("the selected instant is ambiguous in the event time zone")
	}
	return v, nil
}

func eventLocalTime(value, name string) string {
	if name == "" {
		return ""
	}
	zone, err := zoneLocation(name)
	if err != nil {
		return ""
	}
	t, err := time.Parse(time.RFC3339, value)
	if err != nil {
		return ""
	}
	return t.In(zone).Format(time.RFC3339Nano)
}
