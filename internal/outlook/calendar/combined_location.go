package calendar

import (
	"encoding/json"
	"fmt"
)

// combinedLocations preserves every existing location in the captured
// no-source-map shape, including duplicate names unrelated to the target.
func combinedLocations(e *Event, id string, p Params) ([]any, error) {
	if e == nil || e.ID != id || !e.SafeTimedMountain() {
		return nil, bad()
	}
	count := 0
	for _, v := range []*string{p.AddLocation, p.ReplaceLocation, p.RemoveLocation} {
		if v != nil {
			count++
			if !validLocationName(*v) {
				return nil, bad()
			}
		}
	}
	if count != 1 || (p.ReplaceLocation != nil) != (p.WithLocation != nil) || p.WithLocation != nil && !validLocationName(*p.WithLocation) {
		return nil, bad()
	}
	var item map[string]json.RawMessage
	if json.Unmarshal(e.Raw, &item) != nil {
		return nil, bad()
	}
	var locations []any
	if len(item["Locations"]) == 0 || string(item["Locations"]) == "null" || json.Unmarshal(item["Locations"], &locations) != nil {
		return nil, bad()
	}
	var props []struct {
		Field struct {
			Name string `json:"PropertyName"`
		} `json:"ExtendedFieldURI"`
	}
	if v := item["ExtendedProperty"]; len(v) != 0 && string(v) != "null" && json.Unmarshal(v, &props) != nil {
		return nil, bad()
	}
	for _, prop := range props {
		if prop.Field.Name == "LocationSelectionSourceMap" {
			return nil, fmt.Errorf("combined location changes with source metadata are not yet supported; no change made")
		}
	}
	indices := map[string][]int{}
	for i, v := range locations {
		name, ok := obj(v)["DisplayName"].(string)
		if !ok || name == "" {
			return nil, bad()
		}
		indices[name] = append(indices[name], i)
	}
	switch {
	case p.AddLocation != nil:
		name := *p.AddLocation
		if len(indices[name]) != 0 {
			return nil, fmt.Errorf("location is already present")
		}
		locations = append(locations, typedLocation(name))
	case p.ReplaceLocation != nil:
		old, newName := *p.ReplaceLocation, *p.WithLocation
		if len(indices[old]) != 1 {
			return nil, fmt.Errorf("location to replace is missing or ambiguous")
		}
		if len(indices[newName]) != 0 {
			return nil, fmt.Errorf("replacement location is already present")
		}
		locations[indices[old][0]] = typedLocation(newName)
	case p.RemoveLocation != nil:
		old := *p.RemoveLocation
		if len(indices[old]) != 1 {
			return nil, fmt.Errorf("location to remove is missing or ambiguous")
		}
		i := indices[old][0]
		locations = append(locations[:i], locations[i+1:]...)
	}
	return locations, nil
}
