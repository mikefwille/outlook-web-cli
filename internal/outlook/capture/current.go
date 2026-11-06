package capture

import (
	"encoding/json"
	"path/filepath"
	"regexp"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

var exportID = regexp.MustCompile(`^[a-z0-9]+-[a-f0-9]{8}$`)

// LoadCurrentFindFolder consumes the shared adapter's atomic service pointer.
// The pointer contains an opaque export identifier, never credentials.
func LoadCurrentFindFolder(pointer string) (transport.Session, transport.Operation, error) {
	data, err := readPrivateFile(pointer, 4096)
	if err != nil {
		return transport.Session{}, transport.Operation{}, &transport.Error{Kind: transport.KindAuth}
	}
	var p struct {
		Version  int    `json:"version"`
		Service  string `json:"service"`
		ExportID string `json:"export_id"`
	}
	if json.Unmarshal(data, &p) != nil || p.Version != 1 || p.Service != "outlook" || !exportID.MatchString(p.ExportID) {
		return transport.Session{}, transport.Operation{}, &transport.Error{Kind: transport.KindAuth}
	}
	return LoadFindFolder(filepath.Join(filepath.Dir(pointer), "session-outlook-"+p.ExportID+".json"))
}
