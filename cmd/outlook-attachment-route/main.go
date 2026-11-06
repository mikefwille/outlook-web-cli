// Convert one observed attachment request into a credential-free private route.
package main

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/attachments"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/capture"
)

func main() {
	if len(os.Args) != 2 {
		fmt.Fprintln(os.Stderr, "supply the private observed attachment request detail")
		os.Exit(2)
	}
	pointer := os.Getenv("OUTLOOK_WEB_SESSION_POINTER")
	if pointer == "" {
		h, err := os.UserHomeDir()
		if err != nil {
			fmt.Fprintln(os.Stderr, "current session unavailable")
			os.Exit(1)
		}
		pointer = filepath.Join(h, ".local", "share", "outlook-web-dev", "private", "current-outlook.json")
	}
	current, _, err := capture.LoadCurrentFindFolder(pointer)
	if err != nil {
		fmt.Fprintln(os.Stderr, "current session unavailable")
		os.Exit(1)
	}
	route, err := attachments.RouteFromCapture(os.Args[1], current)
	if err != nil {
		fmt.Fprintln(os.Stderr, "attachment route capture incompatible")
		os.Exit(1)
	}
	path := filepath.Join(filepath.Dir(pointer), "outlook-attachment-route-v1.json")
	if err = attachments.SaveDownloadRoute(path, route); err != nil {
		fmt.Fprintln(os.Stderr, "cannot save private attachment route")
		os.Exit(1)
	}
	_ = json.NewEncoder(os.Stdout).Encode(map[string]any{"ok": true, "credential_saved": false, "canary_saved": false})
}
