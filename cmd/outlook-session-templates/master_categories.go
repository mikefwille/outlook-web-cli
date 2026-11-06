package main

import (
	"os"
	"strings"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/capture"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/mastercategories"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func addMasterCategoryTemplates(bundle *capture.RuntimeBundle, current transport.Session) {
	kinds := []string{"create", "update", "delete", "rename-items", "delete-items"}
	configured := 0
	for _, kind := range kinds {
		prefix := "OUTLOOK_WEB_MASTER_CATEGORY_" + strings.ToUpper(strings.ReplaceAll(kind, "-", "_"))
		if os.Getenv(prefix+"_DETAILS") != "" || os.Getenv(prefix+"_REQUEST") != "" {
			configured++
		}
	}
	if configured == 0 {
		return
	}
	if configured != len(kinds) {
		fatal("master category templates must be configured together")
	}
	for _, kind := range kinds {
		name := "master_category_" + strings.ReplaceAll(kind, "-", "_")
		prefix := "OUTLOOK_WEB_" + strings.ToUpper(name)
		detail := os.Getenv(prefix + "_DETAILS")
		if detail == "" || os.Getenv(prefix+"_REQUEST") != "" {
			fatal("master category template missing: " + name)
		}
		p := mastercategories.Params{OldName: "template old category", New: mastercategories.Category{Name: "template new category", Color: 20, ID: "00000000-0000-4000-8000-000000000001", LastTimeUsed: "2026-01-01T00:00:00.000Z", KeyboardShortcut: 0}}
		session, op, err := mastercategories.Prepare(kind, detail, current, p)
		if err != nil {
			fatal("private master category template incompatible: " + name)
		}
		if err = bundle.AddWeb(name, session, op); err != nil {
			fatal("cannot normalize master category template: " + name)
		}
	}
}
