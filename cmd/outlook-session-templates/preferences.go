package main

import (
	"os"
	"strings"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/account"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/capture"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func addPreferenceTemplates(bundle *capture.RuntimeBundle, current transport.Session) {
	operations := map[string]string{"inbox_rules_list": "rules", "preferences_auto_replies": "auto-replies", "preference_focused_update": "focused-update"}
	configured := 0
	for name := range operations {
		prefix := "OUTLOOK_WEB_" + strings.ToUpper(name)
		if os.Getenv(prefix+"_DETAILS") != "" || os.Getenv(prefix+"_REQUEST") != "" {
			configured++
		}
	}
	if configured == 0 {
		return
	}
	if configured != len(operations) {
		fatal("preference templates must be configured together")
	}
	for name, kind := range operations {
		prefix := "OUTLOOK_WEB_" + strings.ToUpper(name)
		detail := os.Getenv(prefix + "_DETAILS")
		if detail == "" || os.Getenv(prefix+"_REQUEST") != "" {
			fatal("preference template missing: " + name)
		}
		s, op, err := account.PreparePreference(kind, detail, current, false)
		if err != nil {
			fatal("private preference template incompatible: " + name)
		}
		if err = bundle.AddWeb(name, s, op); err != nil {
			fatal("cannot normalize preference template: " + name)
		}
	}
}
