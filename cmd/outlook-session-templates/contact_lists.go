package main

import (
	"os"
	"strings"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/capture"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/contactlists"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

// Optional as a group so existing installations can retain their previous
// bundle until all six captures are available. Values are rebuilt as synthetic
// placeholders before storage; raw fixture data is never bundled.
func addContactListTemplates(bundle *capture.RuntimeBundle, current transport.Session) {
	actions := []string{"list", "create", "update", "add-member", "remove-member", "delete"}
	configured := 0
	for _, action := range actions {
		prefix := "OUTLOOK_WEB_CONTACT_LISTS_" + strings.ToUpper(strings.ReplaceAll(action, "-", "_"))
		if os.Getenv(prefix+"_DETAILS") != "" || os.Getenv(prefix+"_REQUEST") != "" {
			configured++
		}
	}
	if configured == 0 {
		return
	}
	if configured != len(actions) {
		fatal("contact list templates must be configured together")
	}
	for _, action := range actions {
		name := "contact_lists_" + strings.ReplaceAll(action, "-", "_")
		prefix := "OUTLOOK_WEB_" + strings.ToUpper(name)
		detail, request := os.Getenv(prefix+"_DETAILS"), os.Getenv(prefix+"_REQUEST")
		if detail == "" || ((action == "list" || action == "delete") && request != "") || ((action != "list" && action != "delete") && request == "") {
			fatal("contact list template missing: " + name)
		}
		p := contactlists.Params{ID: "template-id", Name: "template list", Notes: "template notes", Email: "template@example.invalid", MemberName: "template member"}
		if action == "list" || action == "create" {
			p.ID = ""
		}
		s, op, err := contactlists.Prepare(action, detail, request, current, p)
		if err != nil {
			fatal("private contact list template incompatible: " + name)
		}
		if err = bundle.AddWeb(name, s, op); err != nil {
			fatal("cannot normalize contact list template: " + name)
		}
	}
}
