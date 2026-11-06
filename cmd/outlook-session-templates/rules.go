package main

import (
	"os"
	"strings"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/capture"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/rules"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func addRuleTemplates(bundle *capture.RuntimeBundle, current transport.Session) {
	actions := []string{"create", "disable", "update", "delete"}
	configured := 0
	for _, action := range actions {
		prefix := "OUTLOOK_WEB_RULE_" + strings.ToUpper(action)
		if os.Getenv(prefix+"_DETAILS") != "" || os.Getenv(prefix+"_REQUEST") != "" {
			configured++
		}
	}
	if configured == 0 {
		return
	}
	if configured != len(actions) {
		fatal("rule templates must be configured together")
	}
	for _, action := range actions {
		name := "rule_" + action
		prefix := "OUTLOOK_WEB_" + strings.ToUpper(name)
		detail := os.Getenv(prefix + "_DETAILS")
		if detail == "" || os.Getenv(prefix+"_REQUEST") != "" {
			fatal("rule template missing: " + name)
		}
		p := rules.Params{ID: "template-rule-id", IdentityDisplay: "template identity", Name: "template rule", Subject: "template subject", Category: "template category"}
		s, op, err := rules.Prepare(action, detail, current, p)
		if err != nil {
			fatal("private rule template incompatible: " + name)
		}
		if err = bundle.AddWeb(name, s, op); err != nil {
			fatal("cannot normalize rule template: " + name)
		}
	}
}
