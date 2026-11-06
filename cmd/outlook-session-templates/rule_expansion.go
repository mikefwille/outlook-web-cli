package main

import (
	"os"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/capture"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/rules"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func addRuleExpansionTemplates(bundle *capture.RuntimeBundle, current transport.Session) {
	entries := []struct{ name, prefix string }{{"rule_enable", "OUTLOOK_WEB_RULE_ENABLE"}, {"rules_reorder", "OUTLOOK_WEB_RULES_REORDER"}}
	configured := 0
	for _, entry := range entries {
		if os.Getenv(entry.prefix+"_DETAILS") != "" || os.Getenv(entry.prefix+"_REQUEST") != "" {
			configured++
		}
	}
	if configured == 0 {
		return
	}
	if configured != len(entries) {
		fatal("rule enable and reorder templates must be configured together")
	}
	for _, entry := range entries {
		detail := os.Getenv(entry.prefix + "_DETAILS")
		if detail == "" || os.Getenv(entry.prefix+"_REQUEST") != "" {
			fatal("rule expansion template missing: " + entry.name)
		}
		var s transport.Session
		var op transport.Operation
		var err error
		if entry.name == "rule_enable" {
			s, op, err = rules.Prepare("enable", detail, current, rules.Params{ID: "template-rule-id", IdentityDisplay: "template identity"})
		} else {
			s, op, err = rules.PrepareReorder(detail, current, []rules.OrderMember{{ID: "template-first", IdentityDisplay: "template first", Enabled: true}, {ID: "template-second", IdentityDisplay: "template second", Enabled: false}})
		}
		if err != nil {
			fatal("private rule expansion template incompatible: " + entry.name)
		}
		if err = bundle.AddWeb(entry.name, s, op); err != nil {
			fatal("cannot normalize rule expansion template: " + entry.name)
		}
	}
}
