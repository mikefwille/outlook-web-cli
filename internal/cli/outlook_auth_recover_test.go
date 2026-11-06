package cli

import (
	"strings"
	"testing"
)

func TestRecoveryRequiresVerifiedOutlookRead(t *testing.T) {
	valid := `{"m365_session":1,"action":"recover","service":"outlook","ok":true,"status":"service_read_ok","identity_verified":true,"recovery_browser":"started_and_stopped","secondary_audiences":{"attachments":{"status":"captured_from_cache"}}}`
	result, err := parseRecoverResult([]byte(valid))
	if err != nil || result.Secondary.Attachments.Status != "captured_from_cache" {
		t.Fatal("verified recovery rejected")
	}
	for _, invalid := range []string{
		strings.Replace(valid, `"identity_verified":true`, `"identity_verified":false`, 1),
		strings.Replace(valid, `"status":"service_read_ok"`, `"status":"inconclusive"`, 1),
		strings.Replace(valid, `"service":"outlook"`, `"service":"teams"`, 1),
		strings.Replace(valid, `"recovery_browser":"started_and_stopped"`, `"recovery_browser":"unknown"`, 1),
	} {
		if _, err := parseRecoverResult([]byte(invalid)); err == nil {
			t.Fatal("unverified recovery accepted")
		}
	}
	_, err = parseRecoverResult([]byte(`{"m365_session":1,"action":"recover","service":"outlook","ok":false,"reason":"private@example.invalid","stage":"x\nsecret"}`))
	if err == nil || strings.Contains(err.Error(), "private") || strings.Contains(err.Error(), "secret") {
		t.Fatal("unsafe helper error was exposed")
	}
}

func TestRecoveryPreviewRequiresNoHelperOrBrowser(t *testing.T) {
	cmd := newOutlookRecoverCmd(&rootFlags{dryRun: true})
	cmd.SetArgs([]string{"--execute"})
	if err := cmd.Execute(); err != nil {
		t.Fatalf("dry-run recovery failed: %v", err)
	}
}

func TestRecoveryOutputIsBounded(t *testing.T) {
	var output boundedHelperOutput
	if n, err := output.Write([]byte(strings.Repeat("x", 32769))); n != 32769 || err != nil || !output.overflow || len(output.data) != 0 {
		t.Fatal("oversized helper response not contained")
	}
}

func TestSessionActivationRequiresVerifiedOutlookRead(t *testing.T) {
	good := `{"m365_session":1,"action":"activate","service":"outlook","ok":true,"status":"service_read_ok","identity_verified":true}`
	if _, err := parseSessionHelperResult([]byte(good), "activate"); err != nil {
		t.Fatal("valid session activation rejected")
	}
	if _, err := parseSessionHelperResult([]byte(strings.Replace(good, `"identity_verified":true`, `"identity_verified":false`, 1)), "activate"); err == nil {
		t.Fatal("unverified activation accepted")
	}
	if _, err := parseSessionHelperResult([]byte(good), "recover"); err == nil {
		t.Fatal("different helper action accepted")
	}
}
