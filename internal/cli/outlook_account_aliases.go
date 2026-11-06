package cli

import (
	"context"
	"fmt"
	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/account"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func readAccountAliases(ctx context.Context) ([]string, error) {
	detail, _, err := outlookCapture("OUTLOOK_WEB_ACCOUNT_ALIASES", false)
	if err != nil {
		return nil, err
	}
	current, _, err := currentSession()
	if err != nil {
		return nil, err
	}
	session, op, err := account.PrepareAliases(detail, current)
	if err != nil {
		return nil, err
	}
	response, err := outlookOperation(ctx, session, op)
	if err != nil {
		return nil, err
	}
	a, state := account.ParseAliases(response.Body)
	if state != transport.EnvelopeOK {
		return nil, &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
	}
	return a, nil
}
func newAccountAliasesCmd(flags *rootFlags) *cobra.Command {
	return &cobra.Command{Use: "aliases", Short: "List the current mailbox's SMTP addresses", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live"}, RunE: func(cmd *cobra.Command, args []string) error {
		if len(args) != 0 {
			return usageErr(fmt.Errorf("account aliases takes no arguments"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "account aliases"})
		}
		a, err := readAccountAliases(cmd.Context())
		if err != nil {
			return err
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "addresses": a})
	}}
}
