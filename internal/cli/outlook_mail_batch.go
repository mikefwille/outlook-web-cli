package cli

import (
	"bytes"
	"context"
	"encoding/json"
	"fmt"
	"strings"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

type mailBatchResult struct {
	ID      string         `json:"id"`
	Status  string         `json:"status"`
	Preview map[string]any `json:"preview,omitempty"`
	Result  map[string]any `json:"result,omitempty"`
	Error   string         `json:"error,omitempty"`
}
type mailBatchRunner func(context.Context, string, string, bool) (map[string]any, error)

func runMailBatch(ctx context.Context, ids, keys []string, execute bool, run mailBatchRunner) ([]mailBatchResult, error) {
	if execute && len(keys) != len(ids) {
		return nil, usageErr(fmt.Errorf("one version per message is required"))
	}
	results := make([]mailBatchResult, len(ids))
	ready := true
	for i, id := range ids {
		results[i] = mailBatchResult{ID: id, Status: "not_attempted"}
		preview, err := run(ctx, id, "", false)
		if err != nil {
			results[i].Status = "preview_failed"
			results[i].Error = err.Error()
			ready = false
			continue
		}
		key, ok := preview["change_key"].(string)
		if !ok || key == "" {
			results[i].Status = "preview_failed"
			results[i].Error = "message version unavailable"
			ready = false
			continue
		}
		results[i].Preview = preview
		results[i].Status = "previewed"
		if execute && key != keys[i] {
			results[i].Status = "version_conflict"
			ready = false
		}
	}
	if !ready {
		return results, &transport.Error{Kind: transport.KindPartial}
	}
	if !execute {
		return results, nil
	}
	for i, id := range ids {
		value, err := run(ctx, id, keys[i], true)
		if err != nil {
			results[i].Status = "failed_or_unknown"
			results[i].Error = err.Error()
			for j := i + 1; j < len(results); j++ {
				results[j].Status = "not_attempted"
			}
			return results, &transport.Error{Kind: transport.KindPartial}
		}
		results[i].Status = "accepted"
		results[i].Result = value
	}
	return results, nil
}

func newMailBatchCmd(flags *rootFlags) *cobra.Command {
	var operation string
	var ids, keys []string
	var execute bool
	cmd := &cobra.Command{Use: "batch", Short: "Preview or apply one action to several messages", Long: "Runs the same observed single-message action on up to 50 distinct messages. Preview reads each current version. Execution requires a matching --change-key for each --id in the same order, checks every preview before writing, and stops at the first failure or uncertain result. Earlier accepted items are not rolled back.", Annotations: map[string]string{"pp:data-source": "live"}, RunE: func(cmd *cobra.Command, args []string) error {
		if len(args) != 0 {
			return usageErr(fmt.Errorf("mail batch takes flags only"))
		}
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "mail batch"})
		}
		allowed := map[string]bool{"mark-read": true, "mark-unread": true, "flag": true, "unflag": true, "trash": true}
		if !allowed[operation] {
			return usageErr(fmt.Errorf("--operation must be mark-read, mark-unread, flag, unflag, or trash"))
		}
		if len(ids) == 0 || len(ids) > 50 {
			return usageErr(fmt.Errorf("provide 1–50 --id values"))
		}
		seen := map[string]bool{}
		for _, id := range ids {
			if id == "" || len(id) > 4096 || strings.ContainsAny(id, "\x00\r\n") || seen[id] {
				return usageErr(fmt.Errorf("message IDs must be nonempty and distinct"))
			}
			seen[id] = true
		}
		if execute && !flags.dryRun && len(keys) != len(ids) {
			return usageErr(fmt.Errorf("provide one --change-key per --id in the same order"))
		}
		if !execute && len(keys) > 0 {
			return usageErr(fmt.Errorf("--change-key is used only with --execute"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "mail batch", "action": operation, "count": len(ids)})
		}
		if execute {
			if err := mutationInHarness(); err != nil {
				return err
			}
		}
		runner := func(ctx context.Context, id, key string, write bool) (map[string]any, error) {
			nested := newMailActionCmd(&rootFlags{asJSON: true, noInput: true, noLearn: true}, operation)
			nested.SilenceErrors = true
			nested.SilenceUsage = true
			var out bytes.Buffer
			nested.SetOut(&out)
			nested.SetErr(&bytes.Buffer{})
			argv := []string{id}
			if write {
				argv = append(argv, "--change-key", key, "--execute")
			}
			nested.SetArgs(argv)
			if err := nested.ExecuteContext(ctx); err != nil {
				return nil, err
			}
			var result map[string]any
			if json.Unmarshal(out.Bytes(), &result) != nil {
				return nil, fmt.Errorf("single-message action returned an invalid result")
			}
			return result, nil
		}
		results, err := runMailBatch(cmd.Context(), ids, keys, execute, runner)
		if outputErr := flags.printJSON(cmd, map[string]any{"source": "live", "operation": "mail batch", "action": operation, "preview": !execute, "results": results, "atomic": false, "retries": false}); outputErr != nil {
			return outputErr
		}
		return err
	}}
	cmd.Flags().StringVar(&operation, "operation", "", "Action: mark-read, mark-unread, flag, unflag, or trash")
	cmd.Flags().StringArrayVar(&ids, "id", nil, "Message ID; repeat up to 50 times")
	cmd.Flags().StringArrayVar(&keys, "change-key", nil, "Previewed version for each ID, in the same order")
	cmd.Flags().BoolVar(&execute, "execute", false, "Apply the reviewed action once per message")
	return cmd
}
