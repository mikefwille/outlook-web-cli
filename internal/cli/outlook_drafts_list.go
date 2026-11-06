package cli

import (
	"context"
	"encoding/json"
	"fmt"
	"os"
	"strings"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/capture"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/exchange"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/folders"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

// resolveDistinguishedFolder requires a complete bounded hierarchy and exactly
// one matching system folder, so a stale captured folder ID is never reused.
func resolveDistinguishedFolder(ctx context.Context, name string) (exchange.Folder, error) {
	current, _, err := currentSession()
	if err != nil {
		return exchange.Folder{}, err
	}
	detail, _, err := outlookCapture("OUTLOOK_WEB_FOLDERS_LIST", false)
	if err != nil {
		return exchange.Folder{}, err
	}
	var found *exchange.Folder
	for offset, page := 0, 0; page < 10; page++ {
		s, op, err := folders.PrepareList(detail, current, offset, 1000)
		if err != nil {
			return exchange.Folder{}, err
		}
		response, err := outlookOperation(ctx, s, op)
		if err != nil {
			return exchange.Folder{}, err
		}
		items, state := exchange.ParseFindFolder(response.Body)
		if state != transport.EnvelopeOK || len(items) != 1 || items[0].Page == nil {
			return exchange.Folder{}, fmt.Errorf("folder hierarchy response is incomplete")
		}
		p := items[0].Page
		for _, item := range p.Folders {
			if strings.EqualFold(item.DistinguishedID, name) {
				if found != nil {
					return exchange.Folder{}, fmt.Errorf("system folder is ambiguous")
				}
				candidate := item
				found = &candidate
			}
		}
		if p.IncludesLastItemInRange {
			if found == nil || found.ID.ID == "" {
				return exchange.Folder{}, fmt.Errorf("system folder not found")
			}
			return *found, nil
		}
		if p.IndexedPagingOffset <= offset {
			return exchange.Folder{}, fmt.Errorf("folder hierarchy paging did not advance")
		}
		offset = p.IndexedPagingOffset
	}
	return exchange.Folder{}, fmt.Errorf("folder hierarchy exceeds the 10000-folder safety bound")
}

// pp:data-source live
func newDraftListCmd(flags *rootFlags) *cobra.Command {
	var offset, limit int
	cmd := &cobra.Command{Use: "list", Short: "List a page of server drafts", Long: "Resolves the current mailbox's Drafts folder, then runs the captured mail item query. Results are live and bounded; use the response paging fields to request another page.", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live"}, RunE: func(cmd *cobra.Command, args []string) error {
		if len(args) != 0 {
			return usageErr(fmt.Errorf("drafts list takes no arguments"))
		}
		if limit < 1 || limit > 1000 || offset < 0 || offset > 1000000 {
			return usageErr(fmt.Errorf("--limit must be 1–1000 and --offset must be nonnegative"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "drafts list", "limit": limit, "offset": offset})
		}
		folder, err := resolveDistinguishedFolder(cmd.Context(), "drafts")
		if err != nil {
			return err
		}
		current, _, err := currentSession()
		if err != nil {
			return err
		}
		var session transport.Session
		var op transport.Operation
		if file := os.Getenv("OUTLOOK_WEB_MAIL_LIST_EVIDENCE"); file != "" {
			session, op, err = capture.LoadObservedRead(file, "FindItem")
			if err == nil {
				session, err = capture.RebindRead(session, current)
			}
		} else {
			session, op, err = capture.PrepareRuntimeRead("mail_list", current)
		}
		if err != nil {
			return err
		}
		op, err = exchange.ParameterizeRead(op, "FindItem", exchange.ReadParameters{FolderID: folder.ID.ID, Limit: &limit, Offset: &offset})
		if err != nil {
			return err
		}
		response, err := outlookOperation(cmd.Context(), session, op)
		if err != nil {
			return err
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "folder_id": folder.ID.ID, "response": json.RawMessage(response.Body)})
	}}
	cmd.Flags().IntVar(&offset, "offset", 0, "Draft page offset")
	cmd.Flags().IntVar(&limit, "limit", 100, "Maximum drafts in this page (1–1000)")
	return cmd
}
