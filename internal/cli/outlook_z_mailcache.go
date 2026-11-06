package cli

import (
	"context"
	"fmt"
	"os"
	"path/filepath"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/capture"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/exchange"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/mailcache"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func mailCacheRoot() (string, error) {
	pointer := os.Getenv("OUTLOOK_WEB_SESSION_POINTER")
	if pointer != "" {
		return filepath.Dir(pointer), nil
	}
	home, err := os.UserHomeDir()
	if err != nil {
		return "", err
	}
	return filepath.Join(home, ".local", "share", "outlook-web-dev", "private"), nil
}

func openMailCache(create bool) (*mailcache.Store, error) {
	current, _, err := currentSession()
	if err != nil {
		return nil, err
	}
	root, err := mailCacheRoot()
	if err != nil {
		return nil, err
	}
	return mailcache.Open(root, current.AccountKey, create)
}

func mailPageFetcher(current transport.Session, folderID string) (mailcache.FetchPage, error) {
	var session transport.Session
	var operation transport.Operation
	var err error
	if file := os.Getenv("OUTLOOK_WEB_MAIL_LIST_EVIDENCE"); file != "" {
		session, operation, err = capture.LoadObservedRead(file, "FindItem")
		if err == nil {
			session, err = capture.RebindRead(session, current)
		}
	} else {
		session, operation, err = capture.PrepareRuntimeRead("mail_list", current)
	}
	if err != nil {
		return nil, err
	}
	client, err := transport.New(session.AccountKey, session, nil)
	if err != nil {
		return nil, err
	}
	return func(ctx context.Context, offset, limit int) (mailcache.Page, error) {
		op, err := exchange.ParameterizeRead(operation, "FindItem", exchange.ReadParameters{FolderID: folderID, Offset: &offset, Limit: &limit})
		if err != nil {
			return mailcache.Page{}, err
		}
		response, err := client.Do(ctx, session.AccountKey, op)
		if err != nil {
			return mailcache.Page{}, err
		}
		return mailcache.ParsePage(response.Body, folderID)
	}, nil
}

// pp:data-source live
func newMailSyncCmd(flags *rootFlags) *cobra.Command {
	parent := &cobra.Command{Use: "sync", Short: "Preview or reconcile one mail folder", Long: "Preview or reconcile a bounded folder using the same behavior as sync folder. Run sync status for existing runs. Nothing is written to Outlook.", Annotations: map[string]string{"pp:data-source": "live", "mcp:local-write": "true"}}
	var folderID, resume string
	var maxItems int
	var execute bool
	folder := &cobra.Command{Use: "folder", Short: "Preview or reconcile one mail folder", Long: "Saves metadata from the observed mail-list query to a private account-scoped cache. Pages and checkpoints commit together. Before publishing, a second complete pass checks every item ID and change key; a changed or incomplete scan leaves the prior snapshot intact. This is not a server delta cursor or an atomic server snapshot. Nothing is written to Outlook.", Annotations: map[string]string{"pp:data-source": "live", "mcp:local-write": "true"}, RunE: func(cmd *cobra.Command, args []string) error {
		if len(args) != 0 {
			return usageErr(fmt.Errorf("sync folder takes flags only"))
		}
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "sync folder"})
		}
		if folderID == "" || maxItems < 1 || maxItems > 10000 {
			return usageErr(fmt.Errorf("--folder-id and --max-items between 1 and 10000 are required"))
		}
		if !execute || flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"preview": true, "operation": "sync folder", "folder_id": folderID, "resume_run_id": resume, "max_items": maxItems, "writes": "private local cache only", "server_writes": false, "verification": "two bounded passes; not an atomic server snapshot"})
		}
		if err := mutationInHarness(); err != nil {
			return err
		}
		current, _, err := currentSession()
		if err != nil {
			return err
		}
		root, err := mailCacheRoot()
		if err != nil {
			return err
		}
		store, err := mailcache.Open(root, current.AccountKey, true)
		if err != nil {
			return err
		}
		defer store.Close()
		var run mailcache.Run
		if resume == "" {
			run, err = store.Start(cmd.Context(), folderID)
		} else {
			run, err = store.Get(cmd.Context(), resume)
			if err == nil && run.FolderID != folderID {
				return usageErr(fmt.Errorf("--folder-id does not match the selected run"))
			}
		}
		if err != nil {
			return err
		}
		fetch, err := mailPageFetcher(current, folderID)
		if err != nil {
			return fmt.Errorf("sync run %s was not published: %w", run.ID, err)
		}
		run, err = mailcache.Reconcile(cmd.Context(), store, run, fetch, maxItems)
		if err != nil {
			return fmt.Errorf("sync run %s was not published; inspect sync status before resuming: %w", run.ID, err)
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "cached": true, "run": run, "coverage": "one requested mail folder, list metadata only", "consistency": "verified by a second pass; not an atomic server snapshot"})
	}}
	folder.Flags().StringVar(&folderID, "folder-id", "", "Mailbox folder ID to reconcile")
	folder.Flags().StringVar(&resume, "resume", "", "Resume a previously interrupted run ID")
	folder.Flags().IntVar(&maxItems, "max-items", 1000, "Maximum folder items to reconcile (1–10000)")
	folder.Flags().BoolVar(&execute, "execute", false, "Write the verified snapshot to the private cache")
	// The approved `sync` path runs the same bounded reconciliation as
	// `sync folder`; both declare all flags and default to preview.
	parent.Flags().StringVar(&folderID, "folder-id", "", "Mailbox folder ID to reconcile")
	parent.Flags().StringVar(&resume, "resume", "", "Resume a previously interrupted run ID")
	parent.Flags().IntVar(&maxItems, "max-items", 1000, "Maximum folder items to reconcile (1–10000)")
	parent.Flags().BoolVar(&execute, "execute", false, "Write the verified snapshot to the private cache")
	parent.RunE = folder.RunE
	parent.AddCommand(folder)
	status := &cobra.Command{Use: "status", Short: "List recent local reconciliation runs", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "local"}, RunE: func(cmd *cobra.Command, args []string) error {
		if len(args) != 0 {
			return usageErr(fmt.Errorf("sync status takes no arguments"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "sync status"})
		}
		store, err := openMailCache(false)
		if os.IsNotExist(err) {
			return flags.printJSON(cmd, map[string]any{"source": "local", "runs": []mailcache.Run{}})
		}
		if err != nil {
			return err
		}
		defer store.Close()
		runs, err := store.Runs(cmd.Context())
		if err != nil {
			return err
		}
		return flags.printJSON(cmd, map[string]any{"source": "local", "runs": runs})
	}}
	parent.AddCommand(status)
	return parent
}

// pp:data-source local
func newLocalMailSearchCmd(flags *rootFlags) *cobra.Command {
	var folderID string
	var limit int
	cmd := &cobra.Command{Use: "search <text>", Short: "Search reconciled local mail metadata", Long: "Searches only the last verified local snapshots from sync folder. Results include each folder's sync time. This does not search bodies or unsynced folders; use mail search for a live Outlook search.", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "local", "pp:happy-args": "text=example"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "search"})
		}
		if len(args) != 1 || args[0] == "" || limit < 1 || limit > 100 {
			return usageErr(fmt.Errorf("supply one search text and --limit between 1 and 100"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "search", "limit": limit})
		}
		store, err := openMailCache(false)
		if os.IsNotExist(err) {
			return fmt.Errorf("no local mail snapshot; run sync folder first")
		}
		if err != nil {
			return err
		}
		defer store.Close()
		hits, err := store.Search(cmd.Context(), args[0], folderID, limit)
		if err != nil {
			return err
		}
		return flags.printJSON(cmd, map[string]any{"source": "local", "coverage": "verified snapshots of explicitly synced folders; mail-list metadata only", "results": hits, "limit": limit})
	}}
	cmd.Flags().StringVar(&folderID, "folder-id", "", "Limit results to one synced folder")
	cmd.Flags().IntVar(&limit, "limit", 25, "Maximum matching cached items (1–100)")
	return cmd
}

// pp:data-source local
func newLocalMailSQLCmd(flags *rootFlags) *cobra.Command {
	var limit int
	cmd := &cobra.Command{Use: "sql <query>", Short: "Query published local mail metadata", Long: "Runs one bounded read-only SELECT over items(folder, item_id, change_key, data) and heads(folder, synced_at, item_count). Only explicitly synced folder snapshots are present. The SQL authorizer blocks writes, attachments, staging tables and unsafe functions.", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "local", "pp:happy-args": "query=SELECT count(*) FROM items"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "sql"})
		}
		if len(args) != 1 || args[0] == "" || limit < 1 || limit > 100 {
			return usageErr(fmt.Errorf("supply one SELECT and --limit between 1 and 100"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "sql", "limit": limit})
		}
		store, err := openMailCache(false)
		if os.IsNotExist(err) {
			return fmt.Errorf("no local mail snapshot; run sync folder first")
		}
		if err != nil {
			return err
		}
		defer store.Close()
		result, err := store.Query(cmd.Context(), args[0], limit)
		if err != nil {
			return err
		}
		return flags.printJSON(cmd, map[string]any{"source": "local", "coverage": "published folder metadata only; join heads for snapshot times", "query_result": result})
	}}
	cmd.Flags().IntVar(&limit, "limit", 25, "Maximum SQL rows returned (1–100)")
	return cmd
}

func init() {
	registerNovelCommand(func(root *cobra.Command, flags *rootFlags) {
		for _, cmd := range root.Commands() {
			if cmd.Name() == "sync" || cmd.Name() == "search" || cmd.Name() == "sql" || cmd.Name() == "export" {
				root.RemoveCommand(cmd)
			}
		}
		root.AddCommand(newMailSyncCmd(flags), newLocalMailSearchCmd(flags), newLocalMailSQLCmd(flags), newLocalMailExportCmd(flags))
	})
}
