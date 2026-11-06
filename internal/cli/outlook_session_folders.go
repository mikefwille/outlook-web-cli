// Hand-authored Outlook web extension. This is provisional and must pass the
// remaining discovery, generation and live acceptance gates before promotion.
package cli

import (
	"context"
	"encoding/json"
	"os"
	"path/filepath"
	"sync"
	"time"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/capture"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/exchange"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/folders"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func init() {
	registerNovelCommand(func(root *cobra.Command, flags *rootFlags) {
		// The preliminary generator inferred generic handlers for a shared RPC
		// path. Those handlers do not preserve its envelope and expose a free-form
		// action. Quarantine them until each service-specific command is ready.
		for _, c := range root.Commands() {
			switch c.Name() {
			case "folders", "threads", "mail", "calendar", "calendars", "account", "auth", "doctor":
				root.RemoveCommand(c)
			}
		}
		root.AddCommand(newObservedFoldersCmd(flags))
		threads := newObservedReadGroup("threads", flags, []observedRead{
			{name: "list", action: "FindConversation", env: "OUTLOOK_WEB_THREADS_LIST_EVIDENCE"},
			{name: "get", action: "GetConversationItems", env: "OUTLOOK_WEB_THREADS_GET_EVIDENCE"},
		})
		for _, kind := range []string{"mark-read", "mark-unread", "flag", "complete", "move", "restore"} {
			threads.AddCommand(newConversationActionCmd(flags, kind))
		}
		root.AddCommand(threads)
		mail := newObservedReadGroup("mail", flags, []observedRead{{name: "list", action: "FindItem", env: "OUTLOOK_WEB_MAIL_LIST_EVIDENCE"}})
		mail.AddCommand(newMailComposeCmd(flags))
		mail.AddCommand(newMailGetCmd(flags))
		mail.AddCommand(newMailCopyCmd(flags))
		mail.AddCommand(newMailExportCmd(flags))
		mail.AddCommand(newMailBatchCmd(flags))
		mail.AddCommand(newMailCategoriesCmd(flags))
		mail.AddCommand(newMailSearchCmd(flags))
		for _, kind := range []string{"mark-read", "mark-unread", "flag", "unflag", "complete", "trash"} {
			mail.AddCommand(newMailActionCmd(flags, kind))
		}
		for _, kind := range []string{"move", "archive", "restore"} {
			mail.AddCommand(newMailMoveCmd(flags, kind))
		}
		mail.AddCommand(newResponseDraftCmd(flags, "reply"), newResponseDraftCmd(flags, "reply-all"), newResponseDraftCmd(flags, "forward"))
		root.AddCommand(mail)
		root.AddCommand(newObservedReadGroup("calendars", flags, []observedRead{{name: "list", action: "GetCalendarFoldersV2", env: "OUTLOOK_WEB_CALENDARS_LIST_EVIDENCE"}}))
		root.AddCommand(newObservedReadGroup("calendar", flags, []observedRead{{name: "agenda", action: "GetCalendarView", env: "OUTLOOK_WEB_CALENDAR_AGENDA_EVIDENCE"}}))
		account := newObservedReadGroup("account", flags, []observedRead{{name: "info", action: "GetAccountInformation", env: "OUTLOOK_WEB_ACCOUNT_INFO_EVIDENCE"}})
		account.AddCommand(newAccountAliasesCmd(flags))
		root.AddCommand(account)
		root.AddCommand(newObservedAuthCmd(flags))
		root.AddCommand(newOutlookDoctorCmd(flags))
		root.AddCommand(newDraftsCmd(flags))
	})
}

// An invocation keeps the first loaded session for all of its reads and
// writes. Another process can switch the shared pointer without changing the
// account halfway through an existing CLI or MCP shell-out invocation.
var invocationSession struct {
	sync.Mutex
	active, loaded bool
	session        transport.Session
	operation      transport.Operation
	err            error
}

func beginOutlookInvocation() func() {
	invocationSession.Lock()
	invocationSession.active, invocationSession.loaded = true, false
	invocationSession.session, invocationSession.operation, invocationSession.err = transport.Session{}, transport.Operation{}, nil
	invocationSession.Unlock()
	return func() {
		invocationSession.Lock()
		invocationSession.active, invocationSession.loaded = false, false
		invocationSession.session, invocationSession.operation, invocationSession.err = transport.Session{}, transport.Operation{}, nil
		invocationSession.Unlock()
	}
}

func currentSession() (transport.Session, transport.Operation, error) {
	invocationSession.Lock()
	defer invocationSession.Unlock()
	if !invocationSession.active {
		return loadCurrentSession()
	}
	if !invocationSession.loaded {
		invocationSession.session, invocationSession.operation, invocationSession.err = loadCurrentSession()
		invocationSession.loaded = true
	}
	s, op := invocationSession.session, invocationSession.operation
	s.Headers = s.Headers.Clone()
	op.Headers = op.Headers.Clone()
	op.Body = append([]byte(nil), op.Body...)
	if op.Query != nil {
		query := make(map[string][]string, len(op.Query))
		for key, values := range op.Query {
			query[key] = append([]string(nil), values...)
		}
		op.Query = query
	}
	return s, op, invocationSession.err
}

func loadCurrentSession() (transport.Session, transport.Operation, error) {
	if file := os.Getenv("OUTLOOK_WEB_SESSION_FILE"); file != "" {
		return capture.LoadFindFolder(file)
	}
	pointer := os.Getenv("OUTLOOK_WEB_SESSION_POINTER")
	if pointer == "" {
		home, err := os.UserHomeDir()
		if err != nil {
			return transport.Session{}, transport.Operation{}, &transport.Error{Kind: transport.KindAuth}
		}
		pointer = filepath.Join(home, ".local", "share", "outlook-web-dev", "private", "current-outlook.json")
	}
	return capture.LoadCurrentFindFolder(pointer)
}

func newObservedAuthCmd(flags *rootFlags) *cobra.Command {
	parent := &cobra.Command{Use: "auth", Short: "Inspect the saved Outlook session"}
	parent.AddCommand(&cobra.Command{Use: "status", Short: "Check local session availability and expiry without a network call", Args: cobra.NoArgs,
		Annotations: map[string]string{"mcp:read-only": "true"}, RunE: func(cmd *cobra.Command, _ []string) error {
			if flags.dryRun {
				return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "auth status", "state_change": false})
			}
			s, _, err := currentSession()
			if err != nil {
				return err
			}
			return json.NewEncoder(cmd.OutOrStdout()).Encode(map[string]any{"service": "outlook", "state": "available", "expires_at": s.ExpiresAt.UTC().Format(time.RFC3339), "network_checked": false})
		}})
	parent.AddCommand(newOutlookRecoverCmd(flags))
	parent.AddCommand(newOutlookLogoutCmd(flags))
	parent.AddCommand(newOutlookImportCmd(flags))
	login := newOutlookRecoverCmd(flags)
	login.Use = "login"
	login.Short = "Sign in or recover through the saved Outlook browser profile"
	parent.AddCommand(login)
	return parent
}

func newObservedFoldersCmd(flags *rootFlags) *cobra.Command {
	parent := &cobra.Command{Use: "folders", Short: "Inspect folders in the captured mailbox"}
	var offset, limit int
	list := &cobra.Command{
		Use: "list", Short: "List mailbox folders with names and counts",
		Long:        "List the observed deep mailbox hierarchy with names, parent IDs, counts and paging metadata. Use --offset and --limit to request another page.",
		Args:        cobra.NoArgs,
		Annotations: map[string]string{"mcp:read-only": "true", "pp:method": "POST"},
		RunE: func(cmd *cobra.Command, _ []string) error {
			if flags.dryRun {
				return json.NewEncoder(cmd.OutOrStdout()).Encode(map[string]any{"dry_run": true, "operation": "FindFolder", "source": "live", "state_change": false, "offset": offset, "limit": limit})
			}
			current, _, err := currentSession()
			if err != nil {
				return err
			}
			detail, _, err := outlookCapture("OUTLOOK_WEB_FOLDERS_LIST", false)
			if err != nil {
				return err
			}
			session, operation, err := folders.PrepareList(detail, current, offset, limit)
			if err != nil {
				return err
			}
			client, err := transport.New(session.AccountKey, session, nil)
			if err != nil {
				return err
			}
			ctx, cancel := context.WithTimeout(cmd.Context(), 30*time.Second)
			defer cancel()
			response, err := client.Do(ctx, session.AccountKey, operation)
			if err != nil {
				return err
			}
			items, state := exchange.ParseFindFolder(response.Body)
			if state != transport.EnvelopeOK {
				return &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
			}
			pages := make([]map[string]any, 0, len(items))
			for _, item := range items {
				pages = append(pages, map[string]any{"folders": item.Page.Folders, "indexed_paging_offset": item.Page.IndexedPagingOffset,
					"total_items_in_view": item.Page.TotalItemsInView, "includes_last_item_in_range": item.Page.IncludesLastItemInRange})
			}
			return printJSONFilteredKeep(cmd.OutOrStdout(), map[string]any{"source": "live", "pages": pages}, flags, "source", "pages")
		},
	}
	list.Flags().IntVar(&offset, "offset", 0, "Folder page offset")
	list.Flags().IntVar(&limit, "limit", 1000, "Maximum folders in this page (1–1000)")
	parent.AddCommand(list)
	addFolderActions(parent, flags)
	return parent
}
