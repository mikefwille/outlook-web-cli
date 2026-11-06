package cli

import (
	"context"
	"fmt"
	"strings"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/exchange"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/folders"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

// folderLookup walks bounded pages so a name or version check cannot silently
// operate on an item omitted from the first page.
func folderLookup(ctx context.Context, id string) (exchange.Folder, error) {
	current, _, err := currentSession()
	if err != nil {
		return exchange.Folder{}, err
	}
	detail, _, err := outlookCapture("OUTLOOK_WEB_FOLDERS_LIST", false)
	if err != nil {
		return exchange.Folder{}, err
	}
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
			if item.ID.ID == id {
				return item, nil
			}
		}
		if p.IncludesLastItemInRange {
			return exchange.Folder{}, fmt.Errorf("folder not found in the current mailbox hierarchy")
		}
		next := p.IndexedPagingOffset
		if next <= offset {
			return exchange.Folder{}, fmt.Errorf("folder hierarchy paging did not advance")
		}
		offset = next
	}
	return exchange.Folder{}, fmt.Errorf("folder hierarchy exceeds the 10000-folder safety bound")
}

func addFolderActions(parent *cobra.Command, flags *rootFlags) {
	get := &cobra.Command{Use: "get <id>", Short: "Read a folder's name, parent and counts", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live", "pp:happy-args": "id=example-value"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "folders get"})
		}
		id, err := oneID(args)
		if err != nil {
			return err
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "folders get", "id": id})
		}
		folder, err := folderLookup(cmd.Context(), id)
		if err != nil {
			return err
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "folder": folder})
	}}
	parent.AddCommand(get)
	for _, kind := range []string{"create", "rename", "move", "delete"} {
		parent.AddCommand(newFolderActionCmd(flags, kind))
	}
}

func newFolderActionCmd(flags *rootFlags, kind string) *cobra.Command {
	var name, parentID, key string
	var execute bool
	use := kind + " <id>"
	if kind == "create" {
		use = "create"
	}
	cmd := &cobra.Command{Use: use, Short: "Preview or " + kind + " a mail folder", Long: "Preview is the default. With --execute the action is attempted once. Rename, move and delete require the current change key and perform a non-atomic version check. Only ordinary mail folders can be changed; delete moves an empty folder to Deleted Items.", Annotations: map[string]string{"pp:data-source": "live"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "folders " + kind})
		}
		var id string
		var err error
		if kind == "create" {
			if len(args) != 0 {
				return usageErr(fmt.Errorf("folders create takes flags only"))
			}
		} else {
			id, err = oneID(args)
			if err != nil {
				return err
			}
		}
		if (kind == "create" || kind == "rename") && (strings.TrimSpace(name) != name || name == "" || len(name) > 255 || strings.ContainsAny(name, "\x00\r\n")) {
			return usageErr(fmt.Errorf("--name must be 1–255 characters without leading or trailing whitespace"))
		}
		if (kind == "create" || kind == "move") && (parentID == "" || parentID == id) {
			return usageErr(fmt.Errorf("--parent-id must identify a different destination folder"))
		}
		preview := map[string]any{"preview": true, "operation": "folders " + kind, "id": id, "version_check": "before write; not atomic"}
		if kind == "create" || kind == "rename" {
			preview["name"] = name
		}
		if kind == "create" || kind == "move" {
			preview["parent_id"] = parentID
		}
		if kind == "delete" {
			preview["delete_type"] = "move empty folder to Deleted Items"
		}
		if flags.dryRun {
			return flags.printJSON(cmd, preview)
		}
		if kind != "create" {
			current, err := folderLookup(cmd.Context(), id)
			if err != nil {
				return err
			}
			preview["current"] = current
			if current.DistinguishedID != "" || current.FolderClass != "IPF.Note" {
				return fmt.Errorf("only ordinary non-system mail folders can be changed")
			}
			if kind == "delete" && (current.TotalCount == nil || current.ChildCount == nil || *current.TotalCount != 0 || *current.ChildCount != 0) {
				return fmt.Errorf("only a folder with confirmed zero items and zero child folders can be deleted")
			}
			if !execute {
				return flags.printJSON(cmd, preview)
			}
			if key == "" {
				return usageErr(fmt.Errorf("--change-key is required for execution"))
			}
			if current.ID.ChangeKey == "" || key != current.ID.ChangeKey {
				return &transport.Error{Kind: transport.KindConflict}
			}
		} else if !execute {
			return flags.printJSON(cmd, preview)
		}
		if err := mutationInHarness(); err != nil {
			return err
		}
		prefix := map[string]string{"create": "OUTLOOK_WEB_FOLDER_CREATE", "rename": "OUTLOOK_WEB_FOLDER_RENAME", "move": "OUTLOOK_WEB_FOLDER_MOVE", "delete": "OUTLOOK_WEB_FOLDER_DELETE"}[kind]
		detail, _, err := outlookCapture(prefix, false)
		if err != nil {
			return err
		}
		current, _, err := currentSession()
		if err != nil {
			return err
		}
		s, op, err := folders.Prepare(kind, detail, current, folders.Params{ID: id, ParentID: parentID, Name: name})
		if err != nil {
			return err
		}
		response, err := outlookOperation(cmd.Context(), s, op)
		if err != nil {
			return err
		}
		result, state := folders.ParseResult(kind, response.Body)
		if state != transport.EnvelopeOK {
			return &transport.Error{Kind: transport.KindAmbiguous, Status: response.Status}
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "accepted": true, "operation": "folders " + kind, "previous_id": id, "result": result})
	}}
	if kind == "create" || kind == "rename" {
		cmd.Flags().StringVar(&name, "name", "", "Folder display name")
	}
	if kind == "create" || kind == "move" {
		cmd.Flags().StringVar(&parentID, "parent-id", "", "Destination parent folder ID")
	}
	if kind != "create" {
		cmd.Flags().StringVar(&key, "change-key", "", "Expected folder version from preview")
	}
	cmd.Flags().BoolVar(&execute, "execute", false, "Apply the previewed action once")
	return cmd
}
