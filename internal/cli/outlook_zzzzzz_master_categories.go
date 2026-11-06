package cli

import (
	"context"
	"crypto/rand"
	"fmt"
	"os"
	"strings"
	"time"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/mastercategories"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func categoryUUID() (string, error) {
	var b [16]byte
	if _, err := rand.Read(b[:]); err != nil {
		return "", err
	}
	b[6] = (b[6] & 0x0f) | 0x40
	b[8] = (b[8] & 0x3f) | 0x80
	return fmt.Sprintf("%x-%x-%x-%x-%x", b[0:4], b[4:6], b[6:8], b[8:10], b[10:16]), nil
}
func categoryOperation(kind string, current transport.Session, p mastercategories.Params) (transport.Session, transport.Operation, error) {
	name := "master_category_" + strings.ReplaceAll(kind, "-", "_")
	prefix := "OUTLOOK_WEB_" + strings.ToUpper(name)
	detail := os.Getenv(prefix + "_DETAILS")
	if os.Getenv(prefix+"_REQUEST") != "" {
		return transport.Session{}, transport.Operation{}, fmt.Errorf("master category template uses a header body")
	}
	if detail == "" {
		detail = "runtime:" + name
	}
	return mastercategories.Prepare(kind, detail, current, p)
}
func categoryList(ctx context.Context) (mastercategories.Snapshot, error) {
	data, err := readOutlookConfiguration(ctx)
	if err != nil {
		return mastercategories.Snapshot{}, err
	}
	return mastercategories.ParseConfiguration(data)
}

// pp:data-source live
func init() {
	registerNovelCommand(func(root *cobra.Command, flags *rootFlags) {
		group := findSubcommand(root, "categories")
		if group == nil {
			return
		}
		for _, kind := range []string{"create", "update", "delete"} {
			group.AddCommand(newMasterCategoryAction(flags, kind))
		}
	})
}
func newMasterCategoryAction(flags *rootFlags, kind string) *cobra.Command {
	var name, version string
	var color int
	var execute, allowItems bool
	use := kind + " <id>"
	if kind == "create" {
		use = "create"
	}
	cmd := &cobra.Command{Use: use, Short: "Preview or " + kind + " a master category", Long: "Create adds one category. Update supports the observed rename-and-color change; Outlook assigns a replacement ID. Rename and delete also initiate a separate category-wide operation that may change items using the old name, and completion is not established by acceptance. Execution requires a fresh list fingerprint; the precheck is not atomic. If either write is uncertain, inspect the mailbox before any retry.", Annotations: map[string]string{"pp:data-source": "live", "pp:happy-args": "id=example-value"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "categories " + kind})
		}
		id := ""
		if kind == "create" {
			if len(args) != 0 {
				return usageErr(fmt.Errorf("categories create takes flags only"))
			}
		} else {
			var err error
			id, err = oneID(args)
			if err != nil {
				return err
			}
		}
		if kind == "create" || kind == "update" {
			if !ruleInput(name) || !cmd.Flags().Changed("color") || color < 0 || color > 24 {
				return usageErr(fmt.Errorf("--name and a --color value from 0 to 24 are required"))
			}
		}
		preview := map[string]any{"preview": true, "operation": "categories " + kind, "id": id, "version_check": "entire master list before write; not atomic"}
		if kind == "create" || kind == "update" {
			preview["new_name"], preview["new_color"] = name, color
		}
		if kind != "create" {
			preview["item_effects"] = "Outlook may rename or remove this category across items; completion is not established by acceptance"
			preview["requires_allow_affect_items"] = true
		}
		if flags.dryRun {
			return flags.printJSON(cmd, preview)
		}
		before, err := categoryList(cmd.Context())
		if err != nil {
			return err
		}
		preview["version"] = before.Version
		var selected *mastercategories.Category
		for _, c := range before.Categories {
			if c.ID == id {
				copy := c
				selected = &copy
			}
			if (kind == "create" || kind == "update") && strings.EqualFold(c.Name, name) {
				return fmt.Errorf("the new category name already exists")
			}
		}
		if kind != "create" {
			if selected == nil {
				return fmt.Errorf("category not found in the current mailbox")
			}
			if kind == "update" && !selected.KnownFieldsOnly {
				return fmt.Errorf("category contains unrecognized fields; update refused to avoid dropping them")
			}
			preview["current"] = selected
		}
		if !execute {
			return flags.printJSON(cmd, preview)
		}
		if !ruleVersionPattern.MatchString(version) || version != before.Version {
			return &transport.Error{Kind: transport.KindConflict}
		}
		if kind != "create" && !allowItems {
			return usageErr(fmt.Errorf("--allow-affect-items is required for the category-wide operation"))
		}
		if err = mutationInHarness(); err != nil {
			return err
		}
		p := mastercategories.Params{}
		if selected != nil {
			p.OldName = selected.Name
		}
		if kind == "create" || kind == "update" {
			newID, e := categoryUUID()
			if e != nil {
				return e
			}
			last := time.Now().UTC().Format("2006-01-02T15:04:05.000Z")
			shortcut := 0
			if selected != nil {
				last = selected.LastTimeUsed
				shortcut = selected.KeyboardShortcut
			}
			p.New = mastercategories.Category{Name: name, Color: color, ID: newID, LastTimeUsed: last, KeyboardShortcut: shortcut}
		}
		current, _, err := currentSession()
		if err != nil {
			return err
		}
		firstSession, firstOp, err := categoryOperation(kind, current, p)
		if err != nil {
			return err
		}
		var nextSession transport.Session
		var nextOp transport.Operation
		if kind != "create" {
			nextKind := "delete-items"
			if kind == "update" {
				nextKind = "rename-items"
			}
			nextSession, nextOp, err = categoryOperation(nextKind, current, p)
			if err != nil {
				return err
			}
		}
		response, err := outlookOperation(cmd.Context(), firstSession, firstOp)
		if err != nil {
			return err
		}
		after, err := mastercategories.ParseResponse(response.Body)
		if err != nil || !mastercategories.VerifyDelta(before, after, kind, p) {
			return fmt.Errorf("master category write was accepted but its result could not be reconciled; inspect before retrying")
		}
		if kind != "create" {
			if _, err = outlookOperation(cmd.Context(), nextSession, nextOp); err != nil {
				return fmt.Errorf("master category list changed but the category-wide operation was not confirmed; inspect before retrying")
			}
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "accepted": true, "operation": "categories " + kind, "list_delta_verified": true, "category_wide_operation_accepted": kind != "create", "completion": "check categories and affected items independently; do not retry an uncertain operation"})
	}}
	if kind == "create" || kind == "update" {
		cmd.Flags().StringVar(&name, "name", "", "New unique category name")
		cmd.Flags().IntVar(&color, "color", -1, "Outlook category color index (0–24)")
	}
	cmd.Flags().StringVar(&version, "if-version", "", "Master-list fingerprint returned by preview")
	if kind != "create" {
		cmd.Flags().BoolVar(&allowItems, "allow-affect-items", false, "Acknowledge possible category changes across items")
	}
	cmd.Flags().BoolVar(&execute, "execute", false, "Apply the previewed operation once")
	return cmd
}
