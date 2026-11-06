package cli

import (
	"context"
	"fmt"
	"os"
	"strings"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/contactlists"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/contacts"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func init() {
	registerNovelCommand(func(root *cobra.Command, flags *rootFlags) {
		group := findSubcommand(root, "contacts")
		if group == nil {
			group = &cobra.Command{Use: "contacts", Short: "Work with personal contacts"}
			root.AddCommand(group)
		}
		group.AddCommand(newContactListsCommand(flags))
		direct := newContactListsCommand(flags)
		direct.Use = "contact-lists"
		root.AddCommand(direct)
	})
}

func contactListRequest(ctx context.Context, action string, p contactlists.Params) (transport.Response, error) {
	name := map[string]string{"list": "contact_lists_list", "create": "contact_lists_create", "update": "contact_lists_update", "add-member": "contact_lists_add_member", "remove-member": "contact_lists_remove_member", "delete": "contact_lists_delete"}[action]
	if name == "" {
		return transport.Response{}, fmt.Errorf("unknown contact list operation")
	}
	prefix := "OUTLOOK_WEB_" + strings.ToUpper(name)
	detail, request := os.Getenv(prefix+"_DETAILS"), os.Getenv(prefix+"_REQUEST")
	if detail == "" && request == "" {
		detail = "runtime:" + name
	}
	if detail == "" || ((action == "list" || action == "delete") && request != "") || ((action != "list" && action != "delete") && !strings.HasPrefix(detail, "runtime:") && request == "") {
		return transport.Response{}, fmt.Errorf("contact list operation template is incomplete: %s", name)
	}
	current, _, err := currentSession()
	if err != nil {
		return transport.Response{}, err
	}
	s, op, err := contactlists.Prepare(action, detail, request, current, p)
	if err != nil {
		return transport.Response{}, err
	}
	return outlookOperation(ctx, s, op)
}

func readContactListPage(ctx context.Context) (contactlists.Page, error) {
	response, err := contactListRequest(ctx, "list", contactlists.Params{})
	if err != nil {
		return contactlists.Page{}, err
	}
	page, state := contactlists.ParsePage(response.Body)
	if state != transport.EnvelopeOK {
		return contactlists.Page{}, &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
	}
	return page, nil
}

func readContactList(ctx context.Context, id string) (contactlists.Item, error) {
	if !contactlists.ValidID(id) {
		return contactlists.Item{}, usageErr(fmt.Errorf("list ID has an invalid format"))
	}
	page, err := readContactListPage(ctx)
	if err != nil {
		return contactlists.Item{}, err
	}
	for _, item := range page.Items {
		if item.ID == id {
			return item, nil
		}
	}
	if !page.Complete {
		return contactlists.Item{}, fmt.Errorf("list not found in the first 500 results; the service result may be incomplete")
	}
	return contactlists.Item{}, fmt.Errorf("contact list not found in the current mailbox")
}

// pp:data-source live
func newContactListsCommand(flags *rootFlags) *cobra.Command {
	group := &cobra.Command{Use: "lists", Short: "Work with personal contact lists", Long: "Manage personal contact lists. Reads inspect the observed first 500 lists and report whether that page is complete. No email or invitation is sent."}
	var limit int
	list := &cobra.Command{Use: "list", Short: "List personal contact lists", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live"}, RunE: func(cmd *cobra.Command, args []string) error {
		if len(args) != 0 || limit < 1 || limit > contactlists.ServerLimit {
			return usageErr(fmt.Errorf("list takes no arguments; --limit must be between 1 and 500"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "contacts lists list", "limit": limit})
		}
		page, err := readContactListPage(cmd.Context())
		if err != nil {
			return err
		}
		count := len(page.Items)
		rows := page.Items
		if len(rows) > limit {
			rows = rows[:limit]
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "results": rows, "returned": len(rows), "fetched": count, "server_limit": contactlists.ServerLimit, "complete": page.Complete && count <= limit, "more_in_fetched_page": count > limit, "server_continuation_present": page.ContinuationPresent, "note": "Use --limit 500 to view the full fetched page; server continuations are not followed."})
	}}
	list.Flags().IntVar(&limit, "limit", 100, "Maximum lists to return from the observed first page (1–500)")
	get := &cobra.Command{Use: "get <id>", Short: "Read a personal list and its members", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live", "pp:happy-args": "id=example-value"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "contacts lists get"})
		}
		id, err := oneID(args)
		if err != nil {
			return err
		}
		if !contactlists.ValidID(id) {
			return usageErr(fmt.Errorf("list ID has an invalid format"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "contacts lists get", "id": id})
		}
		item, err := readContactList(cmd.Context(), id)
		if err != nil {
			return err
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "list": item})
	}}
	group.AddCommand(list, get)
	for _, action := range []string{"create", "update", "add-member", "remove-member", "delete"} {
		group.AddCommand(newContactListAction(flags, action))
	}
	return group
}

func newContactListAction(flags *rootFlags, action string) *cobra.Command {
	var name, notes, email, memberName, etag string
	var execute bool
	use := action + " <id>"
	if action == "create" {
		use = "create"
	}
	cmd := &cobra.Command{Use: use, Short: "Preview or " + action + " a personal contact list", Long: "Preview is the default; --execute applies one request. For changes to an existing list, --etag must match a fresh read. This pre-write version check is not atomic. Membership currently supports OneOff SMTP addresses. A list stores addresses and does not send a message.", Annotations: map[string]string{"pp:data-source": "live", "pp:happy-args": "id=example-value"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "contacts lists " + action})
		}
		id := ""
		if action == "create" {
			if len(args) != 0 {
				return usageErr(fmt.Errorf("create takes flags only"))
			}
		} else {
			var err error
			id, err = oneID(args)
			if err != nil {
				return err
			}
			if !contactlists.ValidID(id) {
				return usageErr(fmt.Errorf("list ID has an invalid format"))
			}
		}
		if (action == "create" || cmd.Flags().Changed("name")) && !contactlists.ValidName(name) {
			return usageErr(fmt.Errorf("--name must be 1–255 characters without leading or trailing whitespace"))
		}
		if cmd.Flags().Changed("notes") && !contactlists.ValidNotes(notes) {
			return usageErr(fmt.Errorf("--notes is too long or contains a NUL character"))
		}
		if action == "update" && !cmd.Flags().Changed("name") && !cmd.Flags().Changed("notes") {
			return usageErr(fmt.Errorf("provide --name or --notes"))
		}
		if action == "create" || action == "add-member" || action == "remove-member" {
			if !contacts.ValidEmail(email) {
				return usageErr(fmt.Errorf("--email must be one complete email address"))
			}
			if memberName == "" {
				memberName = email
			}
			if action != "remove-member" && !contactlists.ValidName(memberName) {
				return usageErr(fmt.Errorf("--member-name must be 1–255 characters"))
			}
		}
		preview := map[string]any{"preview": true, "operation": "contacts lists " + action, "id": id}
		if action != "create" {
			preview["version_check"] = "before write; not atomic"
		}
		if action == "create" || cmd.Flags().Changed("name") {
			preview["name"] = name
		}
		if cmd.Flags().Changed("notes") {
			preview["notes"] = notes
		}
		if email != "" {
			preview["email"] = email
		}
		if action == "create" || action == "add-member" {
			preview["member_name"] = memberName
		}
		if flags.dryRun {
			return flags.printJSON(cmd, preview)
		}
		p := contactlists.Params{ID: id, Name: name, Notes: notes, Email: email, MemberName: memberName}
		if action != "create" {
			current, err := readContactList(cmd.Context(), id)
			if err != nil {
				return err
			}
			preview["current"] = current
			if action == "add-member" || action == "remove-member" {
				if err = contactlists.CheckMember(current, email, action == "add-member"); err != nil {
					return err
				}
			}
			if !cmd.Flags().Changed("name") {
				p.Name = current.Name
			}
			if !cmd.Flags().Changed("notes") {
				p.Notes = current.Notes
			}
			if !execute {
				return flags.printJSON(cmd, preview)
			}
			if etag == "" {
				return usageErr(fmt.Errorf("--etag is required with --execute; use the value from a fresh preview"))
			}
			if current.ETag != etag {
				return &transport.Error{Kind: transport.KindConflict}
			}
		} else if !execute {
			return flags.printJSON(cmd, preview)
		}
		if err := mutationInHarness(); err != nil {
			return err
		}
		response, err := contactListRequest(cmd.Context(), action, p)
		if err != nil {
			return err
		}
		if action == "delete" {
			if response.Status != 204 {
				return &transport.Error{Kind: transport.KindAmbiguous, Status: response.Status}
			}
			return flags.printJSON(cmd, map[string]any{"source": "live", "accepted": true, "operation": "contacts lists delete", "id": id, "next_step": "Read the list page independently to verify deletion."})
		}
		item, state := contactlists.ParseItem(response.Body)
		if state != transport.EnvelopeOK {
			return &transport.Error{Kind: transport.KindAmbiguous, Status: response.Status}
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "accepted": true, "operation": "contacts lists " + action, "list": item, "next_step": "Read the list again independently before another change."})
	}}
	if action == "create" {
		cmd.Annotations["pp:happy-args"] = "--name=Example list;--email=member@example.invalid"
	}
	if action == "create" || action == "update" {
		cmd.Flags().StringVar(&name, "name", "", "Personal list display name")
	}
	if action == "update" {
		cmd.Flags().StringVar(&notes, "notes", "", "Replacement notes; an empty value clears them")
	}
	if action == "create" || action == "add-member" || action == "remove-member" {
		cmd.Flags().StringVar(&email, "email", "", "OneOff SMTP address of the member")
	}
	if action == "create" || action == "add-member" {
		cmd.Flags().StringVar(&memberName, "member-name", "", "Member display name; defaults to the email address")
	}
	if action != "create" {
		cmd.Flags().StringVar(&etag, "etag", "", "Expected version from a fresh list read")
	}
	cmd.Flags().BoolVar(&execute, "execute", false, "Apply the previewed action once")
	return cmd
}
