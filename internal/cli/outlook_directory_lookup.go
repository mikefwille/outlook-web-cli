package cli

import (
	"encoding/json"
	"fmt"

	"github.com/spf13/cobra"
	people "github.com/mikefwille/outlook-web-cli/internal/outlook/contacts"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/directory"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

// Both paths expose the same observed exact-email People lookup. Results can
// contain personal contacts as well as directory profiles.
func init() {
	registerNovelCommand(func(root *cobra.Command, flags *rootFlags) {
		for _, entry := range []struct{ group, command, short string }{
			{"directory", "lookup", "Look up an exact email in People and directory results"},
			{"recipients", "resolve", "Resolve one exact email using People lookup"},
		} {
			parent := findSubcommand(root, entry.group)
			if parent == nil {
				parent = &cobra.Command{Use: entry.group, Short: entry.short}
				root.AddCommand(parent)
			}
			contacts := newOutlookPersonalContacts(flags)
			lookup := findSubcommand(contacts, "lookup-email")
			contacts.RemoveCommand(lookup)
			lookup.Use = entry.command
			lookup.Short = entry.short
			parent.AddCommand(lookup)
		}
		findSubcommand(root, "directory").AddCommand(newDirectoryGetCmd(flags), newDirectoryOrgCmd(flags, "manager"), newDirectoryOrgCmd(flags, "reports"))
	})
}

// pp:data-source live
func newDirectoryOrgCmd(flags *rootFlags, kind string) *cobra.Command {
	var objectID, email, upn string
	cmd := &cobra.Command{Use: kind, Short: "Read a profile's observed " + kind + " relationship", Long: "Runs the captured profile-card organization query with an independently scoped, locally captured directory credential. Supply all three exact profile identifiers. Empty returned relationships may reflect unavailable organization data. Reports are bounded to the first 20 and indicate when more exist.", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live", "pp:happy-args": "--object-id=00000000-0000-0000-0000-000000000001;--email=person@example.invalid;--upn=person@example.invalid;--dry-run"}, RunE: func(cmd *cobra.Command, args []string) error {
		if len(args) != 0 {
			return usageErr(fmt.Errorf("directory %s accepts flags only", kind))
		}
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "directory " + kind})
		}
		id := directory.Identity{ObjectID: objectID, Email: email, UPN: upn}
		if !directory.ValidIdentity(id) {
			return usageErr(fmt.Errorf("--object-id, --email and --upn are required exact profile identifiers"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "directory " + kind, "object_id": objectID, "email": email, "upn": upn, "reports_limit": 20})
		}
		pointer, err := outlookPointerPath()
		if err != nil {
			return err
		}
		current, _, err := currentSession()
		if err != nil {
			return err
		}
		session, err := directory.Load(pointer, current)
		if err != nil {
			return fmt.Errorf("directory profile-card session unavailable, expired, or for another account; refresh its observed private capture")
		}
		ctx, cancel := boundCtx(cmd.Context(), flags)
		defer cancel()
		result, err := directory.Fetch(ctx, current, session, id, nil)
		if err != nil {
			return err
		}
		response := map[string]any{"source": "live", "scope": "profile-card organization query", "results": result.Manager, "note": "Empty returned relationships can reflect unavailable organization data."}
		if kind == "reports" {
			response["results"] = result.Reports
			response["limit"] = 20
			response["has_next_page"] = result.HasNextPage
			response["complete"] = !result.HasNextPage
			if result.HasNextPage {
				response["note"] = "More direct reports exist; this captured query returns only the first 20."
			}
		}
		return flags.printJSON(cmd, response)
	}}
	cmd.Flags().StringVar(&objectID, "object-id", "", "Directory object UUID for the requested profile")
	cmd.Flags().StringVar(&email, "email", "", "Exact SMTP address of the requested profile")
	cmd.Flags().StringVar(&upn, "upn", "", "Exact user principal name of the requested profile")
	return cmd
}

// pp:data-source live
func newDirectoryGetCmd(flags *rootFlags) *cobra.Command {
	return &cobra.Command{Use: "get <people-id>", Short: "Read a People profile by its People ID", Long: "Reads the observed People profile-card lookup by People ID. A returned profile can combine directory and personal sources; it is not a manager, direct-reports, or photo query. Obtain the People ID from a lookup result.", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live", "pp:happy-args": "people-id=example-value"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "directory get"})
		}
		id, err := oneID(args)
		if err != nil {
			return err
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "directory get", "people_id": id})
		}
		response, err := contactRequest(cmd.Context(), "get", people.Params{ID: id})
		if err != nil {
			return err
		}
		profile, state := people.ParseProfile(response.Body, id)
		if state != transport.EnvelopeOK {
			return &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
		}
		if profile == nil {
			return fmt.Errorf("People record has no observed Profile source")
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "scope": "People profile; sources may be unified", "profile": json.RawMessage(profile)})
	}}
}
