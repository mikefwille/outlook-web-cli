// Local aliases for already activated, account-bound Outlook session exports.
package cli

import (
	"encoding/hex"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"syscall"
	"time"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/capture"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

var outlookAccountAlias = regexp.MustCompile(`^[a-z0-9][a-z0-9_-]{0,63}$`)
var outlookAccountHash = regexp.MustCompile(`^[a-f0-9]{64}$`)

type outlookAccountProfile struct {
	Alias        string `json:"alias"`
	ExportID     string `json:"export_id"`
	AccountKey   string `json:"account_key"`
	Origin       string `json:"origin"`
	AudienceHash string `json:"audience_hash"`
	AnchorHash   string `json:"anchor_hash"`
}
type outlookAccountRegistry struct {
	Version  int                     `json:"version"`
	Service  string                  `json:"service"`
	Profiles []outlookAccountProfile `json:"profiles"`
}

func init() {
	registerNovelCommand(func(root *cobra.Command, flags *rootFlags) {
		root.AddCommand(newOutlookAccountsCmd(flags))
	})
}

func isAccountHash(s string) bool {
	if len(s) != 64 {
		return false
	}
	_, err := hex.DecodeString(s)
	return err == nil && outlookAccountHash.MatchString(s)
}

func readAccountRegistry(pointer string) (outlookAccountRegistry, error) {
	registry := outlookAccountRegistry{Version: 1, Service: "outlook", Profiles: []outlookAccountProfile{}}
	path := filepath.Join(filepath.Dir(pointer), "outlook-accounts-v1.json")
	if _, err := os.Lstat(path); os.IsNotExist(err) {
		return registry, nil
	}
	data, err := capture.ReadPrivateEvidence(path)
	if err != nil || json.Unmarshal(data, &registry) != nil || registry.Version != 1 || registry.Service != "outlook" || len(registry.Profiles) > 128 {
		return outlookAccountRegistry{}, fmt.Errorf("private Outlook account registry is invalid or unavailable")
	}
	seen := map[string]bool{}
	for _, p := range registry.Profiles {
		if !outlookAccountAlias.MatchString(p.Alias) || seen[p.Alias] || !outlookExportID.MatchString(p.ExportID) || p.Origin != "https://outlook.cloud.microsoft" || !isAccountHash(p.AccountKey) || !isAccountHash(p.AudienceHash) || !isAccountHash(p.AnchorHash) {
			return outlookAccountRegistry{}, fmt.Errorf("private Outlook account registry is invalid or unavailable")
		}
		seen[p.Alias] = true
	}
	if registry.Profiles == nil {
		registry.Profiles = []outlookAccountProfile{}
	}
	return registry, nil
}

func currentExportID(pointer string) (string, error) {
	present, err := safeOutlookPointer(pointer)
	if err != nil || !present {
		return "", fmt.Errorf("active Outlook session pointer is unavailable")
	}
	data, err := capture.ReadPrivateEvidence(pointer)
	var p struct {
		Version  int    `json:"version"`
		Service  string `json:"service"`
		ExportID string `json:"export_id"`
	}
	if err != nil || json.Unmarshal(data, &p) != nil || p.Version != 1 || p.Service != "outlook" || !outlookExportID.MatchString(p.ExportID) {
		return "", fmt.Errorf("active Outlook session pointer is unavailable")
	}
	return p.ExportID, nil
}

func accountProfile(alias, id string, session transport.Session) (outlookAccountProfile, capture.RuntimeBundle, error) {
	bundle, err := capture.ReadBoundRuntimeBundle(session)
	if err != nil {
		return outlookAccountProfile{}, capture.RuntimeBundle{}, fmt.Errorf("account session and private operation bundle do not match")
	}
	// Both basic operations are required; checking just a bundle header would
	// permit switching to a profile that cannot make its basic mailbox reads.
	if _, _, err := capture.LoadRuntimeWeb("folders_list", session); err != nil {
		return outlookAccountProfile{}, capture.RuntimeBundle{}, fmt.Errorf("account folder template is unavailable")
	}
	if _, _, err := capture.PrepareRuntimeRead("mail_list", session); err != nil {
		return outlookAccountProfile{}, capture.RuntimeBundle{}, fmt.Errorf("account mail template is unavailable")
	}
	return outlookAccountProfile{Alias: alias, ExportID: id, AccountKey: bundle.AccountKey, Origin: session.Origin, AudienceHash: bundle.AudienceHash, AnchorHash: bundle.AnchorHash}, bundle, nil
}

func sameAccountBinding(a, b outlookAccountProfile) bool {
	return a.AccountKey == b.AccountKey && a.Origin == b.Origin && a.AudienceHash == b.AudienceHash && a.AnchorHash == b.AnchorHash
}

func loadCurrentAccountProfile(pointer, alias string) (outlookAccountProfile, capture.RuntimeBundle, error) {
	id, err := currentExportID(pointer)
	if err != nil {
		return outlookAccountProfile{}, capture.RuntimeBundle{}, err
	}
	s, _, err := capture.LoadCurrentFindFolder(pointer)
	if err != nil {
		return outlookAccountProfile{}, capture.RuntimeBundle{}, fmt.Errorf("active Outlook session is unavailable or expired")
	}
	if after, err := currentExportID(pointer); err != nil || after != id {
		return outlookAccountProfile{}, capture.RuntimeBundle{}, fmt.Errorf("active Outlook session changed while checking")
	}
	return accountProfile(alias, id, s)
}

// Atomic writes are always inside a validated private directory and never
// follow an existing unsafe destination. The caller holds the helper's lock.
func writeOutlookPrivateAtomic(path string, data []byte) error {
	info, err := os.Lstat(path)
	if err == nil {
		if !info.Mode().IsRegular() || info.Mode().Perm()&0077 != 0 {
			return fmt.Errorf("private account state destination is unsafe")
		}
		owner, ok := info.Sys().(*syscall.Stat_t)
		if !ok || owner.Uid != uint32(os.Getuid()) {
			return fmt.Errorf("private account state destination is unsafe")
		}
	} else if !os.IsNotExist(err) {
		return fmt.Errorf("private account state destination is unavailable")
	}
	f, err := os.CreateTemp(filepath.Dir(path), ".outlook-account-*")
	if err != nil {
		return fmt.Errorf("cannot save private account state")
	}
	tmp := f.Name()
	defer os.Remove(tmp)
	if err = f.Chmod(0600); err == nil {
		_, err = f.Write(data)
	}
	if err == nil {
		err = f.Sync()
	}
	if closeErr := f.Close(); err == nil {
		err = closeErr
	}
	if err != nil {
		return fmt.Errorf("cannot save private account state")
	}
	if err = os.Rename(tmp, path); err != nil {
		return fmt.Errorf("cannot atomically save private account state")
	}
	return nil
}

func withOutlookAccountLock(pointer string, f func() error) error {
	lock := filepath.Join(filepath.Dir(pointer), ".m365-session-outlook.lock")
	if err := os.Mkdir(lock, 0700); err != nil {
		return fmt.Errorf("Outlook session is locked or its private directory is unavailable")
	}
	defer os.Remove(lock)
	if _, err := safeOutlookPointer(pointer); err != nil {
		return err
	}
	return f()
}

func findAccountProfile(r outlookAccountRegistry, alias string) (outlookAccountProfile, bool) {
	for _, p := range r.Profiles {
		if p.Alias == alias {
			return p, true
		}
	}
	return outlookAccountProfile{}, false
}

func registerCurrentAccount(pointer, alias string) error {
	return withOutlookAccountLock(pointer, func() error {
		p, bundle, err := loadCurrentAccountProfile(pointer, alias)
		if err != nil {
			return err
		}
		registry, err := readAccountRegistry(pointer)
		if err != nil {
			return err
		}
		found := false
		for i, existing := range registry.Profiles {
			if existing.Alias == alias {
				if !sameAccountBinding(existing, p) {
					return fmt.Errorf("alias already belongs to a different account, audience or mailbox")
				}
				registry.Profiles[i], found = p, true
			}
		}
		if !found {
			if len(registry.Profiles) >= 128 {
				return fmt.Errorf("private account registry is full")
			}
			registry.Profiles = append(registry.Profiles, p)
		}
		s, _, err := capture.LoadCurrentFindFolder(pointer)
		if err != nil {
			return fmt.Errorf("active Outlook session changed")
		}
		bundlePath, err := capture.RuntimeBundlePathForAccount(pointer, s)
		if err != nil {
			return err
		}
		bundleData, err := json.Marshal(bundle)
		if err != nil {
			return fmt.Errorf("cannot encode private account state")
		}
		// The snapshot contains no session headers and is private. A failed
		// registry write leaves only an unused account-bound bundle snapshot.
		if err = writeOutlookPrivateAtomic(bundlePath, append(bundleData, '\n')); err != nil {
			return err
		}
		data, err := json.Marshal(registry)
		if err != nil {
			return fmt.Errorf("cannot encode private account state")
		}
		return writeOutlookPrivateAtomic(filepath.Join(filepath.Dir(pointer), "outlook-accounts-v1.json"), append(data, '\n'))
	})
}

// accountTarget checks both the saved export and its account-bound templates.
// If a helper has refreshed the currently active account, it is already the
// desired selection even if the old registered export has expired.
func accountTarget(pointer string, p outlookAccountProfile) (outlookAccountProfile, bool, error) {
	if current, _, err := loadCurrentAccountProfile(pointer, p.Alias); err == nil && sameAccountBinding(current, p) {
		return current, true, nil
	}
	s, _, err := capture.LoadFindFolder(filepath.Join(filepath.Dir(pointer), "session-outlook-"+p.ExportID+".json"))
	if err != nil {
		return outlookAccountProfile{}, false, fmt.Errorf("registered account session is unavailable or expired")
	}
	candidate, _, err := accountProfile(p.Alias, p.ExportID, s)
	if err != nil || !sameAccountBinding(candidate, p) {
		return outlookAccountProfile{}, false, fmt.Errorf("registered account, origin, audience or mailbox no longer matches")
	}
	return candidate, false, nil
}

func selectOutlookAccount(pointer, alias string) (bool, error) {
	changed := false
	err := withOutlookAccountLock(pointer, func() error {
		registry, err := readAccountRegistry(pointer)
		if err != nil {
			return err
		}
		p, found := findAccountProfile(registry, alias)
		if !found {
			return fmt.Errorf("unknown Outlook account alias")
		}
		target, active, err := accountTarget(pointer, p)
		if err != nil || active {
			return err
		}
		data, err := json.Marshal(map[string]any{"version": 1, "service": "outlook", "export_id": target.ExportID, "activated_at": time.Now().UTC().Format(time.RFC3339Nano)})
		if err != nil {
			return fmt.Errorf("cannot encode account selection")
		}
		if err = writeOutlookPrivateAtomic(pointer, append(data, '\n')); err != nil {
			return err
		}
		changed = true
		return nil
	})
	return changed, err
}

// pp:data-source local
func newOutlookAccountsCmd(flags *rootFlags) *cobra.Command {
	group := &cobra.Command{Use: "accounts", Short: "Select registered Outlook session profiles"}
	list := &cobra.Command{Use: "list", Short: "List registered Outlook account aliases", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "local"}, RunE: func(cmd *cobra.Command, args []string) error {
		if len(args) != 0 {
			return usageErr(fmt.Errorf("accounts list takes no arguments"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "accounts list"})
		}
		pointer, err := outlookPointerPath()
		if err != nil {
			return err
		}
		if _, err = safeOutlookPointer(pointer); err != nil {
			return err
		}
		registry, err := readAccountRegistry(pointer)
		if err != nil {
			return err
		}
		sort.Slice(registry.Profiles, func(i, j int) bool { return registry.Profiles[i].Alias < registry.Profiles[j].Alias })
		items := make([]map[string]any, 0, len(registry.Profiles))
		for _, p := range registry.Profiles {
			_, active, targetErr := accountTarget(pointer, p)
			state := "available"
			if targetErr != nil {
				state = "unavailable_or_expired"
			}
			items = append(items, map[string]any{"alias": p.Alias, "active": active, "state": state})
		}
		return flags.printJSON(cmd, map[string]any{"source": "local", "accounts": items, "network_checked": false})
	}}
	group.AddCommand(list)
	for _, operation := range []string{"use", "register-current"} {
		operation := operation
		var execute bool
		cmd := &cobra.Command{Use: operation + " <alias>", Short: "Preview or " + operation + " an imported Outlook account", Long: "Operates only on private, previously activated Outlook session exports and matching account-bound templates. Preview is the default; --execute changes local state once. It does not sign in, create accounts, or grant permissions.", Annotations: map[string]string{"pp:data-source": "local", "pp:happy-args": "alias=work;--dry-run"}, RunE: func(cmd *cobra.Command, args []string) error {
			if flags.dryRun && requiredInputWaived() {
				return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "accounts " + operation})
			}
			alias, err := oneID(args)
			if err != nil || !outlookAccountAlias.MatchString(alias) {
				return usageErr(fmt.Errorf("supply one alias of 1–64 lowercase letters, digits, underscores or hyphens"))
			}
			if flags.dryRun {
				return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "accounts " + operation, "alias": alias})
			}
			pointer, err := outlookPointerPath()
			if err != nil {
				return err
			}
			if _, err = safeOutlookPointer(pointer); err != nil {
				return err
			}
			registry, err := readAccountRegistry(pointer)
			if err != nil {
				return err
			}
			active := false
			if operation == "register-current" {
				candidate, _, err := loadCurrentAccountProfile(pointer, alias)
				if err != nil {
					return err
				}
				if prior, ok := findAccountProfile(registry, alias); ok && !sameAccountBinding(prior, candidate) {
					return fmt.Errorf("alias already belongs to a different account, audience or mailbox")
				}
				active = true
			} else {
				p, ok := findAccountProfile(registry, alias)
				if !ok {
					return usageErr(fmt.Errorf("unknown Outlook account alias"))
				}
				_, active, err = accountTarget(pointer, p)
				if err != nil {
					return err
				}
			}
			if !execute {
				return flags.printJSON(cmd, map[string]any{"source": "local", "preview": true, "operation": "accounts " + operation, "alias": alias, "already_active": active, "network_checked": false})
			}
			if err := mutationInHarness(); err != nil {
				return err
			}
			changed := true
			if operation == "register-current" {
				err = registerCurrentAccount(pointer, alias)
			} else {
				changed, err = selectOutlookAccount(pointer, alias)
			}
			if err != nil {
				return err
			}
			return flags.printJSON(cmd, map[string]any{"source": "local", "operation": "accounts " + operation, "alias": alias, "changed": changed, "network_checked": false})
		}}
		cmd.Flags().BoolVar(&execute, "execute", false, "Apply this local account operation once")
		group.AddCommand(cmd)
	}
	return group
}
