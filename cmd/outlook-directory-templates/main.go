// Parent-run helper: import one already observed profile-card organization
// request and successful response into an account-bound private session file.
// It performs no browser or network operations and emits no credentials.
package main

import (
	"encoding/json"
	"errors"
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"syscall"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/capture"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/directory"
)

func run(detail, request, response string) error {
	if detail == "" || request == "" || response == "" || filepath.Dir(filepath.Clean(detail)) != filepath.Dir(filepath.Clean(request)) || filepath.Dir(filepath.Clean(detail)) != filepath.Dir(filepath.Clean(response)) {
		return errors.New("supply_one_same_directory_observed_capture")
	}
	pointer := os.Getenv("OUTLOOK_WEB_SESSION_POINTER")
	if pointer == "" {
		home, err := os.UserHomeDir()
		if err != nil {
			return errors.New("private_session_unavailable")
		}
		pointer = filepath.Join(home, ".local", "share", "outlook-web-dev", "private", "current-outlook.json")
	}
	if filepath.Base(pointer) != "current-outlook.json" || os.Getenv("OUTLOOK_WEB_SESSION_FILE") != "" {
		return errors.New("active_pointer_required")
	}
	private := filepath.Dir(pointer)
	st, err := os.Lstat(private)
	if err != nil || !st.IsDir() || st.Mode()&os.ModeSymlink != 0 || st.Mode().Perm()&0077 != 0 {
		return errors.New("private_directory_unsafe")
	}
	owner, ok := st.Sys().(*syscall.Stat_t)
	if !ok || owner.Uid != uint32(os.Getuid()) {
		return errors.New("private_directory_unsafe")
	}
	lock := filepath.Join(private, ".m365-session-outlook.lock")
	if err := os.Mkdir(lock, 0700); err != nil {
		return errors.New("session_locked")
	}
	defer os.Remove(lock)
	current, _, err := capture.LoadCurrentFindFolder(pointer)
	if err != nil {
		return errors.New("active_session_unavailable")
	}
	d, e1 := capture.ReadPrivateEvidence(detail)
	q, e2 := capture.ReadPrivateEvidence(request)
	r, e3 := capture.ReadPrivateEvidence(response)
	if e1 != nil || e2 != nil || e3 != nil {
		return errors.New("private_capture_unavailable")
	}
	session, err := directory.ExtractObserved(current, d, q, r)
	if err != nil {
		return errors.New("capture_not_bound_to_active_account")
	}
	dest, err := directory.SessionPath(pointer, current)
	if err != nil {
		return errors.New("private_session_path_invalid")
	}
	if info, err := os.Lstat(dest); err == nil {
		if !info.Mode().IsRegular() || info.Mode().Perm()&0077 != 0 {
			return errors.New("directory_session_unsafe")
		}
		if _, err := capture.ReadPrivateEvidence(dest); err != nil {
			return errors.New("directory_session_unsafe")
		}
	} else if !os.IsNotExist(err) {
		return errors.New("directory_session_unavailable")
	}
	data, err := json.Marshal(session)
	if err != nil {
		return errors.New("directory_session_encode_failed")
	}
	f, err := os.CreateTemp(private, ".directory-session-*")
	if err != nil {
		return errors.New("directory_session_write_failed")
	}
	temp := f.Name()
	defer os.Remove(temp)
	if err = f.Chmod(0600); err == nil {
		_, err = f.Write(append(data, '\n'))
	}
	if err == nil {
		err = f.Sync()
	}
	if closeErr := f.Close(); err == nil {
		err = closeErr
	}
	if err != nil {
		return errors.New("directory_session_write_failed")
	}
	if err = os.Rename(temp, dest); err != nil {
		return errors.New("directory_session_write_failed")
	}
	return nil
}

func main() {
	detail := flag.String("details", "", "Private observed request detail file")
	request := flag.String("request", "", "Private observed request body file")
	response := flag.String("response", "", "Private observed response body file")
	flag.Parse()
	if flag.NArg() != 0 {
		fmt.Println(`{"ok":false,"reason":"unexpected_argument","network_called":false}`)
		os.Exit(1)
	}
	if err := run(*detail, *request, *response); err != nil {
		_ = json.NewEncoder(os.Stdout).Encode(map[string]any{"ok": false, "reason": err.Error(), "network_called": false})
		os.Exit(1)
	}
	fmt.Println(`{"ok":true,"directory_session_prepared":true,"network_called":false,"jwt_signature_verified":false}`)
}
