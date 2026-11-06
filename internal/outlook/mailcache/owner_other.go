//go:build !unix

package mailcache

import "os"

// The saved-profile integration is currently validated on Unix-like systems.
// Refuse to persist mailbox content where ownership cannot be checked.
func ownedByCurrentUser(os.FileInfo) bool { return false }
