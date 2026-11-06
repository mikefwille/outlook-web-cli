package cli

import (
	"context"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/attachments"
)

// attachmentFingerprints checks actual bytes rather than Outlook's reported
// size, which can include a server-side envelope. Duplicate names are counted.
func attachmentFingerprints(ctx context.Context, messageID string, items []attachments.Metadata) (map[string]int, error) {
	if len(items) > 20 {
		return nil, fmt.Errorf("forward verification supports at most 20 file attachments")
	}
	result := map[string]int{}
	total := 0
	for _, a := range items {
		if a.Type != "FileAttachment:#Exchange" {
			return nil, fmt.Errorf("this attachment type cannot yet be verified for forwarding")
		}
		fresh, b, err := readAttachmentBytes(ctx, messageID, a.ID, attachments.MaxDownloadBytes)
		if err != nil {
			return nil, err
		}
		if fresh.Name != a.Name || fresh.ContentType != a.ContentType || fresh.ContentID != a.ContentID {
			return nil, fmt.Errorf("attachment metadata changed during review")
		}
		total += len(b)
		if total > 32<<20 {
			return nil, fmt.Errorf("forward verification supports at most 32 MiB in total")
		}
		digest := sha256.Sum256(b)
		key, err := json.Marshal([]string{a.Name, a.ContentType, a.ContentID, hex.EncodeToString(digest[:])})
		if err != nil {
			return nil, err
		}
		result[string(key)]++
	}
	return result, nil
}
func fingerprintsEqual(a, b map[string]int) bool {
	if len(a) != len(b) {
		return false
	}
	for k, n := range a {
		if b[k] != n {
			return false
		}
	}
	return true
}
