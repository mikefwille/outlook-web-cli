package mailcache

import (
	"context"
	"database/sql"
	"encoding/json"
	"fmt"
	"strings"
)

// ExportSnapshot contains only explicitly selected published mail-list records.
// It is not a complete message, MIME export, or server-side snapshot.
type ExportSnapshot struct {
	FolderID string       `json:"folder_id"`
	SyncedAt string       `json:"synced_at"`
	Items    []ExportItem `json:"items"`
}

type ExportItem struct {
	ID        string          `json:"id"`
	ChangeKey string          `json:"change_key"`
	Data      json.RawMessage `json:"metadata"`
}

// Export selects at most 100 exact IDs from one published folder in a single
// read transaction. Missing or repeated IDs fail the entire export.
func (s *Store) Export(ctx context.Context, folder string, ids []string) (ExportSnapshot, error) {
	validID := func(id string) bool { return id != "" && len(id) <= 4096 && !strings.ContainsAny(id, "\x00\r\n") }
	if !validID(folder) || len(ids) == 0 || len(ids) > 100 {
		return ExportSnapshot{}, fmt.Errorf("one folder and 1–100 item IDs are required")
	}
	seen := make(map[string]bool, len(ids))
	for _, id := range ids {
		if !validID(id) || seen[id] {
			return ExportSnapshot{}, fmt.Errorf("item IDs must be valid and unique")
		}
		seen[id] = true
	}
	tx, err := s.db.BeginTx(ctx, &sql.TxOptions{ReadOnly: true})
	if err != nil {
		return ExportSnapshot{}, err
	}
	defer tx.Rollback()
	out := ExportSnapshot{FolderID: folder, Items: make([]ExportItem, 0, len(ids))}
	if err = tx.QueryRowContext(ctx, `SELECT synced_at FROM heads WHERE folder=?`, folder).Scan(&out.SyncedAt); err != nil {
		if err == sql.ErrNoRows {
			return ExportSnapshot{}, fmt.Errorf("folder has no published snapshot; run sync folder first")
		}
		return ExportSnapshot{}, err
	}
	total := 0
	for _, id := range ids {
		item := ExportItem{ID: id}
		var data []byte
		err = tx.QueryRowContext(ctx, `SELECT change_key,data FROM items WHERE folder=? AND item_id=?`, folder, id).Scan(&item.ChangeKey, &data)
		if err == sql.ErrNoRows {
			return ExportSnapshot{}, fmt.Errorf("a selected item is absent from the published folder snapshot")
		}
		if err != nil {
			return ExportSnapshot{}, err
		}
		total += len(data)
		if total > 4<<20 || !json.Valid(data) {
			return ExportSnapshot{}, fmt.Errorf("selected metadata is invalid or exceeds 4 MiB")
		}
		item.Data = append(json.RawMessage(nil), data...)
		out.Items = append(out.Items, item)
	}
	if err = tx.Commit(); err != nil {
		return ExportSnapshot{}, err
	}
	return out, nil
}
