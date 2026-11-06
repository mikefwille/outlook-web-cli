package mailcache

import (
	"context"
	"crypto/rand"
	"crypto/sha256"
	"database/sql"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"time"

	_ "modernc.org/sqlite"
)

const maxBytes = 32 << 20

type Store struct {
	db   *sql.DB
	path string
}
type Run struct {
	ID        string `json:"run_id"`
	FolderID  string `json:"folder_id"`
	Status    string `json:"status"`
	Offset    int    `json:"offset"`
	Count     int    `json:"count"`
	Total     int    `json:"total"`
	Bytes     int    `json:"bytes"`
	StartedNS int64  `json:"started_ns"`
}
type Hit struct {
	FolderID string          `json:"folder_id"`
	SyncedAt string          `json:"synced_at"`
	Data     json.RawMessage `json:"item"`
}

// Open creates an account-partitioned SQLite file only under a private root.
// The caller controls root in tests; the CLI uses its existing private session
// directory. Existing permissive paths or symlinks are refused.
func Open(root, account string, create bool) (*Store, error) {
	if root == "" || account == "" {
		return nil, fmt.Errorf("private account cache is unavailable")
	}
	checkDir := func(p string) error {
		fi, err := os.Lstat(p)
		if err != nil {
			return err
		}
		if !ownedByCurrentUser(fi) || !fi.IsDir() || fi.Mode()&os.ModeSymlink != 0 || fi.Mode().Perm()&0077 != 0 {
			return fmt.Errorf("cache directory must be private")
		}
		return nil
	}
	if err := checkDir(root); err != nil {
		return nil, err
	}
	dir := filepath.Join(root, "mail-cache")
	if create {
		if err := os.Mkdir(dir, 0700); err != nil && !os.IsExist(err) {
			return nil, err
		}
	}
	if err := checkDir(dir); err != nil {
		return nil, err
	}
	digest := sha256.Sum256([]byte(account))
	path := filepath.Join(dir, hex.EncodeToString(digest[:])+".sqlite3")
	if create {
		f, err := os.OpenFile(path, os.O_CREATE|os.O_EXCL|os.O_RDWR, 0600)
		if err == nil {
			if e := f.Close(); e != nil {
				return nil, e
			}
		} else if !os.IsExist(err) {
			return nil, err
		}
	}
	fi, err := os.Lstat(path)
	if err != nil {
		return nil, err
	}
	if !ownedByCurrentUser(fi) || !fi.Mode().IsRegular() || fi.Mode().Perm()&0077 != 0 {
		return nil, fmt.Errorf("cache file must be private")
	}
	db, err := sql.Open("sqlite", path)
	if err != nil {
		return nil, err
	}
	db.SetMaxOpenConns(1)
	if _, err = db.Exec(`PRAGMA busy_timeout=5000; PRAGMA foreign_keys=ON;`); err != nil {
		db.Close()
		return nil, err
	}
	if create {
		_, err = db.Exec(`CREATE TABLE IF NOT EXISTS runs (id TEXT PRIMARY KEY,folder TEXT NOT NULL,status TEXT NOT NULL,offset INTEGER NOT NULL,count INTEGER NOT NULL,total INTEGER NOT NULL,bytes INTEGER NOT NULL,started_ns INTEGER NOT NULL);
CREATE TABLE IF NOT EXISTS staged (run_id TEXT NOT NULL REFERENCES runs(id) ON DELETE CASCADE,item_id TEXT NOT NULL,change_key TEXT NOT NULL,data BLOB NOT NULL,PRIMARY KEY(run_id,item_id));
CREATE TABLE IF NOT EXISTS items (folder TEXT NOT NULL,item_id TEXT NOT NULL,change_key TEXT NOT NULL,data BLOB NOT NULL,PRIMARY KEY(folder,item_id));
CREATE TABLE IF NOT EXISTS heads (folder TEXT PRIMARY KEY,run_started_ns INTEGER NOT NULL,synced_at TEXT NOT NULL,item_count INTEGER NOT NULL);`)
		if err != nil {
			db.Close()
			return nil, err
		}
	}
	return &Store{db: db, path: path}, nil
}
func (s *Store) Close() error { return s.db.Close() }
func runID() (string, error) {
	b := make([]byte, 16)
	if _, err := rand.Read(b); err != nil {
		return "", err
	}
	return hex.EncodeToString(b), nil
}
func (s *Store) Start(ctx context.Context, folder string) (Run, error) {
	if folder == "" || len(folder) > 4096 || strings.ContainsAny(folder, "\x00\r\n") {
		return Run{}, fmt.Errorf("invalid folder ID")
	}
	id, err := runID()
	if err != nil {
		return Run{}, err
	}
	r := Run{ID: id, FolderID: folder, Status: "scan", Total: -1, StartedNS: time.Now().UTC().UnixNano()}
	_, err = s.db.ExecContext(ctx, `INSERT INTO runs VALUES(?,?,?,?,?,?,?,?)`, r.ID, r.FolderID, r.Status, 0, 0, -1, 0, r.StartedNS)
	return r, err
}
func (s *Store) Get(ctx context.Context, id string) (Run, error) {
	var r Run
	err := s.db.QueryRowContext(ctx, `SELECT id,folder,status,offset,count,total,bytes,started_ns FROM runs WHERE id=?`, id).Scan(&r.ID, &r.FolderID, &r.Status, &r.Offset, &r.Count, &r.Total, &r.Bytes, &r.StartedNS)
	return r, err
}
func (s *Store) Runs(ctx context.Context) ([]Run, error) {
	rows, err := s.db.QueryContext(ctx, `SELECT id,folder,status,offset,count,total,bytes,started_ns FROM runs ORDER BY started_ns DESC LIMIT 50`)
	if err != nil {
		return nil, err
	}
	defer rows.Close()
	out := []Run{}
	for rows.Next() {
		var r Run
		if err = rows.Scan(&r.ID, &r.FolderID, &r.Status, &r.Offset, &r.Count, &r.Total, &r.Bytes, &r.StartedNS); err != nil {
			return nil, err
		}
		out = append(out, r)
	}
	return out, rows.Err()
}

// AppendPage commits item rows and their next offset together. A stale or
// concurrently advanced run cannot commit another copy of the same page.
func (s *Store) AppendPage(ctx context.Context, r Run, p Page, maxItems int) (Run, error) {
	if r.Status != "scan" || p.Total > maxItems || p.Total < r.Count+len(p.Items) || !p.Last && p.Next <= r.Offset || p.Last && r.Count+len(p.Items) != p.Total || r.Total >= 0 && r.Total != p.Total {
		return r, fmt.Errorf("folder changed or exceeds the configured item limit; start a new reconciliation")
	}
	n := r.Bytes
	for _, item := range p.Items {
		n += len(item.Data)
	}
	if n > maxBytes {
		return r, fmt.Errorf("folder metadata exceeds the 32 MiB reconciliation limit")
	}
	tx, err := s.db.BeginTx(ctx, nil)
	if err != nil {
		return r, err
	}
	defer tx.Rollback()
	for _, item := range p.Items {
		if _, err = tx.ExecContext(ctx, `INSERT INTO staged(run_id,item_id,change_key,data) VALUES(?,?,?,?)`, r.ID, item.ID, item.ChangeKey, []byte(item.Data)); err != nil {
			return r, fmt.Errorf("duplicate or invalid item in folder pages")
		}
	}
	status := "scan"
	if p.Last {
		status = "verify"
	}
	res, err := tx.ExecContext(ctx, `UPDATE runs SET status=?,offset=?,count=?,total=?,bytes=? WHERE id=? AND status='scan' AND offset=? AND count=?`, status, p.Next, r.Count+len(p.Items), p.Total, n, r.ID, r.Offset, r.Count)
	if err != nil {
		return r, err
	}
	affected, err := res.RowsAffected()
	if err != nil || affected != 1 {
		return r, fmt.Errorf("reconciliation was advanced by another process")
	}
	if err = tx.Commit(); err != nil {
		return r, err
	}
	return s.Get(ctx, r.ID)
}
func (s *Store) Versions(ctx context.Context, id string) (map[string]string, error) {
	rows, err := s.db.QueryContext(ctx, `SELECT item_id,change_key FROM staged WHERE run_id=?`, id)
	if err != nil {
		return nil, err
	}
	defer rows.Close()
	out := map[string]string{}
	for rows.Next() {
		var id, key string
		if err = rows.Scan(&id, &key); err != nil {
			return nil, err
		}
		out[id] = key
	}
	return out, rows.Err()
}
func (s *Store) MarkStale(ctx context.Context, id string) error {
	_, err := s.db.ExecContext(ctx, `UPDATE runs SET status='stale' WHERE id=? AND status IN ('scan','verify')`, id)
	return err
}

// Publish atomically swaps a verified folder snapshot. A run started before an
// already-published run cannot overwrite that newer snapshot.
func (s *Store) Publish(ctx context.Context, r Run) error {
	if r.Status != "verify" {
		return fmt.Errorf("reconciliation is not ready to publish")
	}
	tx, err := s.db.BeginTx(ctx, nil)
	if err != nil {
		return err
	}
	defer tx.Rollback()
	var status, folder string
	var countAtCheckpoint int
	var started int64
	if err = tx.QueryRowContext(ctx, `SELECT status,folder,count,started_ns FROM runs WHERE id=?`, r.ID).Scan(&status, &folder, &countAtCheckpoint, &started); err != nil || status != "verify" || folder != r.FolderID || countAtCheckpoint != r.Count || started != r.StartedNS {
		return fmt.Errorf("reconciliation changed before publish")
	}
	var newer int64
	err = tx.QueryRowContext(ctx, `SELECT run_started_ns FROM heads WHERE folder=?`, r.FolderID).Scan(&newer)
	if err != nil && err != sql.ErrNoRows {
		return err
	}
	if err == nil && newer > r.StartedNS {
		return fmt.Errorf("a newer folder snapshot has already been published")
	}
	var count int
	if err = tx.QueryRowContext(ctx, `SELECT count(*) FROM staged WHERE run_id=?`, r.ID).Scan(&count); err != nil || count != r.Count {
		return fmt.Errorf("staged folder snapshot is incomplete")
	}
	if _, err = tx.ExecContext(ctx, `DELETE FROM items WHERE folder=?`, r.FolderID); err != nil {
		return err
	}
	if _, err = tx.ExecContext(ctx, `INSERT INTO items(folder,item_id,change_key,data) SELECT ?,item_id,change_key,data FROM staged WHERE run_id=?`, r.FolderID, r.ID); err != nil {
		return err
	}
	now := time.Now().UTC().Format(time.RFC3339Nano)
	if _, err = tx.ExecContext(ctx, `INSERT INTO heads(folder,run_started_ns,synced_at,item_count) VALUES(?,?,?,?) ON CONFLICT(folder) DO UPDATE SET run_started_ns=excluded.run_started_ns,synced_at=excluded.synced_at,item_count=excluded.item_count`, r.FolderID, r.StartedNS, now, r.Count); err != nil {
		return err
	}
	if _, err = tx.ExecContext(ctx, `UPDATE runs SET status='complete' WHERE id=? AND status='verify'`, r.ID); err != nil {
		return err
	}
	if _, err = tx.ExecContext(ctx, `DELETE FROM staged WHERE run_id=?`, r.ID); err != nil {
		return err
	}
	return tx.Commit()
}
func (s *Store) Search(ctx context.Context, query, folder string, limit int) ([]Hit, error) {
	if strings.TrimSpace(query) == "" || len(query) > 1000 || limit < 1 || limit > 100 {
		return nil, fmt.Errorf("query and a limit of 1–100 are required")
	}
	rows, err := s.db.QueryContext(ctx, `SELECT i.folder,h.synced_at,i.data FROM items i JOIN heads h ON h.folder=i.folder WHERE (?='' OR i.folder=?) AND instr(lower(CAST(i.data AS TEXT)),lower(?))>0 ORDER BY h.synced_at DESC,i.item_id LIMIT ?`, folder, folder, query, limit)
	if err != nil {
		return nil, err
	}
	defer rows.Close()
	out := []Hit{}
	for rows.Next() {
		var v Hit
		var data []byte
		if err = rows.Scan(&v.FolderID, &v.SyncedAt, &data); err != nil {
			return nil, err
		}
		v.Data = append(json.RawMessage(nil), data...)
		out = append(out, v)
	}
	return out, rows.Err()
}
