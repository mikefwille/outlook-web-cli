package mailcache

import (
	"context"
	"database/sql"
	"encoding/json"
	"fmt"
	"sort"
	"time"
	"unicode/utf8"
)

type BriefItem struct {
	ID               string  `json:"id"`
	Subject          *string `json:"subject,omitempty"`
	SubjectTruncated bool    `json:"subject_truncated,omitempty"`
	ReceivedAt       *string `json:"received_at,omitempty"`
	IsRead           *bool   `json:"is_read,omitempty"`
	FlagStatus       *string `json:"flag_status,omitempty"`
}
type BriefFolder struct {
	FolderID          string      `json:"folder_id"`
	SyncedAt          string      `json:"synced_at"`
	Total             int         `json:"snapshot_item_count"`
	MissingReceivedAt int         `json:"items_without_valid_received_time"`
	UnknownFlagStatus int         `json:"items_with_unknown_flag_status"`
	Recent            []BriefItem `json:"recent"`
	RecentTruncated   bool        `json:"recent_truncated"`
	Flagged           []BriefItem `json:"flagged"`
	FlaggedTruncated  bool        `json:"flagged_truncated"`
}

// BriefFolder reads one already-published folder snapshot in a consistent
// transaction. No missing field is inferred to mean unflagged or old.
func (s *Store) BriefFolder(ctx context.Context, folder string, limit int) (BriefFolder, error) {
	if folder == "" || len(folder) > 4096 || limit < 1 || limit > 50 {
		return BriefFolder{}, fmt.Errorf("one folder and a limit of 1–50 are required")
	}
	tx, err := s.db.BeginTx(ctx, &sql.TxOptions{ReadOnly: true})
	if err != nil {
		return BriefFolder{}, err
	}
	defer tx.Rollback()
	out := BriefFolder{FolderID: folder, Recent: []BriefItem{}, Flagged: []BriefItem{}}
	if err = tx.QueryRowContext(ctx, `SELECT synced_at,item_count FROM heads WHERE folder=?`, folder).Scan(&out.SyncedAt, &out.Total); err != nil {
		if err == sql.ErrNoRows {
			return BriefFolder{}, fmt.Errorf("folder has no published snapshot; run sync folder first")
		}
		return BriefFolder{}, err
	}
	if out.Total < 0 || out.Total > 10000 {
		return BriefFolder{}, fmt.Errorf("snapshot exceeds briefing limit")
	}
	rows, err := tx.QueryContext(ctx, `SELECT item_id,data FROM items WHERE folder=?`, folder)
	if err != nil {
		return BriefFolder{}, err
	}
	all := make([]BriefItem, 0, out.Total)
	count := 0
	for rows.Next() {
		var id string
		var raw []byte
		if err = rows.Scan(&id, &raw); err != nil {
			rows.Close()
			return BriefFolder{}, err
		}
		var x struct {
			Subject  *string `json:"Subject"`
			Received *string `json:"DateTimeReceived"`
			Read     *bool   `json:"IsRead"`
			Flag     *struct {
				Status *string `json:"FlagStatus"`
			} `json:"Flag"`
		}
		if id == "" || json.Unmarshal(raw, &x) != nil {
			rows.Close()
			return BriefFolder{}, fmt.Errorf("cached message metadata is invalid")
		}
		subject, truncated := boundedSubject(x.Subject)
		item := BriefItem{ID: id, Subject: subject, SubjectTruncated: truncated, ReceivedAt: x.Received, IsRead: x.Read}
		count++
		if x.Received == nil {
			out.MissingReceivedAt++
		} else if _, e := time.Parse(time.RFC3339, *x.Received); e != nil {
			out.MissingReceivedAt++
			item.ReceivedAt = nil
		}
		if x.Flag == nil || x.Flag.Status == nil {
			out.UnknownFlagStatus++
		} else {
			switch *x.Flag.Status {
			case "Flagged", "NotFlagged", "Complete":
				item.FlagStatus = x.Flag.Status
			default:
				out.UnknownFlagStatus++
			}
		}
		all = append(all, item)
	}
	if err = rows.Err(); err != nil {
		rows.Close()
		return BriefFolder{}, err
	}
	if err = rows.Close(); err != nil {
		return BriefFolder{}, err
	}
	if count != out.Total {
		return BriefFolder{}, fmt.Errorf("published snapshot count changed or is inconsistent")
	}
	sort.SliceStable(all, func(i, j int) bool {
		if all[i].ReceivedAt == nil {
			return false
		}
		if all[j].ReceivedAt == nil {
			return true
		}
		a, _ := time.Parse(time.RFC3339, *all[i].ReceivedAt)
		b, _ := time.Parse(time.RFC3339, *all[j].ReceivedAt)
		if a.Equal(b) {
			return all[i].ID < all[j].ID
		}
		return a.After(b)
	})
	for _, item := range all {
		if item.ReceivedAt != nil {
			if len(out.Recent) < limit {
				out.Recent = append(out.Recent, item)
			} else {
				out.RecentTruncated = true
			}
		}
		if item.FlagStatus != nil && *item.FlagStatus == "Flagged" {
			if len(out.Flagged) < limit {
				out.Flagged = append(out.Flagged, item)
			} else {
				out.FlaggedTruncated = true
			}
		}
	}
	if err = tx.Commit(); err != nil {
		return BriefFolder{}, err
	}
	return out, nil
}

func boundedSubject(s *string) (*string, bool) {
	if s == nil || len(*s) <= 512 {
		return s, false
	}
	v := (*s)[:512]
	for !utf8.ValidString(v) {
		v = v[:len(v)-1]
	}
	return &v, true
}
