package mailcache

import (
	"bytes"
	"context"
	"encoding/json"
	"os"
	"strings"
	"testing"
	"unicode/utf8"
)

func TestBriefFolderKeepsUnknownsAndBoundsResults(t *testing.T) {
	root := t.TempDir()
	if err := os.Chmod(root, 0700); err != nil {
		t.Fatal(err)
	}
	ctx := context.Background()
	s, err := Open(root, "account", true)
	if err != nil {
		t.Fatal(err)
	}
	defer s.Close()
	item := func(id, when, flag string) Item {
		v := map[string]any{"ItemId": map[string]string{"Id": id, "ChangeKey": "v1"}, "ParentFolderId": map[string]string{"Id": "folder"}, "Subject": id, "IsRead": false, "Body": "omit from briefing"}
		if when != "" {
			v["DateTimeReceived"] = when
		}
		if flag != "" {
			v["Flag"] = map[string]any{"FlagStatus": flag}
		}
		b, _ := json.Marshal(v)
		return Item{ID: id, ChangeKey: "v1", Data: b}
	}
	r, err := s.Start(ctx, "folder")
	if err != nil {
		t.Fatal(err)
	}
	r, err = s.AppendPage(ctx, r, Page{Items: []Item{item("old", "2026-11-01T10:00:00Z", "Flagged"), item("new", "2026-11-02T10:00:00Z", "Flagged"), item("unknown", "bad-date", ""), item("missing", "", "Unexpected")}, Next: 4, Total: 4, Last: true}, 10)
	if err != nil {
		t.Fatal(err)
	}
	if err = s.Publish(ctx, r); err != nil {
		t.Fatal(err)
	}
	b, err := s.BriefFolder(ctx, "folder", 1)
	if err != nil {
		t.Fatal(err)
	}
	if b.Total != 4 || b.SyncedAt == "" || b.MissingReceivedAt != 2 || b.UnknownFlagStatus != 2 || len(b.Recent) != 1 || b.Recent[0].ID != "new" || !b.RecentTruncated || len(b.Flagged) != 1 || b.Flagged[0].ID != "new" || !b.FlaggedTruncated {
		t.Fatalf("incorrect bounded snapshot: %+v", b)
	}
	encoded, _ := json.Marshal(b)
	if string(encoded) == "" || bytes.Contains(encoded, []byte("omit from briefing")) {
		t.Fatal("body leaked into briefing")
	}
	if _, err = s.BriefFolder(ctx, "other", 1); err == nil {
		t.Fatal("unsynced folder accepted")
	}
	if _, err = s.BriefFolder(ctx, "folder", 51); err == nil {
		t.Fatal("oversized limit accepted")
	}
}

func TestBriefSubjectTruncationIsVisibleAndUTF8Safe(t *testing.T) {
	s := strings.Repeat("é", 300)
	got, truncated := boundedSubject(&s)
	if !truncated || got == nil || len(*got) > 512 || !utf8.ValidString(*got) {
		t.Fatal("subject truncation was not safe and visible")
	}
}
