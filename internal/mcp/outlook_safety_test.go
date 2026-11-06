package mcp

import (
	"testing"

	"github.com/mark3labs/mcp-go/server"
)

func TestOutlookMCPOnlyExposesReviewedCommands(t *testing.T) {
	t.Setenv("OUTLOOK_WEB_MCP_READ_ONLY", "")
	s := server.NewMCPServer("outlook-test", "test")
	RegisterTools(s)
	tools := s.ListTools()
	if tools["mail_categories"] != nil || tools["sync"] != nil {
		t.Error("free-form parent aliases must not bypass the reviewed MCP tool allowlist")
	}
	if tools["drafts_attach"] != nil {
		t.Error("local file-reading attachment command exposed to MCP")
	}
	if tools["auth_recover"] != nil || tools["auth_login"] != nil || tools["auth_logout"] != nil || tools["auth_import"] != nil {
		t.Error("browser session recovery exposed to MCP")
	}
	for _, name := range []string{"context", "capabilities", "sync_folder", "sync_status", "search", "sql", "export", "deleted_list", "deleted_restore", "folders_list", "folders_get", "folders_create", "folders_rename", "folders_move", "folders_delete", "categories_list", "categories_get", "settings_get", "mail_batch", "mail_compose", "mail_mark_read", "mail_mark_unread", "mail_flag", "mail_unflag", "mail_complete", "mail_trash", "mail_categories_get", "mail_categories_add", "mail_categories_remove", "mail_move", "mail_archive", "mail_restore", "mail_list", "mail_get", "mail_export", "mail_search", "mail_reply", "mail_forward", "threads_list", "threads_get", "threads_restore", "drafts_list", "drafts_create", "drafts_get", "drafts_update", "drafts_remove_attachment", "drafts_discard", "drafts_send", "calendar_availability", "calendar_get", "calendar_create", "calendar_update", "calendar_delete", "directory_lookup", "recipients_resolve", "contacts_get", "contacts_create", "contacts_update", "contacts_delete", "contacts_lookup_email", "contacts_search"} {
		if tools[name] == nil {
			t.Errorf("reviewed tool %s missing", name)
		}
	}
	for _, key := range []string{"folder-id", "item-id"} {
		if _, ok := tools["export"].Tool.InputSchema.Properties[key]; !ok {
			t.Errorf("export lacks typed %s", key)
		}
	}
	if _, ok := tools["export"].Tool.InputSchema.Properties["output"]; ok {
		t.Error("export exposes local output path")
	}
	if _, ok := tools["sql"].Tool.InputSchema.Properties["query"]; !ok {
		t.Error("SQL tool lacks its typed query input")
	}
	for _, key := range []string{"draft-id", "attachment-id", "execute"} {
		if _, ok := tools["drafts_remove_attachment"].Tool.InputSchema.Properties[key]; !ok {
			t.Errorf("attachment removal lacks typed %s", key)
		}
	}
	if hint := tools["drafts_remove_attachment"].Tool.Annotations.ReadOnlyHint; hint != nil && *hint {
		t.Error("attachment removal marked read-only")
	}
	for name, entry := range tools {
		for _, forbidden := range []string{"action", "base-url", "config", "token", "cookies-file", "db", "input-file", "client-profile", "receipt", "text-file", "html-file"} {
			if _, ok := entry.Tool.InputSchema.Properties[forbidden]; ok {
				t.Errorf("%s exposes unsafe argument %s", name, forbidden)
			}
		}
		if name == "mail_mark_read" || name == "mail_mark_unread" || name == "mail_flag" || name == "mail_unflag" || name == "mail_complete" || name == "mail_trash" || name == "deleted_restore" || name == "mail_categories_get" || name == "mail_categories_add" || name == "mail_categories_remove" || name == "mail_move" || name == "mail_archive" || name == "mail_restore" || name == "folders_get" || name == "folders_rename" || name == "folders_move" || name == "folders_delete" || name == "categories_get" || name == "mail_get" || name == "mail_export" || name == "mail_reply" || name == "mail_forward" || name == "drafts_get" || name == "drafts_update" || name == "drafts_discard" || name == "drafts_send" || name == "calendar_get" || name == "calendar_update" || name == "calendar_delete" || name == "contacts_get" || name == "contacts_update" || name == "contacts_delete" {
			if _, ok := entry.Tool.InputSchema.Properties["id"]; !ok {
				t.Errorf("%s lacks a typed item ID", name)
			}
		}
		if name == "threads_restore" || name == "mail_mark_read" || name == "mail_mark_unread" || name == "mail_flag" || name == "mail_unflag" || name == "mail_complete" || name == "mail_trash" || name == "deleted_restore" || name == "mail_categories_add" || name == "mail_categories_remove" || name == "mail_move" || name == "mail_archive" || name == "mail_restore" || name == "mail_batch" || name == "mail_compose" || name == "sync_folder" || name == "folders_create" || name == "folders_rename" || name == "folders_move" || name == "folders_delete" || name == "mail_reply" || name == "mail_forward" || name == "drafts_create" || name == "drafts_update" || name == "drafts_discard" || name == "drafts_send" || name == "calendar_create" || name == "calendar_update" || name == "calendar_delete" || name == "contacts_create" || name == "contacts_update" || name == "contacts_delete" {
			if entry.Tool.Annotations.ReadOnlyHint != nil && *entry.Tool.Annotations.ReadOnlyHint {
				t.Errorf("%s is a write", name)
			}
			if _, ok := entry.Tool.InputSchema.Properties["execute"]; !ok {
				t.Errorf("%s cannot explicitly execute", name)
			}
		}
	}
}

func TestOutlookMCPReadOnlyModeRemovesDraftWrites(t *testing.T) {
	t.Setenv("OUTLOOK_WEB_MCP_READ_ONLY", "1")
	s := server.NewMCPServer("outlook-test", "test")
	RegisterTools(s)
	tools := s.ListTools()
	for _, name := range []string{"threads_restore", "mail_mark_read", "mail_mark_unread", "mail_flag", "mail_unflag", "mail_complete", "mail_trash", "deleted_restore", "mail_categories_add", "mail_categories_remove", "mail_move", "mail_archive", "mail_restore", "mail_batch", "mail_compose", "sync_folder", "folders_create", "folders_rename", "folders_move", "folders_delete", "mail_reply", "mail_forward", "drafts_create", "drafts_update", "drafts_remove_attachment", "drafts_discard", "drafts_send", "calendar_create", "calendar_update", "calendar_delete", "contacts_create", "contacts_update", "contacts_delete"} {
		if tools[name] != nil {
			t.Errorf("read-only server exposed %s", name)
		}
	}
	for _, name := range []string{"context", "capabilities", "sync_status", "search", "sql", "export", "deleted_list", "folders_list", "folders_get", "categories_list", "categories_get", "settings_get", "calendar_availability", "mail_get", "mail_export", "mail_categories_get", "directory_lookup", "recipients_resolve", "drafts_list", "drafts_get"} {
		if tools[name] == nil {
			t.Errorf("read-only server omitted %s", name)
		}
	}
}
