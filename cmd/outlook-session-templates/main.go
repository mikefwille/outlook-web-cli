// outlook-session-templates converts already observed operation shapes into a
// private, credential-free runtime bundle for the current account. It never
// contacts Outlook, launches a browser, or prints request bodies or identity.
package main

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"syscall"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/account"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/attachments"
	cal "github.com/mikefwille/outlook-web-cli/internal/outlook/calendar"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/capture"
	people "github.com/mikefwille/outlook-web-cli/internal/outlook/contacts"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/conversation"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/drafts"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/exchange"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/folders"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/mailactions"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/search"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func main() {
	checkOnly := len(os.Args) == 2 && os.Args[1] == "--check"
	if len(os.Args) != 1 && !checkOnly {
		fatal("only --check is supported")
	}
	pointer := os.Getenv("OUTLOOK_WEB_SESSION_POINTER")
	if pointer == "" {
		home, err := os.UserHomeDir()
		if err != nil {
			fatal("session unavailable")
		}
		pointer = filepath.Join(home, ".local", "share", "outlook-web-dev", "private", "current-outlook.json")
	}
	current, _, err := capture.LoadCurrentFindFolder(pointer)
	if err != nil {
		fatal("current session unavailable")
	}
	bundle, err := capture.NewRuntimeBundle(current)
	if err != nil {
		fatal("current session incompatible")
	}
	placeholder, err := drafts.ParseRecipients([]string{"template@example.invalid"}, []string{"template@example.invalid"}, []string{"template@example.invalid"})
	if err != nil {
		fatal("template preparation failed")
	}
	changes := &drafts.RecipientChanges{To: &placeholder.To, Cc: &placeholder.Cc, Bcc: &placeholder.Bcc}
	names := []string{"draft_create", "draft_get", "mail_get", "draft_compose", "draft_update", "draft_discard", "draft_send", "reply_create", "reply_all_create", "forward_create"}
	for _, name := range names {
		prefix := map[string]string{"draft_create": "OUTLOOK_WEB_DRAFT_CREATE", "draft_get": "OUTLOOK_WEB_DRAFT_GET", "mail_get": "OUTLOOK_WEB_DRAFT_GET", "draft_compose": "OUTLOOK_WEB_DRAFT_COMPOSE", "draft_update": "OUTLOOK_WEB_DRAFT_UPDATE", "draft_discard": "OUTLOOK_WEB_DRAFT_DISCARD", "draft_send": "OUTLOOK_WEB_DRAFT_SEND", "reply_create": "OUTLOOK_WEB_REPLY_CREATE", "reply_all_create": "OUTLOOK_WEB_REPLY_ALL_CREATE", "forward_create": "OUTLOOK_WEB_FORWARD_CREATE"}[name]
		detail, request := os.Getenv(prefix+"_DETAILS"), os.Getenv(prefix+"_REQUEST")
		if detail == "" || (name != "draft_discard" && request == "") {
			fatal("required private template is not configured: " + name)
		}
		var s transport.Session
		var op transport.Operation
		switch name {
		case "draft_create":
			s, op, err = capture.PrepareDraftCreate(detail, request, current, "template subject", "template body")
		case "draft_get", "mail_get", "draft_compose":
			s, op, err = capture.PrepareDraftOperation(detail, request, "GetItem", current, "template-id", "", "", "")
		case "draft_update":
			s, op, err = capture.PrepareDraftOperationWithRecipients(detail, request, "UpdateItem", current, "template-id", "template-key", "template subject", "template body", changes)
		case "draft_discard":
			s, op, err = capture.PrepareDraftOperation(detail, request, "DeleteItem", current, "template-id", "template-key", "", "")
		case "draft_send":
			s, op, err = capture.PrepareDraftSend(detail, request, current, "template-id", "template-key", "template subject")
		case "reply_create", "reply_all_create", "forward_create":
			kind := "reply"
			if name == "reply_all_create" {
				kind = "reply-all"
			}
			if name == "forward_create" {
				kind = "forward"
			}
			s, op, err = capture.PrepareResponseDraft(detail, request, current, kind, "template-id", "template-key", "template subject", "template body", placeholder)
		}
		if err != nil {
			fatal("private template incompatible: " + name)
		}
		if err = bundle.Add(name, s, op); err != nil {
			fatal("cannot normalize template: " + name)
		}
	}
	addReadTemplates(&bundle, current)
	addWebTemplates(&bundle, current)
	if checkOnly {
		_ = json.NewEncoder(os.Stdout).Encode(map[string]any{"ok": true, "operation_count": len(bundle.Operations), "web_operation_count": len(bundle.WebOperations), "written": false})
		return
	}
	dir := filepath.Dir(pointer)
	info, err := os.Lstat(dir)
	if err != nil || !info.IsDir() || info.Mode().Perm()&0077 != 0 {
		fatal("private session directory unavailable")
	}
	stat, ok := info.Sys().(*syscall.Stat_t)
	if !ok || stat.Uid != uint32(os.Getuid()) {
		fatal("private session directory has wrong owner")
	}
	dest := filepath.Join(dir, "outlook-operations-v1.json")
	if existing, err := os.Lstat(dest); err == nil {
		if !existing.Mode().IsRegular() || existing.Mode().Perm()&0077 != 0 {
			fatal("existing template bundle is not private")
		}
		owner, ok := existing.Sys().(*syscall.Stat_t)
		if !ok || owner.Uid != uint32(os.Getuid()) {
			fatal("existing template bundle has wrong owner")
		}
		data, err := os.ReadFile(dest)
		if err != nil {
			fatal("existing template bundle unavailable")
		}
		var previous capture.RuntimeBundle
		if json.Unmarshal(data, &previous) != nil || previous.Version != 1 || previous.AccountKey != bundle.AccountKey {
			fatal("existing template bundle belongs to another account or format")
		}
	} else if !os.IsNotExist(err) {
		fatal("existing template bundle unavailable")
	}
	tmp, err := os.CreateTemp(dir, ".outlook-operations-*.tmp")
	if err != nil {
		fatal("cannot stage private template bundle")
	}
	name := tmp.Name()
	defer os.Remove(name)
	if tmp.Chmod(0600) != nil || json.NewEncoder(tmp).Encode(bundle) != nil || tmp.Sync() != nil || tmp.Close() != nil {
		fatal("cannot write private template bundle")
	}
	if err = os.Rename(name, dest); err != nil {
		fatal("cannot activate private template bundle")
	}
	_ = json.NewEncoder(os.Stdout).Encode(map[string]any{"ok": true, "operation_count": len(bundle.Operations), "web_operation_count": len(bundle.WebOperations), "credentials_stored": false})
}

func addReadTemplates(bundle *capture.RuntimeBundle, current transport.Session) {
	entries := []struct{ name, action, env string }{
		{"threads_list", "FindConversation", "OUTLOOK_WEB_THREADS_LIST_EVIDENCE"},
		{"threads_get", "GetConversationItems", "OUTLOOK_WEB_THREADS_GET_EVIDENCE"},
		{"mail_list", "FindItem", "OUTLOOK_WEB_MAIL_LIST_EVIDENCE"},
		{"calendars_list", "GetCalendarFoldersV2", "OUTLOOK_WEB_CALENDARS_LIST_EVIDENCE"},
		{"calendar_agenda", "GetCalendarView", "OUTLOOK_WEB_CALENDAR_AGENDA_EVIDENCE"},
		{"account_info", "GetAccountInformation", "OUTLOOK_WEB_ACCOUNT_INFO_EVIDENCE"},
	}
	for _, e := range entries {
		file := os.Getenv(e.env)
		if file == "" {
			fatal("required private read template is not configured: " + e.name)
		}
		var s transport.Session
		var op transport.Operation
		var err error
		if e.name == "mail_list" && os.Getenv("OUTLOOK_WEB_MAIL_LIST_DETAILS") != "" {
			s, op, err = capture.LoadObservedReadDetail(os.Getenv("OUTLOOK_WEB_MAIL_LIST_DETAILS"), e.action)
		} else {
			s, op, err = capture.LoadObservedRead(file, e.action)
		}
		if err != nil {
			fatal("private read template incompatible: " + e.name)
		}
		s, err = capture.RebindRead(s, current)
		if err != nil {
			fatal("private read session incompatible: " + e.name)
		}
		if e.name == "threads_get" || e.name == "mail_list" {
			params := exchange.ReadParameters{ConversationID: "template-id"}
			if e.name == "mail_list" {
				params = exchange.ReadParameters{FolderID: "template-id"}
			}
			op, err = exchange.ParameterizeRead(op, e.action, params)
			if err != nil {
				fatal("cannot sanitize read template: " + e.name)
			}
		}
		if err = bundle.Add(e.name, s, op); err != nil {
			fatal("cannot normalize read template: " + e.name)
		}
	}
}

func addWebTemplates(bundle *capture.RuntimeBundle, current transport.Session) {
	names := []string{"account_aliases", "folders_list", "folder_create", "folder_rename", "folder_move", "folder_delete", "mail_mark_read", "mail_mark_unread", "mail_flag", "mail_unflag", "mail_trash", "mail_move", "mail_copy", "mail_categories", "search_all", "search_folder", "calendar_availability", "calendar_get", "calendar_create_self_meeting", "calendar_update_self_meeting", "calendar_delete_self_meeting", "calendar_create", "calendar_create_series", "calendar_create_all_day", "calendar_create_eastern", "calendar_update", "calendar_update_fields", "calendar_update_time", "calendar_update_combined", "calendar_update_all_day", "calendar_update_eastern", "calendar_update_series_subject", "calendar_update_series_time", "calendar_update_series_recurrence", "calendar_update_occurrence", "calendar_occurrences", "calendar_delete", "calendar_delete_series", "contacts_get", "contacts_create", "contacts_update", "contacts_delete", "people_lookup_email", "contacts_search", "contacts_list", "draft_attach", "draft_remove_attachment", "thread_mark_read", "thread_mark_unread", "thread_flag", "thread_complete", "thread_move", "thread_restore"}
	prefixes := map[string]string{"search_all": "OUTLOOK_WEB_SEARCH_ALL", "search_folder": "OUTLOOK_WEB_SEARCH_FOLDER", "calendar_get": "OUTLOOK_WEB_CALENDAR_GET", "calendar_create": "OUTLOOK_WEB_CALENDAR_CREATE", "calendar_update": "OUTLOOK_WEB_CALENDAR_UPDATE", "calendar_update_fields": "OUTLOOK_WEB_CALENDAR_UPDATE_FIELDS", "calendar_delete": "OUTLOOK_WEB_CALENDAR_DELETE", "contacts_get": "OUTLOOK_WEB_CONTACT_GET", "contacts_create": "OUTLOOK_WEB_CONTACT_CREATE", "contacts_update": "OUTLOOK_WEB_CONTACT_UPDATE", "contacts_delete": "OUTLOOK_WEB_CONTACT_DELETE", "people_lookup_email": "OUTLOOK_WEB_PEOPLE_LOOKUP_EMAIL", "contacts_search": "OUTLOOK_WEB_CONTACTS_SEARCH", "draft_attach": "OUTLOOK_WEB_DRAFT_ATTACH", "draft_remove_attachment": "OUTLOOK_WEB_DRAFT_REMOVE_ATTACHMENT"}
	for name, prefix := range map[string]string{"thread_mark_read": "OUTLOOK_WEB_THREAD_MARK_READ", "thread_mark_unread": "OUTLOOK_WEB_THREAD_MARK_UNREAD", "thread_flag": "OUTLOOK_WEB_THREAD_FLAG", "thread_complete": "OUTLOOK_WEB_THREAD_COMPLETE", "thread_move": "OUTLOOK_WEB_THREAD_MOVE", "thread_restore": "OUTLOOK_WEB_THREAD_RESTORE"} {
		prefixes[name] = prefix
	}
	prefixes["contacts_list"] = "OUTLOOK_WEB_CONTACTS_LIST"
	prefixes["calendar_availability"] = "OUTLOOK_WEB_CALENDAR_AVAILABILITY"
	prefixes["calendar_create_self_meeting"] = "OUTLOOK_WEB_CALENDAR_CREATE_SELF_MEETING"
	prefixes["calendar_update_self_meeting"] = "OUTLOOK_WEB_CALENDAR_UPDATE_SELF_MEETING"
	prefixes["calendar_delete_self_meeting"] = "OUTLOOK_WEB_CALENDAR_DELETE_SELF_MEETING"
	prefixes["calendar_update_time"] = "OUTLOOK_WEB_CALENDAR_UPDATE_TIME"
	prefixes["calendar_update_combined"] = "OUTLOOK_WEB_CALENDAR_UPDATE_COMBINED"
	prefixes["calendar_create_series"] = "OUTLOOK_WEB_CALENDAR_CREATE_SERIES"
	prefixes["calendar_create_all_day"] = "OUTLOOK_WEB_CALENDAR_CREATE_ALL_DAY"
	prefixes["calendar_create_eastern"] = "OUTLOOK_WEB_CALENDAR_CREATE_EASTERN"
	prefixes["calendar_update_all_day"] = "OUTLOOK_WEB_CALENDAR_UPDATE_ALL_DAY"
	prefixes["calendar_update_eastern"] = "OUTLOOK_WEB_CALENDAR_UPDATE_EASTERN"
	prefixes["calendar_update_series_subject"] = "OUTLOOK_WEB_CALENDAR_UPDATE_SERIES_SUBJECT"
	prefixes["calendar_update_series_time"] = "OUTLOOK_WEB_CALENDAR_UPDATE_SERIES_TIME"
	prefixes["calendar_update_series_recurrence"] = "OUTLOOK_WEB_CALENDAR_UPDATE_SERIES_RECURRENCE"
	prefixes["calendar_update_occurrence"] = "OUTLOOK_WEB_CALENDAR_UPDATE_OCCURRENCE"
	prefixes["calendar_occurrences"] = "OUTLOOK_WEB_CALENDAR_OCCURRENCES"
	prefixes["calendar_delete_series"] = "OUTLOOK_WEB_CALENDAR_DELETE_SERIES"
	prefixes["account_aliases"] = "OUTLOOK_WEB_ACCOUNT_ALIASES"
	prefixes["folders_list"] = "OUTLOOK_WEB_FOLDERS_LIST"
	for name, prefix := range map[string]string{"folder_create": "OUTLOOK_WEB_FOLDER_CREATE", "folder_rename": "OUTLOOK_WEB_FOLDER_RENAME", "folder_move": "OUTLOOK_WEB_FOLDER_MOVE", "folder_delete": "OUTLOOK_WEB_FOLDER_DELETE"} {
		prefixes[name] = prefix
	}
	for name, prefix := range map[string]string{"mail_categories": "OUTLOOK_WEB_MAIL_CATEGORIES", "mail_move": "OUTLOOK_WEB_MAIL_MOVE", "mail_copy": "OUTLOOK_WEB_MAIL_COPY", "mail_mark_read": "OUTLOOK_WEB_MAIL_MARK_READ", "mail_mark_unread": "OUTLOOK_WEB_MAIL_MARK_UNREAD", "mail_flag": "OUTLOOK_WEB_MAIL_FLAG", "mail_unflag": "OUTLOOK_WEB_MAIL_UNFLAG", "mail_trash": "OUTLOOK_WEB_MAIL_TRASH"} {
		prefixes[name] = prefix
	}
	no, yes := false, true
	event := cal.Event{ID: "template-id", IsMeeting: &no, IsRecurring: &no, IsOnlineMeeting: &no, IsOrganizer: &yes, IsAllDayEvent: &no, MeetingRequestWasSent: &no, CalendarItemType: "Single", RequiredAttendees: []json.RawMessage{}, OptionalAttendees: []json.RawMessage{}, Resources: []json.RawMessage{}, Raw: json.RawMessage(`{"Locations":[],"ExtendedProperty":[],"StartTimeZoneId":"Mountain Standard Time","EndTimeZoneId":"Mountain Standard Time"}`)}
	allDay := event
	allDay.IsAllDayEvent = &yes
	allDay.Start, allDay.End = "2026-11-06T00:00:00-07:00", "2026-11-08T00:00:00-07:00"
	eastern := event
	eastern.Raw = json.RawMessage(`{"Locations":[],"ExtendedProperty":[],"StartTimeZoneId":"Eastern Standard Time","EndTimeZoneId":"Eastern Standard Time"}`)
	occurrence := event
	occurrence.IsRecurring = &yes
	occurrence.CalendarItemType = "Occurrence"
	occurrence.Raw = json.RawMessage(`{"SeriesMasterItemId":{"Id":"template-master"}}`)
	series := event
	series.IsRecurring = &yes
	series.CalendarItemType = "RecurringMaster"
	series.Raw = json.RawMessage(`{"StartTimeZoneId":"Mountain Standard Time","EndTimeZoneId":"Mountain Standard Time","Recurrence":{"RecurrencePattern":{"__type":"WeeklyRecurrence:#Exchange","Interval":1,"DaysOfWeek":"Friday","FirstDayOfWeek":"Sunday"},"RecurrenceRange":{"__type":"EndDateRecurrence:#Exchange","StartDate":"2026-11-06","EndDate":"2026-11-27"}}}`)
	contactData := []byte(`{"id":"template-id","changeKey":"template-key","names":[{"id":"template-name","first":"Template","last":"Person","source":{"type":"UserProvided"}}],"personSources":[{"type":"UserProvided","isEditable":true,"isDeletable":true}]}`)
	contact, state := people.ParseContact(contactData, false)
	if state != transport.EnvelopeOK || !contact.SafeEditable() {
		fatal("cannot prepare synthetic contact")
	}
	selfEventData, err := capture.ReadPrivateEvidence(os.Getenv("OUTLOOK_SELF_MEETING_BEFORE_RESPONSE"))
	if err != nil {
		fatal("cannot read private self event evidence")
	}
	selfEvent, selfState := cal.ParseEvent(selfEventData)
	if selfState != transport.EnvelopeOK {
		fatal("cannot parse private self event evidence")
	}
	selfEvent.ID = "template-id"
	bodyText, location := "template body", "template location"
	email := "template@example.invalid"
	for _, name := range names {
		prefix := prefixes[name]
		detail, request := os.Getenv(prefix+"_DETAILS"), os.Getenv(prefix+"_REQUEST")
		if detail == "" || name != "account_aliases" && name != "folders_list" && !strings.HasPrefix(name, "folder_") && !strings.HasPrefix(name, "mail_") && name != "calendar_update" && name != "calendar_update_self_meeting" && name != "calendar_delete_self_meeting" && name != "calendar_update_series_subject" && name != "calendar_update_occurrence" && name != "calendar_occurrences" && name != "calendar_delete" && name != "calendar_delete_series" && name != "contacts_delete" && name != "contacts_list" && name != "draft_attach" && name != "draft_remove_attachment" && !strings.HasPrefix(name, "thread_") && request == "" {
			fatal("required private template is not configured: " + name)
		}
		var s transport.Session
		var op transport.Operation
		var err error
		switch name {
		case "account_aliases":
			s, op, err = account.PrepareAliases(detail, current)
		case "folders_list":
			s, op, err = folders.PrepareList(detail, current, 0, 1000)
		case "folder_create", "folder_rename", "folder_move", "folder_delete":
			kind := strings.TrimPrefix(name, "folder_")
			s, op, err = folders.Prepare(kind, detail, current, folders.Params{ID: "template-id", ParentID: "template-parent", Name: "template folder"})
		case "mail_mark_read", "mail_mark_unread", "mail_flag", "mail_unflag", "mail_trash":
			kind := map[string]string{"mail_mark_read": "mark-read", "mail_mark_unread": "mark-unread", "mail_flag": "flag", "mail_unflag": "unflag", "mail_trash": "trash"}[name]
			s, op, err = mailactions.Prepare(kind, detail, current, "template-id")
		case "mail_categories":
			s, op, err = mailactions.PrepareCategories(detail, current, "template-id", []string{"template category"})
		case "mail_move":
			s, op, err = mailactions.PrepareMove(detail, current, "template-id", "template-folder")
		case "mail_copy":
			s, op, err = mailactions.PrepareCopy(detail, current, "template-id", "template-folder")
		case "search_all":
			s, op, err = search.Prepare(detail, request, current, search.Params{Query: "template query", Limit: 25})
		case "search_folder":
			s, op, err = search.Prepare(detail, request, current, search.Params{Query: "template query", FolderID: "template-id", Limit: 25})
		case "calendar_availability":
			var resolved string
			s, op, resolved, err = cal.PrepareAvailability(detail, request, current, "", "2026-11-06T16:00:00Z", "2026-11-06T17:00:00Z")
			if err == nil {
				var calls []map[string]any
				if json.Unmarshal(op.Body, &calls) != nil || len(calls) != 1 {
					fatal("cannot sanitize availability template")
				}
				variables, ok := calls[0]["variables"].(map[string]any)
				if !ok {
					fatal("cannot sanitize availability template")
				}
				input, ok := variables["input"].(map[string]any)
				if !ok || input["userId"] != resolved {
					fatal("cannot sanitize availability template")
				}
				input["userId"], input["schedules"] = "template@example.invalid", []string{"template@example.invalid"}
				op.Body, err = json.Marshal(calls)
				op.Query.Set("cv", "AAAAAAAAAAAAAAAAAAAAAA.0")
			}
		case "calendar_get":
			s, op, err = cal.Prepare("GetCalendarEvent", detail, request, current, cal.Params{ID: "template-id"})
		case "calendar_create_self_meeting":
			s, op, err = cal.Prepare("CreateCalendarEvent", detail, request, current, cal.Params{CalendarID: "template-id", Subject: "template subject", Start: "2026-11-06T09:00:00-07:00", End: "2026-11-06T10:00:00-07:00", SelfMeeting: true})
			if err == nil {
				var payload map[string]any
				if json.Unmarshal(op.Body, &payload) != nil {
					fatal("cannot sanitize self event template")
				}
				body, ok := payload["Body"].(map[string]any)
				if !ok {
					fatal("cannot sanitize self event template")
				}
				items, ok := body["Items"].([]any)
				if !ok || len(items) != 1 {
					fatal("cannot sanitize self event template")
				}
				item, ok := items[0].(map[string]any)
				if !ok {
					fatal("cannot sanitize self event template")
				}
				attendees, ok := item["RequiredAttendees"].([]any)
				if !ok || len(attendees) != 1 {
					fatal("cannot sanitize self event template")
				}
				attendee, ok := attendees[0].(map[string]any)
				if !ok {
					fatal("cannot sanitize self event template")
				}
				mailbox, ok := attendee["Mailbox"].(map[string]any)
				if !ok {
					fatal("cannot sanitize self event template")
				}
				mailbox["EmailAddress"], mailbox["AADObjectId"], mailbox["Name"] = "template@example.invalid", "template-object", "template self"
				op.Body, err = json.Marshal(payload)
			}
		case "calendar_update_self_meeting":
			s, op, err = cal.Prepare("UpdateCalendarEvent", detail, request, current, cal.Params{ID: "template-id", Subject: "template subject", Current: &selfEvent, SelfMeeting: true})
		case "calendar_delete_self_meeting":
			s, op, err = cal.Prepare("DeleteCalendarEvent", detail, request, current, cal.Params{ID: "template-id", Current: &selfEvent, SelfMeeting: true})
		case "calendar_create":
			s, op, err = cal.Prepare("CreateCalendarEvent", detail, request, current, cal.Params{CalendarID: "template-id", Subject: "template subject", Start: "2026-11-06T09:00:00-07:00", End: "2026-11-06T10:00:00-07:00"})
		case "calendar_create_series":
			s, op, err = cal.Prepare("CreateCalendarEvent", detail, request, current, cal.Params{CalendarID: "template-id", Subject: "template subject", Start: "2026-11-06T09:00:00-07:00", End: "2026-11-06T10:00:00-07:00", WeeklyUntil: "2026-11-27"})
		case "calendar_create_all_day":
			s, op, err = cal.Prepare("CreateCalendarEvent", detail, request, current, cal.Params{CalendarID: "template-id", Subject: "template subject", StartDate: "2026-11-06", EndDate: "2026-11-07"})
		case "calendar_create_eastern":
			s, op, err = cal.Prepare("CreateCalendarEvent", detail, request, current, cal.Params{CalendarID: "template-id", Subject: "template subject", Start: "2026-11-06T09:00:00-05:00", End: "2026-11-06T10:00:00-05:00", TimeZone: "Eastern Standard Time"})
		case "calendar_update":
			s, op, err = cal.Prepare("UpdateCalendarEvent", detail, request, current, cal.Params{ID: "template-id", Subject: "template subject"})
		case "calendar_update_fields":
			s, op, err = cal.Prepare("UpdateCalendarEvent", detail, request, current, cal.Params{ID: "template-id", Subject: "template subject", BodyText: &bodyText, AddLocation: &location, Current: &event})
		case "calendar_update_time":
			s, op, err = cal.Prepare("UpdateCalendarEvent", detail, request, current, cal.Params{ID: "template-id", Start: "2026-11-06T09:00:00-07:00", End: "2026-11-06T10:00:00-07:00", Current: &event})
		case "calendar_update_combined":
			s, op, err = cal.Prepare("UpdateCalendarEvent", detail, request, current, cal.Params{ID: "template-id", Start: "2026-11-06T09:00:00-07:00", End: "2026-11-06T10:00:00-07:00", BodyText: &bodyText, AddLocation: &location, Current: &event})
		case "calendar_update_all_day":
			s, op, err = cal.Prepare("UpdateCalendarEvent", detail, request, current, cal.Params{ID: "template-id", StartDate: "2026-11-09", EndDate: "2026-11-10", Current: &allDay})
		case "calendar_update_eastern":
			s, op, err = cal.Prepare("UpdateCalendarEvent", detail, request, current, cal.Params{ID: "template-id", Start: "2026-11-06T09:00:00-05:00", End: "2026-11-06T10:00:00-05:00", TimeZone: "Eastern Standard Time", Current: &eastern})
		case "calendar_update_series_subject":
			s, op, err = cal.Prepare("UpdateCalendarEvent", detail, request, current, cal.Params{ID: "template-id", Subject: "template subject", Current: &series, Series: true})
		case "calendar_update_series_time":
			s, op, err = cal.Prepare("UpdateCalendarEvent", detail, request, current, cal.Params{ID: "template-id", Start: "2026-11-06T09:00:00-07:00", End: "2026-11-06T10:00:00-07:00", Current: &series, Series: true, ResetExceptions: true})
		case "calendar_update_series_recurrence":
			s, op, err = cal.Prepare("UpdateCalendarEvent", detail, request, current, cal.Params{ID: "template-id", WeeklyUntil: "2026-12-04", Current: &series, Series: true, ResetExceptions: true})
		case "calendar_update_occurrence":
			s, op, err = cal.Prepare("UpdateCalendarEvent", detail, request, current, cal.Params{ID: "template-id", Subject: "template subject", Occurrence: true, Current: &occurrence})
		case "calendar_occurrences":
			s, op, err = cal.PrepareOccurrences(detail, current, "template-id", "2026-11-06T09:00:00-07:00", "2026-11-27T10:00:00-07:00")
		case "calendar_delete":
			s, op, err = cal.Prepare("DeleteCalendarEvent", detail, request, current, cal.Params{ID: "template-id"})
		case "calendar_delete_series":
			s, op, err = cal.Prepare("DeleteCalendarEvent", detail, request, current, cal.Params{ID: "template-id", Current: &series, Series: true})
		case "contacts_get":
			s, op, err = people.Prepare("get", detail, request, current, people.Params{ID: "template-id"})
		case "contacts_create":
			s, op, err = people.Prepare("create", detail, request, current, people.Params{First: "Template", Last: "Person"})
		case "contacts_update":
			s, op, err = people.Prepare("update", detail, request, current, people.Params{ID: "template-id", First: "Template", Last: "Person", NameChange: true, Current: &contact})
		case "contacts_delete":
			s, op, err = people.Prepare("delete", detail, request, current, people.Params{ID: "template-id"})
		case "people_lookup_email":
			s, op, err = people.Prepare("lookup-email", detail, request, current, people.Params{Email: &email})
		case "contacts_search":
			s, op, err = people.PrepareSearch(detail, request, current, "template query")
		case "contacts_list":
			s, op, err = people.PrepareList(detail, current, 100, "")
		case "draft_remove_attachment":
			s, op, err = attachments.PrepareRemove(detail, current, "template-id", "template-key", "template-attachment")
		case "draft_attach":
			s, op, err = attachments.PrepareUpload(detail, current, "template-id", "template-key", "template.txt", "text/plain", []byte("template bytes"))
		case "thread_mark_read", "thread_mark_unread", "thread_flag", "thread_complete", "thread_move", "thread_restore":
			kind := map[string]string{"thread_mark_read": "mark-read", "thread_mark_unread": "mark-unread", "thread_flag": "flag", "thread_complete": "complete", "thread_move": "move", "thread_restore": "restore"}[name]
			s, op, err = conversation.Prepare(kind, detail, current, conversation.Params{ConversationID: "template-id", FolderID: "template-folder", DestinationID: "template-destination", Through: "2026-01-01T00:00:00Z", CompletedAt: "2026-01-01T00:00:00Z"})
		}
		if err != nil {
			fatal("private template incompatible: " + name)
		}
		if err = bundle.AddWeb(name, s, op); err != nil {
			fatal("cannot normalize template: " + name)
		}
	}
	addContactListTemplates(bundle, current)
	addPreferenceTemplates(bundle, current)
	addRuleTemplates(bundle, current)
	addRuleExpansionTemplates(bundle, current)
	addMasterCategoryTemplates(bundle, current)
}

func fatal(message string) { fmt.Fprintln(os.Stderr, message); os.Exit(1) }
