// Hand-authored Outlook calendar and personal contacts extension.
package cli

import (
	"context"
	"fmt"
	"os"
	"strings"
	"time"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/cliutil"
	cal "github.com/mikefwille/outlook-web-cli/internal/outlook/calendar"
	people "github.com/mikefwille/outlook-web-cli/internal/outlook/contacts"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func init() {
	registerNovelCommand(func(root *cobra.Command, flags *rootFlags) {
		parent := findSubcommand(root, "calendar")
		if parent == nil {
			parent = &cobra.Command{Use: "calendar", Short: "Work with calendar appointments"}
			root.AddCommand(parent)
		}
		addOutlookAppointments(parent, flags)
		root.AddCommand(newOutlookPersonalContacts(flags))
	})
}
func outlookOperation(ctx context.Context, session transport.Session, op transport.Operation) (transport.Response, error) {
	c, err := transport.New(session.AccountKey, session, nil)
	if err != nil {
		return transport.Response{}, err
	}
	ctx, cancel := context.WithTimeout(ctx, 30*time.Second)
	defer cancel()
	return c.Do(ctx, session.AccountKey, op)
}
func outlookCapture(prefix string, needsBody bool) (string, string, error) {
	detail, request := os.Getenv(prefix+"_DETAILS"), os.Getenv(prefix+"_REQUEST")
	if detail == "" && request == "" {
		name := map[string]string{"OUTLOOK_WEB_ACCOUNT_ALIASES": "account_aliases", "OUTLOOK_WEB_MAIL_MARK_READ": "mail_mark_read", "OUTLOOK_WEB_MAIL_MARK_UNREAD": "mail_mark_unread", "OUTLOOK_WEB_MAIL_FLAG": "mail_flag", "OUTLOOK_WEB_MAIL_UNFLAG": "mail_unflag", "OUTLOOK_WEB_MAIL_TRASH": "mail_trash", "OUTLOOK_WEB_MAIL_MOVE": "mail_move", "OUTLOOK_WEB_MAIL_COPY": "mail_copy", "OUTLOOK_WEB_MAIL_CATEGORIES": "mail_categories", "OUTLOOK_WEB_FOLDERS_LIST": "folders_list", "OUTLOOK_WEB_FOLDER_CREATE": "folder_create", "OUTLOOK_WEB_FOLDER_RENAME": "folder_rename", "OUTLOOK_WEB_FOLDER_MOVE": "folder_move", "OUTLOOK_WEB_FOLDER_DELETE": "folder_delete", "OUTLOOK_WEB_CALENDAR_AVAILABILITY": "calendar_availability", "OUTLOOK_WEB_CALENDAR_GET": "calendar_get", "OUTLOOK_WEB_CALENDAR_CREATE": "calendar_create", "OUTLOOK_WEB_CALENDAR_CREATE_SELF_MEETING": "calendar_create_self_meeting", "OUTLOOK_WEB_CALENDAR_UPDATE_SELF_MEETING": "calendar_update_self_meeting", "OUTLOOK_WEB_CALENDAR_DELETE_SELF_MEETING": "calendar_delete_self_meeting", "OUTLOOK_WEB_CALENDAR_UPDATE": "calendar_update", "OUTLOOK_WEB_CALENDAR_UPDATE_FIELDS": "calendar_update_fields", "OUTLOOK_WEB_CALENDAR_UPDATE_TIME": "calendar_update_time", "OUTLOOK_WEB_CALENDAR_UPDATE_COMBINED": "calendar_update_combined", "OUTLOOK_WEB_CALENDAR_CREATE_SERIES": "calendar_create_series", "OUTLOOK_WEB_CALENDAR_CREATE_ALL_DAY": "calendar_create_all_day", "OUTLOOK_WEB_CALENDAR_CREATE_EASTERN": "calendar_create_eastern", "OUTLOOK_WEB_CALENDAR_UPDATE_ALL_DAY": "calendar_update_all_day", "OUTLOOK_WEB_CALENDAR_UPDATE_EASTERN": "calendar_update_eastern", "OUTLOOK_WEB_CALENDAR_UPDATE_SERIES_SUBJECT": "calendar_update_series_subject", "OUTLOOK_WEB_CALENDAR_UPDATE_SERIES_TIME": "calendar_update_series_time", "OUTLOOK_WEB_CALENDAR_UPDATE_SERIES_RECURRENCE": "calendar_update_series_recurrence", "OUTLOOK_WEB_CALENDAR_UPDATE_OCCURRENCE": "calendar_update_occurrence", "OUTLOOK_WEB_CALENDAR_OCCURRENCES": "calendar_occurrences", "OUTLOOK_WEB_CALENDAR_DELETE": "calendar_delete", "OUTLOOK_WEB_CALENDAR_DELETE_SERIES": "calendar_delete_series", "OUTLOOK_WEB_CONTACT_GET": "contacts_get", "OUTLOOK_WEB_CONTACT_CREATE": "contacts_create", "OUTLOOK_WEB_CONTACT_UPDATE": "contacts_update", "OUTLOOK_WEB_CONTACT_DELETE": "contacts_delete", "OUTLOOK_WEB_PEOPLE_LOOKUP_EMAIL": "people_lookup_email", "OUTLOOK_WEB_CONTACTS_SEARCH": "contacts_search", "OUTLOOK_WEB_CONTACTS_LIST": "contacts_list", "OUTLOOK_WEB_DRAFT_ATTACH": "draft_attach", "OUTLOOK_WEB_DRAFT_REMOVE_ATTACHMENT": "draft_remove_attachment", "OUTLOOK_WEB_THREAD_MARK_READ": "thread_mark_read", "OUTLOOK_WEB_THREAD_MARK_UNREAD": "thread_mark_unread", "OUTLOOK_WEB_THREAD_FLAG": "thread_flag", "OUTLOOK_WEB_THREAD_COMPLETE": "thread_complete", "OUTLOOK_WEB_THREAD_MOVE": "thread_move", "OUTLOOK_WEB_THREAD_RESTORE": "thread_restore"}[prefix]
		if name != "" {
			return "runtime:" + name, "", nil
		}
	}
	if detail == "" || needsBody && request == "" {
		return "", "", fmt.Errorf("configure private %s_DETAILS%s capture paths", prefix, func() string {
			if needsBody {
				return " and " + prefix + "_REQUEST"
			}
			return ""
		}())
	}
	return detail, request, nil
}
func calendarRequest(ctx context.Context, action string, p cal.Params) (transport.Response, error) {
	prefix, body := map[string]string{"GetCalendarEvent": "OUTLOOK_WEB_CALENDAR_GET", "CreateCalendarEvent": "OUTLOOK_WEB_CALENDAR_CREATE", "UpdateCalendarEvent": "OUTLOOK_WEB_CALENDAR_UPDATE", "DeleteCalendarEvent": "OUTLOOK_WEB_CALENDAR_DELETE"}[action], action == "GetCalendarEvent" || action == "CreateCalendarEvent"
	zone, err := cal.CanonicalTimeZone(p.TimeZone)
	if err != nil {
		return transport.Response{}, err
	}
	if action == "CreateCalendarEvent" && zone == "Eastern Standard Time" {
		prefix = "OUTLOOK_WEB_CALENDAR_CREATE_EASTERN"
	}
	if action == "CreateCalendarEvent" && p.WeeklyUntil != "" {
		prefix, body = "OUTLOOK_WEB_CALENDAR_CREATE_SERIES", true
	}
	if action == "CreateCalendarEvent" && (p.StartDate != "" || p.EndDate != "") {
		prefix, body = "OUTLOOK_WEB_CALENDAR_CREATE_ALL_DAY", true
	}
	if action == "CreateCalendarEvent" && p.SelfMeeting {
		prefix, body = "OUTLOOK_WEB_CALENDAR_CREATE_SELF_MEETING", true
	}
	if action == "DeleteCalendarEvent" && p.Series {
		prefix, body = "OUTLOOK_WEB_CALENDAR_DELETE_SERIES", false
	}
	if action == "DeleteCalendarEvent" && p.SelfMeeting {
		prefix, body = "OUTLOOK_WEB_CALENDAR_DELETE_SELF_MEETING", false
	}
	if action == "UpdateCalendarEvent" && p.Occurrence {
		if p.Start != "" || p.End != "" || p.StartDate != "" || p.EndDate != "" || p.BodyText != nil || p.AddLocation != nil || p.ReplaceLocation != nil || p.WithLocation != nil || p.RemoveLocation != nil {
			return transport.Response{}, fmt.Errorf("occurrence updates currently support subject only")
		}
		prefix, body = "OUTLOOK_WEB_CALENDAR_UPDATE_OCCURRENCE", false
	}
	if action == "UpdateCalendarEvent" && (p.StartDate != "" || p.EndDate != "") {
		if p.Start != "" || p.End != "" || p.Subject != "" || p.BodyText != nil || p.AddLocation != nil || p.ReplaceLocation != nil || p.WithLocation != nil || p.RemoveLocation != nil {
			return transport.Response{}, fmt.Errorf("all-day date changes must be made separately from other changes")
		}
		prefix, body = "OUTLOOK_WEB_CALENDAR_UPDATE_ALL_DAY", true
	}
	if action == "UpdateCalendarEvent" && (p.Start != "" || p.End != "") {
		combined := p.BodyText != nil || p.AddLocation != nil || p.ReplaceLocation != nil || p.WithLocation != nil || p.RemoveLocation != nil
		if combined {
			if zone != "Mountain Standard Time" || p.Subject != "" || p.Series || p.Occurrence {
				return transport.Response{}, fmt.Errorf("combined time and content edits require a Mountain-time single appointment and a separate subject change")
			}
			prefix, body = "OUTLOOK_WEB_CALENDAR_UPDATE_COMBINED", true
		} else {
			prefix, body = "OUTLOOK_WEB_CALENDAR_UPDATE_TIME", true
			if zone == "Eastern Standard Time" {
				prefix = "OUTLOOK_WEB_CALENDAR_UPDATE_EASTERN"
			}
		}
	}
	if action == "UpdateCalendarEvent" && p.Start == "" && p.End == "" && (p.BodyText != nil || p.AddLocation != nil || p.ReplaceLocation != nil || p.WithLocation != nil || p.RemoveLocation != nil) {
		prefix = "OUTLOOK_WEB_CALENDAR_UPDATE_FIELDS"
		body = true
	}
	if action == "UpdateCalendarEvent" && p.Series {
		switch {
		case p.Subject != "":
			prefix, body = "OUTLOOK_WEB_CALENDAR_UPDATE_SERIES_SUBJECT", false
		case p.Start != "" || p.End != "":
			prefix, body = "OUTLOOK_WEB_CALENDAR_UPDATE_SERIES_TIME", true
		case p.WeeklyUntil != "":
			prefix, body = "OUTLOOK_WEB_CALENDAR_UPDATE_SERIES_RECURRENCE", true
		default:
			return transport.Response{}, fmt.Errorf("provide a series change")
		}
	}
	if action == "UpdateCalendarEvent" && p.SelfMeeting {
		prefix, body = "OUTLOOK_WEB_CALENDAR_UPDATE_SELF_MEETING", false
	}
	if prefix == "" {
		return transport.Response{}, fmt.Errorf("unsupported calendar operation")
	}
	detail, request, err := outlookCapture(prefix, body)
	if err != nil {
		return transport.Response{}, err
	}
	current, _, err := currentSession()
	if err != nil {
		return transport.Response{}, err
	}
	session, op, err := cal.Prepare(action, detail, request, current, p)
	if err != nil {
		return transport.Response{}, err
	}
	return outlookOperation(ctx, session, op)
}
func calendarRead(ctx context.Context, id string) (cal.Event, error) {
	response, err := calendarRequest(ctx, "GetCalendarEvent", cal.Params{ID: id})
	if err != nil {
		return cal.Event{}, err
	}
	event, state := cal.ParseEvent(response.Body)
	if state != transport.EnvelopeOK || event.ID != id {
		return cal.Event{}, &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
	}
	return event, nil
}
func oneID(args []string) (string, error) {
	if len(args) != 1 || strings.TrimSpace(args[0]) == "" {
		return "", usageErr(fmt.Errorf("supply exactly one item ID"))
	}
	return args[0], nil
}
func mutationInHarness() error {
	if cliutil.IsAnyHarness() {
		return fmt.Errorf("execution of Outlook writes is disabled under the verification harness")
	}
	return nil
}

// pp:data-source live
func addOutlookAppointments(parent *cobra.Command, flags *rootFlags) {
	get := &cobra.Command{Use: "get <id>", Short: "Read one calendar event", Long: "Returns server timestamps, event time zones, and event-local timestamps when the zone is recognized. Personal all-day events include inclusive start_date and end_date.", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live", "pp:happy-args": "id=example-value"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "calendar get"})
		}
		id, err := oneID(args)
		if err != nil {
			return err
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "calendar get", "id": id})
		}
		event, err := calendarRead(cmd.Context(), id)
		if err != nil {
			return err
		}
		return flags.printJSON(cmd, event)
	}}
	parent.AddCommand(get)
	var subject, start, end, startDate, endDate, calendarID, createBody, createLocation, weeklyUntil, createZone string
	var execute, selfMeetingCreate bool
	create := &cobra.Command{Use: "create", Short: "Preview or create a personal appointment", Long: "Creates a no-attendee timed or all-day appointment using the captured calendar time zone. Timed appointments may repeat weekly. All-day start and end dates are inclusive. Supply --execute to create it.", Annotations: map[string]string{"pp:data-source": "live"}, RunE: func(cmd *cobra.Command, args []string) error {
		if len(args) != 0 {
			return usageErr(fmt.Errorf("calendar create takes flags only"))
		}
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "calendar create"})
		}
		if strings.TrimSpace(subject) == "" || calendarID == "" {
			return usageErr(fmt.Errorf("--subject and --calendar-id are required"))
		}
		zoneID, err := cal.CanonicalTimeZone(createZone)
		if err != nil {
			return usageErr(err)
		}
		p := cal.Params{CalendarID: calendarID, Subject: subject, TimeZone: zoneID}
		fields := map[string]any{"calendar_id": calendarID, "subject": subject}
		allDay := cmd.Flags().Changed("start-date") || cmd.Flags().Changed("end-date")
		if selfMeetingCreate && (allDay || zoneID != "Mountain Standard Time" || cmd.Flags().Changed("weekly-until") || cmd.Flags().Changed("body") || cmd.Flags().Changed("location")) {
			return usageErr(fmt.Errorf("--self-meeting currently supports a timed Mountain-time subject only"))
		}
		p.SelfMeeting = selfMeetingCreate
		if selfMeetingCreate {
			fields["self_meeting"] = true
		}
		var a time.Time
		if allDay {
			if zoneID != "Mountain Standard Time" {
				return usageErr(fmt.Errorf("all-day creation currently uses the captured Mountain time zone"))
			}
			if !cmd.Flags().Changed("start-date") || !cmd.Flags().Changed("end-date") || cmd.Flags().Changed("start") || cmd.Flags().Changed("end") || cmd.Flags().Changed("weekly-until") {
				return usageErr(fmt.Errorf("all-day creation requires --start-date and --end-date and no timed or recurrence flags"))
			}
			if _, _, err := cal.AllDayRange(startDate, endDate); err != nil {
				return usageErr(err)
			}
			p.StartDate, p.EndDate = startDate, endDate
			fields["all_day"], fields["start_date"], fields["end_date"] = true, startDate, endDate
		} else {
			if zoneID == "Eastern Standard Time" && cmd.Flags().Changed("weekly-until") {
				return usageErr(fmt.Errorf("Eastern-time recurrence is not yet supported"))
			}
			var b time.Time
			var e, f error
			a, e = time.Parse(time.RFC3339, start)
			b, f = time.Parse(time.RFC3339, end)
			if e != nil || f != nil || !a.Before(b) {
				return usageErr(fmt.Errorf("--start and --end require RFC3339 offsets and end must be after start"))
			}
			p.Start, p.End = start, end
			fields["start"], fields["end"] = start, end
		}
		if cmd.Flags().Changed("weekly-until") {
			if cmd.Flags().Changed("body") || cmd.Flags().Changed("location") {
				return usageErr(fmt.Errorf("weekly creation currently supports a subject and time only"))
			}
			zone, err := time.LoadLocation("America/Denver")
			if err != nil {
				return err
			}
			date, err := time.Parse("2006-01-02", weeklyUntil)
			if err != nil || date.Format("2006-01-02") != weeklyUntil || weeklyUntil < a.In(zone).Format("2006-01-02") {
				return usageErr(fmt.Errorf("--weekly-until must be a date on or after the first occurrence"))
			}
			p.WeeklyUntil = weeklyUntil
			fields["recurrence"] = map[string]any{"pattern": "weekly", "interval": 1, "weekday": a.In(zone).Weekday().String(), "until": weeklyUntil}
		}
		if cmd.Flags().Changed("body") {
			if strings.TrimSpace(createBody) == "" {
				return usageErr(fmt.Errorf("--body cannot be empty"))
			}
			p.BodyText = &createBody
			fields["body"] = createBody
		}
		if cmd.Flags().Changed("location") {
			if strings.TrimSpace(createLocation) == "" {
				return usageErr(fmt.Errorf("--location cannot be empty"))
			}
			p.AddLocation = &createLocation
			fields["location"] = createLocation
		}
		if !execute || flags.dryRun {
			preview := map[string]any{"preview": true, "operation": "calendar create", "fields": fields, "attendees": []string{}, "sends_invitations": false, "time_zone": zoneID}
			if selfMeetingCreate {
				preview["attendees"] = []string{"current account only"}
				delete(preview, "sends_invitations")
				preview["invitation_mode"] = "captured SendToNone; self-delivery not established"
				preview["online_meeting"] = true
			}
			return flags.printJSON(cmd, preview)
		}
		if err := mutationInHarness(); err != nil {
			return err
		}
		response, err := calendarRequest(cmd.Context(), "CreateCalendarEvent", p)
		if err != nil {
			return err
		}
		event, state := cal.ParseEvent(response.Body)
		if state != transport.EnvelopeOK {
			return &transport.Error{Kind: transport.KindAmbiguous, Status: response.Status}
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "created": event})
	}}
	create.Flags().StringVar(&calendarID, "calendar-id", "", "Destination calendar folder ID")
	create.Flags().StringVar(&subject, "subject", "", "Appointment subject")
	create.Flags().StringVar(&start, "start", "", "Start with RFC3339 offset")
	create.Flags().StringVar(&end, "end", "", "End with RFC3339 offset")
	create.Flags().StringVar(&startDate, "start-date", "", "All-day start date (YYYY-MM-DD, inclusive)")
	create.Flags().StringVar(&endDate, "end-date", "", "All-day end date (YYYY-MM-DD, inclusive)")
	create.Flags().StringVar(&createZone, "time-zone", "", "Event time zone: America/Denver (default) or America/New_York")
	create.Flags().StringVar(&createBody, "body", "", "Plain-text appointment body")
	create.Flags().StringVar(&createLocation, "location", "", "Initial plain-text location")
	create.Flags().StringVar(&weeklyUntil, "weekly-until", "", "Repeat weekly on the start weekday through this inclusive YYYY-MM-DD date")
	create.Flags().BoolVar(&selfMeetingCreate, "self-meeting", false, "Create the captured self-only online event; no external attendees")
	create.Flags().BoolVar(&execute, "execute", false, "Create this appointment on the server")
	parent.AddCommand(create)
	var changeKey, newSubject, newBody, addLocation, replaceLocation, withLocation, removeLocation, newStart, newEnd, newStartDate, newEndDate, updateZone, newWeeklyUntil string
	var updateExecute, occurrence, series, resetExceptions, selfMeetingUpdate bool
	update := &cobra.Command{Use: "update <id>", Short: "Preview or change a personal appointment; version check is not atomic", Long: "Changes subject, body, timed start/end, inclusive all-day dates, or a named location. Mountain-time single appointments support a combined time and body/location edit. Series time or recurrence edits require --reset-exceptions. The version check is not atomic.", Annotations: map[string]string{"pp:data-source": "live", "pp:happy-args": "id=example-value"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "calendar update"})
		}
		id, err := oneID(args)
		if err != nil {
			return err
		}
		if changeKey == "" {
			return usageErr(fmt.Errorf("--change-key is required"))
		}
		p := cal.Params{ID: id}
		changes := map[string]string{}
		if cmd.Flags().Changed("subject") {
			if strings.TrimSpace(newSubject) == "" {
				return usageErr(fmt.Errorf("--subject cannot be empty"))
			}
			p.Subject = newSubject
			changes["subject"] = newSubject
		}
		if cmd.Flags().Changed("body") {
			if strings.TrimSpace(newBody) == "" {
				return usageErr(fmt.Errorf("--body cannot be empty"))
			}
			p.BodyText = &newBody
			changes["body"] = newBody
		}
		if cmd.Flags().Changed("add-location") {
			if strings.TrimSpace(addLocation) == "" {
				return usageErr(fmt.Errorf("--add-location cannot be empty"))
			}
			p.AddLocation = &addLocation
			changes["add_location"] = addLocation
		}
		locationOps := 0
		for _, name := range []string{"add-location", "replace-location", "remove-location"} {
			if cmd.Flags().Changed(name) {
				locationOps++
			}
		}
		if locationOps > 1 {
			return usageErr(fmt.Errorf("choose only one location operation"))
		}
		if cmd.Flags().Changed("replace-location") != cmd.Flags().Changed("with-location") {
			return usageErr(fmt.Errorf("--replace-location and --with-location must be supplied together"))
		}
		if cmd.Flags().Changed("replace-location") {
			if strings.TrimSpace(replaceLocation) == "" || strings.TrimSpace(withLocation) == "" {
				return usageErr(fmt.Errorf("replacement location names cannot be empty"))
			}
			p.ReplaceLocation, p.WithLocation = &replaceLocation, &withLocation
			changes["replace_location"], changes["with_location"] = replaceLocation, withLocation
		}
		if cmd.Flags().Changed("remove-location") {
			if strings.TrimSpace(removeLocation) == "" {
				return usageErr(fmt.Errorf("--remove-location cannot be empty"))
			}
			p.RemoveLocation = &removeLocation
			changes["remove_location"] = removeLocation
		}
		if cmd.Flags().Changed("start") != cmd.Flags().Changed("end") {
			return usageErr(fmt.Errorf("--start and --end must be supplied together"))
		}
		if cmd.Flags().Changed("start") {
			start, e1 := time.Parse(time.RFC3339, newStart)
			end, e2 := time.Parse(time.RFC3339, newEnd)
			if e1 != nil || e2 != nil || !start.Before(end) {
				return usageErr(fmt.Errorf("--start and --end require RFC3339 offsets and end must be after start"))
			}
			if (p.BodyText != nil || locationOps > 0) && p.Subject != "" {
				return usageErr(fmt.Errorf("combined time and body/location edits require a separate subject change"))
			}
			p.Start, p.End = newStart, newEnd
			changes["start"], changes["end"] = newStart, newEnd
			if cmd.Flags().Changed("time-zone") {
				zone, err := cal.CanonicalTimeZone(updateZone)
				if err != nil {
					return usageErr(err)
				}
				p.TimeZone = zone
				changes["time_zone"] = zone
			}
		}
		if cmd.Flags().Changed("time-zone") && !cmd.Flags().Changed("start") {
			return usageErr(fmt.Errorf("--time-zone is only valid with --start and --end"))
		}
		if p.Start != "" && (p.BodyText != nil || locationOps > 0) && p.TimeZone == "Eastern Standard Time" {
			return usageErr(fmt.Errorf("combined time and content edits currently use Mountain time"))
		}
		if cmd.Flags().Changed("start-date") != cmd.Flags().Changed("end-date") {
			return usageErr(fmt.Errorf("--start-date and --end-date must be supplied together"))
		}
		if cmd.Flags().Changed("start-date") {
			if len(changes) != 0 || occurrence {
				return usageErr(fmt.Errorf("all-day date changes must be made separately from other changes"))
			}
			if _, _, err := cal.AllDayRange(newStartDate, newEndDate); err != nil {
				return usageErr(err)
			}
			p.StartDate, p.EndDate = newStartDate, newEndDate
			changes["start_date"], changes["end_date"] = newStartDate, newEndDate
		}
		if cmd.Flags().Changed("weekly-until") {
			date, err := time.Parse("2006-01-02", newWeeklyUntil)
			if err != nil || date.Year() < 1 || date.Format("2006-01-02") != newWeeklyUntil {
				return usageErr(fmt.Errorf("--weekly-until requires a valid YYYY-MM-DD date"))
			}
			p.WeeklyUntil = newWeeklyUntil
			changes["weekly_until"] = newWeeklyUntil
		}
		if occurrence {
			if p.Subject == "" || len(changes) != 1 {
				return usageErr(fmt.Errorf("--occurrence currently supports a subject change only"))
			}
			p.Occurrence = true
		}
		if selfMeetingUpdate {
			if occurrence || series || resetExceptions || p.Subject == "" || len(changes) != 1 {
				return usageErr(fmt.Errorf("--self-meeting currently supports a subject change only"))
			}
			p.SelfMeeting = true
		}
		if series {
			seriesTime := p.Start != "" || p.End != ""
			seriesUntil := p.WeeklyUntil != ""
			count := 0
			if p.Subject != "" {
				count++
			}
			if seriesTime {
				count++
			}
			if seriesUntil {
				count++
			}
			if occurrence || p.StartDate != "" || p.EndDate != "" || p.BodyText != nil || locationOps != 0 || count != 1 || p.TimeZone != "" && p.TimeZone != "Mountain Standard Time" {
				return usageErr(fmt.Errorf("--series supports one captured subject, time, or weekly end-date change"))
			}
			if resetExceptions != (seriesTime || seriesUntil) {
				return usageErr(fmt.Errorf("series time and end-date changes require --reset-exceptions; subject changes do not accept it"))
			}
			p.Series, p.ResetExceptions = true, resetExceptions
		} else if resetExceptions || p.WeeklyUntil != "" {
			return usageErr(fmt.Errorf("--reset-exceptions and --weekly-until require --series"))
		}
		if len(changes) == 0 {
			return usageErr(fmt.Errorf("provide a subject, body, time, or location change"))
		}
		if !updateExecute || flags.dryRun {
			scope := "single"
			if selfMeetingUpdate {
				scope = "self-only online event"
			} else if series {
				scope = "entire series"
			} else if occurrence {
				scope = "occurrence"
			}
			preview := map[string]any{"preview": true, "operation": "calendar update", "id": id, "expected_change_key": changeKey, "changes": changes, "version_check": "before write; not atomic", "sends_updates": false, "scope": scope}
			if resetExceptions {
				preview["exception_effect"] = "existing exceptions will be reset"
			}
			return flags.printJSON(cmd, preview)
		}
		if err := mutationInHarness(); err != nil {
			return err
		}
		event, err := calendarRead(cmd.Context(), id)
		if err != nil {
			return err
		}
		if event.ChangeKey != changeKey {
			return &transport.Error{Kind: transport.KindConflict}
		}
		if selfMeetingUpdate {
			// The operation adapter verifies organizer and attendee identity against
			// the current session immediately before preparing the write.
		} else if series {
			if !event.SafeWeeklyMountain() {
				return fmt.Errorf("series is not a verified Mountain-time weekly master; no changes made")
			}
			from, to, expected, err := event.WeeklyDeleteWindow()
			if err != nil {
				return fmt.Errorf("series is not an eligible attendee-free weekly master: %w", err)
			}
			expanded, err := calendarSeriesOccurrences(cmd.Context(), id, from, to)
			if err != nil {
				return err
			}
			if !expanded.SafeWeeklyDeletion(id, expected) {
				return fmt.Errorf("series expansion is incomplete or contains an unsupported occurrence; no changes made")
			}
			if p.Subject != "" && len(expanded.Exceptions) != 0 {
				return fmt.Errorf("series subject changes with exceptions are not yet supported; no changes made")
			}
			event, err = calendarRead(cmd.Context(), id)
			if err != nil {
				return err
			}
			if event.ChangeKey != changeKey {
				return &transport.Error{Kind: transport.KindConflict}
			}
			if !event.SafeWeeklyMountain() {
				return fmt.Errorf("series changed during verification; no changes made")
			}
		} else if occurrence {
			if !event.SafeOccurrence() {
				return fmt.Errorf("event is not a verified attendee-free timed Mountain-time occurrence; no changes made")
			}
		} else if !event.SafePersonal() {
			return fmt.Errorf("event is not a verified single personal appointment; no changes made")
		}
		if p.StartDate != "" && !event.SafeAllDayMountain() {
			return fmt.Errorf("event is not a verified all-day Mountain-time appointment; no changes made")
		}
		if p.Start != "" || p.End != "" {
			zone, ok := event.SafeTimedZone()
			if series {
				zone, ok = "Mountain Standard Time", event.SafeWeeklyMountain()
			}
			if !ok {
				return fmt.Errorf("event is not a verified timed appointment in a supported zone; no changes made")
			}
			if p.TimeZone != "" && p.TimeZone != zone {
				return fmt.Errorf("changing an event's time zone is not yet supported; no changes made")
			}
			if (p.BodyText != nil || locationOps > 0) && zone != "Mountain Standard Time" {
				return fmt.Errorf("combined time and content edits require a Mountain-time appointment; no changes made")
			}
			p.TimeZone = zone
		}
		p.Current = &event
		response, err := calendarRequest(cmd.Context(), "UpdateCalendarEvent", p)
		if err != nil {
			return err
		}
		updated, state := cal.ParseEvent(response.Body)
		if state != transport.EnvelopeOK || updated.ID != id {
			return &transport.Error{Kind: transport.KindAmbiguous, Status: response.Status}
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "updated": updated, "next_step": "Read the event again before a subsequent write; its change key can change after the update."})
	}}
	update.Flags().StringVar(&changeKey, "change-key", "", "Expected version returned by calendar get")
	update.Flags().StringVar(&newSubject, "subject", "", "Replacement appointment subject")
	update.Flags().StringVar(&newBody, "body", "", "Replacement plain-text appointment body")
	update.Flags().StringVar(&addLocation, "add-location", "", "Add a location while retaining existing ones")
	update.Flags().StringVar(&replaceLocation, "replace-location", "", "Existing location name to replace")
	update.Flags().StringVar(&withLocation, "with-location", "", "Replacement location name")
	update.Flags().StringVar(&removeLocation, "remove-location", "", "Existing location name to remove")
	update.Flags().StringVar(&newStart, "start", "", "Replacement start with RFC3339 offset")
	update.Flags().StringVar(&newEnd, "end", "", "Replacement end with RFC3339 offset")
	update.Flags().StringVar(&newStartDate, "start-date", "", "Replacement all-day start date (inclusive YYYY-MM-DD)")
	update.Flags().StringVar(&newEndDate, "end-date", "", "Replacement all-day end date (inclusive YYYY-MM-DD)")
	update.Flags().StringVar(&updateZone, "time-zone", "", "Expected event time zone; a time edit preserves the existing zone")
	update.Flags().BoolVar(&occurrence, "occurrence", false, "Update only a verified attendee-free timed Mountain-time occurrence")
	update.Flags().BoolVar(&series, "series", false, "Update the entire eligible weekly series using its master ID")
	update.Flags().StringVar(&newWeeklyUntil, "weekly-until", "", "Replacement inclusive weekly end date (YYYY-MM-DD)")
	update.Flags().BoolVar(&resetExceptions, "reset-exceptions", false, "Acknowledge that a series time or end-date change resets exceptions")
	update.Flags().BoolVar(&selfMeetingUpdate, "self-meeting", false, "Update the captured self-only online event")
	update.Flags().BoolVar(&updateExecute, "execute", false, "Apply the appointment changes")
	parent.AddCommand(update)
	var deleteKey string
	var deleteExecute, deleteSeries, deleteSelfMeeting bool
	del := &cobra.Command{Use: "delete <id>", Short: "Preview or delete a personal appointment; version check is not atomic", Long: "Deletes a single no-attendee appointment or an eligible weekly series after checking its change key. The checks and delete are separate requests and are not atomic.", Annotations: map[string]string{"pp:data-source": "live", "pp:happy-args": "id=example-value"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "calendar delete"})
		}
		id, err := oneID(args)
		if err != nil {
			return err
		}
		if deleteKey == "" {
			return usageErr(fmt.Errorf("--change-key is required"))
		}
		if deleteSeries && deleteSelfMeeting {
			return usageErr(fmt.Errorf("choose --series or --self-meeting, not both"))
		}
		if !deleteExecute || flags.dryRun {
			scope := "single"
			if deleteSeries {
				scope = "entire series"
			}
			if deleteSelfMeeting {
				scope = "self-only online event"
			}
			return flags.printJSON(cmd, map[string]any{"preview": true, "operation": "calendar delete", "id": id, "scope": scope, "expected_change_key": deleteKey, "version_check": "before write; not atomic", "sends_cancellations": false})
		}
		if err := mutationInHarness(); err != nil {
			return err
		}
		event, err := calendarRead(cmd.Context(), id)
		if err != nil {
			return err
		}
		if event.ChangeKey != deleteKey {
			return &transport.Error{Kind: transport.KindConflict}
		}
		if deleteSelfMeeting {
			// The operation adapter performs the account-bound self-only guard.
		} else if deleteSeries {
			from, to, expected, err := event.WeeklyDeleteWindow()
			if err != nil {
				return fmt.Errorf("series is not a verified attendee-free weekly master: %w", err)
			}
			expanded, err := calendarSeriesOccurrences(cmd.Context(), id, from, to)
			if err != nil {
				return err
			}
			if !expanded.SafeWeeklyDeletion(id, expected) {
				return fmt.Errorf("series expansion is incomplete or contains an unsupported occurrence; no deletion made")
			}
			event, err = calendarRead(cmd.Context(), id)
			if err != nil {
				return err
			}
			if event.ChangeKey != deleteKey {
				return &transport.Error{Kind: transport.KindConflict}
			}
			if !event.SafeWeeklySeries() {
				return fmt.Errorf("series changed during verification; no deletion made")
			}
		} else if !event.SafePersonal() {
			return fmt.Errorf("event is not a verified single personal appointment; no deletion made")
		}
		_, err = calendarRequest(cmd.Context(), "DeleteCalendarEvent", cal.Params{ID: id, Current: &event, Series: deleteSeries, SelfMeeting: deleteSelfMeeting})
		if err != nil {
			return err
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "deleted": true, "id": id})
	}}
	del.Flags().StringVar(&deleteKey, "change-key", "", "Expected version returned by calendar get")
	del.Flags().BoolVar(&deleteSeries, "series", false, "Delete the entire eligible weekly series using its master ID")
	del.Flags().BoolVar(&deleteSelfMeeting, "self-meeting", false, "Delete the captured self-only online event")
	del.Flags().BoolVar(&deleteExecute, "execute", false, "Delete this appointment")
	parent.AddCommand(del)
}

func contactRequest(ctx context.Context, action string, p people.Params) (transport.Response, error) {
	prefix, body := map[string]string{"get": "OUTLOOK_WEB_CONTACT_GET", "lookup-email": "OUTLOOK_WEB_PEOPLE_LOOKUP_EMAIL", "create": "OUTLOOK_WEB_CONTACT_CREATE", "update": "OUTLOOK_WEB_CONTACT_UPDATE", "delete": "OUTLOOK_WEB_CONTACT_DELETE"}[action], action != "delete"
	if prefix == "" {
		return transport.Response{}, fmt.Errorf("unsupported contact operation")
	}
	detail, request, err := outlookCapture(prefix, body)
	if err != nil {
		return transport.Response{}, err
	}
	current, _, err := currentSession()
	if err != nil {
		return transport.Response{}, err
	}
	session, op, err := people.Prepare(action, detail, request, current, p)
	if err != nil {
		return transport.Response{}, err
	}
	return outlookOperation(ctx, session, op)
}
func contactRead(ctx context.Context, id string) (people.Contact, error) {
	response, err := contactRequest(ctx, "get", people.Params{ID: id})
	if err != nil {
		return people.Contact{}, err
	}
	contact, state := people.ParseContact(response.Body, true)
	if state != transport.EnvelopeOK || contact.ID != id {
		return people.Contact{}, &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
	}
	return contact, nil
}

// pp:data-source live
func newOutlookPersonalContacts(flags *rootFlags) *cobra.Command {
	parent := &cobra.Command{Use: "contacts", Short: "Work with personal contacts"}
	var lookupEmail string
	lookup := &cobra.Command{Use: "lookup-email", Short: "Look up people by exact email address", Long: "Looks up an exact email address in People; results may include organization directory profiles and personal contacts.", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live", "pp:happy-args": "--email=example@example.invalid"}, RunE: func(cmd *cobra.Command, args []string) error {
		if len(args) != 0 {
			return usageErr(fmt.Errorf("contacts lookup-email takes --email only"))
		}
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "people lookup-email"})
		}
		if !people.ValidEmail(lookupEmail) {
			return usageErr(fmt.Errorf("--email must be one complete email address"))
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "people lookup-email", "email": lookupEmail, "scope": "people including directory profiles"})
		}
		response, err := contactRequest(cmd.Context(), "lookup-email", people.Params{Email: &lookupEmail})
		if err != nil {
			return err
		}
		values, state := people.ParseEmailLookup(response.Body)
		if state != transport.EnvelopeOK {
			return &transport.Error{Kind: transport.KindUpstream, Status: response.Status}
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "scope": "people including directory profiles", "results": values})
	}}
	lookup.Flags().StringVar(&lookupEmail, "email", "", "Exact email address to look up")
	parent.AddCommand(lookup)
	get := &cobra.Command{Use: "get <id>", Short: "Read a personal contact by ID", Annotations: map[string]string{"mcp:read-only": "true", "pp:data-source": "live", "pp:happy-args": "id=example-value"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "contacts get"})
		}
		id, err := oneID(args)
		if err != nil {
			return err
		}
		if flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "contacts get", "id": id})
		}
		c, err := contactRead(cmd.Context(), id)
		if err != nil {
			return err
		}
		return flags.printJSON(cmd, c)
	}}
	parent.AddCommand(get)
	var first, last, email, phone, company, notes string
	var execute bool
	create := &cobra.Command{Use: "create", Short: "Preview or create a personal contact", Long: "Creates a personal contact with a name and optional work email, business phone, company, and plain-text notes. Supply --execute to create it.", Annotations: map[string]string{"pp:data-source": "live"}, RunE: func(cmd *cobra.Command, args []string) error {
		if len(args) != 0 {
			return usageErr(fmt.Errorf("contacts create takes flags only"))
		}
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "contacts create"})
		}
		if strings.TrimSpace(first+" "+last) == "" {
			return usageErr(fmt.Errorf("supply --first or --last"))
		}
		p := people.Params{First: first, Last: last}
		values := map[string]string{"first": first, "last": last}
		for _, v := range []struct {
			name   string
			value  *string
			target **string
		}{{"email", &email, &p.Email}, {"phone", &phone, &p.Phone}, {"company", &company, &p.Company}, {"notes", &notes, &p.Notes}} {
			if cmd.Flags().Changed(v.name) {
				if strings.TrimSpace(*v.value) == "" {
					return usageErr(fmt.Errorf("--%s cannot be empty", v.name))
				}
				*v.target = v.value
				values[v.name] = *v.value
			}
		}
		if !execute || flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"preview": true, "operation": "contacts create", "fields": values})
		}
		if err := mutationInHarness(); err != nil {
			return err
		}
		response, err := contactRequest(cmd.Context(), "create", p)
		if err != nil {
			return err
		}
		c, state := people.ParseContact(response.Body, false)
		if state != transport.EnvelopeOK {
			return &transport.Error{Kind: transport.KindAmbiguous, Status: response.Status}
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "created": c})
	}}
	create.Flags().StringVar(&first, "first", "", "Contact first name")
	create.Flags().StringVar(&last, "last", "", "Contact last name")
	create.Flags().StringVar(&email, "email", "", "Work email address")
	create.Flags().StringVar(&phone, "phone", "", "Business phone number")
	create.Flags().StringVar(&company, "company", "", "Company name")
	create.Flags().StringVar(&notes, "notes", "", "Plain-text contact notes")
	create.Flags().BoolVar(&execute, "execute", false, "Create this personal contact")
	parent.AddCommand(create)
	var updateFirst, updateLast, updateEmail, updatePhone, updateCompany, updateNotes, key string
	var updateExecute, clearEmail, clearPhone, clearCompany, clearNotes bool
	update := &cobra.Command{Use: "update <id>", Short: "Preview or change personal contact fields; version check is not atomic", Long: "Changes names or sets/clears one work email, business phone, company, or plain-text note. Clearing a company is refused when its position contains other information. The change key check is not atomic.", Annotations: map[string]string{"pp:data-source": "live", "pp:happy-args": "id=example-value"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "contacts update"})
		}
		id, err := oneID(args)
		if err != nil {
			return err
		}
		if key == "" {
			return usageErr(fmt.Errorf("--change-key is required"))
		}
		changes := map[string]any{}
		if cmd.Flags().Changed("first") {
			changes["first"] = updateFirst
		}
		if cmd.Flags().Changed("last") {
			changes["last"] = updateLast
		}
		p := people.Params{ID: id, NameChange: cmd.Flags().Changed("first") || cmd.Flags().Changed("last")}
		for _, v := range []struct {
			name   string
			value  *string
			target **string
		}{{"email", &updateEmail, &p.Email}, {"phone", &updatePhone, &p.Phone}, {"company", &updateCompany, &p.Company}, {"notes", &updateNotes, &p.Notes}} {
			if cmd.Flags().Changed(v.name) {
				if strings.TrimSpace(*v.value) == "" {
					return usageErr(fmt.Errorf("--%s cannot be empty", v.name))
				}
				*v.target = v.value
				changes[v.name] = *v.value
			}
		}
		for _, v := range []struct {
			name       string
			clear, set bool
			target     *bool
		}{{"email", clearEmail, p.Email != nil, &p.ClearEmail}, {"phone", clearPhone, p.Phone != nil, &p.ClearPhone}, {"company", clearCompany, p.Company != nil, &p.ClearCompany}, {"notes", clearNotes, p.Notes != nil, &p.ClearNotes}} {
			if v.clear && v.set {
				return usageErr(fmt.Errorf("--%s and --clear-%s cannot be combined", v.name, v.name))
			}
			if v.clear {
				*v.target = true
				changes["clear_"+v.name] = true
			}
		}
		if len(changes) == 0 {
			return usageErr(fmt.Errorf("provide at least one contact field to change"))
		}
		if !updateExecute || flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"preview": true, "operation": "contacts update", "id": id, "expected_change_key": key, "changes": changes, "version_check": "before write; not atomic"})
		}
		if err := mutationInHarness(); err != nil {
			return err
		}
		old, err := contactRead(cmd.Context(), id)
		if err != nil {
			return err
		}
		if old.ChangeKey != key {
			return &transport.Error{Kind: transport.KindConflict}
		}
		if !old.SafeEditable() {
			return fmt.Errorf("contact is not an editable personal contact; no changes made")
		}
		f, l := old.First, old.Last
		if cmd.Flags().Changed("first") {
			f = updateFirst
		}
		if cmd.Flags().Changed("last") {
			l = updateLast
		}
		p.First, p.Last, p.Current = f, l, &old
		response, err := contactRequest(cmd.Context(), "update", p)
		if err != nil {
			return err
		}
		c, state := people.ParseContact(response.Body, false)
		if state != transport.EnvelopeOK || c.ID != id {
			return &transport.Error{Kind: transport.KindAmbiguous, Status: response.Status}
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "updated": c, "next_step": "Read the contact again before a subsequent write; its change key can change after the update."})
	}}
	update.Flags().StringVar(&key, "change-key", "", "Expected version returned by contacts get")
	update.Flags().StringVar(&updateFirst, "first", "", "Replacement first name")
	update.Flags().StringVar(&updateLast, "last", "", "Replacement last name")
	update.Flags().StringVar(&updateEmail, "email", "", "Work email")
	update.Flags().StringVar(&updatePhone, "phone", "", "Business phone")
	update.Flags().StringVar(&updateCompany, "company", "", "Company name")
	update.Flags().StringVar(&updateNotes, "notes", "", "Plain-text notes")
	update.Flags().BoolVar(&clearEmail, "clear-email", false, "Remove the matching work email")
	update.Flags().BoolVar(&clearPhone, "clear-phone", false, "Remove the matching business phone")
	update.Flags().BoolVar(&clearCompany, "clear-company", false, "Remove an otherwise empty company position")
	update.Flags().BoolVar(&clearNotes, "clear-notes", false, "Remove the matching plain-text note")
	update.Flags().BoolVar(&updateExecute, "execute", false, "Apply the contact changes")
	parent.AddCommand(update)
	var deleteKey string
	var deleteExecute bool
	del := &cobra.Command{Use: "delete <id>", Short: "Preview or delete a personal contact; version check is not atomic", Long: "Deletes a verified personal contact after checking its change key. The check and delete are separate requests and are not atomic.", Annotations: map[string]string{"pp:data-source": "live", "pp:happy-args": "id=example-value"}, RunE: func(cmd *cobra.Command, args []string) error {
		if flags.dryRun && requiredInputWaived() {
			return flags.printJSON(cmd, map[string]any{"dry_run": true, "operation": "contacts delete"})
		}
		id, err := oneID(args)
		if err != nil {
			return err
		}
		if deleteKey == "" {
			return usageErr(fmt.Errorf("--change-key is required"))
		}
		if !deleteExecute || flags.dryRun {
			return flags.printJSON(cmd, map[string]any{"preview": true, "operation": "contacts delete", "id": id, "expected_change_key": deleteKey, "version_check": "before write; not atomic"})
		}
		if err := mutationInHarness(); err != nil {
			return err
		}
		old, err := contactRead(cmd.Context(), id)
		if err != nil {
			return err
		}
		if old.ChangeKey != deleteKey {
			return &transport.Error{Kind: transport.KindConflict}
		}
		if !old.SafeDeletable() {
			return fmt.Errorf("contact is not a deletable personal contact; no deletion made")
		}
		response, err := contactRequest(cmd.Context(), "delete", people.Params{ID: id})
		if err != nil {
			return err
		}
		if response.Status != 204 {
			return &transport.Error{Kind: transport.KindAmbiguous, Status: response.Status}
		}
		return flags.printJSON(cmd, map[string]any{"source": "live", "deleted": true, "id": id})
	}}
	del.Flags().StringVar(&deleteKey, "change-key", "", "Expected version returned by contacts get")
	del.Flags().BoolVar(&deleteExecute, "execute", false, "Delete this personal contact")
	parent.AddCommand(del)
	return parent
}
