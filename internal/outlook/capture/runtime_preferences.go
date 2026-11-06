package capture

import "encoding/json"

var webPreferenceActions = map[string]string{"inbox_rules_list": "GetInboxRule", "preferences_auto_replies": "GetMailboxAutoReplyConfiguration", "preference_focused_update": "UpdateUserConfiguration", "rule_create": "NewInboxRule", "rule_enable": "EnableInboxRule", "rule_disable": "DisableInboxRule", "rule_update": "SetInboxRule", "rule_delete": "RemoveInboxRule", "rules_reorder": "SetInboxAndSweepRules", "master_category_create": "UpdateMasterCategoryList", "master_category_update": "UpdateMasterCategoryList", "master_category_delete": "UpdateMasterCategoryList", "master_category_rename_items": "ApplyCategoryOperation", "master_category_delete_items": "ApplyCategoryOperation"}

func validPreferenceTemplate(name string, t RuntimeWebTemplate) (bool, bool) {
	action := webPreferenceActions[name]
	if action == "" {
		return false, false
	}
	q := t.Query
	valid := t.Method == "POST" && t.Path == "/owa/service.svc" && t.HeaderCarrier && json.Valid(t.Body) && len(q) == 3 && len(q["action"]) == 1 && q.Get("action") == action && len(q["app"]) == 1 && q.Get("app") == "Mail" && len(q["n"]) == 1 && digits.MatchString(q.Get("n"))
	return true, valid
}
