package exchange

import (
	"encoding/json"
	"fmt"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

// IncludeReplyTo requests the observed ReplyTo property on the same bounded
// GetItem shape, so its absence in the response means no explicit Reply-To.
func IncludeReplyTo(op transport.Operation) (transport.Operation, error) {
	return includeMessageProperty(op, "ReplyTo")
}

// IncludeCategories explicitly requests the current message categories.
func IncludeCategories(op transport.Operation) (transport.Operation, error) {
	return includeMessageProperty(op, "Categories")
}

func includeMessageProperty(op transport.Operation, field string) (transport.Operation, error) {
	if op.Class != transport.Read || op.Method != "POST" || op.Path != "/owa/service.svc" || op.Query.Get("action") != "GetItem" || len(op.Body) == 0 {
		return transport.Operation{}, fmt.Errorf("unsupported message read shape")
	}
	var root map[string]any
	if json.Unmarshal(op.Body, &root) != nil {
		return transport.Operation{}, fmt.Errorf("unsupported message read shape")
	}
	body, ok := root["Body"].(map[string]any)
	if !ok {
		return transport.Operation{}, fmt.Errorf("unsupported message read shape")
	}
	shape, ok := body["ItemShape"].(map[string]any)
	if !ok {
		return transport.Operation{}, fmt.Errorf("unsupported message read shape")
	}
	props, ok := shape["AdditionalProperties"].([]any)
	if !ok || len(props) > 100 {
		return transport.Operation{}, fmt.Errorf("unsupported message read shape")
	}
	for _, v := range props {
		if p, ok := v.(map[string]any); ok && p["FieldURI"] == field {
			return op, nil
		}
	}
	shape["AdditionalProperties"] = append(props, map[string]any{"__type": "PropertyUri:#Exchange", "FieldURI": field})
	b, err := json.Marshal(root)
	if err != nil {
		return transport.Operation{}, err
	}
	op.Body = b
	return op, nil
}
