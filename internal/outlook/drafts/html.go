package drafts

import (
	"encoding/json"
	"strings"
	"unicode/utf8"

	nethtml "golang.org/x/net/html"
)

// HTMLText derives the plain-text compose companion while retaining the exact
// supplied HTML separately. It is not an HTML sanitizer.
func HTMLText(source string) (string, error) {
	if len(source) > 1<<20 || !utf8.ValidString(source) || strings.ContainsRune(source, 0) {
		return "", ErrInvalid
	}
	root, err := nethtml.Parse(strings.NewReader(source))
	if err != nil {
		return "", ErrInvalid
	}
	var out strings.Builder
	nodes := 0
	var walk func(*nethtml.Node, int) bool
	walk = func(n *nethtml.Node, depth int) bool {
		nodes++
		if nodes > 100000 || depth > 512 {
			return false
		}
		if n.Type == nethtml.ElementNode && (n.Data == "head" || n.Data == "script" || n.Data == "style") {
			return true
		}
		if n.Type == nethtml.TextNode {
			out.WriteString(n.Data)
		}
		if n.Type == nethtml.ElementNode && n.Data == "br" {
			out.WriteByte('\n')
		}
		for child := n.FirstChild; child != nil; child = child.NextSibling {
			if !walk(child, depth+1) {
				return false
			}
		}
		if n.Type == nethtml.ElementNode {
			switch n.Data {
			case "p", "div", "li", "blockquote", "h1", "h2", "h3", "h4", "h5", "h6", "tr", "pre":
				out.WriteByte('\n')
			}
		}
		return true
	}
	if !walk(root, 0) {
		return "", ErrInvalid
	}
	return strings.TrimSpace(out.String()), nil
}

// SetHTMLBody replaces exactly the HTML body in an already-validated SaveOnly
// create or update request. Other requested fields remain unchanged.
func SetHTMLBody(encoded []byte, action, source string) ([]byte, error) {
	if _, err := HTMLText(source); err != nil {
		return nil, err
	}
	var root map[string]any
	if json.Unmarshal(encoded, &root) != nil || root["__type"] != action+"JsonRequest:#Exchange" {
		return nil, ErrInvalid
	}
	body, ok := root["Body"].(map[string]any)
	if !ok || body["MessageDisposition"] != "SaveOnly" {
		return nil, ErrInvalid
	}
	var item map[string]any
	switch action {
	case "CreateItem":
		items, ok := body["Items"].([]any)
		if !ok || len(items) != 1 {
			return nil, ErrInvalid
		}
		item, ok = items[0].(map[string]any)
		if !ok {
			return nil, ErrInvalid
		}
	case "UpdateItem":
		changes, ok := body["ItemChanges"].([]any)
		if !ok || len(changes) != 1 {
			return nil, ErrInvalid
		}
		change, ok := changes[0].(map[string]any)
		if !ok {
			return nil, ErrInvalid
		}
		updates, ok := change["Updates"].([]any)
		if !ok {
			return nil, ErrInvalid
		}
		count := 0
		for _, u := range updates {
			v, ok := u.(map[string]any)
			if !ok {
				return nil, ErrInvalid
			}
			path, ok := v["Path"].(map[string]any)
			if !ok {
				continue
			}
			if path["FieldURI"] == "Body" {
				count++
				item, ok = v["Item"].(map[string]any)
				if !ok {
					return nil, ErrInvalid
				}
			}
		}
		if count != 1 {
			return nil, ErrInvalid
		}
	default:
		return nil, ErrInvalid
	}
	if item["__type"] != "Message:#Exchange" {
		return nil, ErrInvalid
	}
	content, ok := item["Body"].(map[string]any)
	if !ok || content["BodyType"] != "HTML" {
		return nil, ErrInvalid
	}
	content["Value"] = source
	return json.Marshal(root)
}
