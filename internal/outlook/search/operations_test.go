package search

import (
	"strings"
	"testing"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

func TestSearchResponseRejectsErrorsAndReportsPartial(t *testing.T) {
	const good = `{"EntitySets":[{"IsPartial":false,"Properties":{"HasParseException":false},"ResultSets":[{"Total":0,"MoreResultsAvailable":false}]}]}`
	if Validate([]byte(good)) != transport.EnvelopeOK {
		t.Fatal("valid empty search rejected")
	}
	if Validate([]byte(strings.Replace(good, `"IsPartial":false`, `"IsPartial":true`, 1))) != transport.EnvelopePartial {
		t.Fatal("partial search hidden")
	}
	if Validate([]byte(strings.Replace(good, `"HasParseException":false`, `"HasParseException":true`, 1))) == transport.EnvelopeOK {
		t.Fatal("invalid query accepted")
	}
}
