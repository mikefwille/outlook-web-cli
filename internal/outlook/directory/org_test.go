package directory

import (
	"bytes"
	"io"
	"mime/multipart"
	"net/textproto"
	"strings"
	"testing"
)

func syntheticMultipart(t *testing.T, docs ...string) []byte {
	t.Helper()
	var body bytes.Buffer
	w := multipart.NewWriter(&body)
	if err := w.SetBoundary("graphql"); err != nil {
		t.Fatal(err)
	}
	for _, doc := range docs {
		header := textproto.MIMEHeader{}
		header.Set("Content-Type", "application/json")
		part, err := w.CreatePart(header)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := io.WriteString(part, doc); err != nil {
			t.Fatal(err)
		}
	}
	if err := w.Close(); err != nil {
		t.Fatal(err)
	}
	return body.Bytes()
}

func TestSyntheticOrgMultipartRequiresCompleteMatchingIdentity(t *testing.T) {
	id := Identity{ObjectID: "00000000-0000-0000-0000-000000000001", Email: "person@example.invalid", UPN: "person@example.invalid"}
	first := `{"data":{"orgexplorer":{"personInFocusV2":{"aadObjectId":"00000000-0000-0000-0000-000000000001","email":"person@example.invalid","manager":{"nodes":[{"id":"manager-id","aadObjectId":"manager-object","fullName":"Sample Manager"}]},"directs":{"edges":[{"node":{"id":"report-id","aadObjectId":"report-object","fullName":"Sample Report"}}],"pageInfo":{"endCursor":"next","hasNextPage":true}}}}},"hasNext":true}`
	last := `{"incremental":[{"path":["orgexplorer","personInFocusV2"],"data":{"location":"synthetic"}}],"hasNext":false}`
	data := syntheticMultipart(t, first, last)
	got, err := ParseOrg(data, id)
	if err != nil || len(got.Manager) != 1 || len(got.Reports) != 1 || !got.HasNextPage || got.EndCursor != "next" || !strings.Contains(string(got.Reports[0]), "Sample Report") {
		t.Fatal("complete synthetic result was not preserved")
	}
	for _, incomplete := range [][]byte{
		syntheticMultipart(t, first),
		syntheticMultipart(t, first, `{"errors":[{"message":"error"}],"hasNext":false}`),
		bytes.Replace(data, []byte("person@example.invalid"), []byte("other@example.invalid"), 1),
	} {
		if _, err := ParseOrg(incomplete, id); err == nil {
			t.Fatal("incomplete or foreign response was accepted")
		}
	}
}
