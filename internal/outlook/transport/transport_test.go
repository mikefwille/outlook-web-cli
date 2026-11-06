package transport

import (
	"context"
	"errors"
	"io"
	"net/http"
	"strings"
	"testing"
	"time"
)

type roundTripFunc func(*http.Request) (*http.Response, error)

func (f roundTripFunc) RoundTrip(r *http.Request) (*http.Response, error) { return f(r) }

func session() Session {
	return Session{AccountKey: "account-a", Origin: "https://outlook.example", ExpiresAt: time.Now().Add(time.Hour), Headers: http.Header{"Authorization": {"Bearer private-test-value"}, "Cookie": {"session=private-test-value"}}}
}

func okEnvelope([]byte) EnvelopeResult { return EnvelopeOK }

func response(code int, body string, h http.Header) *http.Response {
	return &http.Response{StatusCode: code, Header: h, Body: io.NopCloser(strings.NewReader(body))}
}

func TestPostReadRetriesRateLimitAndPreservesRequest(t *testing.T) {
	var calls int
	c, err := New("account-a", session(), roundTripFunc(func(r *http.Request) (*http.Response, error) {
		calls++
		b, _ := io.ReadAll(r.Body)
		if r.Method != "POST" || string(b) != `{"synthetic":true}` || r.Header.Get("Authorization") != "Bearer private-test-value" {
			t.Errorf("request not preserved")
		}
		if calls == 1 {
			return response(429, `{"error":"limited"}`, http.Header{"Retry-After": {"0"}}), nil
		}
		return response(200, `{"items":[]}`, nil), nil
	}))
	if err != nil {
		t.Fatal(err)
	}
	r, err := c.Do(context.Background(), "account-a", Operation{Method: "POST", Path: "/read", Body: []byte(`{"synthetic":true}`), Class: Read, Validate: okEnvelope})
	if err != nil || calls != 2 || string(r.Body) != `{"items":[]}` {
		t.Fatalf("result: calls=%d body=%q err=%v", calls, r.Body, err)
	}
}

func TestWriteIsNeverRetriedAndErrorsDoNotLeak(t *testing.T) {
	calls := 0
	c, err := New("account-a", session(), roundTripFunc(func(*http.Request) (*http.Response, error) {
		calls++
		return response(503, "private-test-value", http.Header{"Retry-After": {"0"}}), nil
	}))
	if err != nil {
		t.Fatal(err)
	}
	_, err = c.Do(context.Background(), "account-a", Operation{Method: "POST", Path: "/write", Class: Write, Validate: okEnvelope})
	var e *Error
	if !errors.As(err, &e) || e.Kind != KindAmbiguous || calls != 1 || strings.Contains(err.Error(), "private-test-value") {
		t.Fatalf("calls=%d err=%v", calls, err)
	}
}

func TestWriteDistinguishesUnknownAndRejectedOutcomes(t *testing.T) {
	c, err := New("account-a", session(), roundTripFunc(func(*http.Request) (*http.Response, error) { return response(200, `{}`, nil), nil }))
	if err != nil {
		t.Fatal(err)
	}
	for _, tc := range []struct {
		result EnvelopeResult
		want   Kind
	}{{EnvelopeError, KindAmbiguous}, {EnvelopeConflict, KindConflict}, {EnvelopeRejected, KindUpstream}} {
		_, err := c.Do(context.Background(), "account-a", Operation{Method: "POST", Path: "/write", Class: Write, Validate: func([]byte) EnvelopeResult { return tc.result }})
		var e *Error
		if !errors.As(err, &e) || e.Kind != tc.want {
			t.Errorf("result=%v error=%v", tc.result, err)
		}
	}
}

func TestRejectsForeignTargetsAndAccountMixing(t *testing.T) {
	calls := 0
	c, err := New("account-a", session(), roundTripFunc(func(*http.Request) (*http.Response, error) { calls++; return response(200, `{}`, nil), nil }))
	if err != nil {
		t.Fatal(err)
	}
	for _, path := range []string{"https://foreign.example/steal", "//foreign.example/x", "/../foreign", "/x?token=bad", `/x\foreign`} {
		_, err := c.Do(context.Background(), "account-a", Operation{Method: "GET", Path: path, Class: Read, Validate: okEnvelope})
		if err == nil {
			t.Errorf("accepted unsafe path %q", path)
		}
	}
	_, err = c.Do(context.Background(), "account-b", Operation{Method: "GET", Path: "/safe", Class: Read, Validate: okEnvelope})
	if err == nil || calls != 0 {
		t.Fatalf("cross-account request reached transport, calls=%d", calls)
	}
}

func TestRedirectDoesNotForwardSession(t *testing.T) {
	calls := 0
	c, err := New("account-a", session(), roundTripFunc(func(*http.Request) (*http.Response, error) {
		calls++
		return response(302, "", http.Header{"Location": {"https://foreign.example/file"}}), nil
	}))
	if err != nil {
		t.Fatal(err)
	}
	_, err = c.Do(context.Background(), "account-a", Operation{Method: "GET", Path: "/file", Class: Read, Validate: okEnvelope})
	if err == nil || calls != 1 {
		t.Fatalf("redirect followed or accepted: calls=%d err=%v", calls, err)
	}
}

func TestEnvelopeFailureAndPartialAreNotSuccess(t *testing.T) {
	c, err := New("account-a", session(), roundTripFunc(func(*http.Request) (*http.Response, error) {
		return response(200, `{"synthetic":"response"}`, nil), nil
	}))
	if err != nil {
		t.Fatal(err)
	}
	for _, tc := range []struct {
		result EnvelopeResult
		want   Kind
	}{{EnvelopeError, KindUpstream}, {EnvelopePartial, KindPartial}, {0, KindUpstream}} {
		_, err := c.Do(context.Background(), "account-a", Operation{Method: "POST", Path: "/read", Class: Read, Validate: func([]byte) EnvelopeResult { return tc.result }})
		var e *Error
		if !errors.As(err, &e) || e.Kind != tc.want {
			t.Errorf("result=%v err=%v", tc.result, err)
		}
	}
}

func TestExpiryCancellationAndResponseLimit(t *testing.T) {
	expired := session()
	expired.ExpiresAt = time.Now().Add(-time.Second)
	if _, err := New("account-a", expired, nil); err == nil {
		t.Fatal("accepted expired session")
	}
	c, err := New("account-a", session(), roundTripFunc(func(r *http.Request) (*http.Response, error) {
		if r.Context().Err() != nil {
			return nil, r.Context().Err()
		}
		return response(200, strings.Repeat("x", 33), nil), nil
	}))
	if err != nil {
		t.Fatal(err)
	}
	c.maxBody = 32
	_, err = c.Do(context.Background(), "account-a", Operation{Method: "GET", Path: "/read", Class: Read, Validate: okEnvelope})
	var e *Error
	if !errors.As(err, &e) || e.Kind != KindLimit {
		t.Fatalf("expected limit, got %v", err)
	}
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	_, err = c.Do(ctx, "account-a", Operation{Method: "GET", Path: "/read", Class: Read, Validate: okEnvelope})
	if !errors.As(err, &e) || e.Kind != KindCancelled {
		t.Fatalf("expected cancellation, got %v", err)
	}
}
