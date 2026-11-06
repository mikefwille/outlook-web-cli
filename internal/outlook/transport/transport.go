// Package transport is a capture-independent seam for an Outlook web client.
// It does not implement or assert any Outlook endpoint or session format.
package transport

import (
	"bytes"
	"context"
	"errors"
	"fmt"
	"io"
	"net"
	"net/http"
	"net/url"
	"strconv"
	"strings"
	"time"
)

type Kind string

const (
	KindAuth       Kind = "authentication_required"
	KindPermission Kind = "permission_denied"
	KindRateLimit  Kind = "rate_limited"
	KindUpstream   Kind = "upstream_error"
	KindPartial    Kind = "partial_failure"
	KindNetwork    Kind = "network_error"
	KindAmbiguous  Kind = "write_outcome_unknown"
	KindCancelled  Kind = "cancelled"
	KindLimit      Kind = "response_limit"
	KindInput      Kind = "invalid_input"
	KindConflict   Kind = "version_conflict"
)

// Error intentionally excludes URLs, upstream bodies, headers and wrapped
// errors: any of those can contain account data or session material.
type Error struct {
	Kind       Kind
	Status     int
	RetryAfter time.Duration
}

func (e *Error) Error() string {
	if e.Status != 0 {
		return fmt.Sprintf("outlook web: %s (HTTP %d)", e.Kind, e.Status)
	}
	return "outlook web: " + string(e.Kind)
}

type Session struct {
	AccountKey string
	Origin     string
	ExpiresAt  time.Time
	Headers    http.Header
}

// Class is semantic, not inferred from the HTTP method. A POST can be a read.
type Class uint8

const (
	Read Class = iota + 1
	Write
)

// EnvelopeResult lets a service-specific decoder detect errors inside HTTP 2xx
// without allowing raw server error text into diagnostic messages.
type EnvelopeResult uint8

const (
	EnvelopeOK EnvelopeResult = iota + 1
	EnvelopeError
	EnvelopePartial
	EnvelopeRejected
	EnvelopeConflict
)

type Operation struct {
	Method   string
	Path     string
	Query    url.Values
	Headers  http.Header
	Body     []byte
	Class    Class
	Validate func([]byte) EnvelopeResult
}

type Response struct {
	Status int
	Body   []byte
}

type Client struct {
	account string
	origin  *url.URL
	expires time.Time
	headers http.Header
	http    *http.Client
	maxBody int64
}

const maxResponseBytes int64 = 16 << 20

// New binds a client to one account and HTTPS origin. rt may be supplied for
// testing or a selected HTTP transport. The caller must keep Session private.
func New(account string, session Session, rt http.RoundTripper) (*Client, error) {
	if account == "" || session.AccountKey != account || session.ExpiresAt.IsZero() || !time.Now().Before(session.ExpiresAt) || len(session.Headers) == 0 {
		return nil, &Error{Kind: KindAuth}
	}
	u, err := url.Parse(session.Origin)
	if err != nil || u.Scheme != "https" || u.Host == "" || u.User != nil || u.Path != "" || u.RawQuery != "" || u.Fragment != "" || !validHeaders(session.Headers, true, nil) {
		return nil, &Error{Kind: KindInput}
	}
	if rt == nil {
		rt = http.DefaultTransport
	}
	headers := make(http.Header, len(session.Headers))
	for k, values := range session.Headers {
		for _, value := range values {
			headers.Add(k, value)
		}
	}
	return &Client{
		account: account, origin: u, expires: session.ExpiresAt, headers: headers, maxBody: maxResponseBytes,
		http: &http.Client{Transport: rt, Timeout: 30 * time.Second, CheckRedirect: func(*http.Request, []*http.Request) error { return http.ErrUseLastResponse }},
	}, nil
}

func (c *Client) Do(ctx context.Context, account string, op Operation) (Response, error) {
	if c == nil || account != c.account || !time.Now().Before(c.expires) {
		return Response{}, &Error{Kind: KindAuth}
	}
	if ctx == nil || op.Validate == nil || (op.Class != Read && op.Class != Write) || !validMethod(op.Method) || !validPath(op.Path) || !validHeaders(op.Headers, false, c.headers) {
		return Response{}, &Error{Kind: KindInput}
	}
	u := *c.origin
	u.Path = op.Path
	u.RawQuery = op.Query.Encode()
	for attempt := 0; ; attempt++ {
		req, err := http.NewRequestWithContext(ctx, op.Method, u.String(), bytes.NewReader(op.Body))
		if err != nil {
			return Response{}, &Error{Kind: KindInput}
		}
		req.Header = op.Headers.Clone()
		if req.Header == nil {
			req.Header = make(http.Header)
		}
		for k, values := range c.headers {
			req.Header[k] = append([]string(nil), values...)
		}
		res, err := c.http.Do(req)
		if err != nil {
			if op.Class == Write {
				return Response{}, &Error{Kind: KindAmbiguous}
			}
			if errors.Is(ctx.Err(), context.Canceled) || errors.Is(ctx.Err(), context.DeadlineExceeded) {
				return Response{}, &Error{Kind: KindCancelled}
			}
			var netErr net.Error
			if errors.As(err, &netErr) && netErr.Timeout() {
				return Response{}, &Error{Kind: KindNetwork}
			}
			return Response{}, &Error{Kind: KindNetwork}
		}
		body, readErr := io.ReadAll(io.LimitReader(res.Body, c.maxBody+1))
		_ = res.Body.Close()
		if readErr != nil {
			if op.Class == Write {
				return Response{}, &Error{Kind: KindAmbiguous}
			}
			return Response{}, &Error{Kind: KindNetwork}
		}
		if int64(len(body)) > c.maxBody {
			if op.Class == Write {
				return Response{}, &Error{Kind: KindAmbiguous, Status: res.StatusCode}
			}
			return Response{}, &Error{Kind: KindLimit, Status: res.StatusCode}
		}
		if (res.StatusCode == http.StatusTooManyRequests || res.StatusCode == http.StatusServiceUnavailable) && op.Class == Read && attempt < 2 {
			if err := wait(ctx, retryDelay(res.Header.Get("Retry-After"), attempt)); err != nil {
				return Response{}, &Error{Kind: KindCancelled}
			}
			continue
		}
		if res.StatusCode < 200 || res.StatusCode >= 300 {
			if op.Class == Write && res.StatusCode >= 500 {
				return Response{}, &Error{Kind: KindAmbiguous, Status: res.StatusCode}
			}
			return Response{}, statusError(res.StatusCode, res.Header.Get("Retry-After"))
		}
		out := Response{Status: res.StatusCode, Body: body}
		switch op.Validate(body) {
		case EnvelopeOK:
			return out, nil
		case EnvelopePartial:
			return out, &Error{Kind: KindPartial, Status: res.StatusCode}
		case EnvelopeRejected:
			return Response{}, &Error{Kind: KindUpstream, Status: res.StatusCode}
		case EnvelopeConflict:
			return Response{}, &Error{Kind: KindConflict, Status: res.StatusCode}
		default:
			if op.Class == Write {
				return Response{}, &Error{Kind: KindAmbiguous, Status: res.StatusCode}
			}
			return Response{}, &Error{Kind: KindUpstream, Status: res.StatusCode}
		}
	}
}

func validMethod(s string) bool {
	switch s {
	case http.MethodGet, http.MethodPost, http.MethodPut, http.MethodPatch, http.MethodDelete:
		return true
	}
	return false
}

func validPath(s string) bool {
	if !strings.HasPrefix(s, "/") || strings.HasPrefix(s, "//") || strings.ContainsAny(s, "?#\\\r\n") {
		return false
	}
	for _, segment := range strings.Split(s, "/") {
		if segment == "." || segment == ".." {
			return false
		}
	}
	return true
}

func validHeaders(h http.Header, session bool, credentialHeaders http.Header) bool {
	for key, values := range h {
		if key == "" || http.CanonicalHeaderKey(key) == "" {
			return false
		}
		for _, ch := range key {
			if !(ch >= 'a' && ch <= 'z' || ch >= 'A' && ch <= 'Z' || ch >= '0' && ch <= '9' || strings.ContainsRune("!#$%&'*+-.^_`|~", ch)) {
				return false
			}
		}
		switch strings.ToLower(key) {
		case "host", "content-length", "transfer-encoding", "connection", "proxy-authorization":
			return false
		case "authorization", "cookie":
			if !session {
				return false
			}
		}
		if !session {
			for credentialKey := range credentialHeaders {
				if strings.EqualFold(key, credentialKey) {
					return false
				}
			}
		}
		for _, v := range values {
			if strings.ContainsAny(v, "\r\n\x00") {
				return false
			}
		}
	}
	return true
}

func statusError(status int, after string) *Error {
	e := &Error{Kind: KindUpstream, Status: status}
	switch status {
	case http.StatusUnauthorized:
		e.Kind = KindAuth
	case http.StatusForbidden:
		e.Kind = KindPermission
	case http.StatusTooManyRequests:
		e.Kind = KindRateLimit
		e.RetryAfter = retryDelay(after, 0)
	}
	return e
}

func retryDelay(value string, attempt int) time.Duration {
	const capDelay = 2 * time.Second
	if seconds, err := strconv.Atoi(strings.TrimSpace(value)); err == nil && seconds >= 0 {
		if seconds > int(capDelay/time.Second) {
			return capDelay
		}
		return time.Duration(seconds) * time.Second
	}
	if at, err := http.ParseTime(value); err == nil {
		d := time.Until(at)
		if d < 0 {
			return 0
		}
		if d > capDelay {
			return capDelay
		}
		return d
	}
	return time.Duration(1<<attempt) * 100 * time.Millisecond
}

func wait(ctx context.Context, d time.Duration) error {
	t := time.NewTimer(d)
	defer t.Stop()
	select {
	case <-ctx.Done():
		return ctx.Err()
	case <-t.C:
		return nil
	}
}
