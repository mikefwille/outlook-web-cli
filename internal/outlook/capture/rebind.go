package capture

import (
	"encoding/base64"
	"encoding/json"
	"strings"
	"time"

	"github.com/mikefwille/outlook-web-cli/internal/outlook/transport"
)

// RebindRead uses fresh same-account, same-audience browser credentials while
// preserving the operation's independently observed routing and envelope.
// It never changes the request origin or follows an external redirect.
func RebindRead(observed, current transport.Session) (transport.Session, error) {
	fail := func() (transport.Session, error) {
		return transport.Session{}, &transport.Error{Kind: transport.KindAuth}
	}
	if observed.Origin != "https://outlook.cloud.microsoft" || current.Origin != observed.Origin || observed.AccountKey == "" || current.AccountKey != observed.AccountKey || !time.Now().Before(current.ExpiresAt) {
		return fail()
	}
	if observed.Headers.Get("X-AnchorMailbox") == "" || observed.Headers.Get("X-AnchorMailbox") != current.Headers.Get("X-AnchorMailbox") {
		return fail()
	}
	a, ok1 := jwtAudience(observed.Headers.Get("Authorization"))
	b, ok2 := jwtAudience(current.Headers.Get("Authorization"))
	if !ok1 || !ok2 || a != b || current.Headers.Get("Cookie") == "" {
		return fail()
	}
	observed.Headers = observed.Headers.Clone()
	observed.Headers.Set("Authorization", current.Headers.Get("Authorization"))
	observed.Headers.Set("Cookie", current.Headers.Get("Cookie"))
	observed.ExpiresAt = current.ExpiresAt
	return observed, nil
}

func jwtAudience(auth string) (string, bool) {
	parts := strings.SplitN(auth, " ", 2)
	if len(parts) != 2 || !strings.EqualFold(parts[0], "Bearer") {
		return "", false
	}
	seg := strings.Split(parts[1], ".")
	if len(seg) != 3 {
		return "", false
	}
	p, err := base64.RawURLEncoding.DecodeString(seg[1])
	if err != nil {
		return "", false
	}
	var claims struct {
		Audience string `json:"aud"`
	}
	if json.Unmarshal(p, &claims) != nil || claims.Audience == "" {
		return "", false
	}
	return claims.Audience, true
}
