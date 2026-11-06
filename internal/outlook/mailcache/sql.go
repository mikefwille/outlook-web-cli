package mailcache

import (
	"context"
	"database/sql"
	"fmt"
	"math"
	"net/url"
	"strings"
	"time"
	"unicode/utf8"

	"modernc.org/sqlite"
)

type QueryResult struct {
	Columns   []string `json:"columns"`
	Rows      [][]any  `json:"rows"`
	Truncated bool     `json:"truncated"`
}

// Query allows a single bounded SELECT over published snapshots only. A
// dedicated read-only SQLite connection and authorizer deny writes, schema
// access, ATTACH, staging rows, arbitrary functions, and recursive queries.
func (s *Store) Query(ctx context.Context, query string, limit int) (QueryResult, error) {
	bad := func() (QueryResult, error) {
		return QueryResult{}, fmt.Errorf("SQL must be one bounded read-only SELECT over items and heads")
	}
	q := strings.TrimSpace(query)
	fields := strings.Fields(q)
	if len(fields) == 0 || !strings.EqualFold(fields[0], "select") || len(q) > 4096 || strings.Contains(q, ";") || strings.Contains(q, "--") || strings.Contains(q, "/*") || strings.Contains(q, "*/") || limit < 1 || limit > 100 {
		return bad()
	}
	u := url.URL{Scheme: "file", Path: s.path, RawQuery: "mode=ro"}
	db, err := sql.Open("sqlite", u.String())
	if err != nil {
		return QueryResult{}, err
	}
	defer db.Close()
	db.SetMaxOpenConns(1)
	ctx, cancel := context.WithTimeout(ctx, 5*time.Second)
	defer cancel()
	conn, err := db.Conn(ctx)
	if err != nil {
		return QueryResult{}, err
	}
	defer conn.Close()
	allowFunctions := map[string]bool{"count": true, "min": true, "max": true, "sum": true, "avg": true, "total": true, "length": true, "lower": true, "upper": true, "substr": true, "substring": true, "coalesce": true, "ifnull": true, "nullif": true, "json_extract": true, "json_type": true, "json_valid": true, "date": true, "datetime": true, "strftime": true}
	if err = conn.Raw(func(driver any) error {
		a, ok := driver.(sqlite.AuthorizerRegisterer)
		if !ok {
			return fmt.Errorf("SQLite read authorizer unavailable")
		}
		return a.RegisterAuthorizer(func(action sqlite.AuthorizerActionCode, arg1, arg2, database, trigger string) sqlite.AuthorizerReturnCode {
			switch action {
			case sqlite.AuthSelect:
				return sqlite.AuthorizerOK
			case sqlite.AuthRead:
				if (database == "main" || database == "" && arg2 == "") && (arg1 == "items" || arg1 == "heads") {
					return sqlite.AuthorizerOK
				}
			case sqlite.AuthFunction:
				if allowFunctions[strings.ToLower(arg2)] {
					return sqlite.AuthorizerOK
				}
			}
			return sqlite.AuthorizerDeny
		})
	}); err != nil {
		return QueryResult{}, err
	}
	rows, err := conn.QueryContext(ctx, "SELECT * FROM ("+q+") LIMIT ?", limit+1)
	if err != nil {
		return QueryResult{}, fmt.Errorf("query rejected or invalid")
	}
	defer rows.Close()
	columns, err := rows.Columns()
	if err != nil || len(columns) > 64 {
		return bad()
	}
	out := QueryResult{Columns: columns, Rows: [][]any{}}
	bytes := 0
	for rows.Next() {
		if len(out.Rows) == limit {
			out.Truncated = true
			break
		}
		values := make([]any, len(columns))
		ptrs := make([]any, len(columns))
		for i := range values {
			ptrs[i] = &values[i]
		}
		if err = rows.Scan(ptrs...); err != nil {
			return QueryResult{}, fmt.Errorf("query result could not be read")
		}
		for i, v := range values {
			switch x := v.(type) {
			case []byte:
				if !utf8.Valid(x) {
					return QueryResult{}, fmt.Errorf("binary SQL result is not supported")
				}
				values[i] = string(x)
				bytes += len(x)
			case string:
				bytes += len(x)
			case float64:
				if math.IsNaN(x) || math.IsInf(x, 0) {
					return QueryResult{}, fmt.Errorf("non-finite SQL result is not supported")
				}
			}
		}
		if bytes > 4<<20 {
			return QueryResult{}, fmt.Errorf("SQL result exceeds 4 MiB")
		}
		out.Rows = append(out.Rows, values)
	}
	if err = rows.Err(); err != nil {
		return QueryResult{}, fmt.Errorf("query did not complete")
	}
	return out, nil
}
