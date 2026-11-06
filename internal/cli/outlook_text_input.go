package cli

import (
	"fmt"
	"io"
	"os"
	"unicode/utf8"

	"github.com/spf13/cobra"
	"github.com/mikefwille/outlook-web-cli/internal/outlook/drafts"
)

// draftTextInput supports a bounded local file or stdin without exposing file
// paths through MCP. It retains exact user-supplied text including newlines.
func draftTextInput(cmd *cobra.Command, text, path string) (string, error) {
	if cmd.Flags().Changed("text") && cmd.Flags().Changed("text-file") {
		return "", usageErr(fmt.Errorf("choose --text or --text-file, not both"))
	}
	if !cmd.Flags().Changed("text-file") {
		return text, nil
	}
	if path == "" {
		return "", usageErr(fmt.Errorf("--text-file requires a path or - for stdin"))
	}
	var r io.Reader = cmd.InOrStdin()
	if path != "-" {
		f, err := os.Open(path)
		if err != nil {
			return "", fmt.Errorf("cannot read text input file")
		}
		defer f.Close()
		r = f
	}
	data, err := io.ReadAll(io.LimitReader(r, (1<<20)+1))
	if err != nil {
		return "", fmt.Errorf("cannot read text input")
	}
	if len(data) > 1<<20 || !utf8.Valid(data) {
		return "", usageErr(fmt.Errorf("text input must be valid UTF-8 and at most 1 MiB"))
	}
	return string(data), nil
}

// draftBodyInput accepts plain text, HTML, or both; when text is omitted the
// HTML's visible text is used for Outlook's plain-text compose companion.
func draftBodyInput(cmd *cobra.Command, text, textFile, htmlText, htmlFile string) (string, *string, error) {
	if cmd.Flags().Changed("html") && cmd.Flags().Changed("html-file") {
		return "", nil, usageErr(fmt.Errorf("choose --html or --html-file, not both"))
	}
	if cmd.Flags().Changed("text-file") && textFile == "-" && cmd.Flags().Changed("html-file") && htmlFile == "-" {
		return "", nil, usageErr(fmt.Errorf("only one body input may read from stdin"))
	}
	plain, err := draftTextInput(cmd, text, textFile)
	if err != nil {
		return "", nil, err
	}
	if !cmd.Flags().Changed("html") && !cmd.Flags().Changed("html-file") {
		return plain, nil, nil
	}
	source := htmlText
	if cmd.Flags().Changed("html-file") {
		if htmlFile == "" {
			return "", nil, usageErr(fmt.Errorf("--html-file requires a path or - for stdin"))
		}
		var reader io.Reader = cmd.InOrStdin()
		if htmlFile != "-" {
			f, e := os.Open(htmlFile)
			if e != nil {
				return "", nil, fmt.Errorf("cannot read HTML input file")
			}
			defer f.Close()
			reader = f
		}
		data, e := io.ReadAll(io.LimitReader(reader, (1<<20)+1))
		if e != nil {
			return "", nil, fmt.Errorf("cannot read HTML input")
		}
		source = string(data)
	}
	derived, err := drafts.HTMLText(source)
	if err != nil {
		return "", nil, usageErr(fmt.Errorf("HTML input must be valid UTF-8, contain no NUL and be at most 1 MiB"))
	}
	if !cmd.Flags().Changed("text") && !cmd.Flags().Changed("text-file") {
		plain = derived
	}
	return plain, &source, nil
}
