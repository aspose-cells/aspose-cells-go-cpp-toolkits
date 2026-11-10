package editor

import (
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
)

// Option configures how an edit entry point loads its source workbook.
type Option func(*options)

type options struct {
	load cells.LoadOptions
}

func defaultOptions() *options { return &options{} }

func applyOptions(cfg *options, opts []Option) {
	for _, opt := range opts {
		if opt != nil {
			opt(cfg)
		}
	}
}

// WithPassword opens a password-protected source workbook with password.
//
// This is the read half of the encryption workflow: create an encrypted
// workbook with Encrypt, then edit or inspect it with WithPassword.
func WithPassword(password string) Option {
	return func(o *options) { o.load.Password = password }
}
