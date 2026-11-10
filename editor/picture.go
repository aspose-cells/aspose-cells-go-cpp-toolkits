package editor

import (
	"fmt"

	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	cells "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/internal/aspose/cells"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// PictureOption configures AddPicture.
type PictureOption func(*pictureConfig)

type pictureConfig struct {
	rows int
	cols int
}

func defaultPictureConfig() *pictureConfig {
	return &pictureConfig{rows: 1, cols: 1}
}

// WithPictureSpan sets the inclusive zero-based cell rectangle the picture is
// scaled to: it starts at the AddPicture cell and spans rows rows and cols
// columns. Both values must be at least 1.
func WithPictureSpan(rows, cols int) PictureOption {
	return func(c *pictureConfig) { c.rows, c.cols = rows, cols }
}

// AddPicture creates a WorksheetAction that embeds data as an image anchored at
// the zero-based cell (row, column). By default the picture covers that one
// cell; WithPictureSpan stretches it over a larger rectangle.
//
// data must be a decodable image (PNG, JPEG, ...). The engine rejects anything
// else, which surfaces as ErrPictureAddFailed.
//
// Example:
//
//	editor.InWorksheet(0, editor.AddPicture(1, 1, pngBytes,
//		editor.WithPictureSpan(4, 3)))
func AddPicture(row, column int, data []byte, opts ...PictureOption) WorksheetAction {
	cfg := defaultPictureConfig()
	for _, opt := range opts {
		if opt != nil {
			opt(cfg)
		}
	}
	return func(worksheet *asposecells.Worksheet) error {
		if cfg.rows < 1 || cfg.cols < 1 {
			return fmt.Errorf("picture span %dx%d: %w", cfg.rows, cfg.cols, toolkiterrors.ErrInvalidCount)
		}
		return cells.AddPictureAt(
			worksheet,
			row, column,
			row+cfg.rows-1, column+cfg.cols-1,
			data,
		)
	}
}
