package cells

import (
	"fmt"

	toolkiterrors "github.com/aspose-cells/aspose-cells-go-cpp-toolkits/v26/errors"
	asposecells "github.com/aspose-cells/aspose-cells-go-cpp/v26"
)

// Pictures returns the worksheet's picture collection. A nil or null handle
// becomes an error rather than a collection that reports zero pictures, so a
// caller can tell "this sheet has no pictures" from "the handle is unusable".
func Pictures(ws *asposecells.Worksheet) (*asposecells.PictureCollection, error) {
	pics, err := ws.GetPictures()
	if err != nil {
		return nil, err
	}
	if pics == nil {
		return nil, fmt.Errorf("worksheet returned no picture collection: %w", toolkiterrors.ErrPictureAddFailed)
	}
	null, err := pics.IsNull()
	if err != nil {
		return nil, err
	}
	if null {
		return nil, fmt.Errorf("worksheet returned no picture collection: %w", toolkiterrors.ErrPictureAddFailed)
	}
	return pics, nil
}

// AddPictureAt embeds data as an image on the worksheet, scaled to span the
// inclusive zero-based cell rectangle (topRow, leftColumn)-(bottomRow,
// rightColumn).
//
// The rectangle must be the same shape as the chart placement rectangle: top
// and left must not exceed bottom and right, and all four must be on the grid.
// Data must be a decodable image; the engine rejects anything else, which
// surfaces here as ErrPictureAddFailed.
func AddPictureAt(ws *asposecells.Worksheet, topRow, leftColumn, bottomRow, rightColumn int, data []byte) error {
	if len(data) == 0 {
		return fmt.Errorf("picture data is empty: %w", toolkiterrors.ErrPictureAddFailed)
	}
	if topRow < 0 || leftColumn < 0 || bottomRow < 0 || rightColumn < 0 {
		return fmt.Errorf("picture rectangle (%d,%d)-(%d,%d) has a negative edge: %w",
			topRow, leftColumn, bottomRow, rightColumn, toolkiterrors.ErrInvalidRange)
	}
	if topRow > bottomRow || leftColumn > rightColumn {
		return fmt.Errorf("picture rectangle (%d,%d)-(%d,%d) is reversed: %w",
			topRow, leftColumn, bottomRow, rightColumn, toolkiterrors.ErrInvalidRange)
	}
	pics, err := Pictures(ws)
	if err != nil {
		return err
	}
	index, err := pics.Add_Int_Int_Int_Int_Stream(
		int32(topRow), int32(leftColumn), int32(bottomRow), int32(rightColumn), data)
	if err != nil {
		return fmt.Errorf("add picture: %w", err)
	}
	if index < 0 {
		return fmt.Errorf("add picture returned no index: %w", toolkiterrors.ErrPictureAddFailed)
	}
	return nil
}
