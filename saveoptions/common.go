// CommonConfig and CommonSetter remove the option fields and Set* plumbing that
// every saveoptions subpackage duplicated. Each subpackage's Config embeds
// CommonConfig, and its Apply delegates the shared fields to ApplyCommon, so
// only format-specific fields and their With* options remain in the subpackage.

package saveoptions

// CommonConfig holds the option fields shared by every saveoptions format
// (clearData, cachedFileFolder, validateMergedAreas, mergeAreas,
// createDirectory, sortNames, sortExternalNames, refreshChartCache,
// checkExcelRestriction, updateSmartArt, encryptDocumentProperties).
//
// The fields are exported so subpackages can set them from their With* options;
// the pointer-presence convention is unchanged (nil = keep the native default).
type CommonConfig struct {
	ClearData                 *bool
	CachedFileFolder          *string
	ValidateMergedAreas       *bool
	MergeAreas                *bool
	CreateDirectory           *bool
	SortNames                 *bool
	SortExternalNames         *bool
	RefreshChartCache         *bool
	CheckExcelRestriction     *bool
	UpdateSmartArt            *bool
	EncryptDocumentProperties *bool
}

// CommonSetter is the subset of the engine's save-options API implemented by
// every native option type the toolkit wraps (TxtSaveOptions, HtmlSaveOptions,
// PdfSaveOptions, ImageSaveOptions, OoxmlSaveOptions, ...). ApplyCommon drives
// it, so a format package's Apply only has to handle its own fields.
type CommonSetter interface {
	SetClearData(value bool) error
	SetCachedFileFolder(value string) error
	SetValidateMergedAreas(value bool) error
	SetMergeAreas(value bool) error
	SetCreateDirectory(value bool) error
	SetSortNames(value bool) error
	SetSortExternalNames(value bool) error
	SetRefreshChartCache(value bool) error
	SetCheckExcelRestriction(value bool) error
	SetUpdateSmartArt(value bool) error
	SetEncryptDocumentProperties(value bool) error
}

// ApplyCommon pushes every non-nil field onto o, preserving the presence
// semantics and call order used by the previous per-package implementations.
func (c *CommonConfig) ApplyCommon(o CommonSetter) error {
	if c.ClearData != nil {
		if err := o.SetClearData(*c.ClearData); err != nil {
			return err
		}
	}
	if c.CachedFileFolder != nil {
		if err := o.SetCachedFileFolder(*c.CachedFileFolder); err != nil {
			return err
		}
	}
	if c.ValidateMergedAreas != nil {
		if err := o.SetValidateMergedAreas(*c.ValidateMergedAreas); err != nil {
			return err
		}
	}
	if c.MergeAreas != nil {
		if err := o.SetMergeAreas(*c.MergeAreas); err != nil {
			return err
		}
	}
	if c.CreateDirectory != nil {
		if err := o.SetCreateDirectory(*c.CreateDirectory); err != nil {
			return err
		}
	}
	if c.SortNames != nil {
		if err := o.SetSortNames(*c.SortNames); err != nil {
			return err
		}
	}
	if c.SortExternalNames != nil {
		if err := o.SetSortExternalNames(*c.SortExternalNames); err != nil {
			return err
		}
	}
	if c.RefreshChartCache != nil {
		if err := o.SetRefreshChartCache(*c.RefreshChartCache); err != nil {
			return err
		}
	}
	if c.CheckExcelRestriction != nil {
		if err := o.SetCheckExcelRestriction(*c.CheckExcelRestriction); err != nil {
			return err
		}
	}
	if c.UpdateSmartArt != nil {
		if err := o.SetUpdateSmartArt(*c.UpdateSmartArt); err != nil {
			return err
		}
	}
	if c.EncryptDocumentProperties != nil {
		if err := o.SetEncryptDocumentProperties(*c.EncryptDocumentProperties); err != nil {
			return err
		}
	}
	return nil
}
