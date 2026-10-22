//! The worksheet features read from a worksheet stream without decoding its
//! cells (unified-ops record 0322).

use crate::conditional_formatting::Formatting;
use crate::hyperlinks::Hyperlink;
use crate::package::data_validation::Validation;
use crate::package::error::Result;
use crate::page_breaks::PageBreaks;
use crate::part_edit::SheetLayout;
use crate::sheet::AutoFilter;

bitflags::bitflags! {
    /// The worksheet features one read decodes. Hyperlinks, the auto-filter
    /// and page breaks are read from the record layout alone; validations
    /// and conditional formats also render their formulas.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    pub struct FeatureSet: u8 {
        /// [`Features::hyperlinks`].
        const HYPERLINKS = 1;
        /// [`Features::auto_filter`].
        const AUTO_FILTER = 1 << 1;
        /// [`Features::validations`].
        const VALIDATIONS = 1 << 2;
        /// [`Features::formats`].
        const FORMATS = 1 << 3;
        /// [`Features::page_breaks`].
        const PAGE_BREAKS = 1 << 4;
    }
}

/// The features stored in one worksheet stream, in stream order: the same
/// values the worksheet edit families list and edit.
///
/// Conditional formats and validations are the classic (`BrtBegin…`)
/// records; Office 2013 formats and `BrtDVal14` validations are not listed.
/// A hyperlink's external target is resolved through the worksheet's
/// relationships. A feature its read did not select is empty.
#[derive(Debug, Clone)]
#[non_exhaustive]
pub struct Features {
    /// Hyperlinks (`BrtHLink`).
    pub hyperlinks: Vec<Hyperlink>,
    /// The auto-filter range (`BrtBeginAFilter`).
    pub auto_filter: Option<AutoFilter>,
    /// Classic data validations (`BrtDVal`).
    pub validations: Vec<Validation>,
    /// Classic conditional formats (`BrtBeginConditionalFormatting`).
    pub formats: Vec<Formatting>,
    /// Manual page breaks.
    pub page_breaks: PageBreaks,
}

/// Read the `select`ed features of a worksheet stream at complete-catalog
/// position `sheet`.
pub(crate) fn read(
    bytes: &[u8],
    relationships: &litchi_opc::Relationships,
    context: &crate::package::formula::Context,
    sheet: usize,
    leniency: crate::leniency::Leniency,
    select: FeatureSet,
    raw_limits: crate::raw::Limits,
) -> Result<Features> {
    let layout = SheetLayout::scan_with_limits(bytes, raw_limits)?;
    let (validations, formats) = if select.intersects(FeatureSet::VALIDATIONS | FeatureSet::FORMATS)
    {
        let pass = crate::part_edit::feature_pass(bytes, context, sheet, raw_limits, leniency)?;
        (
            if select.contains(FeatureSet::VALIDATIONS) {
                crate::data_validation::edit::classic(pass.validations)
            } else {
                Vec::new()
            },
            if select.contains(FeatureSet::FORMATS) {
                crate::conditional_formatting::edit::classic(pass.formats)
            } else {
                Vec::new()
            },
        )
    } else {
        (Vec::new(), Vec::new())
    };
    Ok(Features {
        hyperlinks: if select.contains(FeatureSet::HYPERLINKS) {
            crate::hyperlinks::edit::read_layout(bytes, &layout, relationships, raw_limits)?
        } else {
            Vec::new()
        },
        auto_filter: if select.contains(FeatureSet::AUTO_FILTER) {
            crate::auto_filter::read_layout(bytes, &layout, raw_limits)?
        } else {
            None
        },
        validations,
        formats,
        page_breaks: if select.contains(FeatureSet::PAGE_BREAKS) {
            crate::page_breaks::read_layout(bytes, &layout, raw_limits)?
        } else {
            PageBreaks::default()
        },
    })
}
