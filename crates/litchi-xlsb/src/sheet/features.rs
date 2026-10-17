//! The worksheet features read from a worksheet stream without decoding its
//! cells (unified-ops record 0322).

use crate::conditional_formatting::Formatting;
use crate::hyperlinks::Hyperlink;
use crate::package::data_validation::Validation;
use crate::package::error::Result;
use crate::page_breaks::PageBreaks;
use crate::part_edit::SheetLayout;
use crate::sheet::AutoFilter;

/// The features stored in one worksheet stream, in stream order: the same
/// values the worksheet edit families list and edit.
///
/// Conditional formats and validations are the classic (`BrtBegin…`)
/// records; Office 2013 formats and `BrtDVal14` validations are not listed.
/// A hyperlink's external target is resolved through the worksheet's
/// relationships.
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

/// Read the features of a worksheet stream at complete-catalog position
/// `sheet`.
pub(crate) fn read(
    bytes: &[u8],
    relationships: &litchi_opc::Relationships,
    context: &crate::package::formula::Context,
    sheet: usize,
    leniency: crate::leniency::Leniency,
) -> Result<Features> {
    let layout = SheetLayout::scan(bytes)?;
    let pass = crate::part_edit::feature_pass(bytes, context, sheet, leniency)?;
    Ok(Features {
        hyperlinks: crate::hyperlinks::edit::read_layout(bytes, &layout, relationships)?,
        auto_filter: crate::auto_filter::read_layout(bytes, &layout)?,
        validations: crate::data_validation::edit::classic(pass.validations),
        formats: crate::conditional_formatting::edit::classic(pass.formats),
        page_breaks: crate::page_breaks::read_layout(bytes, &layout)?,
    })
}
