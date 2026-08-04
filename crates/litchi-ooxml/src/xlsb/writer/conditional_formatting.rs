//! Compatibility adapter for XLSB conditional-formatting serialization.
//!
//! The bounded classic and Office 2013 record writer is owned by
//! litchi-xlsb. This host module keeps the historical result type and writer
//! path while leaving worksheet orchestration in litchi-ooxml.

use crate::xlsb::conditional_formatting::ConditionalFormatting;
use crate::xlsb::error::{XlsbError, XlsbResult};
use std::io::Write;

/// Write all classic and Office 2013 conditional-formatting collections.
pub fn write_conditional_formattings<W: Write>(
    writer: &mut litchi_xlsb::raw::Writer<W>,
    conditional_formattings: &[ConditionalFormatting],
) -> XlsbResult<()> {
    litchi_xlsb::conditional_formatting::write_conditional_formattings(
        writer,
        conditional_formattings,
    )
    .map_err(XlsbError::from)
}
