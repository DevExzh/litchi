/// TAP (Table Properties) parser with arena allocator support.
///
/// This facade owns the stable parser API. Semantic state transitions live in
/// `model`, while binary SPRM operand decoding and [MS-DOC] validation live
/// in `codec`.
///
/// Reference: Apache POI's org.apache.poi.hwpf.sprm.TableSprmUncompressor
mod codec;
mod model;
#[cfg(test)]
mod tests;

use crate::leniency::ToleranceReport;
use crate::package::Result;
use crate::parts::styles::StyleSheet;
use crate::parts::tap::TableProperties;
use bumpalo::Bump;

/// TAP parser with arena allocation for temporary structures.
///
/// Uses bumpalo arena allocator for zero-cost temporary allocations
/// during TAP parsing. The arena is automatically cleaned up when
/// the parser is dropped.
pub struct TapParser<'arena> {
    /// Arena allocator for temporary parsing data (reserved for future use).
    #[allow(
        dead_code,
        reason = "reserved DOC structure retained for format completeness or future round-trip support"
    )]
    arena: &'arena Bump,
}

impl<'arena> TapParser<'arena> {
    /// Create a new TAP parser with an arena allocator.
    pub fn new(arena: &'arena Bump) -> Self {
        Self { arena }
    }

    /// Parse table properties from SPRM list.
    ///
    /// Based on Apache POI's uncompressTAP method.
    ///
    /// Word-compatible repairs (see [`crate::SprmDefect`]) are applied
    /// unconditionally; this convenience entry point discards the per-parse
    /// [`ToleranceReport`]. Document reads surface it through
    /// [`crate::Document::tolerance_report`].
    pub fn parse_tap(&self, grpprl: &[u8]) -> Result<TableProperties> {
        self.parse_tap_context(grpprl, false, None, &mut ToleranceReport::default())
    }

    /// [`Self::parse_tap`] that records any Word-compatible repairs it applies.
    pub(crate) fn parse_tap_with_report(
        &self,
        grpprl: &[u8],
        tolerance: &mut ToleranceReport,
    ) -> Result<TableProperties> {
        self.parse_tap_context(grpprl, false, None, tolerance)
    }

    /// Parse direct table properties while resolving sprmTIstd against a stylesheet.
    pub(crate) fn parse_tap_with_stylesheet(
        &self,
        grpprl: &[u8],
        stylesheet: &StyleSheet,
        tolerance: &mut ToleranceReport,
    ) -> Result<TableProperties> {
        self.parse_tap_context(grpprl, false, Some(stylesheet), tolerance)
    }

    pub(crate) fn parse_conditional_tap(
        &self,
        grpprl: &[u8],
        tolerance: &mut ToleranceReport,
    ) -> Result<TableProperties> {
        self.parse_tap_context(grpprl, true, None, tolerance)
    }
}
