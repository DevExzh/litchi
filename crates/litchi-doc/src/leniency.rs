//! Opt-in tolerance for non-structural stylesheet defects in binary documents.
//!
//! MS-DOC 2.9 states that a style's name and each of its aliases "MUST NOT be
//! empty and MUST be unique within all names in the stylesheet". Word enforces
//! that when it writes, but files produced by other tools — and by localized
//! Word builds — do reach users with a duplicated name, and Word opens them
//! without complaint. Rejecting such a file costs the caller the entire
//! document: its text, tables, and everything else, over metadata that only
//! affects how a style is labelled.
//!
//! [`Leniency::TolerateStylesheetDefects`] trades strict conformance for
//! reach. Every repair substitutes the value documented on the corresponding
//! [`StylesheetDefect`] variant and is *recorded*, so a caller can enumerate
//! exactly what was tolerated through [`ToleranceReport`]. Nothing is
//! swallowed silently, and nothing outside that enumeration changes: piece
//! tables, FIB structure, stream grammar, and encryption remain hard errors
//! under every variant.
//!
//! This mirrors the contract `XlsLeniency` provides for legacy workbooks so the
//! two legacy readers behave the same way.

/// Upper bound on individually recorded defects.
///
/// A hostile or merely enormous stylesheet must not be able to make the report
/// grow without limit. Defects beyond this bound are counted, not stored; see
/// [`ToleranceReport::unrecorded`].
const MAX_RECORDED_DEFECTS: usize = 1024;

/// How the reader treats non-structural stylesheet defects.
///
/// Structural defects — malformed FIB or piece-table framing, a broken stream
/// grammar, or unsupported/absent decryption material — are hard errors under
/// every variant.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Leniency {
    /// Every deviation from MS-DOC is a hard error.
    ///
    /// This is the default and the historical behaviour: a document either
    /// parses exactly as specified or it does not parse at all.
    #[default]
    Strict,
    /// Repair and record the defects listed by [`StylesheetDefect`].
    TolerateStylesheetDefects,
}

impl Leniency {
    /// Whether stylesheet defects are repaired rather than rejected.
    #[inline]
    #[must_use]
    pub fn tolerates_stylesheet_defects(self) -> bool {
        matches!(self, Self::TolerateStylesheetDefects)
    }
}

/// The class of stylesheet defect a lenient read is allowed to repair.
///
/// This enumeration is closed by design: it is the exhaustive contract of what
/// [`Leniency::TolerateStylesheetDefects`] may change about a document.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum StylesheetDefect {
    /// A style name or alias repeats one already used in the stylesheet.
    ///
    /// MS-DOC 2.9 requires uniqueness. The duplicate is kept as-is on the style
    /// that carries it — no name is rewritten or discarded — because the name
    /// only labels the style; lookup by index, which is how every stored
    /// reference works, is unaffected.
    DuplicateStyleName,
}

impl StylesheetDefect {
    /// A short, stable description of what a lenient read did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::DuplicateStyleName => "kept the duplicated name; styles are resolved by index",
        }
    }
}

/// One repaired defect, with enough context to locate it in the file.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ToleratedDefect {
    /// Which defect class was repaired.
    pub defect: StylesheetDefect,
    /// Index of the style within the stylesheet that carried the defect.
    pub style_index: u16,
}

/// Everything a lenient read repaired, in the order the defects were found.
///
/// An empty report means the document conformed and no substitution was made,
/// which is always the case after a [`Leniency::Strict`] read.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ToleranceReport {
    defects: Vec<ToleratedDefect>,
    unrecorded: u32,
}

impl ToleranceReport {
    /// Whether the document parsed without any repair.
    #[inline]
    #[must_use]
    pub fn is_clean(&self) -> bool {
        self.defects.is_empty() && self.unrecorded == 0
    }

    /// The individually recorded defects, in discovery order.
    #[inline]
    #[must_use]
    pub fn defects(&self) -> &[ToleratedDefect] {
        &self.defects
    }

    /// Defects that were repaired but not stored, once the bound was reached.
    #[inline]
    #[must_use]
    pub fn unrecorded(&self) -> u32 {
        self.unrecorded
    }

    /// Total repairs performed, recorded or not.
    #[inline]
    #[must_use]
    pub fn total(&self) -> u64 {
        self.defects.len() as u64 + u64::from(self.unrecorded)
    }

    /// Record one repair, saturating at [`MAX_RECORDED_DEFECTS`] stored entries.
    pub(crate) fn record(&mut self, defect: StylesheetDefect, style_index: u16) {
        if self.defects.len() < MAX_RECORDED_DEFECTS {
            self.defects.push(ToleratedDefect {
                defect,
                style_index,
            });
        } else {
            self.unrecorded = self.unrecorded.saturating_add(1);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn strict_is_the_default() {
        assert_eq!(Leniency::default(), Leniency::Strict);
        assert!(!Leniency::default().tolerates_stylesheet_defects());
    }

    #[test]
    fn a_fresh_report_is_clean() {
        let report = ToleranceReport::default();
        assert!(report.is_clean());
        assert_eq!(report.total(), 0);
        assert!(report.defects().is_empty());
    }

    #[test]
    fn recorded_defects_are_returned_in_discovery_order() {
        let mut report = ToleranceReport::default();
        report.record(StylesheetDefect::DuplicateStyleName, 15);
        report.record(StylesheetDefect::DuplicateStyleName, 42);

        assert!(!report.is_clean());
        assert_eq!(report.total(), 2);
        assert_eq!(
            report
                .defects()
                .iter()
                .map(|d| d.style_index)
                .collect::<Vec<_>>(),
            vec![15, 42]
        );
    }

    /// A hostile stylesheet must not be able to grow the report without limit.
    #[test]
    fn defects_past_the_bound_are_counted_not_stored() {
        let mut report = ToleranceReport::default();
        for index in 0..(MAX_RECORDED_DEFECTS + 5) {
            report.record(StylesheetDefect::DuplicateStyleName, index as u16);
        }

        assert_eq!(report.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(report.unrecorded(), 5);
        assert_eq!(report.total(), MAX_RECORDED_DEFECTS as u64 + 5);
    }
}
