//! Opt-in tolerance for non-structural formatting defects in legacy workbooks.
//!
//! Real-world `.xls` producers emit workbooks whose *cosmetic* formatting
//! metadata contradicts MS-XLS while the workbook itself is perfectly readable:
//! a `Font` record with a `bFamily` byte outside the enumeration, an `XF` whose
//! `fJustLast` bit is set without distributed alignment, an `XFCRC` whose
//! `cxfs` disagrees with the number of `XF` records actually present. Excel
//! opens all of these; by default this reader does not, because rejecting them
//! is the only way to guarantee a caller never silently reads a value the file
//! did not contain.
//!
//! [`XlsLeniency::TolerateFormattingDefects`] trades that guarantee for reach:
//! the defects enumerated by [`XlsFormattingDefect`] are repaired with a
//! documented substitute value and *recorded*, so the caller can enumerate
//! exactly what was tolerated through [`XlsToleranceReport`]. Nothing is
//! swallowed silently, and nothing outside that enumeration is affected —
//! record framing, stream grammar, and encryption remain hard errors in both
//! modes.

use super::error::{XlsError, XlsResult};

/// Upper bound on individually recorded defects.
///
/// A hostile or merely enormous workbook must not be able to make the report
/// grow without limit. Defects beyond this bound are counted, not stored; see
/// [`XlsToleranceReport::unrecorded`].
const MAX_RECORDED_DEFECTS: usize = 1024;

/// How the reader treats non-structural formatting defects.
///
/// Structural defects — malformed record framing, a broken workbook stream
/// grammar, or unsupported/absent decryption material — are hard errors under
/// every variant.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum XlsLeniency {
    /// Every deviation from MS-XLS is a hard error.
    ///
    /// This is the default and the historical behaviour: a workbook either
    /// parses exactly as specified or it does not parse at all.
    #[default]
    Strict,
    /// Repair and record the cosmetic defects listed by [`XlsFormattingDefect`].
    ///
    /// Each repair substitutes the value documented on the corresponding
    /// [`XlsFormattingDefect`] variant and appends an [`XlsToleratedDefect`] to
    /// the workbook's [`XlsToleranceReport`].
    TolerateFormattingDefects,
}

impl XlsLeniency {
    /// Whether formatting defects are repaired rather than rejected.
    pub fn tolerates_formatting_defects(self) -> bool {
        matches!(self, XlsLeniency::TolerateFormattingDefects)
    }
}

/// The class of cosmetic formatting defect a lenient read is allowed to repair.
///
/// This enumeration is closed by design: it is the exhaustive contract of what
/// [`XlsLeniency::TolerateFormattingDefects`] may change about a workbook.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum XlsFormattingDefect {
    /// `Font.bFamily` (MS-XLS 2.4.122) held a value outside `0..=5`.
    ///
    /// Repaired to `NotApplicable`, the value MS-XLS assigns to a font whose
    /// family is unspecified. `bFamily` is a substitution hint for absent
    /// fonts; it never changes the glyphs a conforming reader selects.
    FontFamily,
    /// `Font.cch` was zero, so the record declared a nameless font.
    ///
    /// Repaired to an empty font name. Callers that render text fall back to
    /// their own default face exactly as they would for an unknown name.
    FontNameEmpty,
    /// `XF.fJustLast` was set without `alcH` being `Distributed`.
    ///
    /// Repaired by clearing the justify-last-line flag, which is the only
    /// reading under which the surviving horizontal alignment is meaningful.
    AlignmentJustifyLastLine,
    /// `XFCRC.cxfs` (MS-XLS 2.4.354) disagreed with the number of `XF` records.
    ///
    /// Repaired by trusting the records that were actually parsed. `XFCRC` is a
    /// redundant integrity summary; the `XF` records themselves are the data.
    ExtendedFormatCountMismatch,
    /// A `Format` record's `cch` claimed more characters than its payload held.
    ///
    /// Repaired by decoding only the characters the record actually contains,
    /// which yields a truncated but well-formed number-format code.
    FormatStringOverrun,
}

impl XlsFormattingDefect {
    /// The BIFF record type whose payload carried the defect.
    pub fn record_type(self) -> u16 {
        match self {
            XlsFormattingDefect::FontFamily | XlsFormattingDefect::FontNameEmpty => {
                super::font::FONT_RECORD_TYPE
            },
            XlsFormattingDefect::AlignmentJustifyLastLine => super::number_format::XF_RECORD,
            XlsFormattingDefect::ExtendedFormatCountMismatch => super::number_format::XFCRC_RECORD,
            XlsFormattingDefect::FormatStringOverrun => super::number_format::FORMAT_RECORD,
        }
    }

    /// A stable, allocation-free description of the defect class.
    pub fn description(self) -> &'static str {
        match self {
            XlsFormattingDefect::FontFamily => "Font family byte is outside the enumeration",
            XlsFormattingDefect::FontNameEmpty => "Font record declares an empty name",
            XlsFormattingDefect::AlignmentJustifyLastLine => {
                "XF justify-last-line is set without distributed horizontal alignment"
            },
            XlsFormattingDefect::ExtendedFormatCountMismatch => {
                "XFCRC disagrees with the number of XF records"
            },
            XlsFormattingDefect::FormatStringOverrun => {
                "Format record character count overstates its payload"
            },
        }
    }
}

/// One repaired formatting defect, located within the workbook that carried it.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct XlsToleratedDefect {
    defect: XlsFormattingDefect,
    ordinal: u32,
    observed: u32,
}

impl XlsToleratedDefect {
    /// The class of defect that was repaired.
    pub fn defect(&self) -> XlsFormattingDefect {
        self.defect
    }

    /// Where the defect was found.
    ///
    /// For [`XlsFormattingDefect::FontFamily`] and
    /// [`XlsFormattingDefect::FontNameEmpty`] this is the font's logical index.
    /// For [`XlsFormattingDefect::AlignmentJustifyLastLine`] it is the `XF`
    /// index. For [`XlsFormattingDefect::FormatStringOverrun`] it is the
    /// zero-based ordinal of the `Format` record among the `Format` records
    /// parsed so far. For [`XlsFormattingDefect::ExtendedFormatCountMismatch`]
    /// it is the number of `XF` records that were actually parsed.
    pub fn ordinal(&self) -> u32 {
        self.ordinal
    }

    /// The offending value read from the file.
    ///
    /// For [`XlsFormattingDefect::FontFamily`] this is the out-of-range
    /// `bFamily` byte. For [`XlsFormattingDefect::AlignmentJustifyLastLine`] it
    /// is the `alcH` value that accompanied the flag. For
    /// [`XlsFormattingDefect::ExtendedFormatCountMismatch`] it is the declared
    /// `cxfs`. For [`XlsFormattingDefect::FormatStringOverrun`] it is the
    /// declared `cch`. For [`XlsFormattingDefect::FontNameEmpty`] it is always
    /// zero, since the defect *is* the value.
    pub fn observed(&self) -> u32 {
        self.observed
    }

    /// The BIFF record type whose payload carried the defect.
    pub fn record_type(&self) -> u16 {
        self.defect.record_type()
    }
}

/// Everything a lenient read repaired, in the order the defects were found.
///
/// A strict read always produces an empty report. An empty report from a
/// lenient read means the workbook needed no repairs at all.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct XlsToleranceReport {
    defects: Vec<XlsToleratedDefect>,
    unrecorded: u32,
}

impl XlsToleranceReport {
    /// Whether the workbook parsed without any repair.
    pub fn is_clean(&self) -> bool {
        self.defects.is_empty() && self.unrecorded == 0
    }

    /// The individually recorded defects, in discovery order.
    ///
    /// At most [`XlsToleranceReport::RECORD_LIMIT`] entries; any excess is
    /// reported by [`XlsToleranceReport::unrecorded`].
    pub fn defects(&self) -> &[XlsToleratedDefect] {
        &self.defects
    }

    /// Defects that were repaired but not individually stored.
    ///
    /// Non-zero only for workbooks with more than
    /// [`XlsToleranceReport::RECORD_LIMIT`] defects, where storing every one
    /// would let the file dictate unbounded memory use.
    pub fn unrecorded(&self) -> u32 {
        self.unrecorded
    }

    /// Total number of repairs, including those not individually stored.
    pub fn total(&self) -> u64 {
        self.defects.len() as u64 + u64::from(self.unrecorded)
    }

    /// How many recorded defects belong to `defect`.
    ///
    /// Counts only individually recorded entries; compare
    /// [`XlsToleranceReport::unrecorded`] before treating a zero as absence.
    pub fn count(&self, defect: XlsFormattingDefect) -> usize {
        self.defects
            .iter()
            .filter(|entry| entry.defect == defect)
            .count()
    }

    /// The maximum number of individually recorded defects.
    pub const RECORD_LIMIT: usize = MAX_RECORDED_DEFECTS;
}

/// Reader-side policy holder that decides and records each repair.
///
/// Parsers call [`XlsToleranceLog::tolerate`] at the exact point they would
/// otherwise have returned an error, and continue with the documented
/// substitute value only when it returns `Ok`.
#[derive(Debug, Clone, Default)]
pub(crate) struct XlsToleranceLog {
    leniency: XlsLeniency,
    report: XlsToleranceReport,
}

impl XlsToleranceLog {
    /// Build a log that applies `leniency`.
    pub(crate) fn new(leniency: XlsLeniency) -> Self {
        Self {
            leniency,
            report: XlsToleranceReport::default(),
        }
    }

    /// Decide whether `defect` may be repaired, recording it if so.
    ///
    /// Returns `Err(on_strict())` under [`XlsLeniency::Strict`], so a caller
    /// that propagates with `?` keeps byte-identical strict behaviour. Under
    /// [`XlsLeniency::TolerateFormattingDefects`] it returns `Ok(())` and the
    /// caller substitutes the value documented on the `defect` variant.
    pub(crate) fn tolerate(
        &mut self,
        defect: XlsFormattingDefect,
        ordinal: u32,
        observed: u32,
        on_strict: impl FnOnce() -> XlsError,
    ) -> XlsResult<()> {
        if !self.leniency.tolerates_formatting_defects() {
            return Err(on_strict());
        }
        if self.report.defects.len() < MAX_RECORDED_DEFECTS {
            self.report.defects.push(XlsToleratedDefect {
                defect,
                ordinal,
                observed,
            });
        } else {
            self.report.unrecorded = self.report.unrecorded.saturating_add(1);
        }
        Ok(())
    }

    /// Consume the log and yield the report handed to the caller.
    pub(crate) fn into_report(self) -> XlsToleranceReport {
        self.report
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn strict_error() -> XlsError {
        XlsError::InvalidData("strict".to_string())
    }

    #[test]
    fn strict_mode_rejects_every_defect_and_records_nothing() {
        let mut log = XlsToleranceLog::new(XlsLeniency::Strict);
        let error = log
            .tolerate(XlsFormattingDefect::FontFamily, 3, 32, strict_error)
            .expect_err("strict mode must reject");
        assert!(error.to_string().contains("strict"));
        assert!(log.into_report().is_clean());
    }

    #[test]
    fn default_leniency_is_strict() {
        assert_eq!(XlsLeniency::default(), XlsLeniency::Strict);
        assert!(!XlsLeniency::default().tolerates_formatting_defects());
        assert!(XlsToleranceLog::default().into_report().is_clean());
    }

    #[test]
    fn lenient_mode_records_defects_in_discovery_order() {
        let mut log = XlsToleranceLog::new(XlsLeniency::TolerateFormattingDefects);
        log.tolerate(XlsFormattingDefect::FontFamily, 3, 32, strict_error)
            .expect("lenient mode tolerates");
        log.tolerate(
            XlsFormattingDefect::FormatStringOverrun,
            7,
            40,
            strict_error,
        )
        .expect("lenient mode tolerates");
        log.tolerate(XlsFormattingDefect::FontFamily, 5, 108, strict_error)
            .expect("lenient mode tolerates");

        let report = log.into_report();
        assert!(!report.is_clean());
        assert_eq!(report.total(), 3);
        assert_eq!(report.unrecorded(), 0);
        assert_eq!(report.count(XlsFormattingDefect::FontFamily), 2);
        assert_eq!(report.count(XlsFormattingDefect::FontNameEmpty), 0);

        let first = report.defects()[0];
        assert_eq!(first.defect(), XlsFormattingDefect::FontFamily);
        assert_eq!(first.ordinal(), 3);
        assert_eq!(first.observed(), 32);
        assert_eq!(first.record_type(), super::super::font::FONT_RECORD_TYPE);
        assert_eq!(report.defects()[1].ordinal(), 7);
        assert_eq!(report.defects()[2].observed(), 108);
    }

    #[test]
    fn recorded_defects_are_bounded_and_the_excess_is_counted() {
        let mut log = XlsToleranceLog::new(XlsLeniency::TolerateFormattingDefects);
        let total = XlsToleranceReport::RECORD_LIMIT + 5;
        for ordinal in 0..total {
            log.tolerate(
                XlsFormattingDefect::FontNameEmpty,
                ordinal as u32,
                0,
                strict_error,
            )
            .expect("lenient mode tolerates");
        }

        let report = log.into_report();
        assert_eq!(report.defects().len(), XlsToleranceReport::RECORD_LIMIT);
        assert_eq!(report.unrecorded(), 5);
        assert_eq!(report.total(), total as u64);
    }

    #[test]
    fn every_defect_class_maps_to_its_record_type_and_description() {
        let expectations = [
            (
                XlsFormattingDefect::FontFamily,
                super::super::font::FONT_RECORD_TYPE,
            ),
            (
                XlsFormattingDefect::FontNameEmpty,
                super::super::font::FONT_RECORD_TYPE,
            ),
            (
                XlsFormattingDefect::AlignmentJustifyLastLine,
                super::super::number_format::XF_RECORD,
            ),
            (
                XlsFormattingDefect::ExtendedFormatCountMismatch,
                super::super::number_format::XFCRC_RECORD,
            ),
            (
                XlsFormattingDefect::FormatStringOverrun,
                super::super::number_format::FORMAT_RECORD,
            ),
        ];
        for (defect, record_type) in expectations {
            assert_eq!(defect.record_type(), record_type);
            assert!(!defect.description().is_empty());
        }
    }
}
