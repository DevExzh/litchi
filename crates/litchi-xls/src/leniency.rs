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
//! [`Leniency::TolerateFormattingDefects`] trades that guarantee for reach:
//! the defects enumerated by [`FormattingDefect`] are repaired with a
//! documented substitute value and *recorded*, so the caller can enumerate
//! exactly what was tolerated through [`ToleranceReport`]. Nothing is
//! swallowed silently, and nothing outside that enumeration is affected —
//! record framing, stream grammar, and encryption remain hard errors in both
//! modes.

use super::error::{Error, Result};

/// Upper bound on individually recorded defects.
///
/// A hostile or merely enormous workbook must not be able to make the report
/// grow without limit. Defects beyond this bound are counted, not stored; see
/// [`ToleranceReport::unrecorded`].
const MAX_RECORDED_DEFECTS: usize = 1024;

/// How the reader treats non-structural formatting defects.
///
/// Structural defects — malformed record framing, a broken workbook stream
/// grammar, or unsupported/absent decryption material — are hard errors under
/// every variant.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Leniency {
    /// Every deviation from MS-XLS is a hard error.
    ///
    /// This is the default and the historical behaviour: a workbook either
    /// parses exactly as specified or it does not parse at all.
    #[default]
    Strict,
    /// Repair and record the cosmetic defects listed by [`FormattingDefect`].
    ///
    /// Each repair substitutes the value documented on the corresponding
    /// [`FormattingDefect`] variant and appends an [`ToleratedDefect`] to
    /// the workbook's [`ToleranceReport`].
    TolerateFormattingDefects,
}

impl Leniency {
    /// Whether formatting defects are repaired rather than rejected.
    pub fn tolerates_formatting_defects(self) -> bool {
        matches!(self, Leniency::TolerateFormattingDefects)
    }
}

/// The class of cosmetic formatting defect a lenient read is allowed to repair.
///
/// This enumeration is closed by design: it is the exhaustive contract of what
/// [`Leniency::TolerateFormattingDefects`] may change about a workbook.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FormattingDefect {
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

impl FormattingDefect {
    /// The BIFF record type whose payload carried the defect.
    pub fn record_type(self) -> u16 {
        match self {
            FormattingDefect::FontFamily | FormattingDefect::FontNameEmpty => {
                super::font::FONT_RECORD_TYPE
            },
            FormattingDefect::AlignmentJustifyLastLine => super::number_format::XF_RECORD,
            FormattingDefect::ExtendedFormatCountMismatch => super::number_format::XFCRC_RECORD,
            FormattingDefect::FormatStringOverrun => super::number_format::FORMAT_RECORD,
        }
    }

    /// A stable, allocation-free description of the defect class.
    pub fn description(self) -> &'static str {
        match self {
            FormattingDefect::FontFamily => "Font family byte is outside the enumeration",
            FormattingDefect::FontNameEmpty => "Font record declares an empty name",
            FormattingDefect::AlignmentJustifyLastLine => {
                "XF justify-last-line is set without distributed horizontal alignment"
            },
            FormattingDefect::ExtendedFormatCountMismatch => {
                "XFCRC disagrees with the number of XF records"
            },
            FormattingDefect::FormatStringOverrun => {
                "Format record character count overstates its payload"
            },
        }
    }
}

/// One repaired formatting defect, located within the workbook that carried it.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ToleratedDefect {
    defect: FormattingDefect,
    ordinal: u32,
    observed: u32,
}

impl ToleratedDefect {
    /// The class of defect that was repaired.
    pub fn defect(&self) -> FormattingDefect {
        self.defect
    }

    /// Where the defect was found.
    ///
    /// For [`FormattingDefect::FontFamily`] and
    /// [`FormattingDefect::FontNameEmpty`] this is the font's logical index.
    /// For [`FormattingDefect::AlignmentJustifyLastLine`] it is the `XF`
    /// index. For [`FormattingDefect::FormatStringOverrun`] it is the
    /// zero-based ordinal of the `Format` record among the `Format` records
    /// parsed so far. For [`FormattingDefect::ExtendedFormatCountMismatch`]
    /// it is the number of `XF` records that were actually parsed.
    pub fn ordinal(&self) -> u32 {
        self.ordinal
    }

    /// The offending value read from the file.
    ///
    /// For [`FormattingDefect::FontFamily`] this is the out-of-range
    /// `bFamily` byte. For [`FormattingDefect::AlignmentJustifyLastLine`] it
    /// is the `alcH` value that accompanied the flag. For
    /// [`FormattingDefect::ExtendedFormatCountMismatch`] it is the declared
    /// `cxfs`. For [`FormattingDefect::FormatStringOverrun`] it is the
    /// declared `cch`. For [`FormattingDefect::FontNameEmpty`] it is always
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
pub struct ToleranceReport {
    defects: Vec<ToleratedDefect>,
    unrecorded: u32,
}

impl ToleranceReport {
    /// Whether the workbook parsed without any repair.
    pub fn is_clean(&self) -> bool {
        self.defects.is_empty() && self.unrecorded == 0
    }

    /// The individually recorded defects, in discovery order.
    ///
    /// At most [`ToleranceReport::RECORD_LIMIT`] entries; any excess is
    /// reported by [`ToleranceReport::unrecorded`].
    pub fn defects(&self) -> &[ToleratedDefect] {
        &self.defects
    }

    /// Defects that were repaired but not individually stored.
    ///
    /// Non-zero only for workbooks with more than
    /// [`ToleranceReport::RECORD_LIMIT`] defects, where storing every one
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
    /// [`ToleranceReport::unrecorded`] before treating a zero as absence.
    pub fn count(&self, defect: FormattingDefect) -> usize {
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
/// Parsers call [`ToleranceLog::tolerate`] at the exact point they would
/// otherwise have returned an error, and continue with the documented
/// substitute value only when it returns `Ok`.
#[derive(Debug, Clone, Default)]
pub(crate) struct ToleranceLog {
    leniency: Leniency,
    report: ToleranceReport,
}

impl ToleranceLog {
    /// Build a log that applies `leniency`.
    pub(crate) fn new(leniency: Leniency) -> Self {
        Self {
            leniency,
            report: ToleranceReport::default(),
        }
    }

    /// Decide whether `defect` may be repaired, recording it if so.
    ///
    /// Returns `Err(on_strict())` under [`Leniency::Strict`], so a caller
    /// that propagates with `?` keeps byte-identical strict behaviour. Under
    /// [`Leniency::TolerateFormattingDefects`] it returns `Ok(())` and the
    /// caller substitutes the value documented on the `defect` variant.
    pub(crate) fn tolerate(
        &mut self,
        defect: FormattingDefect,
        ordinal: u32,
        observed: u32,
        on_strict: impl FnOnce() -> Error,
    ) -> Result<()> {
        if !self.leniency.tolerates_formatting_defects() {
            return Err(on_strict());
        }
        if self.report.defects.len() < MAX_RECORDED_DEFECTS {
            self.report.defects.push(ToleratedDefect {
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
    pub(crate) fn into_report(self) -> ToleranceReport {
        self.report
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn strict_error() -> Error {
        Error::InvalidData("strict".to_string())
    }

    #[test]
    fn strict_mode_rejects_every_defect_and_records_nothing() {
        let mut log = ToleranceLog::new(Leniency::Strict);
        let error = log
            .tolerate(FormattingDefect::FontFamily, 3, 32, strict_error)
            .expect_err("strict mode must reject");
        assert!(error.to_string().contains("strict"));
        assert!(log.into_report().is_clean());
    }

    #[test]
    fn default_leniency_is_strict() {
        assert_eq!(Leniency::default(), Leniency::Strict);
        assert!(!Leniency::default().tolerates_formatting_defects());
        assert!(ToleranceLog::default().into_report().is_clean());
    }

    #[test]
    fn lenient_mode_records_defects_in_discovery_order() {
        let mut log = ToleranceLog::new(Leniency::TolerateFormattingDefects);
        log.tolerate(FormattingDefect::FontFamily, 3, 32, strict_error)
            .expect("lenient mode tolerates");
        log.tolerate(FormattingDefect::FormatStringOverrun, 7, 40, strict_error)
            .expect("lenient mode tolerates");
        log.tolerate(FormattingDefect::FontFamily, 5, 108, strict_error)
            .expect("lenient mode tolerates");

        let report = log.into_report();
        assert!(!report.is_clean());
        assert_eq!(report.total(), 3);
        assert_eq!(report.unrecorded(), 0);
        assert_eq!(report.count(FormattingDefect::FontFamily), 2);
        assert_eq!(report.count(FormattingDefect::FontNameEmpty), 0);

        let first = report.defects()[0];
        assert_eq!(first.defect(), FormattingDefect::FontFamily);
        assert_eq!(first.ordinal(), 3);
        assert_eq!(first.observed(), 32);
        assert_eq!(first.record_type(), super::super::font::FONT_RECORD_TYPE);
        assert_eq!(report.defects()[1].ordinal(), 7);
        assert_eq!(report.defects()[2].observed(), 108);
    }

    #[test]
    fn recorded_defects_are_bounded_and_the_excess_is_counted() {
        let mut log = ToleranceLog::new(Leniency::TolerateFormattingDefects);
        let total = ToleranceReport::RECORD_LIMIT + 5;
        for ordinal in 0..total {
            log.tolerate(
                FormattingDefect::FontNameEmpty,
                ordinal as u32,
                0,
                strict_error,
            )
            .expect("lenient mode tolerates");
        }

        let report = log.into_report();
        assert_eq!(report.defects().len(), ToleranceReport::RECORD_LIMIT);
        assert_eq!(report.unrecorded(), 5);
        assert_eq!(report.total(), total as u64);
    }

    #[test]
    fn every_defect_class_maps_to_its_record_type_and_description() {
        let expectations = [
            (
                FormattingDefect::FontFamily,
                super::super::font::FONT_RECORD_TYPE,
            ),
            (
                FormattingDefect::FontNameEmpty,
                super::super::font::FONT_RECORD_TYPE,
            ),
            (
                FormattingDefect::AlignmentJustifyLastLine,
                super::super::number_format::XF_RECORD,
            ),
            (
                FormattingDefect::ExtendedFormatCountMismatch,
                super::super::number_format::XFCRC_RECORD,
            ),
            (
                FormattingDefect::FormatStringOverrun,
                super::super::number_format::FORMAT_RECORD,
            ),
        ];
        for (defect, record_type) in expectations {
            assert_eq!(defect.record_type(), record_type);
            assert!(!defect.description().is_empty());
        }
    }
}
