//! Tolerance for non-structural defects in legacy workbooks.
//!
//! Real-world `.xls` producers emit workbooks whose metadata contradicts
//! MS-XLS while the workbook itself is perfectly readable: a `Font` record
//! with a `bFamily` byte outside the enumeration, an `XF` whose `fJustLast`
//! bit is set without distributed alignment, a cell whose `ixfe` names an
//! XF index that does not exist. Excel opens all of these; this reader
//! repairs them the same way and *records* every repair, so a caller can
//! enumerate exactly what changed through [`ToleranceReport`]. Nothing is
//! swallowed silently, and nothing outside the two closed enumerations
//! changes — record framing, stream grammar, and encryption remain hard
//! errors under every [`Leniency`] variant.
//!
//! **Excel-compatible repairs** are applied under every [`Leniency`] variant,
//! including the default [`Leniency::Strict`]. They cover references MS-XLS
//! makes a producer obligation but defines no reader-side refusal for —
//! bytes Excel and LibreOffice read past without complaint, so rejecting
//! them would cost the caller a whole worksheet over metadata that cannot
//! change what a conforming workbook means. Each substitutes the value
//! documented on the corresponding [`CellFormatDefect`] or [`EncodingDefect`]
//! variant.
//!
//! **Opt-in repairs** trade strictness for reach behind explicit [`Leniency`]
//! variants. The defects enumerated by [`FormattingDefect`] — cosmetic
//! formatting metadata that contradicts itself — are repaired with a
//! documented substitute value only under
//! [`Leniency::TolerateFormattingDefects`]; under [`Leniency::Strict`] they
//! remain hard errors, because rejecting them is the only way to guarantee a
//! caller never silently reads a value the file did not contain.

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
/// every variant. The Excel-compatible repairs listed by [`CellFormatDefect`]
/// are applied and recorded under every variant, including
/// [`Leniency::Strict`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub enum Leniency {
    /// Every deviation outside the Excel-compatible [`CellFormatDefect`]
    /// repairs is a hard error.
    ///
    /// This is the default: the [`FormattingDefect`] classes are rejected,
    /// while the [`CellFormatDefect`] repairs Excel itself applies are still
    /// performed and recorded in the [`ToleranceReport`].
    #[default]
    Strict,
    /// Repair and record the cosmetic defects listed by [`FormattingDefect`].
    ///
    /// Each repair substitutes the value documented on the corresponding
    /// [`FormattingDefect`] variant and appends a [`ToleratedDefect`] to
    /// the workbook's [`ToleranceReport`].
    TolerateFormattingDefects,
}

impl Leniency {
    /// Whether formatting defects are repaired rather than rejected.
    #[must_use]
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
    #[must_use]
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
    #[must_use]
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

/// The class of cell-format defect the reader repairs the way Excel repairs it.
///
/// This enumeration is closed by design: it is the exhaustive contract of the
/// repairs a read applies under every [`Leniency`] variant, including
/// [`Leniency::Strict`]. Each one matches what Excel and LibreOffice do with
/// references MS-XLS leaves a producer obligation, so applying it never
/// changes what a conforming workbook means.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum CellFormatDefect {
    /// A cell, row, or column record's `ixfe` named an XF index that cannot
    /// format a cell: it was out of range, named a reserved style slot
    /// (1–14), or named a style XF.
    ///
    /// MS-XLS 2.5.115 CellIXF makes a valid reference a producer obligation
    /// and defines no reader-side refusal. Excel renders the cell with
    /// default formatting and LibreOffice clamps to the default XF, so the
    /// reader substitutes the default cell XF — the mandatory first cell XF
    /// at index 15 (MS-XLS 2.4.353). Only styling falls back; the cell's
    /// value is never touched.
    XfReference,
}

impl CellFormatDefect {
    /// The BIFF record type whose payload carried the defect, or `None` when
    /// the carrying record varies per occurrence.
    #[must_use]
    pub fn record_type(self) -> Option<u16> {
        match self {
            CellFormatDefect::XfReference => None,
        }
    }

    /// A stable, allocation-free description of the defect class.
    #[must_use]
    pub fn description(self) -> &'static str {
        match self {
            CellFormatDefect::XfReference => {
                "Cell, row, or column references an XF that cannot format a cell"
            },
        }
    }
}

/// The class of text-encoding defect the reader repairs the way Excel and
/// LibreOffice repair it.
///
/// Like [`CellFormatDefect`], every class here is repaired and recorded under
/// every [`Leniency`] variant, including [`Leniency::Strict`]: the repair
/// decodes text lossily instead of discarding the workbook, and BIFF8 stores
/// almost all text as UTF-16 regardless.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum EncodingDefect {
    /// The `CODEPAGE` record (MS-XLS 2.4.52) named a code page this reader
    /// has no exact codec for.
    ///
    /// The raw value is kept in the report, and byte-stream text decodes
    /// through Windows-1252 — the default `CODEPAGE` MS-XLS assumes when the
    /// record is absent, and the page LibreOffice and POI fall back to. The
    /// page is total, so every byte still decodes to some character.
    UnsupportedCodePage,
}

impl EncodingDefect {
    /// The BIFF record type whose payload carried the defect.
    #[must_use]
    pub fn record_type(self) -> Option<u16> {
        match self {
            EncodingDefect::UnsupportedCodePage => Some(CODEPAGE_RECORD),
        }
    }

    /// A stable, allocation-free description of the defect class.
    #[must_use]
    pub fn description(self) -> &'static str {
        match self {
            EncodingDefect::UnsupportedCodePage => {
                "CODEPAGE names a code page without a supported codec"
            },
        }
    }
}

/// `CODEPAGE` (MS-XLS 2.4.52).
const CODEPAGE_RECORD: u16 = 0x0042;

/// One class of defect repaired while a workbook was read.
///
/// This enumeration is closed by design: it is the exhaustive contract of what
/// a read may change about a workbook. [`Defect::CellFormat`] and
/// [`Defect::Encoding`] classes are repaired under every [`Leniency`]
/// variant; [`Defect::Formatting`] classes
/// require [`Leniency::TolerateFormattingDefects`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Defect {
    /// A cosmetic formatting defect repaired by the opt-in [`Leniency`]
    /// variant.
    Formatting(FormattingDefect),
    /// A cell-format defect repaired the way Excel repairs it.
    CellFormat(CellFormatDefect),
    /// A text-encoding defect repaired the way Excel repairs it.
    Encoding(EncodingDefect),
}

impl Defect {
    /// The BIFF record type whose payload carried the defect.
    ///
    /// `None` for [`CellFormatDefect::XfReference`], whose carrying record
    /// varies per occurrence: any cell record, `ROW` (0x0208), or `COLINFO`
    /// (0x007D).
    #[must_use]
    pub fn record_type(self) -> Option<u16> {
        match self {
            Defect::Formatting(defect) => Some(defect.record_type()),
            Defect::CellFormat(defect) => defect.record_type(),
            Defect::Encoding(defect) => defect.record_type(),
        }
    }

    /// A stable, allocation-free description of the defect class.
    #[must_use]
    pub fn description(self) -> &'static str {
        match self {
            Defect::Formatting(defect) => defect.description(),
            Defect::CellFormat(defect) => defect.description(),
            Defect::Encoding(defect) => defect.description(),
        }
    }
}

impl From<FormattingDefect> for Defect {
    fn from(defect: FormattingDefect) -> Self {
        Self::Formatting(defect)
    }
}

impl From<CellFormatDefect> for Defect {
    fn from(defect: CellFormatDefect) -> Self {
        Self::CellFormat(defect)
    }
}

impl From<EncodingDefect> for Defect {
    fn from(defect: EncodingDefect) -> Self {
        Self::Encoding(defect)
    }
}

/// One repaired defect, located within the workbook that carried it.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ToleratedDefect {
    defect: Defect,
    ordinal: u32,
    observed: u32,
}

impl ToleratedDefect {
    /// The class of defect that was repaired.
    #[must_use]
    pub fn defect(&self) -> Defect {
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
    /// it is the number of `XF` records that were actually parsed. For
    /// [`CellFormatDefect::XfReference`] it is the zero-based row of the
    /// cell or `ROW` record, or the first column of the `COLINFO` record.
    /// For [`EncodingDefect::UnsupportedCodePage`] it is the byte offset of
    /// the `CODEPAGE` record within the Workbook stream.
    #[must_use]
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
    /// zero, since the defect *is* the value. For
    /// [`CellFormatDefect::XfReference`] it is the `ixfe` the record carried.
    /// For [`EncodingDefect::UnsupportedCodePage`] it is the raw `cv` code
    /// page identifier.
    #[must_use]
    pub fn observed(&self) -> u32 {
        self.observed
    }

    /// The BIFF record type whose payload carried the defect.
    ///
    /// `None` when the carrying record type varies per occurrence; see
    /// [`Defect::record_type`].
    #[must_use]
    pub fn record_type(&self) -> Option<u16> {
        self.defect.record_type()
    }
}

/// Everything a read repaired, in the order the defects were found.
///
/// An empty report means the workbook conformed and no substitution was made.
/// After a [`Leniency::Strict`] read the report can still list
/// [`Defect::CellFormat`] and [`Defect::Encoding`] repairs, which Excel
/// itself applies; it never lists
/// [`Defect::Formatting`] repairs.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ToleranceReport {
    defects: Vec<ToleratedDefect>,
    unrecorded: u32,
}

impl ToleranceReport {
    /// Whether the workbook parsed without any repair.
    #[must_use]
    pub fn is_clean(&self) -> bool {
        self.defects.is_empty() && self.unrecorded == 0
    }

    /// The individually recorded defects, in discovery order.
    ///
    /// At most [`ToleranceReport::RECORD_LIMIT`] entries; any excess is
    /// reported by [`ToleranceReport::unrecorded`].
    #[must_use]
    pub fn defects(&self) -> &[ToleratedDefect] {
        &self.defects
    }

    /// Defects that were repaired but not individually stored.
    ///
    /// Non-zero only for workbooks with more than
    /// [`ToleranceReport::RECORD_LIMIT`] defects, where storing every one
    /// would let the file dictate unbounded memory use.
    #[must_use]
    pub fn unrecorded(&self) -> u32 {
        self.unrecorded
    }

    /// Total number of repairs, including those not individually stored.
    #[must_use]
    pub fn total(&self) -> u64 {
        self.defects.len() as u64 + u64::from(self.unrecorded)
    }

    /// How many recorded defects belong to `defect`.
    ///
    /// Counts only individually recorded entries; compare
    /// [`ToleranceReport::unrecorded`] before treating a zero as absence.
    #[must_use]
    pub fn count(&self, defect: impl Into<Defect>) -> usize {
        let defect = defect.into();
        self.defects
            .iter()
            .filter(|entry| entry.defect == defect)
            .count()
    }

    /// The maximum number of individually recorded defects.
    pub const RECORD_LIMIT: usize = MAX_RECORDED_DEFECTS;

    /// Record one repair, saturating at [`MAX_RECORDED_DEFECTS`] stored
    /// entries; the excess is counted in [`ToleranceReport::unrecorded`].
    pub(crate) fn record(&mut self, defect: impl Into<Defect>, ordinal: u32, observed: u32) {
        if self.defects.len() < MAX_RECORDED_DEFECTS {
            self.defects.push(ToleratedDefect {
                defect: defect.into(),
                ordinal,
                observed,
            });
        } else {
            self.unrecorded = self.unrecorded.saturating_add(1);
        }
    }
}

/// Reader-side policy holder that decides and records each opt-in repair.
///
/// Parsers call [`ToleranceLog::tolerate`] at the exact point they would
/// otherwise have returned an error, and continue with the documented
/// substitute value only when it returns `Ok`. Repairs that Excel itself
/// applies bypass the policy decision and go straight to
/// [`ToleranceReport::record`] through [`ToleranceLog::report_mut`].
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
        self.report.record(defect, ordinal, observed);
        Ok(())
    }

    /// The report under construction, for repairs that need no policy
    /// decision because Excel applies them under every [`Leniency`] variant.
    pub(crate) fn report_mut(&mut self) -> &mut ToleranceReport {
        &mut self.report
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
    fn strict_mode_rejects_every_opt_in_defect_and_records_nothing() {
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
        assert_eq!(
            first.defect(),
            Defect::Formatting(FormattingDefect::FontFamily)
        );
        assert_eq!(first.ordinal(), 3);
        assert_eq!(first.observed(), 32);
        assert_eq!(
            first.record_type(),
            Some(super::super::font::FONT_RECORD_TYPE)
        );
        assert_eq!(report.defects()[1].ordinal(), 7);
        assert_eq!(report.defects()[2].observed(), 108);
    }

    #[test]
    fn excel_compatible_repairs_are_recorded_under_every_leniency() {
        // The `CellFormatDefect` classes need no policy decision: they record
        // straight into the report, Strict included.
        for leniency in [Leniency::Strict, Leniency::TolerateFormattingDefects] {
            let mut log = ToleranceLog::new(leniency);
            log.report_mut()
                .record(CellFormatDefect::XfReference, 10, 4095);
            let report = log.into_report();
            assert_eq!(report.total(), 1, "{leniency:?} must record the repair");

            let entry = report.defects()[0];
            assert_eq!(
                entry.defect(),
                Defect::CellFormat(CellFormatDefect::XfReference)
            );
            assert_eq!(entry.ordinal(), 10);
            assert_eq!(entry.observed(), 4095);
            // The carrying record varies per occurrence (cell records, ROW,
            // COLINFO), so the class has no single record type.
            assert_eq!(entry.record_type(), None);
        }
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
            let defect = Defect::from(defect);
            assert_eq!(defect.record_type(), Some(record_type));
            assert!(!defect.description().is_empty());
        }
        let cell = Defect::from(CellFormatDefect::XfReference);
        assert_eq!(cell.record_type(), None);
        assert!(!cell.description().is_empty());
        let encoding = Defect::from(EncodingDefect::UnsupportedCodePage);
        assert_eq!(encoding.record_type(), Some(0x0042));
        assert!(!encoding.description().is_empty());
    }
}
