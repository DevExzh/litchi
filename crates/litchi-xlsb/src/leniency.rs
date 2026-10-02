//! Repair reporting for non-structural defects in Excel Binary Workbooks.
//!
//! Real-world BIFF12 producers — newer Excel builds, add-ins, and third-party
//! writers — emit workbooks whose records contradict a literal reading of
//! MS-XLSB while the workbook itself is perfectly readable: a reserved bit set
//! in a cell's flag byte, a style record with a few unconsumed trailing bytes,
//! a direct color whose `fValidRGB` bit is clear. Excel reads BIFF12 with
//! ignore-and-advance semantics, and MS-XLSB 2.1.4/2.1.6 builds the format
//! around applications that "can ignore but preserve" what they do not
//! understand. This reader repairs such bytes the same way and *records* every
//! repair, so a caller can enumerate exactly what changed through
//! [`ToleranceReport`]. Nothing is swallowed silently: an empty report means
//! the workbook conformed and no substitution was made.
//!
//! **Excel-compatible repairs** are applied under every [`Leniency`] variant,
//! including the default [`Leniency::Strict`]. They cover bytes MS-XLSB
//! describes as "MUST be 0 and MUST be ignored", payload bytes past the end of
//! a record's defined layout, and producer obligations for which MS-XLSB
//! defines no reader-side refusal. Each substitutes the value documented on
//! the corresponding [`ReservedFieldDefect`] or [`LayoutDefect`] variant.
//!
//! **Opt-in repairs** trade strictness for reach behind explicit [`Leniency`]
//! variants. None exists yet: the [`CachedNumberDefect`] class (MS-XLSB
//! 2.5.153 `Xnum` caches that are NaN, infinite, denormalized, or negative
//! zero) is enumerated so the contract is complete, and will be repaired only
//! under a dedicated opt-in variant because a NaN cache has no unambiguous
//! ground truth. Likewise [`ValueEncodingDefect`] enumerates the value-domain
//! repairs (Boolean bytes, `BErr` codes, UTF-16, data-validation enums) that a
//! later iteration will apply. A class whose repair has not landed is never
//! recorded; the corresponding input is still a hard error.
//!
//! **Structural corruption and security limits stay hard errors** under every
//! variant: BIFF12 record framing (header varints, declared lengths, truncated
//! payloads), stream grammar (begin/end nesting), OPC package structure, and
//! every configurable resource limit.
//!
//! # Where repairs are recorded
//!
//! The report travels with the object whose parse performed the repair:
//!
//! - [`Workbook::tolerance_report`](crate::Workbook::tolerance_report) covers
//!   every part parsed while the workbook opens — the workbook stream
//!   (`workbook.bin`, including `BrtCalcProp`, `BrtBundleSh`, and `BrtName`)
//!   and the styles part (`styles.bin`).
//! - [`Package::tolerance_report`](crate::Package::tolerance_report) is the
//!   report of the validating open that admitted the package.
//! - [`Worksheet::tolerance_report`](crate::Worksheet::tolerance_report)
//!   covers one materialized worksheet stream. Worksheets are parsed lazily,
//!   so each materialization carries its own report instead of mutating the
//!   workbook's.
//! - [`SourceBackedWorkbook::tolerance_report`](crate::SourceBackedWorkbook::tolerance_report)
//!   returns a snapshot of the workbook-scope repairs performed so far: the
//!   workbook stream at open, then the styles part once a worksheet
//!   materialization first needs it. Each lazily parsed part contributes its
//!   repairs exactly once.
//! - [`cell_values::Snapshot::tolerance_report`](crate::cell_values::Snapshot::tolerance_report)
//!   covers the worksheet stream behind one cell-value edit snapshot.
//!
//! Internally, parsers receive `&mut ToleranceReport` (or own one, as the
//! worksheet cell reader does) and call `record` at the exact point they
//! would otherwise have returned an error, then continue with the documented
//! substitute value. Low-level single-record codec functions that cannot carry
//! a report (for example [`styles::parse_border`](crate::styles::parse_border))
//! remain strict, so a repair can never happen without being recorded; the
//! write path keeps validating authored payloads strictly as well.

/// Upper bound on individually recorded defects.
///
/// A hostile or merely enormous workbook must not be able to make the report
/// grow without limit. Defects beyond this bound are counted, not stored; see
/// [`ToleranceReport::unrecorded`].
pub const MAX_RECORDED_DEFECTS: usize = 1024;

/// How the reader treats non-structural workbook defects.
///
/// Structural defects — malformed record framing, a broken stream grammar, an
/// invalid package, or an exceeded resource limit — are hard errors under
/// every variant. The Excel-compatible repairs listed by
/// [`ReservedFieldDefect`] and [`LayoutDefect`] are applied and recorded under
/// every variant, including [`Leniency::Strict`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum Leniency {
    /// Reject every defect whose repair requires opting in.
    ///
    /// This is the default: the Excel-compatible repairs are still performed
    /// and recorded in the [`ToleranceReport`], while every other deviation
    /// is a hard error.
    #[default]
    Strict,
}

/// A reserved or undefined bit field that carried a nonzero value.
///
/// This enumeration is closed by design: it is the exhaustive contract of the
/// reserved-field repairs a read applies under every [`Leniency`] variant
/// (audit finding F3.11). In each case MS-XLSB says the bits "MUST be 0 and
/// MUST be ignored" (or are "undefined and MUST be ignored"); Excel reads
/// past them, so the reader masks the field down to its defined bits and
/// keeps every defined bit. The recorded [`ToleratedDefect::observed`] value
/// is the offending bits that were masked off.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ReservedFieldDefect {
    /// The flag byte of a `Cell` structure (MS-XLSB 2.5.10) had bits other
    /// than `fPhShow` set. Those 7 reserved bits are masked off.
    CellFlags,
    /// `BrtRowHdr` (MS-XLSB 2.4.726) had reserved or unused flag bits set.
    /// They are masked off; outline level, visibility, height, and style
    /// flags are kept.
    RowHeaderFlags,
    /// `BrtColInfo` (MS-XLSB 2.4.323) had reserved or unused flag bits set.
    /// They are masked off; width, visibility, and outline flags are kept.
    ColumnInfoFlags,
    /// A formula cell's `GrbitFmla` (MS-XLSB 2.5.97.3) had bits other than
    /// `fAlwaysCalc` set. They are masked off.
    FormulaFlags,
    /// `BrtBorder` (MS-XLSB 2.4.43) had flag bits other than the two diagonal
    /// bits set. They are masked off; the diagonal flags are kept.
    BorderFlags,
    /// A `Blxf` border side (MS-XLSB 2.5.5) had a nonzero reserved byte. The
    /// byte is ignored and the side is read normally.
    BorderSideReserved,
    /// `BrtCalcProp` (MS-XLSB 2.4.318) had option bits outside the defined
    /// set. They are masked off; every defined calculation switch is kept.
    CalcOptions,
}

impl ReservedFieldDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::CellFlags => "masked the reserved Cell flag bits; fPhShow was kept",
            Self::RowHeaderFlags => {
                "masked the reserved BrtRowHdr flag bits; every defined row flag was kept"
            },
            Self::ColumnInfoFlags => {
                "masked the reserved BrtColInfo flag bits; every defined column flag was kept"
            },
            Self::FormulaFlags => "masked the reserved GrbitFmla bits; fAlwaysCalc was kept",
            Self::BorderFlags => {
                "masked the reserved BrtBorder flag bits; the diagonal flags were kept"
            },
            Self::BorderSideReserved => "ignored the reserved Blxf byte",
            Self::CalcOptions => {
                "masked the undefined BrtCalcProp option bits; every defined switch was kept"
            },
        }
    }
}

/// A record whose layout deviated from MS-XLSB in a way Excel reads past.
///
/// This enumeration is closed by design: it is the exhaustive contract of the
/// layout repairs a read applies under every [`Leniency`] variant (audit
/// finding F3.11). Record framing itself is never repaired: a record whose
/// payload is *shorter* than its layout requires stays a hard error.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LayoutDefect {
    /// A record's payload continued past the end of its defined layout.
    ///
    /// MS-XLSB 2.1.4 lets newer producers append fields that older readers
    /// ignore; Excel advances by the record length. The trailing bytes are
    /// ignored (they remain in the package untouched). The recorded
    /// [`ToleratedDefect::observed`] value is the number of ignored bytes.
    TrailingBytes,
    /// A direct `BrtColor` (MS-XLSB 2.5.17, `xColorType = 2`) had `fValidRGB`
    /// clear, so its RGB bytes carry no color.
    ///
    /// The color is read as absent ("no color"), exactly as for a color of
    /// any other non-RGB type. The recorded observed value is the color's
    /// flag byte.
    ColorWithoutValidRgb,
    /// A `Blxf` border side (MS-XLSB 2.5.5) named a line style outside the
    /// defined `0..=13` range.
    ///
    /// The side is read as having no border. The recorded observed value is
    /// the style byte.
    BorderStyle,
    /// A `BrtFmt` (MS-XLSB 2.4.648) `ifmt` lay outside the identifier ranges
    /// MS-XLSB reserves for workbook-defined formats.
    ///
    /// The format code is kept verbatim under the identifier the file gave
    /// it, as Excel and LibreOffice do. The recorded observed value is the
    /// identifier.
    NumberFormatId,
    /// A `BrtFmt` `stFmtCode` was empty or longer than 255 characters.
    ///
    /// The format code is kept verbatim. The recorded observed value is its
    /// length in UTF-16 code units (saturated to `u32::MAX`).
    NumberFormatCodeLength,
}

impl LayoutDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::TrailingBytes => "ignored payload bytes past the end of the record layout",
            Self::ColorWithoutValidRgb => {
                "read a direct color whose fValidRGB bit is clear as no color"
            },
            Self::BorderStyle => "read a border side with an undefined line style as no border",
            Self::NumberFormatId => {
                "kept a number format whose identifier is outside the reserved ranges"
            },
            Self::NumberFormatCodeLength => {
                "kept a number format code whose length is outside 1..=255"
            },
        }
    }
}

/// A cached `Xnum` value outside the domain MS-XLSB 2.5.153 allows.
///
/// Audit finding F3.10. This class is enumerated to complete the contract but
/// is **not yet repaired**: such a value is still a hard error under every
/// current [`Leniency`] variant, and will be repaired only behind a dedicated
/// opt-in variant because a NaN or infinite cache has no unambiguous ground
/// truth.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum CachedNumberDefect {
    /// A cell value or formula cache was NaN, infinite, denormalized, or
    /// negative zero. The planned repair degrades that single cell.
    NonNormalNumber,
}

impl CachedNumberDefect {
    /// A short, stable description of what the repair does.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::NonNormalNumber => {
                "degraded a cell whose cached number is NaN, infinite, denormalized, or negative zero"
            },
        }
    }
}

/// A value-domain defect in an individual value (audit finding F3.12).
///
/// This class is enumerated to complete the contract but is **not yet
/// repaired**: each input is still a hard error. The documented repairs are
/// the ones a later iteration will apply under every [`Leniency`] variant.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ValueEncodingDefect {
    /// A Boolean byte held a value other than 0 or 1; Excel reads any nonzero
    /// byte as TRUE.
    BooleanByte,
    /// A `BErr` byte held a code outside the defined error values; Excel
    /// renders it as `#NULL!`.
    ErrorCode,
    /// A wide string contained an unpaired UTF-16 surrogate; it decodes with
    /// U+FFFD in its place.
    Utf16,
    /// A `BrtDVal`/`BrtDVal14` record carried an undefined enumeration value
    /// or reserved bit; the rule falls back to its default.
    DataValidation,
}

impl ValueEncodingDefect {
    /// A short, stable description of what the repair does.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::BooleanByte => "read a nonzero Boolean byte as TRUE",
            Self::ErrorCode => "read an undefined BErr code as #NULL!",
            Self::Utf16 => "decoded an unpaired UTF-16 surrogate as U+FFFD",
            Self::DataValidation => "fell back to the default for an undefined validation field",
        }
    }
}

/// One class of defect repaired while a workbook was read.
///
/// This enumeration is closed by design: it is the exhaustive contract of what
/// a read may change about a workbook.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Defect {
    /// A reserved field masked the way Excel ignores it.
    ReservedField(ReservedFieldDefect),
    /// A record layout deviation read past the way Excel reads past it.
    Layout(LayoutDefect),
    /// A cached number outside the `Xnum` domain (not yet repaired).
    CachedNumber(CachedNumberDefect),
    /// A value-domain defect in an individual value (not yet repaired).
    ValueEncoding(ValueEncodingDefect),
}

impl Defect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::ReservedField(defect) => defect.repair(),
            Self::Layout(defect) => defect.repair(),
            Self::CachedNumber(defect) => defect.repair(),
            Self::ValueEncoding(defect) => defect.repair(),
        }
    }
}

impl From<ReservedFieldDefect> for Defect {
    fn from(defect: ReservedFieldDefect) -> Self {
        Self::ReservedField(defect)
    }
}

impl From<LayoutDefect> for Defect {
    fn from(defect: LayoutDefect) -> Self {
        Self::Layout(defect)
    }
}

impl From<CachedNumberDefect> for Defect {
    fn from(defect: CachedNumberDefect) -> Self {
        Self::CachedNumber(defect)
    }
}

impl From<ValueEncodingDefect> for Defect {
    fn from(defect: ValueEncodingDefect) -> Self {
        Self::ValueEncoding(defect)
    }
}

/// One repaired defect, with enough context to locate it in the workbook.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ToleratedDefect {
    defect: Defect,
    record_type: u16,
    location: u32,
    observed: u32,
}

impl ToleratedDefect {
    /// The class of defect that was repaired.
    #[must_use]
    pub const fn defect(&self) -> Defect {
        self.defect
    }

    /// The BIFF12 record type whose payload carried the defect.
    #[must_use]
    pub const fn record_type(&self) -> u16 {
        self.record_type
    }

    /// Where the defect was found within its part.
    ///
    /// For cell records, `BrtRowHdr`, and formula cells this is the
    /// zero-based row. For `BrtColInfo` it is the first column. For style
    /// records (`BrtFont`, `BrtFill`, `BrtBorder`, `BrtXF`, `BrtFmt`) it is
    /// the zero-based ordinal of the record within its collection. For
    /// `BrtBundleSh` and `BrtName` it is the zero-based ordinal of the record
    /// among records of its type. For singleton records such as
    /// `BrtCalcProp` it is zero.
    #[must_use]
    pub const fn location(&self) -> u32 {
        self.location
    }

    /// The offending value read from the file.
    ///
    /// For a [`ReservedFieldDefect`] this is the masked-off bits. For a
    /// [`LayoutDefect`] it is documented on each variant (for
    /// [`LayoutDefect::TrailingBytes`], the number of ignored bytes).
    #[must_use]
    pub const fn observed(&self) -> u32 {
        self.observed
    }
}

/// Everything a read repaired, in the order the defects were found.
///
/// An empty report means the workbook conformed and no substitution was made.
/// After a [`Leniency::Strict`] read the report can list the Excel-compatible
/// [`Defect::ReservedField`] and [`Defect::Layout`] repairs.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ToleranceReport {
    defects: Vec<ToleratedDefect>,
    unrecorded: u32,
}

impl ToleranceReport {
    /// The maximum number of individually recorded defects.
    pub const RECORD_LIMIT: usize = MAX_RECORDED_DEFECTS;

    /// Whether the workbook parsed without any repair.
    #[inline]
    #[must_use]
    pub fn is_clean(&self) -> bool {
        self.defects.is_empty() && self.unrecorded == 0
    }

    /// The individually recorded defects, in discovery order.
    ///
    /// At most [`ToleranceReport::RECORD_LIMIT`] entries; any excess is
    /// reported by [`ToleranceReport::unrecorded`].
    #[inline]
    #[must_use]
    pub fn defects(&self) -> &[ToleratedDefect] {
        &self.defects
    }

    /// Iterate the individually recorded defects in discovery order.
    pub fn iter(&self) -> std::slice::Iter<'_, ToleratedDefect> {
        self.defects.iter()
    }

    /// Defects that were repaired but not individually stored, once the
    /// bound was reached.
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

    /// How many individually recorded defects belong to `defect`.
    ///
    /// Counts only stored entries; check [`ToleranceReport::unrecorded`]
    /// before treating a zero as absence.
    #[must_use]
    pub fn count(&self, defect: impl Into<Defect>) -> usize {
        let defect = defect.into();
        self.defects
            .iter()
            .filter(|entry| entry.defect == defect)
            .count()
    }

    /// Record one repair, saturating at [`MAX_RECORDED_DEFECTS`] stored
    /// entries; the excess is counted in [`ToleranceReport::unrecorded`].
    #[allow(dead_code, reason = "PHASE0: repair sites land with F3.11")]
    pub(crate) fn record(
        &mut self,
        defect: impl Into<Defect>,
        record_type: crate::raw::Kind,
        location: u32,
        observed: u32,
    ) {
        self.push(ToleratedDefect {
            defect: defect.into(),
            record_type: record_type.get(),
            location,
            observed,
        });
    }

    /// Record trailing payload bytes ignored past a record's defined layout.
    #[allow(dead_code, reason = "PHASE0: repair sites land with F3.11")]
    pub(crate) fn record_trailing(
        &mut self,
        record_type: crate::raw::Kind,
        location: u32,
        trailing: usize,
    ) {
        self.record(
            LayoutDefect::TrailingBytes,
            record_type,
            location,
            u32::try_from(trailing).unwrap_or(u32::MAX),
        );
    }

    fn push(&mut self, entry: ToleratedDefect) {
        if self.defects.len() < MAX_RECORDED_DEFECTS {
            self.defects.push(entry);
        } else {
            self.unrecorded = self.unrecorded.saturating_add(1);
        }
    }

    /// Merge another parse's report into this one, keeping discovery order.
    ///
    /// The stored-entry bound applies to the merged report as a whole;
    /// entries that no longer fit join the unrecorded count.
    pub(crate) fn append(&mut self, other: &ToleranceReport) {
        for entry in &other.defects {
            self.push(*entry);
        }
        self.unrecorded = self.unrecorded.saturating_add(other.unrecorded);
    }
}

/// The repair channel threaded through one parse.
///
/// A parser holding a `Tolerance` calls [`Tolerance::tolerate`] at the exact
/// point it would otherwise have returned an error. When the parse records
/// into a [`ToleranceReport`] the defect is recorded and the parser continues
/// with the documented substitute value; when no report can carry the repair
/// (a low-level public codec function returning a bare value), the original
/// error is returned unchanged, so a repair is never applied unrecorded.
#[allow(dead_code, reason = "PHASE0: repair sites land with F3.11")]
#[derive(Debug)]
pub(crate) struct Tolerance<'a> {
    report: Option<&'a mut ToleranceReport>,
}

#[allow(dead_code, reason = "PHASE0: repair sites land with F3.11")]
impl<'a> Tolerance<'a> {
    /// A channel that refuses every defect with the parser's original error.
    pub(crate) const fn refusing() -> Self {
        Self { report: None }
    }

    /// A channel that repairs and records into `report`.
    pub(crate) fn recording(report: &'a mut ToleranceReport) -> Self {
        Self {
            report: Some(report),
        }
    }

    /// Borrow this channel for a nested parse.
    pub(crate) fn reborrow(&mut self) -> Tolerance<'_> {
        Tolerance {
            report: self.report.as_deref_mut(),
        }
    }

    /// Record `defect` and allow the repair, or return `on_refusal()` when
    /// this channel cannot record.
    pub(crate) fn tolerate<E>(
        &mut self,
        defect: impl Into<Defect>,
        record_type: crate::raw::Kind,
        location: u32,
        observed: u32,
        on_refusal: impl FnOnce() -> E,
    ) -> Result<(), E> {
        match self.report.as_deref_mut() {
            Some(report) => {
                report.record(defect, record_type, location, observed);
                Ok(())
            },
            None => Err(on_refusal()),
        }
    }

    /// Allow `trailing` unconsumed payload bytes past a record's defined
    /// layout, recording [`LayoutDefect::TrailingBytes`]; a zero count is
    /// conformant and records nothing.
    pub(crate) fn tolerate_trailing<E>(
        &mut self,
        record_type: crate::raw::Kind,
        location: u32,
        trailing: usize,
        on_refusal: impl FnOnce() -> E,
    ) -> Result<(), E> {
        if trailing == 0 {
            return Ok(());
        }
        match self.report.as_deref_mut() {
            Some(report) => {
                report.record_trailing(record_type, location, trailing);
                Ok(())
            },
            None => Err(on_refusal()),
        }
    }
}

impl<'a> IntoIterator for &'a ToleranceReport {
    type Item = &'a ToleratedDefect;
    type IntoIter = std::slice::Iter<'a, ToleratedDefect>;

    fn into_iter(self) -> Self::IntoIter {
        self.iter()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::raw::kind;

    #[test]
    fn strict_is_the_default() {
        assert_eq!(Leniency::default(), Leniency::Strict);
    }

    #[test]
    fn a_fresh_report_is_clean() {
        let report = ToleranceReport::default();
        assert!(report.is_clean());
        assert_eq!(report.total(), 0);
        assert!(report.defects().is_empty());
        assert_eq!(report.iter().count(), 0);
    }

    #[test]
    fn recorded_defects_are_returned_in_discovery_order() {
        let mut report = ToleranceReport::default();
        report.record(ReservedFieldDefect::CellFlags, kind::CELL_REAL, 4, 0x80);
        report.record_trailing(kind::FMT, 2, 3);
        report.record(ReservedFieldDefect::CellFlags, kind::CELL_BLANK, 9, 0x02);

        assert!(!report.is_clean());
        assert_eq!(report.total(), 3);
        assert_eq!(report.count(ReservedFieldDefect::CellFlags), 2);
        assert_eq!(report.count(LayoutDefect::TrailingBytes), 1);
        assert_eq!(report.count(LayoutDefect::BorderStyle), 0);

        let entries: Vec<_> = (&report)
            .into_iter()
            .map(|entry| {
                (
                    entry.defect(),
                    entry.record_type(),
                    entry.location(),
                    entry.observed(),
                )
            })
            .collect();
        assert_eq!(
            entries,
            vec![
                (
                    Defect::ReservedField(ReservedFieldDefect::CellFlags),
                    kind::CELL_REAL.get(),
                    4,
                    0x80
                ),
                (
                    Defect::Layout(LayoutDefect::TrailingBytes),
                    kind::FMT.get(),
                    2,
                    3
                ),
                (
                    Defect::ReservedField(ReservedFieldDefect::CellFlags),
                    kind::CELL_BLANK.get(),
                    9,
                    0x02
                ),
            ]
        );
    }

    #[test]
    fn trailing_byte_counts_saturate() {
        let mut report = ToleranceReport::default();
        report.record_trailing(kind::NAME, 0, usize::MAX);
        assert_eq!(report.defects()[0].observed(), u32::MAX);
    }

    /// A hostile workbook must not be able to grow the report without limit.
    #[test]
    fn defects_past_the_bound_are_counted_not_stored() {
        let mut report = ToleranceReport::default();
        for row in 0..(MAX_RECORDED_DEFECTS + 5) {
            report.record(
                ReservedFieldDefect::CellFlags,
                kind::CELL_BLANK,
                u32::try_from(row).unwrap(),
                2,
            );
        }
        assert_eq!(report.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(report.unrecorded(), 5);
        assert_eq!(report.total(), MAX_RECORDED_DEFECTS as u64 + 5);
        assert_eq!(ToleranceReport::RECORD_LIMIT, MAX_RECORDED_DEFECTS);
    }

    #[test]
    fn appended_reports_merge_within_the_bound() {
        let mut base = ToleranceReport::default();
        base.record(ReservedFieldDefect::CalcOptions, kind::CALC_PROP, 0, 0x8000);

        let mut parsed = ToleranceReport::default();
        for row in 0..MAX_RECORDED_DEFECTS {
            parsed.record(
                ReservedFieldDefect::CellFlags,
                kind::CELL_BLANK,
                u32::try_from(row).unwrap(),
                2,
            );
        }
        parsed.unrecorded = 2;

        base.append(&parsed);
        assert_eq!(base.defects().len(), MAX_RECORDED_DEFECTS);
        assert_eq!(base.unrecorded(), 3);
        assert_eq!(base.total(), MAX_RECORDED_DEFECTS as u64 + 3);
        assert_eq!(
            base.defects()[0].defect(),
            Defect::ReservedField(ReservedFieldDefect::CalcOptions)
        );

        let mut clean = ToleranceReport::default();
        clean.append(&ToleranceReport::default());
        assert!(clean.is_clean());
    }

    #[test]
    fn a_refusing_channel_returns_the_original_error_and_records_nothing() {
        let mut tolerance = Tolerance::refusing();
        let refused = tolerance.tolerate(
            ReservedFieldDefect::CellFlags,
            kind::CELL_BLANK,
            0,
            2,
            || "strict",
        );
        assert_eq!(refused, Err("strict"));
        assert_eq!(
            tolerance.tolerate_trailing(kind::FMT, 0, 1, || "strict"),
            Err("strict")
        );
        // A conformant layout is never a defect, even without a report.
        assert_eq!(
            tolerance.tolerate_trailing(kind::FMT, 0, 0, || "strict"),
            Ok(())
        );
    }

    #[test]
    fn a_recording_channel_records_through_reborrows() {
        let mut report = ToleranceReport::default();
        {
            let mut tolerance = Tolerance::recording(&mut report);
            tolerance
                .tolerate(
                    ReservedFieldDefect::BorderFlags,
                    kind::BORDER,
                    1,
                    0x04,
                    || "strict",
                )
                .unwrap();
            let mut nested = tolerance.reborrow();
            nested
                .tolerate_trailing(kind::FMT, 3, 2, || "strict")
                .unwrap();
            nested
                .tolerate_trailing(kind::FMT, 4, 0, || "strict")
                .unwrap();
        }
        assert_eq!(report.total(), 2);
        assert_eq!(report.count(ReservedFieldDefect::BorderFlags), 1);
        assert_eq!(report.defects()[1].location(), 3);
        assert_eq!(report.defects()[1].observed(), 2);
    }

    #[test]
    fn every_defect_class_describes_its_repair() {
        let defects = [
            Defect::from(ReservedFieldDefect::CellFlags),
            Defect::from(ReservedFieldDefect::RowHeaderFlags),
            Defect::from(ReservedFieldDefect::ColumnInfoFlags),
            Defect::from(ReservedFieldDefect::FormulaFlags),
            Defect::from(ReservedFieldDefect::BorderFlags),
            Defect::from(ReservedFieldDefect::BorderSideReserved),
            Defect::from(ReservedFieldDefect::CalcOptions),
            Defect::from(LayoutDefect::TrailingBytes),
            Defect::from(LayoutDefect::ColorWithoutValidRgb),
            Defect::from(LayoutDefect::BorderStyle),
            Defect::from(LayoutDefect::NumberFormatId),
            Defect::from(LayoutDefect::NumberFormatCodeLength),
            Defect::from(CachedNumberDefect::NonNormalNumber),
            Defect::from(ValueEncodingDefect::BooleanByte),
            Defect::from(ValueEncodingDefect::ErrorCode),
            Defect::from(ValueEncodingDefect::Utf16),
            Defect::from(ValueEncodingDefect::DataValidation),
        ];
        for defect in defects {
            assert!(!defect.repair().is_empty(), "{defect:?}");
        }
        assert_eq!(
            Defect::Layout(LayoutDefect::TrailingBytes).repair(),
            LayoutDefect::TrailingBytes.repair()
        );
    }
}
