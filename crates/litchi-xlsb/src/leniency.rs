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
//! the corresponding [`ReservedFieldDefect`], [`LayoutDefect`], or
//! [`ValueEncodingDefect`] variant. The value-encoding repairs (audit finding
//! F3.12) cover Boolean bytes, `BErr` codes, and UTF-16 in cell values,
//! formula caches, shared strings, sheet/defined/font/format names and
//! comments, plus `BrtDVal`/`BrtDVal14` flag words.
//! They are applied in the workbook stream (`BrtCalcProp`, `BrtBundleSh`,
//! `BrtName`), the styles part (`BrtFont`/`BrtFill` colors, `BrtBorder`,
//! `BrtFmt`), and worksheet streams (`Cell` flags, `BrtRowHdr`, `BrtColInfo`,
//! formula-cell `GrbitFmla` and trailing bytes), to trailing payload bytes in
//! the adjunct parts parsed with the workbook — table parts, PivotCache
//! definitions, External Data Connections, and comments — and to `Cell` flags
//! on the binary-index point lookup
//! ([`SourceBackedIndexedWorksheet::cached_value`](crate::SourceBackedIndexedWorksheet::cached_value)).
//!
//! **Opt-in repairs** trade strictness for reach behind explicit [`Leniency`]
//! variants. [`Leniency::TolerateCachedNumbers`] repairs the
//! [`CachedNumberDefect`] class (MS-XLSB 2.5.153 `Xnum` caches that are NaN,
//! infinite, denormalized, or negative zero, audit finding F3.10): a NaN cache
//! has no unambiguous ground truth, so the repair is never applied by
//! default. Select it with [`Workbook::new_with_leniency`](crate::Workbook::new_with_leniency),
//! [`Workbook::from_opc_package_with_leniency`](crate::Workbook::from_opc_package_with_leniency),
//! or [`SourceBackedWorkbook::from_reader_with_leniency`](crate::SourceBackedWorkbook::from_reader_with_leniency).
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
//!   the OPC package read that admitted the workbook and every part parsed
//!   while the workbook opens — the workbook stream
//!   (`workbook.bin`, including `BrtCalcProp`, `BrtBundleSh`, and `BrtName`),
//!   the styles part (`styles.bin`), and the adjunct parts parsed at open:
//!   table parts, PivotCache definitions, and External Data Connections.
//! - [`Package::tolerance_report`](crate::Package::tolerance_report) is the
//!   report of the validating open that admitted the package.
//! - [`Worksheet::tolerance_report`](crate::Worksheet::tolerance_report)
//!   covers one materialized worksheet stream and its comments part.
//!   Worksheets are parsed lazily,
//!   so each materialization carries its own report instead of mutating the
//!   workbook's.
//! - [`SourceBackedWorkbook::tolerance_report`](crate::SourceBackedWorkbook::tolerance_report)
//!   returns a snapshot of the workbook-scope repairs performed so far: the
//!   OPC package read and the workbook stream at open, then the styles and table parts once a
//!   worksheet materialization first needs them. Each lazily parsed part
//!   contributes its repairs exactly once. Each binary-index lookup through
//!   `SourceBackedIndexedWorksheet::cached_value` that performs a repair
//!   records it here as well.
//! - [`cell_values::Snapshot::tolerance_report`](crate::cell_values::Snapshot::tolerance_report)
//!   covers the worksheet stream behind one cell-value edit snapshot.
//!
//! # Packaging repairs
//!
//! The OPC layer applies its own reference-compatible repairs while the
//! package is read — duplicate relationship `Id`s, unresolvable internal
//! targets, `TargetMode` casing, missing `Target`/`Type`, duplicate
//! content-type mappings, and the manifest findings documented in
//! `litchi_opc::leniency`. Every open path (eager [`Workbook`](crate::Workbook),
//! validating [`Package`](crate::Package), and
//! [`SourceBackedWorkbook`](crate::SourceBackedWorkbook), including an OPC
//! package materialized from a source-backed one) folds that report into the
//! workbook-scope report first, as [`Defect::Package`] entries that keep the
//! OPC location ([`ToleratedDefect::part`], [`ToleratedDefect::item`],
//! [`ToleratedDefect::value`]), and carries the OPC unrecorded count over.
//! The XLSB package does not read core or custom properties at open. A caller
//! that reads them alongside the workbook (the `litchi` facade does) folds
//! the metadata report in with [`ToleranceReport::append_properties`], as
//! [`Defect::Properties`] entries. Worksheet-scope reports never repeat
//! packaging defects.
//!
//! Internally, parsers receive `&mut ToleranceReport` (or own one, as the
//! worksheet cell reader does) and call `record` at the exact point they
//! would otherwise have returned an error, then continue with the documented
//! substitute value. Parsers built on the shared record walker (tables,
//! PivotCache definitions, connections) carry the channel in the walker,
//! which attributes trailing bytes to the record it yielded last. Low-level
//! single-record codec functions that cannot carry
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
    /// Also repair and record the cached-number defects listed by
    /// [`CachedNumberDefect`] (audit finding F3.10).
    ///
    /// MS-XLSB 2.5.153 makes a finite, normalized `Xnum` a producer
    /// obligation, but add-ins and RTD feeds cache NaN and infinities and
    /// third-party writers serialize `f64` verbatim; Excel opens such files.
    /// A NaN or infinite cache has no unambiguous ground truth, so this
    /// repair is opt-in: the affected cell reads as the `#NUM!` error Excel
    /// and LibreOffice display for a number they cannot represent, and a
    /// denormalized or negative-zero cache reads as `0`, the value Excel
    /// flushes it to. Every other cell keeps its value.
    TolerateCachedNumbers,
}

impl Leniency {
    /// Whether non-normal cached numbers are repaired rather than rejected.
    #[inline]
    #[must_use]
    pub const fn tolerates_cached_numbers(self) -> bool {
        matches!(self, Self::TolerateCachedNumbers)
    }
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
    /// `BrtName` (MS-XLSB 2.4.647) had any of its 14 reserved flag bits set.
    /// They are masked off; every defined name flag is kept.
    NameFlags,
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
            Self::NameFlags => {
                "masked the reserved BrtName flag bits; every defined name flag was kept"
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
/// Audit finding F3.10. These repairs are **opt-in**: they are applied only
/// under [`Leniency::TolerateCachedNumbers`]; under [`Leniency::Strict`] the
/// cell-value snapshot and the binary-index lookup reject such a value as
/// before, while materialized worksheets keep passing the stored `f64`
/// through verbatim as they always have. Under the opt-in all three read
/// the documented substitute and record it. The recorded [`ToleratedDefect::observed`] value is the upper 32
/// bits of the IEEE-754 bit pattern (sign, exponent, and high mantissa), and
/// the location is the cell's zero-based row.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum CachedNumberDefect {
    /// A cell value or formula cache was NaN or infinite. The cell reads as
    /// the `#NUM!` error (a formula keeps its formula with a `#NUM!` cache).
    NonFinite,
    /// A cell value or formula cache was denormalized or negative zero. The
    /// cell reads as `0`.
    Denormalized,
}

impl CachedNumberDefect {
    /// A short, stable description of what the repair did.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::NonFinite => "read a NaN or infinite cached number as the #NUM! error",
            Self::Denormalized => "read a denormalized or negative-zero cached number as 0",
        }
    }
}

/// The value a cached number reads as after [`Tolerance::cached_number`].
#[derive(Debug, Clone, Copy, PartialEq)]
pub(crate) enum CachedNumber {
    /// A conformant number, or a denormalized one flushed to `0`.
    Number(f64),
    /// A NaN or infinite cache, read as `#NUM!`.
    NumError,
}

/// A value-domain defect in an individual value (audit finding F3.12).
///
/// This enumeration is closed by design: it is the exhaustive contract of the
/// value repairs a read applies under every [`Leniency`] variant. Each matches
/// what Excel does with a value MS-XLSB enumerates as a producer obligation,
/// and only the one value changes.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ValueEncodingDefect {
    /// A cell or formula-cache Boolean byte held a value other than 0 or 1;
    /// Excel reads any nonzero byte as TRUE, and so does the reader. Observed:
    /// the byte.
    BooleanByte,
    /// A cell or formula-cache `BErr` byte held a code outside the defined
    /// error values. Materialized worksheets show the generic `#ERR!` text
    /// they have always used for it; the typed cell-value model, which has
    /// no generic error, reads it as `#NULL!`, as Excel renders it. Observed:
    /// the byte.
    ErrorCode,
    /// A wide string contained unpaired UTF-16 surrogates; each decodes as
    /// U+FFFD. Observed: the number of replacements.
    Utf16,
    /// A `BrtDVal`/`BrtDVal14` flag word carried an undefined `valType`,
    /// `errStyle`, `mdImeMode`, or `typOperator` value, or reserved bits. The
    /// undefined fields fall back to AnyValue, Stop, NoControl, and Between,
    /// reserved bits are masked, and the rule's stored formulas are kept
    /// without the type/operator presence cross-check. Observed: the original
    /// flag word; location: the rule's ordinal on its worksheet.
    DataValidation,
}

impl ValueEncodingDefect {
    /// A short, stable description of what the repair does.
    #[must_use]
    pub const fn repair(self) -> &'static str {
        match self {
            Self::BooleanByte => "read a nonzero Boolean byte as TRUE",
            Self::ErrorCode => "read an undefined BErr code as a generic error (#ERR!, or #NULL!)",
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
    /// A packaging defect the OPC layer repaired while the workbook's package
    /// was read (`litchi_opc::leniency`): relationship manifests,
    /// `[Content_Types].xml`, and the relationship graph.
    ///
    /// The entry's [`ToleratedDefect::part`], [`ToleratedDefect::item`], and
    /// [`ToleratedDefect::value`] carry the package location the OPC report
    /// recorded; its record type, location, and observed value are zero.
    Package(litchi_opc::leniency::Defect),
    /// A core- or custom-properties defect `litchi_ooxml_common` repaired
    /// while the package metadata was read (`litchi_ooxml_common::leniency`,
    /// audit findings 8.1–8.4).
    ///
    /// The XLSB reader does not read metadata at open; a caller that reads it
    /// alongside the workbook folds the metadata report in with
    /// [`ToleranceReport::append_properties`]. The entry's
    /// [`ToleratedDefect::part`], [`ToleratedDefect::item`], and
    /// [`ToleratedDefect::value`] carry the metadata location; its record
    /// type, location, and observed value are zero.
    Properties(litchi_ooxml_common::leniency::Defect),
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
            Self::Package(defect) => defect.repair(),
            Self::Properties(defect) => defect.repair(),
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

impl From<litchi_opc::leniency::Defect> for Defect {
    fn from(defect: litchi_opc::leniency::Defect) -> Self {
        Self::Package(defect)
    }
}

impl From<litchi_ooxml_common::leniency::Defect> for Defect {
    fn from(defect: litchi_ooxml_common::leniency::Defect) -> Self {
        Self::Properties(defect)
    }
}

/// The package location of a [`Defect::Package`] or [`Defect::Properties`]
/// entry, as the lower layer's report recorded it.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
struct PackageLocation {
    part: Box<str>,
    item: Option<Box<str>>,
    value: Option<Box<str>>,
}

/// One repaired defect, with enough context to locate it in the workbook.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct ToleratedDefect {
    defect: Defect,
    record_type: u16,
    location: u32,
    observed: u32,
    package: Option<Box<PackageLocation>>,
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
    /// `BrtCalcProp` it is zero. For records in adjunct parts (table parts,
    /// PivotCache definitions, connections, comments) it is the byte offset
    /// of the record header within its part, saturated to `u32::MAX`.
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

    /// For a [`Defect::Package`] entry, the OPC part name the defect belongs
    /// to: the relationship source part (`/` for package relationships) or
    /// `/[Content_Types].xml`. For a [`Defect::Properties`] entry, the
    /// metadata part. `None` for every BIFF12 record defect.
    #[must_use]
    pub fn part(&self) -> Option<&str> {
        self.package.as_deref().map(|location| &*location.part)
    }

    /// For a [`Defect::Package`] entry, the item within the part — a
    /// relationship `Id`, or the extension or part name of a content-type
    /// mapping — when the OPC report recorded one; for a
    /// [`Defect::Properties`] entry, the property or element name.
    #[must_use]
    pub fn item(&self) -> Option<&str> {
        self.package
            .as_deref()
            .and_then(|location| location.item.as_deref())
    }

    /// For a [`Defect::Package`] entry, the raw value the OPC repair replaced
    /// or discarded, when the OPC report recorded one; for a
    /// [`Defect::Properties`] entry, the raw metadata value.
    #[must_use]
    pub fn value(&self) -> Option<&str> {
        self.package
            .as_deref()
            .and_then(|location| location.value.as_deref())
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
            package: None,
        });
    }

    /// Fold the report of the OPC read that admitted the workbook's package,
    /// keeping its discovery order and carrying its unrecorded count over so
    /// the merged total stays exact.
    pub(crate) fn append_package(&mut self, package: &litchi_opc::leniency::ToleranceReport) {
        for tolerated in package {
            self.push(ToleratedDefect {
                defect: Defect::Package(tolerated.defect()),
                record_type: 0,
                location: 0,
                observed: 0,
                package: Some(Box::new(PackageLocation {
                    part: tolerated.part().into(),
                    item: tolerated.item().map(Into::into),
                    value: tolerated.value().map(Into::into),
                })),
            });
        }
        self.unrecorded = self.unrecorded.saturating_add(package.unrecorded());
    }

    /// Fold a core- or custom-properties report (`litchi_ooxml_common`) read
    /// alongside the workbook into this one.
    ///
    /// Every metadata repair becomes a [`Defect::Properties`] entry that keeps
    /// the metadata location ([`ToleratedDefect::part`],
    /// [`ToleratedDefect::item`], [`ToleratedDefect::value`]), in discovery
    /// order, and the metadata report's unrecorded count is carried over.
    /// Callers that read package metadata with
    /// `litchi_ooxml_common::properties::read_with_report` use this to keep
    /// one report for the whole document.
    pub fn append_properties(
        &mut self,
        properties: &litchi_ooxml_common::leniency::ToleranceReport,
    ) {
        for tolerated in properties {
            self.push(ToleratedDefect {
                defect: Defect::Properties(tolerated.defect()),
                record_type: 0,
                location: 0,
                observed: 0,
                package: Some(Box::new(PackageLocation {
                    part: tolerated.part().into(),
                    item: tolerated.item().map(Into::into),
                    value: tolerated.value().map(Into::into),
                })),
            });
        }
        self.unrecorded = self.unrecorded.saturating_add(properties.unrecorded());
    }

    /// A workbook-scope report seeded with the packaging repairs of the OPC
    /// read that admitted the workbook's package.
    pub(crate) fn from_package(package: &litchi_opc::leniency::ToleranceReport) -> Self {
        let mut report = Self::default();
        report.append_package(package);
        report
    }

    /// Record trailing payload bytes ignored past a record's defined layout.
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
    pub fn append(&mut self, other: &ToleranceReport) {
        for entry in &other.defects {
            self.push(entry.clone());
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
#[derive(Debug)]
pub(crate) struct Tolerance<'a> {
    report: Option<&'a mut ToleranceReport>,
    leniency: Leniency,
}

impl<'a> Tolerance<'a> {
    /// A channel that refuses every defect with the parser's original error.
    pub(crate) const fn refusing() -> Self {
        Self {
            report: None,
            leniency: Leniency::Strict,
        }
    }

    /// A channel that applies the default repairs and records into `report`.
    pub(crate) fn recording(report: &'a mut ToleranceReport) -> Self {
        Self::recording_with(report, Leniency::Strict)
    }

    /// A channel that applies the default repairs plus the opt-in repairs
    /// `leniency` selects, recording into `report`.
    pub(crate) fn recording_with(report: &'a mut ToleranceReport, leniency: Leniency) -> Self {
        Self {
            report: Some(report),
            leniency,
        }
    }

    /// Borrow this channel for a nested parse.
    pub(crate) fn reborrow(&mut self) -> Tolerance<'_> {
        Tolerance {
            report: self.report.as_deref_mut(),
            leniency: self.leniency,
        }
    }

    /// Read a cached `Xnum`, applying the opt-in [`CachedNumberDefect`]
    /// repair under [`Leniency::TolerateCachedNumbers`] and refusing a
    /// non-normal value with `on_refusal()` otherwise.
    pub(crate) fn cached_number<E>(
        &mut self,
        value: f64,
        record_type: crate::raw::Kind,
        row: u32,
        on_refusal: impl FnOnce() -> E,
    ) -> Result<CachedNumber, E> {
        if value.is_normal() || value.to_bits() == 0 {
            return Ok(CachedNumber::Number(value));
        }
        let Some(report) = self.report.as_deref_mut() else {
            return Err(on_refusal());
        };
        if !self.leniency.tolerates_cached_numbers() {
            return Err(on_refusal());
        }
        let observed = u32::try_from(value.to_bits() >> 32).unwrap_or(u32::MAX);
        if value.is_finite() {
            report.record(CachedNumberDefect::Denormalized, record_type, row, observed);
            Ok(CachedNumber::Number(0.0))
        } else {
            report.record(CachedNumberDefect::NonFinite, record_type, row, observed);
            Ok(CachedNumber::NumError)
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

impl Tolerance<'_> {
    /// Finish a payload cursor: trailing bytes past the decoded layout are
    /// ignored and recorded as [`LayoutDefect::TrailingBytes`], or refused
    /// with the cursor's own error when this channel cannot record.
    pub(crate) fn finish(
        &mut self,
        cursor: &crate::raw::Cursor<'_>,
        record_type: crate::raw::Kind,
        location: u32,
    ) -> Result<(), crate::raw::Error> {
        let trailing = cursor.remaining();
        match self.report.as_deref_mut() {
            Some(report) if trailing != 0 => {
                report.record_trailing(record_type, location, trailing);
                Ok(())
            },
            Some(_) | None => cursor.finish(),
        }
    }
}

impl Tolerance<'_> {
    /// Decode one `XLWideString`, replacing unpaired UTF-16 surrogates with
    /// U+FFFD and recording [`ValueEncodingDefect::Utf16`] (observed: the
    /// number of replacements) when this channel records; strict otherwise.
    pub(crate) fn read_wide_string(
        &mut self,
        cursor: &mut crate::raw::Cursor<'_>,
        record_type: crate::raw::Kind,
        location: u32,
    ) -> Result<String, crate::raw::Error> {
        let Some(report) = self.report.as_deref_mut() else {
            return cursor.read_wide_string();
        };
        let (value, replaced) = cursor.read_wide_string_lossy()?;
        record_utf16(report, record_type, location, replaced);
        Ok(value)
    }

    /// Decode one `XLNullableWideString`; see [`Self::read_wide_string`].
    pub(crate) fn read_nullable_wide_string(
        &mut self,
        cursor: &mut crate::raw::Cursor<'_>,
        record_type: crate::raw::Kind,
        location: u32,
    ) -> Result<Option<String>, crate::raw::Error> {
        let Some(report) = self.report.as_deref_mut() else {
            return cursor.read_nullable_wide_string();
        };
        let (value, replaced) = cursor.read_nullable_wide_string_lossy()?;
        record_utf16(report, record_type, location, replaced);
        Ok(value)
    }

    /// Read a byte Boolean: any nonzero byte reads as TRUE, recording
    /// [`ValueEncodingDefect::BooleanByte`] (observed: the byte) for values
    /// other than 0 and 1 when this channel records; strict otherwise.
    pub(crate) fn read_bool8(
        &mut self,
        cursor: &mut crate::raw::Cursor<'_>,
        record_type: crate::raw::Kind,
        location: u32,
    ) -> Result<bool, crate::raw::Error> {
        if self.report.is_none() {
            return cursor.read_bool8();
        }
        let offset = cursor.position();
        let value = cursor.read_u8()?;
        self.bool_byte(value, record_type, location, || {
            crate::raw::Error::InvalidBool {
                value: u32::from(value),
                offset,
            }
        })
    }

    /// Interpret a Boolean byte the way Excel does (nonzero is TRUE),
    /// recording [`ValueEncodingDefect::BooleanByte`] for values other than 0
    /// and 1, or return `on_refusal()` when this channel cannot record.
    pub(crate) fn bool_byte<E>(
        &mut self,
        value: u8,
        record_type: crate::raw::Kind,
        location: u32,
        on_refusal: impl FnOnce() -> E,
    ) -> Result<bool, E> {
        if value > 1 {
            self.tolerate(
                ValueEncodingDefect::BooleanByte,
                record_type,
                location,
                u32::from(value),
                on_refusal,
            )?;
        }
        Ok(value != 0)
    }
}

fn record_utf16(
    report: &mut ToleranceReport,
    record_type: crate::raw::Kind,
    location: u32,
    replaced: usize,
) {
    if replaced != 0 {
        report.record(
            ValueEncodingDefect::Utf16,
            record_type,
            location,
            u32::try_from(replaced).unwrap_or(u32::MAX),
        );
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
    fn cached_numbers_are_repaired_only_under_the_opt_in() {
        let mut report = ToleranceReport::default();
        for value in [1.5, 0.0, -2.0e300] {
            let mut tolerance = Tolerance::refusing();
            assert_eq!(
                tolerance.cached_number(value, kind::CELL_REAL, 0, || "strict"),
                Ok(CachedNumber::Number(value))
            );
        }
        assert_eq!(
            Tolerance::recording(&mut report)
                .cached_number(f64::NAN, kind::CELL_REAL, 0, || { "strict" }),
            Err("strict")
        );
        assert!(report.is_clean());

        let mut tolerance = Tolerance::recording_with(&mut report, Leniency::TolerateCachedNumbers);
        assert_eq!(
            tolerance.cached_number(f64::INFINITY, kind::CELL_REAL, 7, || "strict"),
            Ok(CachedNumber::NumError)
        );
        assert_eq!(
            tolerance
                .reborrow()
                .cached_number(-0.0, kind::FMLA_NUM, 8, || "strict"),
            Ok(CachedNumber::Number(0.0))
        );
        assert_eq!(report.count(CachedNumberDefect::NonFinite), 1);
        assert_eq!(report.count(CachedNumberDefect::Denormalized), 1);
        assert_eq!(report.defects()[0].observed(), 0x7FF0_0000);
        assert_eq!(report.defects()[1].observed(), 0x8000_0000);
        assert!(Leniency::TolerateCachedNumbers.tolerates_cached_numbers());
        assert!(!Leniency::Strict.tolerates_cached_numbers());
    }

    #[test]
    fn value_encodings_are_repaired_when_recording_and_refused_otherwise() {
        // "A", an unpaired high surrogate, "B".
        let data = [3, 0, 0, 0, 0x41, 0, 0x00, 0xD8, 0x42, 0];
        let mut cursor = crate::raw::Cursor::new(&data, "test");
        assert!(
            Tolerance::refusing()
                .read_wide_string(&mut cursor, kind::CELL_ST, 0)
                .is_err()
        );

        let mut report = ToleranceReport::default();
        let mut tolerance = Tolerance::recording(&mut report);
        let mut cursor = crate::raw::Cursor::new(&data, "test");
        assert_eq!(
            tolerance
                .read_wide_string(&mut cursor, kind::CELL_ST, 4)
                .unwrap(),
            "A\u{FFFD}B"
        );
        let null = [0xFF; 4];
        let mut cursor = crate::raw::Cursor::new(&null, "test");
        assert_eq!(
            tolerance
                .read_nullable_wide_string(&mut cursor, kind::NAME, 0)
                .unwrap(),
            None
        );
        let mut cursor = crate::raw::Cursor::new(&[5], "test");
        assert!(
            tolerance
                .read_bool8(&mut cursor, kind::CELL_BOOL, 6)
                .unwrap()
        );
        assert_eq!(
            tolerance.bool_byte(0, kind::CELL_BOOL, 6, || "strict"),
            Ok(false)
        );
        assert_eq!(
            Tolerance::refusing().bool_byte(2, kind::CELL_BOOL, 6, || "strict"),
            Err("strict")
        );
        assert_eq!(report.count(ValueEncodingDefect::Utf16), 1);
        assert_eq!(report.count(ValueEncodingDefect::BooleanByte), 1);
        assert_eq!(report.defects()[0].observed(), 1);
        assert_eq!(report.defects()[1].observed(), 5);
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
            Defect::from(ReservedFieldDefect::NameFlags),
            Defect::from(LayoutDefect::TrailingBytes),
            Defect::from(LayoutDefect::ColorWithoutValidRgb),
            Defect::from(LayoutDefect::BorderStyle),
            Defect::from(LayoutDefect::NumberFormatId),
            Defect::from(LayoutDefect::NumberFormatCodeLength),
            Defect::from(CachedNumberDefect::NonFinite),
            Defect::from(CachedNumberDefect::Denormalized),
            Defect::from(ValueEncodingDefect::BooleanByte),
            Defect::from(ValueEncodingDefect::ErrorCode),
            Defect::from(ValueEncodingDefect::Utf16),
            Defect::from(ValueEncodingDefect::DataValidation),
            Defect::from(litchi_opc::leniency::Defect::from(
                litchi_opc::leniency::RelationshipDefect::DuplicateId,
            )),
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
