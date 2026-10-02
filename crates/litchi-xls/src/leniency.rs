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
//! documented on the corresponding [`CellFormatDefect`], [`EncodingDefect`],
//! [`PaletteDefect`], [`SheetDefect`], or [`SharedStringDefect`] variant.
//! The compound-file repairs the container reader applied while the
//! workbook's OLE file was opened join the same report as
//! [`Defect::Container`] entries, ahead of every BIFF-level repair, and the
//! property-set repairs made while the eager [`crate::Workbook`] open reads
//! the standard metadata streams join it as [`Defect::PropertySet`] entries,
//! after them. (The source-backed owner reads no metadata stream, so its
//! report never lists a property-set repair.)
//!
//! **Opt-in repairs** trade strictness for reach behind explicit [`Leniency`]
//! variants. The defects enumerated by [`FormattingDefect`] — cosmetic
//! formatting metadata that contradicts itself — are repaired with a
//! documented substitute value only under
//! [`Leniency::TolerateFormattingDefects`]; under [`Leniency::Strict`] they
//! remain hard errors, because rejecting them is the only way to guarantee a
//! caller never silently reads a value the file did not contain. Worksheet
//! names that break the MS-XLS naming rules ([`SheetNameDefect`]) are kept
//! verbatim only under [`Leniency::TolerateSheetNameDefects`];
//! [`Leniency::TolerateAllDefects`] enables both opt-in classes.

use super::error::{Error, Result};

/// The standard OLE property-set streams carrying the workbook metadata.
const PROPERTY_SET_STREAMS: [&str; 2] = [
    "\u{0005}SummaryInformation",
    "\u{0005}DocumentSummaryInformation",
];

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
    /// Repair and record the worksheet-name defects listed by
    /// [`SheetNameDefect`].
    ///
    /// Each name is kept verbatim and the repair is recorded as a
    /// [`ToleratedDefect`]; sheets stay addressable by index. The
    /// [`FormattingDefect`] classes are still rejected.
    TolerateSheetNameDefects,
    /// Repair and record every opt-in class: both the [`FormattingDefect`]
    /// and the [`SheetNameDefect`] classes.
    TolerateAllDefects,
}

impl Leniency {
    /// Whether formatting defects are repaired rather than rejected.
    #[must_use]
    pub fn tolerates_formatting_defects(self) -> bool {
        matches!(
            self,
            Leniency::TolerateFormattingDefects | Leniency::TolerateAllDefects
        )
    }

    /// Whether worksheet-name defects are repaired rather than rejected.
    #[must_use]
    pub fn tolerates_sheet_name_defects(self) -> bool {
        matches!(
            self,
            Leniency::TolerateSheetNameDefects | Leniency::TolerateAllDefects
        )
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
    /// `Font.dyHeight` was neither zero nor within `20..=8191` twips.
    ///
    /// Repaired by clamping into `20..=8191`.
    FontHeight,
    /// `Font.icv` was not a valid `Icv`.
    ///
    /// Repaired to 0x7FFF, the automatic (window-text) font color.
    FontColor,
    /// `Font.bls` was neither zero nor within `100..=1000`.
    ///
    /// Repaired by clamping into `100..=1000`.
    FontWeight,
    /// `Font.sss` (escapement) was outside `0..=2`.
    ///
    /// Repaired to `Normal`: no superscript or subscript.
    FontEscapement,
    /// `Font.uls` was not one of the five defined underline styles.
    ///
    /// Repaired to `Single`, keeping the record's claim that the text is
    /// underlined.
    FontUnderline,
    /// `Font.bCharSet` was outside the character-set enumeration.
    ///
    /// Repaired to `Default` (1). The name and glyphs are unaffected; BIFF8
    /// font names are stored as Unicode.
    FontCharset,
    /// `Font.cch` exceeded 31 characters, though the payload held them all.
    ///
    /// Repaired by keeping the whole name as stored.
    FontNameLength,
    /// The font name contained a NUL character.
    ///
    /// Repaired by keeping the name as stored, NUL included.
    FontNameNul,
    /// The workbook held fewer than the four `Font` records MS-XLS requires.
    ///
    /// Repaired by indexing the font table as read; references to an absent
    /// font resolve to no font.
    FontCount,
}

impl FormattingDefect {
    /// The BIFF record type whose payload carried the defect.
    #[must_use]
    pub fn record_type(self) -> u16 {
        match self {
            FormattingDefect::FontFamily
            | FormattingDefect::FontNameEmpty
            | FormattingDefect::FontHeight
            | FormattingDefect::FontColor
            | FormattingDefect::FontWeight
            | FormattingDefect::FontEscapement
            | FormattingDefect::FontUnderline
            | FormattingDefect::FontCharset
            | FormattingDefect::FontNameLength
            | FormattingDefect::FontNameNul
            | FormattingDefect::FontCount => super::font::FONT_RECORD_TYPE,
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
            FormattingDefect::FontHeight => "Font height is outside 20..=8191 twips",
            FormattingDefect::FontColor => "Font color index is not a valid Icv",
            FormattingDefect::FontWeight => "Font weight is outside 100..=1000",
            FormattingDefect::FontEscapement => "Font escapement is outside the enumeration",
            FormattingDefect::FontUnderline => "Font underline style is outside the enumeration",
            FormattingDefect::FontCharset => "Font character set is outside the enumeration",
            FormattingDefect::FontNameLength => "Font name is longer than 31 characters",
            FormattingDefect::FontNameNul => "Font name contains a NUL character",
            FormattingDefect::FontCount => "Workbook has fewer than four Font records",
        }
    }
}

/// The class of worksheet-name defect an opt-in read is allowed to repair.
///
/// MS-XLS 2.4.28 constrains `BoundSheet8.stName` only as a producer
/// obligation; POI performs no name validation on read and LibreOffice opens
/// and renames. Because a name is how formulas and callers address a sheet,
/// accepting one that breaks the rules is a judgment call, so these repairs
/// apply only under [`Leniency::TolerateSheetNameDefects`] (or
/// [`Leniency::TolerateAllDefects`]). Every repair keeps the name exactly as
/// stored; nothing is renamed.
///
/// This enumeration is closed by design: it is the exhaustive contract of
/// what the opt-in may change about a workbook.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SheetNameDefect {
    /// The name was empty.
    Empty,
    /// The name had more than 31 UTF-16 code units.
    TooLong,
    /// The name contained NUL, U+0003, or one of `: \ * ? / [ ]`.
    ForbiddenCharacter,
    /// The name began or ended with an apostrophe.
    EdgeApostrophe,
    /// The name repeated, ignoring case, the name of an earlier sheet.
    ///
    /// Both sheets are kept; a lookup by name resolves to the earlier one,
    /// while every sheet stays addressable by index.
    Duplicate,
}

impl SheetNameDefect {
    /// The BIFF record type whose payload carried the defect: `BoundSheet8`.
    #[must_use]
    pub fn record_type(self) -> u16 {
        BOUND_SHEET_RECORD
    }

    /// A stable, allocation-free description of the defect class.
    #[must_use]
    pub fn description(self) -> &'static str {
        match self {
            SheetNameDefect::Empty => "BoundSheet8 declares an empty sheet name",
            SheetNameDefect::TooLong => "Sheet name is longer than 31 UTF-16 code units",
            SheetNameDefect::ForbiddenCharacter => "Sheet name contains a forbidden character",
            SheetNameDefect::EdgeApostrophe => "Sheet name begins or ends with an apostrophe",
            SheetNameDefect::Duplicate => "Sheet name repeats an earlier sheet name, ignoring case",
        }
    }
}

/// `BoundSheet8` (MS-XLS 2.4.28).
const BOUND_SHEET_RECORD: u16 = 0x0085;

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
    /// `XF.alcV` (MS-XLS 2.4.353) was 5–7, outside the enumeration.
    ///
    /// Repaired to bottom alignment, the default LibreOffice substitutes.
    AlignmentVertical,
    /// `XF.trot` was 181–254, outside the defined rotations.
    ///
    /// Repaired to no rotation.
    AlignmentRotation,
    /// The reserved bit of `XF.cIndent`'s byte (0x20) was set.
    ///
    /// Ignored, as MS-XLS directs: "MUST be zero, and MUST be ignored".
    AlignmentReservedBit,
    /// `XF.iReadOrder` was 3, outside the enumeration.
    ///
    /// Repaired to context reading order.
    AlignmentReadingOrder,
    /// An `XF` border line style was greater than 13.
    ///
    /// Repaired to no border, the default LibreOffice substitutes.
    BorderStyle,
    /// `XF.fls` was greater than 18.
    ///
    /// Repaired to no fill, the default LibreOffice substitutes.
    FillPattern,
    /// `XF` diagonal direction flags and diagonal style contradicted each
    /// other: a direction without a style, or a style without a direction.
    ///
    /// Kept exactly as stored; such a diagonal draws no line, as in
    /// LibreOffice.
    DiagonalBorder,
    /// An `XF` border or fill color index named no palette entry
    /// (0x42–0x7E).
    ///
    /// Repaired to 0x7F, the automatic color.
    ColorIndex,
    /// `XF.ifmt` named neither a built-in format nor a `Format` record.
    ///
    /// Repaired to General (0), the format Excel renders such a cell with.
    NumberFormatReference,
    /// A `DXF` number-format property named neither a built-in format nor a
    /// `Format` record.
    ///
    /// Repaired to General (0).
    DxfNumberFormatReference,
    /// A `Format` record's `ifmt` (MS-XLS 2.4.126) was outside the ranges a
    /// workbook may define.
    ///
    /// The record is kept, as LibreOffice keeps it, so an `XF` naming the
    /// identifier still resolves to its code.
    FormatIdentifier,
    /// A style `XF` had the `f123Prefix` bit set.
    ///
    /// Ignored: the quote prefix is a cell-XF property.
    StylePrefix,
    /// `XF.ifnt` was the reserved index 4 or above 1022.
    ///
    /// Repaired to font 0, the default font Excel renders the cell with.
    FontReference,
    /// The `XFCRC` record's future-record header type or reserved bytes were
    /// wrong.
    ///
    /// Ignored, as MS-XLS directs ("MUST be zero, and MUST be ignored").
    XfcrcReserved,
    /// `XFCRC.cxfs` was outside `16..=4050`.
    ///
    /// The redundant summary is discarded; the `XF` records themselves are
    /// the data.
    XfcrcCount,
}

impl CellFormatDefect {
    /// The BIFF record type whose payload carried the defect, or `None` when
    /// the carrying record varies per occurrence.
    #[must_use]
    pub fn record_type(self) -> Option<u16> {
        match self {
            CellFormatDefect::XfReference => None,
            CellFormatDefect::AlignmentVertical
            | CellFormatDefect::AlignmentRotation
            | CellFormatDefect::AlignmentReservedBit
            | CellFormatDefect::AlignmentReadingOrder
            | CellFormatDefect::BorderStyle
            | CellFormatDefect::FillPattern
            | CellFormatDefect::DiagonalBorder
            | CellFormatDefect::ColorIndex
            | CellFormatDefect::NumberFormatReference
            | CellFormatDefect::StylePrefix
            | CellFormatDefect::FontReference => Some(super::number_format::XF_RECORD),
            CellFormatDefect::DxfNumberFormatReference => {
                Some(super::differential_format::DXF_RECORD_TYPE)
            },
            CellFormatDefect::FormatIdentifier => Some(super::number_format::FORMAT_RECORD),
            CellFormatDefect::XfcrcReserved | CellFormatDefect::XfcrcCount => {
                Some(super::number_format::XFCRC_RECORD)
            },
        }
    }

    /// A stable, allocation-free description of the defect class.
    #[must_use]
    pub fn description(self) -> &'static str {
        match self {
            CellFormatDefect::XfReference => {
                "Cell, row, or column references an XF that cannot format a cell"
            },
            CellFormatDefect::AlignmentVertical => {
                "XF vertical alignment is outside the enumeration"
            },
            CellFormatDefect::AlignmentRotation => "XF text rotation is outside 0..=180 and 255",
            CellFormatDefect::AlignmentReservedBit => "XF indentation reserved bit is set",
            CellFormatDefect::AlignmentReadingOrder => {
                "XF reading order is outside the enumeration"
            },
            CellFormatDefect::BorderStyle => "XF border style is outside the enumeration",
            CellFormatDefect::FillPattern => "XF fill pattern is outside the enumeration",
            CellFormatDefect::DiagonalBorder => {
                "XF diagonal border direction and style contradict each other"
            },
            CellFormatDefect::ColorIndex => "XF border or fill color index names no palette entry",
            CellFormatDefect::NumberFormatReference => {
                "XF number format identifier resolves to no format"
            },
            CellFormatDefect::DxfNumberFormatReference => {
                "DXF number format identifier resolves to no format"
            },
            CellFormatDefect::FormatIdentifier => {
                "Format record identifier is outside the permitted ranges"
            },
            CellFormatDefect::StylePrefix => "Style XF has the f123Prefix bit set",
            CellFormatDefect::FontReference => "XF font index is reserved or out of range",
            CellFormatDefect::XfcrcReserved => "XFCRC header or reserved fields are nonzero",
            CellFormatDefect::XfcrcCount => "XFCRC count is outside 16..=4050",
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
    /// A `Label` cell's text held an unpaired UTF-16 surrogate.
    ///
    /// Each unpaired surrogate decodes to U+FFFD, exactly as POI and
    /// LibreOffice decode it; every well-formed character survives.
    LabelSurrogate,
    /// A `Font` record's name held an unpaired UTF-16 surrogate.
    ///
    /// Each unpaired surrogate decodes to U+FFFD.
    FontNameSurrogate,
}

impl EncodingDefect {
    /// The BIFF record type whose payload carried the defect.
    #[must_use]
    pub fn record_type(self) -> Option<u16> {
        match self {
            EncodingDefect::UnsupportedCodePage => Some(CODEPAGE_RECORD),
            EncodingDefect::LabelSurrogate => Some(LABEL_RECORD),
            EncodingDefect::FontNameSurrogate => Some(super::font::FONT_RECORD_TYPE),
        }
    }

    /// A stable, allocation-free description of the defect class.
    #[must_use]
    pub fn description(self) -> &'static str {
        match self {
            EncodingDefect::UnsupportedCodePage => {
                "CODEPAGE names a code page without a supported codec"
            },
            EncodingDefect::LabelSurrogate => "Label cell text has an unpaired UTF-16 surrogate",
            EncodingDefect::FontNameSurrogate => "Font name has an unpaired UTF-16 surrogate",
        }
    }
}

/// `Label` (MS-XLS 2.4.148).
const LABEL_RECORD: u16 = 0x0204;

/// The class of shared-string defect the reader repairs the way Excel, POI,
/// and LibreOffice repair it.
///
/// Repaired and recorded under every [`Leniency`] variant: each repair keeps
/// every well-formed string and the cell grid intact instead of discarding
/// the whole shared-string table — and with it the workbook — over one
/// string.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SharedStringDefect {
    /// A shared string's text held an unpaired UTF-16 surrogate.
    ///
    /// Each unpaired surrogate decodes to U+FFFD.
    UnpairedSurrogate,
    /// A shared string's phonetic (`ExtRst`) text held an unpaired UTF-16
    /// surrogate.
    ///
    /// Each unpaired surrogate decodes to U+FFFD.
    PhoneticSurrogate,
    /// A rich-text formatting run started past the string's text or did not
    /// start after the previous run.
    ///
    /// The run is dropped; the surrounding runs and the text are kept.
    FormattingRun,
    /// A phonetic (`ExtRst`) run pointed outside its texts or did not follow
    /// the previous run, or the runs together overran the base string.
    ///
    /// The offending run is dropped — every run, for an overrun — and the
    /// phonetic text is kept.
    PhoneticRun,
}

impl SharedStringDefect {
    /// The BIFF record type whose payload carried the defect: `SST`.
    #[must_use]
    pub fn record_type(self) -> u16 {
        SST_RECORD
    }

    /// A stable, allocation-free description of the defect class.
    #[must_use]
    pub fn description(self) -> &'static str {
        match self {
            SharedStringDefect::UnpairedSurrogate => {
                "Shared string has an unpaired UTF-16 surrogate"
            },
            SharedStringDefect::PhoneticSurrogate => {
                "Shared string phonetic text has an unpaired UTF-16 surrogate"
            },
            SharedStringDefect::FormattingRun => {
                "Shared string formatting run is out of order or past its text"
            },
            SharedStringDefect::PhoneticRun => "Shared string phonetic run is invalid",
        }
    }
}

/// `SST` (MS-XLS 2.4.265).
const SST_RECORD: u16 = 0x00FC;

/// The first unpaired surrogate in `units`, if any: the code unit an
/// [`EncodingDefect`] or [`SharedStringDefect`] surrogate repair reports.
pub(crate) fn first_unpaired_surrogate(units: impl IntoIterator<Item = u16>) -> Option<u16> {
    char::decode_utf16(units).find_map(|unit| unit.err().map(|error| error.unpaired_surrogate()))
}

/// `CODEPAGE` (MS-XLS 2.4.52).
const CODEPAGE_RECORD: u16 = 0x0042;

/// The class of `Palette` defect the reader repairs the way Excel repairs it.
///
/// Repaired and recorded under every [`Leniency`] variant: each repair
/// follows a MS-XLS "MUST be ignored" directive or the reading LibreOffice
/// and POI apply, and none can change a conforming workbook.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum PaletteDefect {
    /// One or more `LongRGB.reserved` bytes (MS-XLS 2.5.177) were nonzero.
    ///
    /// The bytes are ignored, as the specification directs; the colors are
    /// read from the red, green, and blue bytes alone.
    ReservedByte,
    /// `Palette.ccv` (MS-XLS 2.4.188) was not 56, though the payload held
    /// exactly `ccv` entries.
    ///
    /// The first 56 entries are used and any missing entry keeps its default
    /// color, as LibreOffice and POI read a short or long palette.
    ColorCount,
    /// The workbook globals held more than one `Palette` record.
    ///
    /// The last record wins, as in Excel and LibreOffice.
    Duplicate,
}

impl PaletteDefect {
    /// The BIFF record type whose payload carried the defect: `Palette`.
    #[must_use]
    pub fn record_type(self) -> u16 {
        PALETTE_RECORD
    }

    /// A stable, allocation-free description of the defect class.
    #[must_use]
    pub fn description(self) -> &'static str {
        match self {
            PaletteDefect::ReservedByte => "Palette LongRGB reserved byte is nonzero",
            PaletteDefect::ColorCount => "Palette declares a color count other than 56",
            PaletteDefect::Duplicate => "Workbook globals contain more than one Palette record",
        }
    }
}

/// `Palette` (MS-XLS 2.4.188).
const PALETTE_RECORD: u16 = 0x0092;

/// The class of `BoundSheet8` defect the reader repairs the way Excel and
/// LibreOffice repair it.
///
/// Repaired and recorded under every [`Leniency`] variant. Worksheet-name
/// rule violations are a separate, opt-in class: see [`SheetNameDefect`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SheetDefect {
    /// Reserved `stName` option bits (MS-XLS 2.5.294 `ShortXLUnicodeString`)
    /// were set.
    ///
    /// The bits are ignored, as the specification directs and as the
    /// `Format` and `SST` parsers already ignore the identical bits.
    ReservedNameFlags,
    /// `hsState` (MS-XLS 2.4.28) was 3, outside the visibility enumeration.
    ///
    /// Repaired to very hidden — the most restrictive state — so a sheet the
    /// producer meant to hide is never surfaced.
    Visibility,
    /// `dt` (MS-XLS 2.4.28) was outside the sheet-type enumeration.
    ///
    /// Repaired to worksheet, as LibreOffice reads it.
    SheetType,
    /// The record's payload continued past the sheet name.
    ///
    /// The trailing bytes are ignored, as POI and LibreOffice ignore them.
    TrailingBytes,
}

impl SheetDefect {
    /// The BIFF record type whose payload carried the defect: `BoundSheet8`.
    #[must_use]
    pub fn record_type(self) -> u16 {
        BOUND_SHEET_RECORD
    }

    /// A stable, allocation-free description of the defect class.
    #[must_use]
    pub fn description(self) -> &'static str {
        match self {
            SheetDefect::ReservedNameFlags => "BoundSheet8 name has reserved option bits set",
            SheetDefect::Visibility => "BoundSheet8 visibility is outside the enumeration",
            SheetDefect::SheetType => "BoundSheet8 sheet type is outside the enumeration",
            SheetDefect::TrailingBytes => "BoundSheet8 payload continues past the sheet name",
        }
    }
}

/// One class of defect repaired while a workbook was read.
///
/// This enumeration is closed by design: it is the exhaustive contract of what
/// a read may change about a workbook. Every class except the two opt-in
/// ones is an Excel-compatible repair applied under every [`Leniency`]
/// variant; [`Defect::Formatting`] classes require
/// [`Leniency::TolerateFormattingDefects`] and [`Defect::SheetName`] classes
/// [`Leniency::TolerateSheetNameDefects`] (either is implied by
/// [`Leniency::TolerateAllDefects`]).
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
    /// A `Palette` defect repaired the way Excel repairs it.
    Palette(PaletteDefect),
    /// A `BoundSheet8` defect repaired the way Excel repairs it.
    Sheet(SheetDefect),
    /// A shared-string defect repaired the way Excel repairs it.
    SharedString(SharedStringDefect),
    /// A worksheet-name defect repaired by the opt-in [`Leniency`] variant.
    SheetName(SheetNameDefect),
    /// A compound-file (CFB) repair the container reader applied while the
    /// workbook's OLE file was opened; see [`litchi_cfb::Defect`] for the
    /// closed set of container repairs and their substitutions. Applied under
    /// every [`Leniency`] variant.
    Container(litchi_cfb::Defect),
    /// A property-set repair the metadata reader applied while the
    /// workbook's `\u{5}SummaryInformation` or
    /// `\u{5}DocumentSummaryInformation` stream was read by the eager
    /// [`crate::Workbook`] open (the source-backed owner reads neither); see
    /// [`litchi_ole_common::leniency::Defect`] for the closed set of
    /// property-set repairs and their substitutions. Applied under every
    /// [`Leniency`] variant.
    PropertySet(litchi_ole_common::leniency::Defect),
}

impl Defect {
    /// The BIFF record type whose payload carried the defect.
    ///
    /// `None` for [`CellFormatDefect::XfReference`], whose carrying record
    /// varies per occurrence: any cell record, `ROW` (0x0208), or `COLINFO`
    /// (0x007D), and for [`Defect::Container`] and [`Defect::PropertySet`],
    /// which no BIFF record carries.
    #[must_use]
    pub fn record_type(self) -> Option<u16> {
        match self {
            Defect::Formatting(defect) => Some(defect.record_type()),
            Defect::CellFormat(defect) => defect.record_type(),
            Defect::Encoding(defect) => defect.record_type(),
            Defect::Palette(defect) => Some(defect.record_type()),
            Defect::Sheet(defect) => Some(defect.record_type()),
            Defect::SharedString(defect) => Some(defect.record_type()),
            Defect::SheetName(defect) => Some(defect.record_type()),
            Defect::Container(_) | Defect::PropertySet(_) => None,
        }
    }

    /// A stable, allocation-free description of the defect class.
    #[must_use]
    pub fn description(self) -> &'static str {
        match self {
            Defect::Formatting(defect) => defect.description(),
            Defect::CellFormat(defect) => defect.description(),
            Defect::Encoding(defect) => defect.description(),
            Defect::Palette(defect) => defect.description(),
            Defect::Sheet(defect) => defect.description(),
            Defect::SharedString(defect) => defect.description(),
            Defect::SheetName(defect) => defect.description(),
            Defect::Container(defect) => defect.repair(),
            Defect::PropertySet(defect) => defect.repair(),
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

impl From<SheetNameDefect> for Defect {
    fn from(defect: SheetNameDefect) -> Self {
        Self::SheetName(defect)
    }
}

impl From<SharedStringDefect> for Defect {
    fn from(defect: SharedStringDefect) -> Self {
        Self::SharedString(defect)
    }
}

impl From<SheetDefect> for Defect {
    fn from(defect: SheetDefect) -> Self {
        Self::Sheet(defect)
    }
}

impl From<PaletteDefect> for Defect {
    fn from(defect: PaletteDefect) -> Self {
        Self::Palette(defect)
    }
}

impl From<EncodingDefect> for Defect {
    fn from(defect: EncodingDefect) -> Self {
        Self::Encoding(defect)
    }
}

impl From<litchi_cfb::Defect> for Defect {
    fn from(defect: litchi_cfb::Defect) -> Self {
        Self::Container(defect)
    }
}

impl From<litchi_ole_common::leniency::Defect> for Defect {
    fn from(defect: litchi_ole_common::leniency::Defect) -> Self {
        Self::PropertySet(defect)
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
    /// For every `Font*` [`FormattingDefect`] except
    /// [`FormattingDefect::FontCount`] this is the font's logical index; for
    /// [`FormattingDefect::FontCount`] it is the number of fonts read.
    /// For [`FormattingDefect::AlignmentJustifyLastLine`] it is the `XF`
    /// index. For [`FormattingDefect::FormatStringOverrun`] it is the
    /// zero-based ordinal of the `Format` record among the `Format` records
    /// parsed so far. For [`FormattingDefect::ExtendedFormatCountMismatch`]
    /// it is the number of `XF` records that were actually parsed. For
    /// [`CellFormatDefect::XfReference`] it is the zero-based row of the
    /// cell or `ROW` record, or the first column of the `COLINFO` record.
    /// For [`CellFormatDefect::DxfNumberFormatReference`] it is the `DXF`
    /// index, for [`CellFormatDefect::FormatIdentifier`] the `Format`
    /// record's ordinal, and for [`CellFormatDefect::XfcrcReserved`] and
    /// [`CellFormatDefect::XfcrcCount`] the `XFCRC` record's byte offset in
    /// the Workbook stream. For the other [`CellFormatDefect`] classes it is
    /// the `XF` index.
    /// For [`EncodingDefect::UnsupportedCodePage`] it is the byte offset of
    /// the `CODEPAGE` record within the Workbook stream. For
    /// [`EncodingDefect::LabelSurrogate`] it is the cell's zero-based row; for
    /// [`EncodingDefect::FontNameSurrogate`] the font's logical index. For
    /// every [`SharedStringDefect`] it is the zero-based shared-string index.
    /// For every [`SheetNameDefect`] and [`SheetDefect`] it is the zero-based workbook sheet index of the
    /// `BoundSheet8` record. For every [`PaletteDefect`] it is the byte
    /// offset of the `Palette` record within the Workbook stream. For a
    /// [`Defect::Container`] it is the compound-file context (a directory
    /// SID, sector index, or header field value — see
    /// [`litchi_cfb::ToleratedDefect::context`]). For a
    /// [`Defect::PropertySet`] it is the property-set context (a property
    /// identifier, stored field value, or section index — see
    /// [`litchi_ole_common::leniency::ToleratedDefect::context`]).
    #[must_use]
    pub fn ordinal(&self) -> u32 {
        self.ordinal
    }

    /// The offending value read from the file.
    ///
    /// For [`FormattingDefect::FontFamily`] this is the out-of-range
    /// `bFamily` byte; for the other scalar `Font*` classes it is likewise the
    /// raw field value, for [`FormattingDefect::FontNameLength`] the declared
    /// `cch`, for [`FormattingDefect::FontCount`] the number of fonts read,
    /// and zero for [`FormattingDefect::FontNameNul`]. For [`FormattingDefect::AlignmentJustifyLastLine`] it
    /// is the `alcH` value that accompanied the flag. For
    /// [`FormattingDefect::ExtendedFormatCountMismatch`] it is the declared
    /// `cxfs`. For [`FormattingDefect::FormatStringOverrun`] it is the
    /// declared `cch`. For [`FormattingDefect::FontNameEmpty`] it is always
    /// zero, since the defect *is* the value. For
    /// [`CellFormatDefect::XfReference`] it is the `ixfe` the record carried.
    /// For the alignment, border-style, fill-pattern, and color-index
    /// [`CellFormatDefect`] classes it is the raw field value (the whole
    /// indentation byte for [`CellFormatDefect::AlignmentReservedBit`]); for
    /// [`CellFormatDefect::DiagonalBorder`] it is the stored flags and style
    /// packed as `fDgDown | fDgUp << 1 | dgDiag << 2`. For the number-format,
    /// identifier, and font-reference classes it is the raw identifier or
    /// index; for [`CellFormatDefect::StylePrefix`] the XF's flags word; for
    /// [`CellFormatDefect::XfcrcReserved`] the number of nonzero reserved
    /// bytes (zero when only the header type was wrong); for
    /// [`CellFormatDefect::XfcrcCount`] the declared `cxfs`.
    /// For [`EncodingDefect::UnsupportedCodePage`] it is the raw `cv` code
    /// page identifier; for [`EncodingDefect::LabelSurrogate`] the cell's
    /// zero-based column; for [`EncodingDefect::FontNameSurrogate`],
    /// [`SharedStringDefect::UnpairedSurrogate`], and
    /// [`SharedStringDefect::PhoneticSurrogate`] the first unpaired surrogate
    /// code unit. For [`SharedStringDefect::FormattingRun`] it is the dropped
    /// run's first character index; for [`SharedStringDefect::PhoneticRun`]
    /// the zero-based index of the dropped run, or the run count when an
    /// overrun dropped every run. For [`SheetNameDefect::TooLong`] it is the name's
    /// length in UTF-16 code units; for [`SheetNameDefect::ForbiddenCharacter`]
    /// the first forbidden character's scalar value; for
    /// [`SheetNameDefect::EdgeApostrophe`] the apostrophe (0x27); for
    /// [`SheetNameDefect::Duplicate`] the sheet index of the earlier sheet it
    /// repeats; and zero for [`SheetNameDefect::Empty`]. For
    /// [`PaletteDefect::ReservedByte`] it is the number of entries whose
    /// reserved byte was nonzero; for [`PaletteDefect::ColorCount`] the
    /// declared `ccv`; for [`PaletteDefect::Duplicate`] the one-based count
    /// of `Palette` records seen so far. For [`SheetDefect::ReservedNameFlags`]
    /// it is the `stName` option byte, for [`SheetDefect::Visibility`] the
    /// `hsState` byte, for [`SheetDefect::SheetType`] the `dt` byte, and for
    /// [`SheetDefect::TrailingBytes`] the number of bytes ignored. For a
    /// [`Defect::Container`] or [`Defect::PropertySet`] it is always zero;
    /// the substrate's context is the [`Self::ordinal`].
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
/// After a [`Leniency::Strict`] read the report can still list the
/// Excel-compatible repairs Excel itself applies; it never lists the opt-in
/// [`Defect::Formatting`] or [`Defect::SheetName`] repairs.
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

    /// Append every repair `other` holds, in order, after this report's own.
    pub(crate) fn absorb(&mut self, other: ToleranceReport) {
        for entry in other.defects {
            self.record(entry.defect, entry.ordinal, entry.observed);
        }
        self.unrecorded = self.unrecorded.saturating_add(other.unrecorded);
    }

    /// Merge the repairs the compound-file reader applied while opening the
    /// workbook's container, as [`Defect::Container`] entries in discovery
    /// order. The container's own overflow count joins this report's
    /// unrecorded count.
    pub(crate) fn append_container(&mut self, container: &litchi_cfb::ToleranceReport) {
        for tolerated in container.defects() {
            self.record(Defect::Container(tolerated.defect), tolerated.context, 0);
        }
        self.unrecorded = self.unrecorded.saturating_add(container.unrecorded());
    }

    /// Merge the repairs the property-set reader applies to the workbook's
    /// metadata streams, as [`Defect::PropertySet`] entries in discovery
    /// order, each stream's overflow count joining this report's.
    ///
    /// `read` parses one stream by path. An absent or unreadable stream
    /// records nothing and does not fail the open: no workbook content
    /// depends on the metadata, and the property-set accessors still report
    /// the failure when the metadata is asked for.
    pub(crate) fn append_property_sets(
        &mut self,
        mut read: impl FnMut(
            &[&str],
        ) -> std::result::Result<
            litchi_ole_common::property_set::Stream,
            litchi_cfb::OleError,
        >,
    ) {
        for name in PROPERTY_SET_STREAMS {
            let Ok(stream) = read(&[name]) else {
                continue;
            };
            let parsed = stream.tolerance_report();
            for tolerated in parsed.defects() {
                self.record(Defect::PropertySet(tolerated.defect), tolerated.context, 0);
            }
            self.unrecorded = self.unrecorded.saturating_add(parsed.unrecorded());
        }
    }

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

    /// Decide whether the worksheet-name `defect` may be repaired, recording
    /// it if so; the counterpart of [`ToleranceLog::tolerate`] for
    /// [`Leniency::TolerateSheetNameDefects`].
    pub(crate) fn tolerate_sheet_name(
        &mut self,
        defect: SheetNameDefect,
        ordinal: u32,
        observed: u32,
        on_strict: impl FnOnce() -> Error,
    ) -> Result<()> {
        if !self.leniency.tolerates_sheet_name_defects() {
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
    fn each_opt_in_variant_enables_exactly_its_classes() {
        let cases = [
            (Leniency::Strict, false, false),
            (Leniency::TolerateFormattingDefects, true, false),
            (Leniency::TolerateSheetNameDefects, false, true),
            (Leniency::TolerateAllDefects, true, true),
        ];
        for (leniency, formatting, sheet_names) in cases {
            assert_eq!(leniency.tolerates_formatting_defects(), formatting);
            assert_eq!(leniency.tolerates_sheet_name_defects(), sheet_names);
            let mut log = ToleranceLog::new(leniency);
            let outcome = log.tolerate_sheet_name(SheetNameDefect::Duplicate, 2, 0, strict_error);
            assert_eq!(outcome.is_ok(), sheet_names, "{leniency:?}");
            assert_eq!(log.into_report().total(), u64::from(sheet_names));
        }
        assert_eq!(
            Defect::from(SheetNameDefect::TooLong).record_type(),
            Some(0x0085)
        );
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
