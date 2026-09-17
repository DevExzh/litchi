//! Deterministic OOXML corpora in the shape real Office producers emit.
//!
//! Every generated corpus this harness had before change 0601 is *marker
//! free*: the XLSX generator writes worksheets with no Markup Compatibility
//! root, no `x14ac:dyDescent`, no `<cols>`, no shared-string part and no
//! worksheet relationships, and change 0032 recorded that explicitly.  Real
//! Excel, Word and PowerPoint files carry all of these by default, and the
//! 0587 survey measured that the difference is the single largest term in the
//! OOXML read path (`XML-1`: an eager open plus one cell costs 12.9x the
//! marker-free control on one real worksheet).
//!
//! This module adds opt-in corpora that carry the producer signature, so that
//! the candidates ranked `XML-1`, `XML-2`, `XML-3` and `XLSX-2` can be priced
//! on the path real files take.  Nothing here changes an existing selector, an
//! existing corpus identity or the default matrix; every case below is opt-in
//! and absent from [`crate::Case::DEFAULT`].
//!
//! The shapes are built by rewriting a package the production writers
//! themselves produced, so the surrounding package stays exactly as valid as
//! the corpora that already exist; only the parts whose markup is under test
//! are authored here.  Generation is a pure function of the shape (and, for
//! `--real-file`, of the named file's bytes): no clock, no PRNG, no ambient
//! state.

use std::{
    error::Error,
    fs,
    path::Path,
    sync::Arc,
    time::{Duration, Instant},
};

use litchi_core::{OwnedSource, ReadAt};
use serde::Serialize;
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

use crate::{
    Case, CaseResult, Corpus, CorpusManifest, CountingSink, SemanticShape, XlsxCorpus,
    XlsxManifest, iteration_count, record_elapsed, semantic_docx_bytes, semantic_pptx_bytes,
    sha256_hex, statistics, xlsx_source_layout,
};

/// Generator identity of the producer-shape XLSX family.
pub(crate) const XLSX_PRODUCER_SHAPE_GENERATOR: &str = "litchi-xlsx-producer-shape-v1";
/// Generator identity of the producer-shape DOCX family.
pub(crate) const DOCX_PRODUCER_SHAPE_GENERATOR: &str = "litchi-docx-producer-shape-v1";
/// Generator identity of the producer-shape PPTX family.
pub(crate) const PPTX_PRODUCER_SHAPE_GENERATOR: &str = "litchi-pptx-producer-shape-v1";
/// Generator identity used when `--real-file` supplies the bytes.  The family
/// is deliberately distinct: nothing about such a corpus is generated.
pub(crate) const XLSX_REAL_FILE_GENERATOR: &str = "litchi-xlsx-real-file-v1";

/// Largest `--real-file` input this harness will read.  A caller-named file is
/// still a bounded resource; the limit matches the XLSX planning guard's.
const MAX_REAL_FILE_BYTES: u64 = 32 * 1024 * 1024;

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const OPC_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const MCE_NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const X14AC_NAMESPACE: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
const XR_NAMESPACE: &str = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision";
const XR2_NAMESPACE: &str = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2";
const XR3_NAMESPACE: &str = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3";
const XR6_NAMESPACE: &str = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6";
const XR10_NAMESPACE: &str = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10";
const X15_NAMESPACE: &str = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
const A14_NAMESPACE: &str = "http://schemas.microsoft.com/office/drawing/2010/main";
const PRINTER_SETTINGS_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings";
const PRINTER_SETTINGS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings";
const SHARED_STRINGS_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const SHARED_STRINGS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
const XML_DECLARATION: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;

/// `mc:Ignorable` token list Excel writes on a worksheet root by default.
const WORKSHEET_IGNORABLE: &str = "x14ac xr xr2 xr3";
/// `mc:Ignorable` token list Word 2013 writes on a `document.xml` root.
const DOCUMENT_IGNORABLE: &str = "w14 w15 wp14";

/// Share of worksheet cells stored as shared strings on the shared-string
/// sheet, in percent.  40% is the working figure the 0601 brief names for a
/// realistic Excel worksheet.
const SHARED_STRING_CELL_PERCENT: usize = 40;
/// Distinct entries in the generated shared-string table.  Excel de-duplicates,
/// so the table is much smaller than the number of string cells.
const SHARED_STRING_UNIQUE_COUNT: usize = 64;
/// Fixed size of the generated `printerSettings` binary.
const PRINTER_SETTINGS_BYTES: usize = 1_024;
/// Fixed `x14ac:dyDescent` value Excel writes for the default font.
const DY_DESCENT: &str = "0.25";

/// The two producer-shape worksheet sizes.  Both mirror an existing
/// [`crate::XlsxShape`] size so the new selectors can be read next to the
/// marker-free ones they are meant to be compared against.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum XlsxProducerShape {
    /// 4 worksheets of 32 rows x 32 columns, the `XlsxShape::Medium` size.
    Medium,
    /// 3 worksheets of 128 rows x 128 columns: the dense size the XLSX
    /// planning guard already uses (`dense-sparse`), with one worksheet per
    /// producer role.  `XlsxShape::DenseWide`'s 256 x 256 grid was measured
    /// and rejected for this family: with the markup-compatibility codec on
    /// the path it costs seconds per sample, so a 100-sample baseline would
    /// take half an hour per selector for no extra signal.
    Dense,
}

impl XlsxProducerShape {
    pub(crate) const ALL: [Self; 2] = [Self::Medium, Self::Dense];

    pub(crate) const fn name(self) -> &'static str {
        match self {
            Self::Medium => "producer-medium",
            Self::Dense => "producer-dense",
        }
    }

    const fn sheet_count(self) -> usize {
        match self {
            Self::Medium => 4,
            Self::Dense => 3,
        }
    }

    const fn row_count(self) -> usize {
        match self {
            Self::Medium => 32,
            Self::Dense => 128,
        }
    }

    const fn column_count(self) -> usize {
        match self {
            Self::Medium => 32,
            Self::Dense => 128,
        }
    }
}

/// Which producer-shape package a scenario runs on.
///
/// One archive cannot serve both. Excel writes a shared-string part, worksheet
/// relationships and markup-compatibility roots together. The value-only
/// editor now retains and resolves the shared-string part, while it still
/// refuses the relationship-bearing worksheet and qualified root attributes.
/// Splitting the family keeps the read scenarios on the complete producer
/// signature while giving the planning and edit/save scenarios the largest
/// producer-shaped package the library admits; the remaining refusals are
/// proven, untimed, at construction.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum XlsxProducerVariant {
    /// The complete producer signature: markup-compatibility worksheet roots,
    /// `x14ac:dyDescent`, `<cols>`, `pageMargins`, a shared-string part with
    /// 40% string cells, worksheet relationships and an `mc:Ignorable`
    /// workbook root.
    Read,
    /// The markup-compatibility namespace declarations and the `<cols>` block
    /// only: the largest part of the producer signature the value-only editor
    /// admits.  The declarations alone still defeat 0546's fused traversal and
    /// still send the worksheet through the MCE codec's whole-part rewrite,
    /// and `<cols>` alone still makes the selected-cell stream ineligible, so
    /// this variant is producer-shaped where it matters for the ranked
    /// candidates.
    Edit,
    /// The same grid with none of the producer signature: the marker-free
    /// control the existing corpora are, byte-comparable with `Edit` except
    /// for the namespace declarations and the `<cols>` block.  Without them
    /// the worksheet is `source_stream_eligible`, 0546's fused traversal
    /// applies and the selected-cell stream stays eligible, so the control
    /// prices exactly what the producer signature costs.
    Control,
}

impl XlsxProducerVariant {
    pub(crate) const ALL: [Self; 3] = [Self::Read, Self::Edit, Self::Control];

    const fn suffix(self) -> &'static str {
        match self {
            Self::Read => "read",
            Self::Edit => "edit",
            Self::Control => "control",
        }
    }
}

/// Which producer facts one generated archive carries.
///
/// The producer facts are split into switches so the value-only editor's
/// dependency and preservation verdict can be observed one at a time. That is
/// what makes the census exact instead of narrative.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct ArchiveOptions {
    /// `xmlns:mc`, `xmlns:x14ac`, `xmlns:xr`, `xmlns:xr2` and `xmlns:xr3` on
    /// the worksheet root.  Declarations alone already defeat
    /// `source_stream_eligible` and already send the part through the MCE
    /// codec's whole-part rewrite, because both gates key on the namespace
    /// string occurring anywhere in the bytes.
    namespace_declarations: bool,
    /// `mc:Ignorable` and `xr:uid` on the worksheet root, and
    /// `x14ac:dyDescent` on `sheetFormatPr` and on every row.
    markup_compatibility_attributes: bool,
    /// The `<cols>` block, which precedes `sheetData` and therefore makes the
    /// selected-cell stream ineligible on its own.
    cols: bool,
    /// `pageMargins`, and `pageSetup r:id` on the relationship role.
    page_setup: bool,
    /// `xl/sharedStrings.xml`, its workbook relationship, and 40% `t="s"`
    /// cells on the shared-string worksheet.
    shared_strings: bool,
    /// `xl/worksheets/_rels/sheetN.xml.rels` and its `printerSettings` target.
    worksheet_relationships: bool,
    /// `mc:Ignorable` and the x15/xr namespace set on the workbook root.
    workbook_root_markers: bool,
}

impl ArchiveOptions {
    /// The complete signature Excel writes.
    const ALL: Self = Self {
        namespace_declarations: true,
        markup_compatibility_attributes: true,
        cols: true,
        page_setup: true,
        shared_strings: true,
        worksheet_relationships: true,
        workbook_root_markers: true,
    };
    /// No producer facts at all: the marker-free control.
    const NONE: Self = Self {
        namespace_declarations: false,
        markup_compatibility_attributes: false,
        cols: false,
        page_setup: false,
        shared_strings: false,
        worksheet_relationships: false,
        workbook_root_markers: false,
    };
    /// The largest subset of that signature the value-only editor admits.
    /// Shared strings are retained even though the planning worksheet uses
    /// numeric cells; the selected shared-string worksheet proves its read
    /// path separately.
    const ADMITTED: Self = Self {
        namespace_declarations: true,
        markup_compatibility_attributes: false,
        cols: true,
        page_setup: false,
        shared_strings: true,
        worksheet_relationships: false,
        workbook_root_markers: false,
    };

    const fn of(variant: XlsxProducerVariant) -> Self {
        match variant {
            XlsxProducerVariant::Read => Self::ALL,
            XlsxProducerVariant::Edit => Self::ADMITTED,
            XlsxProducerVariant::Control => Self::NONE,
        }
    }
}

/// What a worksheet of the producer shape carries beyond the plain grid.
///
/// The roles are deliberately separated so one corpus can prove each library
/// gate independently: the value-only editor admits a shared-string worksheet
/// but refuses a relationship-bearing worksheet before it parses, so a sheet
/// carrying both can only witness the latter.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
#[serde(rename_all = "kebab-case")]
enum SheetRole {
    /// Markup-compatibility root, `x14ac:dyDescent` rows and a `<cols>` block;
    /// numeric cells only and no worksheet relationships.
    Markup,
    /// `Markup`, plus 40% of the cells stored as shared strings.
    SharedStrings,
    /// `Markup`, plus a `printerSettings` worksheet relationship.
    Relationship,
}

impl SheetRole {
    const fn of(index: usize, sheet_count: usize, options: ArchiveOptions) -> Self {
        match index {
            SELECTED_SHEET if options.shared_strings => Self::SharedStrings,
            RELATIONSHIP_SHEET if options.worksheet_relationships && sheet_count > 2 => {
                Self::Relationship
            },
            _ => Self::Markup,
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::Markup => "producer-markup",
            Self::SharedStrings => "producer-markup-shared-strings",
            Self::Relationship => "producer-markup-worksheet-relationship",
        }
    }
}

/// Worksheet index the planning and edit/save selectors act on: the one role
/// the value-only editor admits.
const PLANNING_SHEET: usize = 0;
/// Worksheet index the selected-cell selector reads: the shared-string role.
const SELECTED_SHEET: usize = 1;
/// Worksheet index that carries the `printerSettings` relationship.
const RELATIONSHIP_SHEET: usize = 2;

// ---------------------------------------------------------------------------
// Eligibility oracle
// ---------------------------------------------------------------------------

/// Mirror of `litchi_xlsx::raw::worksheet::source_stream_eligible`
/// (`crates/litchi-xlsx/src/raw/worksheet/mod.rs:60`), which is crate private.
///
/// This is an oracle, not a re-implementation the library consumes: it exists
/// so a corpus can *state* which gate its parts trip and so a future change to
/// the gate is visible as a disagreement between this census and the measured
/// behaviour.  The constants are copied from the same file.
const MAX_SHARED_SOURCE_BYTES: usize = 8 * 1024 * 1024;
const MCE_MAX_INPUT_BYTES: usize = 256 * 1024 * 1024;
const MCE_MAX_OUTPUT_BYTES: usize = 512 * 1024 * 1024;

fn source_stream_ineligibility(content: &[u8]) -> Vec<&'static str> {
    let mut reasons = Vec::new();
    if content.len() > MAX_SHARED_SOURCE_BYTES {
        reasons.push("exceeds-shared-source-byte-limit");
    }
    if content.len() > MCE_MAX_INPUT_BYTES || content.len() > MCE_MAX_OUTPUT_BYTES {
        reasons.push("exceeds-mce-limits");
    }
    if std::str::from_utf8(content).is_err() {
        reasons.push("not-utf8");
    }
    if contains(content, MCE_NAMESPACE.as_bytes()) {
        reasons.push("markup-compatibility-namespace");
    }
    if contains(content, b"AlternateContent") {
        reasons.push("alternate-content");
    }
    if contains(content, X14AC_NAMESPACE.as_bytes()) {
        reasons.push("x14ac-namespace");
    }
    if contains(content, b"dyDescent") {
        reasons.push("dy-descent");
    }
    reasons
}

fn contains(content: &[u8], marker: &[u8]) -> bool {
    memchr_find(content, marker).is_some()
}

fn memchr_find(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    if needle.is_empty() || needle.len() > haystack.len() {
        return None;
    }
    haystack
        .windows(needle.len())
        .position(|window| window == needle)
}

/// The producer-signature census of one part, and the library gates it trips.
#[derive(Clone, Debug, Serialize)]
pub(crate) struct PartCensus {
    pub(crate) part: String,
    pub(crate) role: Option<&'static str>,
    pub(crate) bytes: usize,
    pub(crate) sha256: String,
    pub(crate) markup_compatibility_namespace: bool,
    pub(crate) mc_ignorable: Option<String>,
    pub(crate) x14ac_namespace: bool,
    pub(crate) dy_descent_occurrences: usize,
    pub(crate) alternate_content_occurrences: usize,
    pub(crate) cols_block: bool,
    pub(crate) shared_string_cells: usize,
    pub(crate) numeric_cells: usize,
    pub(crate) worksheet_relationship_part: Option<String>,
    pub(crate) worksheet_relationship_targets: Vec<String>,
    /// `false` whenever the library would take the four-pass fallback rather
    /// than 0546's fused validate-and-parse traversal.
    pub(crate) source_stream_eligible: bool,
    pub(crate) source_stream_ineligible_reasons: Vec<&'static str>,
    /// `<cols>` precedes `sheetData`, so its presence alone makes the
    /// selected-cell stream ineligible (`NotEligible(Styles)`).
    pub(crate) selected_cell_stream_ineligible_on_cols: bool,
}

fn census_of(
    part: &str,
    role: Option<&'static str>,
    bytes: &[u8],
    relationship_part: Option<String>,
    relationship_targets: Vec<String>,
) -> PartCensus {
    let reasons = source_stream_ineligibility(bytes);
    PartCensus {
        part: part.to_owned(),
        role,
        bytes: bytes.len(),
        sha256: sha256_hex(bytes),
        markup_compatibility_namespace: contains(bytes, MCE_NAMESPACE.as_bytes()),
        mc_ignorable: extract_attribute(bytes, "mc:Ignorable=\""),
        x14ac_namespace: contains(bytes, X14AC_NAMESPACE.as_bytes()),
        dy_descent_occurrences: count_occurrences(bytes, b"dyDescent"),
        alternate_content_occurrences: count_occurrences(bytes, b"<mc:AlternateContent"),
        cols_block: contains(bytes, b"<cols>"),
        shared_string_cells: count_occurrences(bytes, br#" t="s""#),
        numeric_cells: count_occurrences(bytes, b"<c r=\"")
            .saturating_sub(count_occurrences(bytes, br#" t="s""#)),
        worksheet_relationship_part: relationship_part,
        worksheet_relationship_targets: relationship_targets,
        source_stream_eligible: reasons.is_empty(),
        source_stream_ineligible_reasons: reasons,
        selected_cell_stream_ineligible_on_cols: contains(bytes, b"<cols>"),
    }
}

fn count_occurrences(haystack: &[u8], needle: &[u8]) -> usize {
    if needle.is_empty() {
        return 0;
    }
    let mut count = 0;
    let mut offset = 0;
    while let Some(position) = memchr_find(&haystack[offset..], needle) {
        count += 1;
        offset += position + 1;
        if offset >= haystack.len() {
            break;
        }
    }
    count
}

fn extract_attribute(bytes: &[u8], prefix: &str) -> Option<String> {
    let start = memchr_find(bytes, prefix.as_bytes())? + prefix.len();
    let rest = bytes.get(start..)?;
    let end = memchr_find(rest, b"\"")?;
    std::str::from_utf8(rest.get(..end)?)
        .ok()
        .map(str::to_owned)
}

/// Everything a producer-shape corpus states about itself beyond the ordinary
/// [`CorpusManifest`]: the marker census per selected part, the library gates
/// each part trips, and the typed refusals the corpus proves at build time.
#[derive(Clone, Debug, Serialize)]
pub(crate) struct ProducerEvidence {
    pub(crate) generator: &'static str,
    pub(crate) shape: &'static str,
    pub(crate) shape_size: Option<&'static str>,
    pub(crate) variant: Option<&'static str>,
    pub(crate) sheet_roles: Vec<&'static str>,
    /// `uniqueCount` as the shared-string part declares it, read back from the
    /// archive rather than asserted by the generator, so a real file reports
    /// its own table size.
    pub(crate) shared_string_unique_count: Option<usize>,
    /// Share of the selected worksheet's cells stored as shared strings,
    /// derived from the census of that part.
    pub(crate) shared_string_cell_percent: Option<usize>,
    pub(crate) shared_string_part: Option<String>,
    pub(crate) parts: Vec<PartCensus>,
    pub(crate) planning_sheet: usize,
    pub(crate) selected_sheet: usize,
    pub(crate) selected_target: String,
    pub(crate) selected_target_text: String,
    /// Typed refusals the corpus proves once, untimed, at construction.
    pub(crate) proven_refusals: Vec<ProvenRefusal>,
    pub(crate) real_file: Option<RealFileProvenance>,
}

#[derive(Clone, Debug, Serialize)]
pub(crate) struct ProvenRefusal {
    pub(crate) scenario: &'static str,
    pub(crate) sheet: usize,
    pub(crate) role: &'static str,
    pub(crate) message: String,
}

#[derive(Clone, Debug, Serialize)]
pub(crate) struct RealFileProvenance {
    pub(crate) path: String,
    pub(crate) bytes: u64,
    pub(crate) sha256: String,
}

/// A producer-shape corpus: an ordinary [`Corpus`] plus its evidence block.
#[derive(Debug)]
pub(crate) struct ProducerCorpus {
    pub(crate) corpus: Corpus,
    pub(crate) evidence: ProducerEvidence,
}

// ---------------------------------------------------------------------------
// XLSX generation
// ---------------------------------------------------------------------------

fn column_label(column: usize) -> String {
    let mut value = column + 1;
    let mut label = Vec::new();
    while value != 0 {
        let remainder = (value - 1) % 26;
        label.push(b'A' + u8::try_from(remainder).expect("column remainder is below 26"));
        value = (value - 1) / 26;
    }
    label.reverse();
    String::from_utf8(label).expect("column label is ASCII")
}

fn cell_address(row: usize, column: usize) -> String {
    format!("{}{}", column_label(column), row + 1)
}

/// Deterministic numeric payload of one cell.
const fn numeric_value(sheet: usize, row: usize, column: usize) -> usize {
    sheet * 1_000_003 + row * 1_009 + column
}

/// Deterministic shared-string membership: two cells in every five, i.e. 40%.
const fn is_shared_string(ordinal: usize) -> bool {
    ordinal % 5 < (SHARED_STRING_CELL_PERCENT / 20)
}

fn shared_string_text(index: usize) -> String {
    format!("litchi-perf-producer-shared-{index:04}")
}

fn shared_strings_part(unique_count: usize, total_references: usize) -> String {
    let mut xml = String::new();
    xml.push_str(XML_DECLARATION);
    xml.push_str(&format!(
        r#"<sst xmlns="{SML}" count="{total_references}" uniqueCount="{unique_count}">"#
    ));
    for index in 0..unique_count {
        xml.push_str(&format!(r#"<si><t>{}</t></si>"#, shared_string_text(index)));
    }
    xml.push_str("</sst>");
    xml
}

/// Author one worksheet in the shape Excel writes by default.
///
/// The root carries `mc:Ignorable="x14ac xr xr2 xr3"` with the matching
/// namespace declarations, every `<row>` carries `x14ac:dyDescent`, a `<cols>`
/// block precedes `sheetData`, and the shared-string role stores 40% of its
/// cells as `t="s"` references into `xl/sharedStrings.xml`.
fn worksheet_part(
    sheet: usize,
    rows: usize,
    columns: usize,
    role: SheetRole,
    options: ArchiveOptions,
) -> Result<(String, usize), Box<dyn Error>> {
    let last = cell_address(rows - 1, columns - 1);
    let mut xml = String::new();
    xml.try_reserve(4_096 + rows * columns * 40)
        .map_err(|source| format!("producer worksheet reservation failed: {source}"))?;
    xml.push_str(XML_DECLARATION);
    xml.push_str(&format!(r#"<worksheet xmlns="{SML}" xmlns:r="{REL}""#));
    if options.namespace_declarations {
        xml.push_str(&format!(
            concat!(
                r#" xmlns:mc="{mce}" xmlns:x14ac="{x14ac}" xmlns:xr="{xr}""#,
                r#" xmlns:xr2="{xr2}" xmlns:xr3="{xr3}""#
            ),
            mce = MCE_NAMESPACE,
            x14ac = X14AC_NAMESPACE,
            xr = XR_NAMESPACE,
            xr2 = XR2_NAMESPACE,
            xr3 = XR3_NAMESPACE,
        ));
    }
    if options.markup_compatibility_attributes {
        xml.push_str(&format!(
            r#" mc:Ignorable="{WORKSHEET_IGNORABLE}" xr:uid="{{00000000-0001-0000-{:04X}-000000000000}}""#,
            sheet + 1
        ));
    }
    xml.push('>');
    xml.push_str(&format!(r#"<dimension ref="A1:{last}"/>"#));
    xml.push_str(
        r#"<sheetViews><sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>"#,
    );
    if options.markup_compatibility_attributes {
        xml.push_str(&format!(
            r#"<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="{DY_DESCENT}"/>"#
        ));
    } else {
        xml.push_str(r#"<sheetFormatPr defaultRowHeight="15"/>"#);
    }
    if options.cols {
        xml.push_str(&format!(
            r#"<cols><col min="1" max="{columns}" width="11.5703125" bestFit="1" customWidth="1"/></cols>"#
        ));
    }
    xml.push_str("<sheetData>");
    let mut shared_references = 0;
    for row in 0..rows {
        if options.markup_compatibility_attributes {
            xml.push_str(&format!(
                r#"<row r="{}" spans="1:{columns}" x14ac:dyDescent="{DY_DESCENT}">"#,
                row + 1
            ));
        } else {
            xml.push_str(&format!(r#"<row r="{}" spans="1:{columns}">"#, row + 1));
        }
        for column in 0..columns {
            let address = cell_address(row, column);
            let ordinal = row
                .checked_mul(columns)
                .and_then(|value| value.checked_add(column))
                .ok_or("producer worksheet cell ordinal overflows usize")?;
            if role == SheetRole::SharedStrings && is_shared_string(ordinal) {
                let index = ordinal % SHARED_STRING_UNIQUE_COUNT;
                shared_references += 1;
                xml.push_str(&format!(r#"<c r="{address}" t="s"><v>{index}</v></c>"#));
            } else {
                xml.push_str(&format!(
                    r#"<c r="{address}"><v>{}</v></c>"#,
                    numeric_value(sheet, row, column)
                ));
            }
        }
        xml.push_str("</row>");
    }
    xml.push_str("</sheetData>");
    if options.page_setup {
        xml.push_str(
            r#"<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>"#,
        );
        if role == SheetRole::Relationship {
            xml.push_str(r#"<pageSetup paperSize="9" orientation="portrait" r:id="rIdPrinter"/>"#);
        }
    }
    xml.push_str("</worksheet>");
    Ok((xml, shared_references))
}

fn printer_settings_bytes(sheet: usize) -> Vec<u8> {
    // A fixed, inert binary: the part exists to give the worksheet a
    // relationship, which is the fact under test.  It is never interpreted.
    (0..PRINTER_SETTINGS_BYTES)
        .map(|index| u8::try_from((index + sheet * 7) % 251).expect("modulus is below 256"))
        .collect()
}

fn worksheet_relationships_part() -> String {
    format!(
        concat!(
            r#"{declaration}<Relationships xmlns="{opc}">"#,
            r#"<Relationship Id="rIdPrinter" Type="{printer}" "#,
            r#"Target="../printerSettings/printerSettings1.bin"/></Relationships>"#
        ),
        declaration = XML_DECLARATION,
        opc = OPC_REL,
        printer = PRINTER_SETTINGS_RELATIONSHIP,
    )
}

/// Rewrite a workbook root so it carries the markup-compatibility declarations
/// Excel writes.  The `<sheets>` catalog the library authored is preserved
/// verbatim: only the root start tag changes.
///
/// This is **not** applied to the measured archive.  The value-only editor
/// refuses `mc:Ignorable` on `<workbook>` outright
/// (`value-only edits refuse attribute 'mc:Ignorable' on 'workbook'`), so a
/// workbook root in the shape Excel writes would make the planning and
/// edit/save selectors measure a refusal rather than the work.  The rewrite is
/// kept and applied once, untimed, to a derived variant so the corpus can
/// *prove* that refusal instead of asserting it; the measured archive keeps
/// the plain root the production writer authored.
fn producer_workbook_root(xml: &str) -> Result<String, Box<dyn Error>> {
    let open = xml
        .find("<workbook")
        .ok_or("producer workbook rewrite found no workbook root")?;
    let close = xml[open..]
        .find('>')
        .ok_or("producer workbook root start tag is unterminated")?
        + open;
    let original = &xml[open..=close];
    if original.contains("mc:Ignorable") {
        return Err("producer workbook root already declares mc:Ignorable".into());
    }
    let injected = format!(
        concat!(
            r#" xmlns:mc="{mce}" mc:Ignorable="x15 xr xr6 xr10 xr2" "#,
            r#"xmlns:x15="{x15}" xmlns:xr="{xr}" xmlns:xr2="{xr2}" "#,
            r#"xmlns:xr6="{xr6}" xmlns:xr10="{xr10}"{close}"#
        ),
        mce = MCE_NAMESPACE,
        x15 = X15_NAMESPACE,
        xr = XR_NAMESPACE,
        xr2 = XR2_NAMESPACE,
        xr6 = XR6_NAMESPACE,
        xr10 = XR10_NAMESPACE,
        close = if original.ends_with("/>") { "/>" } else { ">" }
    );
    let trimmed = original
        .strip_suffix("/>")
        .or_else(|| original.strip_suffix('>'))
        .ok_or("producer workbook root start tag is malformed")?;
    Ok(format!(
        "{}{}{}{}",
        &xml[..open],
        trimmed,
        injected,
        &xml[close + 1..]
    ))
}

fn patch_content_types(xml: &str, options: ArchiveOptions) -> Result<String, Box<dyn Error>> {
    let mut additions = String::new();
    if options.worksheet_relationships && !xml.contains(r#"Extension="bin""#) {
        additions.push_str(&format!(
            r#"<Default Extension="bin" ContentType="{PRINTER_SETTINGS_CONTENT_TYPE}"/>"#
        ));
    }
    if options.shared_strings && !xml.contains("/xl/sharedStrings.xml") {
        additions.push_str(&format!(
            r#"<Override PartName="/xl/sharedStrings.xml" ContentType="{SHARED_STRINGS_CONTENT_TYPE}"/>"#
        ));
    }
    if additions.is_empty() {
        return Ok(xml.to_owned());
    }
    let patched = xml.replacen("</Types>", &format!("{additions}</Types>"), 1);
    if patched == xml {
        return Err("producer content-type closure has no </Types>".into());
    }
    Ok(patched)
}

fn patch_workbook_relationships(xml: &str) -> Result<String, Box<dyn Error>> {
    if xml.contains("sharedStrings.xml") {
        return Ok(xml.to_owned());
    }
    let relationship = format!(
        r#"<Relationship Id="rIdProducerSharedStrings" Type="{SHARED_STRINGS_RELATIONSHIP}" Target="sharedStrings.xml"/>"#
    );
    let patched = xml.replacen(
        "</Relationships>",
        &format!("{relationship}</Relationships>"),
        1,
    );
    if patched == xml {
        return Err("producer workbook relationships have no </Relationships>".into());
    }
    Ok(patched)
}

/// Build one producer-shape XLSX corpus.
fn xlsx_producer_archive(
    shape: XlsxProducerShape,
    options: ArchiveOptions,
) -> Result<(Vec<u8>, usize), Box<dyn Error>> {
    let sheet_count = shape.sheet_count();
    let rows = shape.row_count();
    let columns = shape.column_count();

    // Start from a package the production writer produced, so everything this
    // module does not author is exactly as valid as the existing corpora.
    let spec = XlsxCorpus {
        sheet_count,
        row_count: 1,
        column_count: 1,
        one_percent_updates: Vec::new(),
        cell_inventory: None,
    };
    let skeleton = crate::build_xlsx_workbook(&spec)?.to_bytes()?;

    let mut worksheets = Vec::with_capacity(sheet_count);
    let mut shared_references = 0;
    for sheet in 0..sheet_count {
        let role = SheetRole::of(sheet, sheet_count, options);
        let (xml, references) = worksheet_part(sheet, rows, columns, role, options)?;
        shared_references += references;
        worksheets.push(xml);
    }
    if options.shared_strings && shared_references == 0 {
        return Err("producer XLSX shape has no shared-string references".into());
    }

    let reader = ArchiveReader::new(&skeleton)?;
    let names = reader.file_names().map(str::to_owned).collect::<Vec<_>>();
    let mut writer = StreamingArchiveWriter::new();
    let mut rewritten_sheets = 0;
    for name in &names {
        let bytes = reader.read(name.as_str())?;
        match name.as_str() {
            "[Content_Types].xml" => {
                let xml = std::str::from_utf8(&bytes)?;
                writer.write_deflated(name, patch_content_types(xml, options)?.as_bytes())?;
            },
            "xl/workbook.xml" if options.workbook_root_markers => {
                let xml = std::str::from_utf8(&bytes)?;
                writer.write_deflated(name, producer_workbook_root(xml)?.as_bytes())?;
            },
            "xl/_rels/workbook.xml.rels" if options.shared_strings => {
                let xml = std::str::from_utf8(&bytes)?;
                writer.write_deflated(name, patch_workbook_relationships(xml)?.as_bytes())?;
            },
            _ if name.starts_with("xl/worksheets/sheet") && name.ends_with(".xml") => {
                let index = name
                    .strip_prefix("xl/worksheets/sheet")
                    .and_then(|suffix| suffix.strip_suffix(".xml"))
                    .and_then(|index| index.parse::<usize>().ok())
                    .and_then(|index| index.checked_sub(1))
                    .ok_or("producer worksheet member name is not sheetN.xml")?;
                let xml = worksheets
                    .get(index)
                    .ok_or("producer worksheet index is outside the shape")?;
                writer.write_deflated(name, xml.as_bytes())?;
                rewritten_sheets += 1;
            },
            _ => writer.write_deflated(name, &bytes)?,
        }
    }
    if rewritten_sheets != sheet_count {
        return Err("producer XLSX rewrite did not replace every worksheet".into());
    }
    if options.shared_strings {
        let shared_strings = shared_strings_part(SHARED_STRING_UNIQUE_COUNT, shared_references);
        writer.write_deflated("xl/sharedStrings.xml", shared_strings.as_bytes())?;
    }
    if options.worksheet_relationships {
        let relationship_sheet = (0..sheet_count)
            .find(|&index| SheetRole::of(index, sheet_count, options) == SheetRole::Relationship)
            .ok_or("producer XLSX shape has no relationship worksheet")?;
        writer.write_deflated(
            format!(
                "xl/worksheets/_rels/sheet{}.xml.rels",
                relationship_sheet + 1
            )
            .as_str(),
            worksheet_relationships_part().as_bytes(),
        )?;
        writer.write_deflated(
            "xl/printerSettings/printerSettings1.bin",
            &printer_settings_bytes(relationship_sheet),
        )?;
    }
    Ok((writer.finish_to_bytes()?, shared_references))
}

/// Build one producer-shape XLSX corpus.
pub(crate) fn build_xlsx_producer_corpus(
    shape: XlsxProducerShape,
    variant: XlsxProducerVariant,
) -> Result<ProducerCorpus, Box<dyn Error>> {
    let options = ArchiveOptions::of(variant);
    let (archive, _shared_references) = xlsx_producer_archive(shape, options)?;
    finish_xlsx_producer_corpus(
        archive,
        XLSX_PRODUCER_SHAPE_GENERATOR,
        Some((shape, variant)),
        None,
    )
}

/// Build a corpus from a caller-named real Office file.
///
/// This is the `--real-file` opt-in.  It is the only corpus in this harness
/// whose bytes are not produced in process, so the provenance block records
/// the path, the size and the SHA-256, and the case is absent from the default
/// matrix.  Nothing about the file is assumed: the marker census, the target
/// cell and the expected value are all derived from the bytes.
pub(crate) fn build_xlsx_real_file_corpus(path: &Path) -> Result<ProducerCorpus, Box<dyn Error>> {
    let metadata = fs::metadata(path)
        .map_err(|source| format!("--real-file {} is unreadable: {source}", path.display()))?;
    if !metadata.is_file() {
        return Err(format!("--real-file {} is not a regular file", path.display()).into());
    }
    if metadata.len() > MAX_REAL_FILE_BYTES {
        return Err(format!(
            "--real-file {} is {} bytes, above the {MAX_REAL_FILE_BYTES}-byte bound",
            path.display(),
            metadata.len()
        )
        .into());
    }
    let archive = fs::read(path)
        .map_err(|source| format!("--real-file {} could not be read: {source}", path.display()))?;
    let provenance = RealFileProvenance {
        path: path.display().to_string(),
        bytes: metadata.len(),
        sha256: sha256_hex(&archive),
    };
    let workbook = litchi_xlsx::SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(
        archive.clone(),
    )))?;
    let sheet_count = workbook.len();
    if sheet_count == 0 {
        return Err("--real-file workbook declares no worksheet".into());
    }
    drop(workbook);
    let _ = sheet_count;
    finish_xlsx_producer_corpus(archive, XLSX_REAL_FILE_GENERATOR, None, Some(provenance))
}

/// Shared tail of both XLSX corpus builders: reopen the archive through the
/// library, derive the census and the scenario target, and prove the typed
/// refusals or the admission the variant is meant to witness.  Everything
/// here is untimed and happens once per corpus.
fn finish_xlsx_producer_corpus(
    archive: Vec<u8>,
    generator: &'static str,
    generated: Option<(XlsxProducerShape, XlsxProducerVariant)>,
    real_file: Option<RealFileProvenance>,
) -> Result<ProducerCorpus, Box<dyn Error>> {
    let shape_label = match generated {
        Some((shape, variant)) => match (shape, variant) {
            (XlsxProducerShape::Medium, XlsxProducerVariant::Read) => "producer-medium-read",
            (XlsxProducerShape::Medium, XlsxProducerVariant::Edit) => "producer-medium-edit",
            (XlsxProducerShape::Medium, XlsxProducerVariant::Control) => "producer-medium-control",
            (XlsxProducerShape::Dense, XlsxProducerVariant::Read) => "producer-dense-read",
            (XlsxProducerShape::Dense, XlsxProducerVariant::Edit) => "producer-dense-edit",
            (XlsxProducerShape::Dense, XlsxProducerVariant::Control) => "producer-dense-control",
        },
        None => "real-file",
    };
    let options = generated.map(|(_shape, variant)| ArchiveOptions::of(variant));
    let reader = ArchiveReader::new(&archive)?;
    let member_names = reader.file_names().map(str::to_owned).collect::<Vec<_>>();
    let archive_member_count = member_names.len();

    let workbook = litchi_xlsx::SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(
        archive.clone(),
    )))?;
    let sheet_count = generated.map_or_else(|| workbook.len(), |(shape, _)| shape.sheet_count());
    if workbook.len() != sheet_count {
        return Err("producer XLSX worksheet count differs from the shape".into());
    }
    let (rows, columns) = generated.map_or((0, 0), |(shape, _)| {
        (shape.row_count(), shape.column_count())
    });

    let mut parts = Vec::new();
    let mut sheet_roles = Vec::new();
    for sheet in 0..sheet_count {
        let member = format!("xl/worksheets/sheet{}.xml", sheet + 1);
        if !member_names.iter().any(|name| name == &member) {
            continue;
        }
        let role = options.map(|options| SheetRole::of(sheet, sheet_count, options).name());
        if let Some(role) = role {
            sheet_roles.push(role);
        }
        let relationship_member = format!("xl/worksheets/_rels/sheet{}.xml.rels", sheet + 1);
        let (relationship_part, relationship_targets) =
            if member_names.iter().any(|name| name == &relationship_member) {
                let bytes = reader.read(relationship_member.as_str())?;
                let xml = String::from_utf8_lossy(&bytes).into_owned();
                let targets = xml
                    .split("Target=\"")
                    .skip(1)
                    .filter_map(|rest| rest.split('"').next().map(str::to_owned))
                    .collect();
                (Some(relationship_member), targets)
            } else {
                (None, Vec::new())
            };
        let bytes = reader.read(member.as_str())?;
        parts.push(census_of(
            member.as_str(),
            role,
            &bytes,
            relationship_part,
            relationship_targets,
        ));
    }
    let shared_string_part = member_names
        .iter()
        .find(|name| name.as_str() == "xl/sharedStrings.xml")
        .cloned();
    let mut shared_string_unique_count = None;
    if let Some(member) = shared_string_part.as_ref() {
        let bytes = reader.read(member.as_str())?;
        shared_string_unique_count =
            extract_attribute(&bytes, "uniqueCount=\"").and_then(|value| value.parse().ok());
        parts.push(census_of(member.as_str(), None, &bytes, None, Vec::new()));
    }
    let workbook_member = reader.read("xl/workbook.xml")?;
    parts.push(census_of(
        "xl/workbook.xml",
        None,
        &workbook_member,
        None,
        Vec::new(),
    ));

    // Derive the selected-cell target from the package itself, so the same
    // code serves generated and real files.
    let selected_sheet = SELECTED_SHEET.min(sheet_count.saturating_sub(1));
    let sheet_names = workbook
        .sheets()
        .map(|sheet| sheet.name().to_owned())
        .collect::<Vec<_>>();
    let selected_name = sheet_names
        .get(selected_sheet)
        .ok_or("producer XLSX corpus has no selected worksheet")?
        .clone();
    let (selected_target, selected_target_text) = {
        let sheet = workbook
            .sheet(selected_name.as_str())?
            .ok_or("producer XLSX selected worksheet is missing")?;
        let selected_role =
            options.map(|options| SheetRole::of(selected_sheet, sheet_count, options));
        let candidate = scenario_target_address(&sheet, rows, columns, selected_role)?;
        let view = sheet.cell(candidate.as_str())?;
        let text = describe_cell(view.stored());
        (format!("{selected_name}!{candidate}"), text)
    };
    drop(workbook);

    let proven_refusals = match generated {
        None => Vec::new(),
        Some((shape, XlsxProducerVariant::Edit | XlsxProducerVariant::Control)) => {
            prove_edit_variant_is_admitted(&archive)?;
            let _ = shape;
            Vec::new()
        },
        Some((shape, XlsxProducerVariant::Read)) => prove_value_editor_refusals(shape, &archive)?,
    };

    let selected_member = format!("xl/worksheets/sheet{}.xml", selected_sheet + 1);
    let shared_string_cell_percent = parts
        .iter()
        .find(|part| part.part == selected_member)
        .and_then(|part| {
            let total = part.shared_string_cells.checked_add(part.numeric_cells)?;
            (total != 0).then(|| part.shared_string_cells.saturating_mul(100) / total)
        });
    let worksheet_bytes = parts
        .iter()
        .filter(|part| part.part.starts_with("xl/worksheets/"))
        .try_fold(0usize, |total, part| total.checked_add(part.bytes))
        .ok_or("producer XLSX worksheet byte total overflows usize")?;
    let target_payload = selected_target_text.clone().into_bytes();
    let cell_count = sheet_count
        .checked_mul(rows)
        .and_then(|value| value.checked_mul(columns))
        .ok_or("producer XLSX cell count overflows usize")?;
    let (_ranges, source_members) = xlsx_source_layout(&archive, sheet_count)?;

    let evidence = ProducerEvidence {
        generator,
        shape: shape_label,
        shape_size: generated.map(|(shape, _variant)| shape.name()),
        variant: generated.map(|(_shape, variant)| variant.suffix()),
        sheet_roles,
        shared_string_unique_count,
        shared_string_cell_percent,
        shared_string_part,
        parts,
        planning_sheet: PLANNING_SHEET,
        selected_sheet,
        selected_target: selected_target.clone(),
        selected_target_text: selected_target_text.clone(),
        proven_refusals,
        real_file,
    };

    let corpus = Corpus {
        manifest: CorpusManifest {
            name: format!("xlsx-{shape_label}"),
            generator,
            package_format: "XLSX/OPC/ZIP",
            shape: shape_label,
            payload_kind: "producer-shape-markup-compatibility-grid",
            compression: "deflate",
            entry_count: cell_count,
            archive_member_count,
            entry_bytes: std::mem::size_of::<i32>(),
            uncompressed_payload_bytes: worksheet_bytes,
            archive_bytes: archive.len(),
            archive_sha256: sha256_hex(&archive),
            target_entry: selected_target.clone(),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: None,
            xlsx: Some(XlsxManifest {
                sheet_count,
                rows_per_sheet: rows,
                columns_per_sheet: columns,
                one_percent_update_count: 0,
                source_members,
            }),
        },
        archive,
        target_name: selected_target,
        target_payload,
        xlsx: Some(XlsxCorpus {
            sheet_count,
            row_count: rows,
            column_count: columns,
            one_percent_updates: Vec::new(),
            cell_inventory: None,
        }),
    };
    Ok(ProducerCorpus { corpus, evidence })
}

/// The `Edit` variant exists because the value-only editor admits it; assert
/// that, so the split stays honest if a gate moves.
fn prove_edit_variant_is_admitted(archive: &[u8]) -> Result<(), Box<dyn Error>> {
    let editor = litchi_xlsx::cell_values::SourceBackedEditor::from_read_at(Arc::new(
        OwnedSource::new(archive.to_vec()),
    ))?;
    let transaction = editor.edit_sheets([litchi_xlsx::Selector::from(PLANNING_SHEET)])?;
    if transaction.worksheet_count() != 1 {
        return Err("producer XLSX edit variant planned an unexpected worksheet count".into());
    }
    Ok(())
}

/// Record, once and untimed, the value-only editor's verdict on each producer
/// fact the complete signature carries.
///
/// Each row is the admitted package plus exactly one producer fact, because
/// the gates fire in package-then-part order and an archive carrying several
/// facts can only witness the first.  This is the first-class census the 0587
/// survey's evidence gap 3 asks for.
///
/// Change 0601 wrote this as a *refusal* census and treated an admitted row as
/// a corpus-build failure, because at the time the editor refused all five
/// facts and its reach on Excel output was zero. Changes 0657 and 0667
/// replaced those allow-lists with a dependency rule and a retained shared
/// string table. The census records the verdict rather than asserting it: an
/// admitted row carries [`ADMITTED_VERDICT`] as its message, and the corpus
/// still builds. The schema is unchanged.
fn prove_value_editor_refusals(
    shape: XlsxProducerShape,
    read_archive: &[u8],
) -> Result<Vec<ProvenRefusal>, Box<dyn Error>> {
    let ladder: [(&'static str, &'static str, usize, ArchiveOptions); 5] = [
        (
            "value-only-planning-worksheet-markup-attributes",
            "producer-worksheet-mc-attributes",
            PLANNING_SHEET,
            ArchiveOptions {
                markup_compatibility_attributes: true,
                ..ArchiveOptions::ADMITTED
            },
        ),
        (
            "value-only-planning-page-setup",
            "producer-page-setup",
            PLANNING_SHEET,
            ArchiveOptions {
                page_setup: true,
                ..ArchiveOptions::ADMITTED
            },
        ),
        (
            "value-only-planning-workbook-shared-string-relationship",
            "producer-shared-strings",
            PLANNING_SHEET,
            ArchiveOptions {
                shared_strings: true,
                ..ArchiveOptions::ADMITTED
            },
        ),
        (
            "value-only-planning-worksheet-relationship",
            "producer-worksheet-relationship",
            RELATIONSHIP_SHEET,
            ArchiveOptions {
                worksheet_relationships: true,
                ..ArchiveOptions::ADMITTED
            },
        ),
        (
            "value-only-planning-workbook-root-markers",
            "producer-workbook-root",
            PLANNING_SHEET,
            ArchiveOptions {
                workbook_root_markers: true,
                ..ArchiveOptions::ADMITTED
            },
        ),
    ];
    let mut refusals = Vec::with_capacity(ladder.len() + 1);
    for (scenario, role, sheet, options) in ladder {
        let (archive, _references) = xlsx_producer_archive(shape, options)?;
        refusals.push(refuse(scenario, role, sheet, &archive)?);
    }
    // And the complete signature, which is what the read selectors measure.
    refusals.push(refuse(
        "value-only-planning-complete-producer-signature",
        "producer-complete",
        PLANNING_SHEET,
        read_archive,
    )?);
    Ok(refusals)
}

/// The message recorded for a producer fact the value-only editor admits.
pub(crate) const ADMITTED_VERDICT: &str = "admitted by the value-only editor";

fn refuse(
    scenario: &'static str,
    role: &'static str,
    sheet: usize,
    archive: &[u8],
) -> Result<ProvenRefusal, Box<dyn Error>> {
    let editor = litchi_xlsx::cell_values::SourceBackedEditor::from_read_at(Arc::new(
        OwnedSource::new(archive.to_vec()),
    ))?;
    let message = match editor.edit_sheets([litchi_xlsx::Selector::from(sheet)]) {
        Ok(_) => ADMITTED_VERDICT.to_owned(),
        Err(error) => error.to_string(),
    };
    Ok(ProvenRefusal {
        scenario,
        sheet,
        role,
        message,
    })
}

/// Address the selected-cell selector reads.
///
/// Generated shapes know their extent, so the target is the first cell at or
/// after the middle of the grid that carries the role's payload: a shared
/// string on the shared-string worksheet, a number otherwise.  A real file has
/// no declared extent here, so it is probed over a bounded lattice.
fn scenario_target_address(
    sheet: &litchi_xlsx::SourceWorksheet,
    rows: usize,
    columns: usize,
    role: Option<SheetRole>,
) -> Result<String, Box<dyn Error>> {
    if rows != 0 && columns != 0 {
        let total = rows
            .checked_mul(columns)
            .ok_or("producer worksheet cell count overflows usize")?;
        let middle = total / 2;
        let want_shared = role == Some(SheetRole::SharedStrings);
        for ordinal in middle..total {
            if is_shared_string(ordinal) == want_shared {
                return Ok(cell_address(ordinal / columns, ordinal % columns));
            }
        }
        return Err("producer worksheet has no target cell of the required kind".into());
    }
    const PROBE_ROWS: usize = 64;
    const PROBE_COLUMNS: usize = 32;
    for row in 0..PROBE_ROWS {
        for column in 0..PROBE_COLUMNS {
            let address = cell_address(row, column);
            let view = sheet.cell(address.as_str())?;
            if view.stored().is_some() {
                return Ok(address);
            }
        }
    }
    Err("--real-file worksheet has no stored cell in the probed lattice".into())
}

fn describe_cell(cell: Option<&litchi_xlsx::Cell>) -> String {
    match cell {
        None => "absent".to_owned(),
        Some(litchi_xlsx::Cell::Empty) => "empty".to_owned(),
        Some(litchi_xlsx::Cell::Value(value)) => match value {
            litchi_xlsx::Value::Bool(value) => format!("bool:{value}"),
            litchi_xlsx::Value::Number(value) => format!("number:{}", value.as_str()),
            litchi_xlsx::Value::Text(value) => format!("text:{}", value.as_str()),
            litchi_xlsx::Value::Date(value) => format!("date:{}", value.as_str()),
            litchi_xlsx::Value::Error(value) => format!("error:{value:?}"),
            other => format!("value:{other:?}"),
        },
        Some(litchi_xlsx::Cell::Formula(_)) => "formula".to_owned(),
        Some(other) => format!("other:{other:?}"),
    }
}

// ---------------------------------------------------------------------------
// DOCX generation
// ---------------------------------------------------------------------------

/// The Word 2013 root namespace set, as `word/document.xml` carries it.
fn document_root_attributes() -> String {
    [
        (
            "wpc",
            "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
        ),
        (
            "cx",
            "http://schemas.microsoft.com/office/drawing/2014/chartex",
        ),
        ("mc", MCE_NAMESPACE),
        ("o", "urn:schemas-microsoft-com:office:office"),
        ("r", REL),
        (
            "m",
            "http://schemas.openxmlformats.org/officeDocument/2006/math",
        ),
        ("v", "urn:schemas-microsoft-com:vml"),
        (
            "wp14",
            "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
        ),
        (
            "wp",
            "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        ),
        ("w10", "urn:schemas-microsoft-com:office:word"),
        (
            "w",
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        ),
        (
            "w14",
            "http://schemas.microsoft.com/office/word/2010/wordml",
        ),
        (
            "w15",
            "http://schemas.microsoft.com/office/word/2012/wordml",
        ),
        (
            "wpg",
            "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        ),
        (
            "wpi",
            "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
        ),
        (
            "wne",
            "http://schemas.microsoft.com/office/word/2006/wordml",
        ),
        (
            "wps",
            "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        ),
    ]
    .into_iter()
    .map(|(prefix, uri)| format!(r#" xmlns:{prefix}="{uri}""#))
    .collect::<String>()
}

/// Build the producer-shape DOCX corpus.
///
/// The body the production writer authored is preserved verbatim; only the
/// root start tag is replaced, and every `<w:p>` gains the `w14:paraId`,
/// `w14:textId` and `w:rsidR` attributes Word writes.
pub(crate) fn build_docx_producer_corpus() -> Result<ProducerCorpus, Box<dyn Error>> {
    let shape = SemanticShape::Medium;
    let skeleton = semantic_docx_bytes(shape)?;
    let reader = ArchiveReader::new(&skeleton)?;
    let names = reader.file_names().map(str::to_owned).collect::<Vec<_>>();
    let mut writer = StreamingArchiveWriter::new();
    let mut rewritten = 0;
    let mut paragraph_edits = 0;
    for name in &names {
        let bytes = reader.read(name.as_str())?;
        if name == "word/document.xml" {
            let xml = std::str::from_utf8(&bytes)?;
            let (patched, edits) = producer_document_part(xml)?;
            paragraph_edits = edits;
            writer.write_deflated(name, patched.as_bytes())?;
            rewritten += 1;
        } else {
            writer.write_deflated(name, &bytes)?;
        }
    }
    if rewritten != 1 {
        return Err("producer DOCX rewrite found no word/document.xml".into());
    }
    if paragraph_edits == 0 {
        return Err("producer DOCX rewrite marked no paragraph".into());
    }
    let archive = writer.finish_to_bytes()?;

    // Prove the package still reads, and derive the scenario facts from it.
    let package = litchi_docx::source_backed::Package::from_read_at(Arc::new(OwnedSource::new(
        archive.clone(),
    )))?;
    let document = package.document()?;
    let paragraph_count = document.paragraph_count()?;
    if paragraph_count != shape.docx_paragraphs() {
        return Err("producer DOCX paragraph count differs from the skeleton".into());
    }
    let target_index = paragraph_count / 2;
    let target_text = document
        .paragraph(target_index)?
        .ok_or("producer DOCX target paragraph is missing")?
        .text()?;
    drop(document);
    drop(package);

    let reader = ArchiveReader::new(&archive)?;
    let main = reader.read("word/document.xml")?;
    let parts = vec![census_of(
        "word/document.xml",
        Some("producer-markup"),
        &main,
        None,
        Vec::new(),
    )];
    let archive_member_count = reader.file_names().count();
    let target_payload = target_text.clone().into_bytes();
    Ok(ProducerCorpus {
        evidence: ProducerEvidence {
            generator: DOCX_PRODUCER_SHAPE_GENERATOR,
            shape: "producer-medium",
            shape_size: Some("producer-medium"),
            variant: None,
            sheet_roles: Vec::new(),
            shared_string_unique_count: None,
            shared_string_cell_percent: None,
            shared_string_part: None,
            parts,
            planning_sheet: 0,
            selected_sheet: target_index,
            selected_target: format!("paragraph:{target_index}"),
            selected_target_text: target_text,
            proven_refusals: Vec::new(),
            real_file: None,
        },
        corpus: Corpus {
            manifest: CorpusManifest {
                name: "docx-producer-medium".to_owned(),
                generator: DOCX_PRODUCER_SHAPE_GENERATOR,
                package_format: "DOCX/OPC/ZIP",
                shape: "producer-medium",
                payload_kind: "producer-shape-word-2013-namespaces",
                compression: "deflate",
                entry_count: paragraph_count,
                archive_member_count,
                entry_bytes: target_payload.len(),
                uncompressed_payload_bytes: main.len(),
                archive_bytes: archive.len(),
                archive_sha256: sha256_hex(&archive),
                target_entry: format!("paragraph:{target_index}"),
                target_payload_bytes: target_payload.len(),
                target_payload_sha256: sha256_hex(&target_payload),
                rtf_variant: None,
                xlsx: None,
            },
            archive,
            target_name: format!("paragraph:{target_index}"),
            target_payload,
            xlsx: None,
        },
    })
}

fn producer_document_part(xml: &str) -> Result<(String, usize), Box<dyn Error>> {
    let open = xml
        .find("<w:document")
        .ok_or("producer DOCX rewrite found no document root")?;
    let close = xml[open..]
        .find('>')
        .ok_or("producer DOCX document root start tag is unterminated")?
        + open;
    if xml[open..=close].contains("mc:Ignorable") {
        return Err("producer DOCX root already declares mc:Ignorable".into());
    }
    let root = format!(
        r#"<w:document{attributes} mc:Ignorable="{DOCUMENT_IGNORABLE}">"#,
        attributes = document_root_attributes()
    );
    let body = &xml[close + 1..];
    let mut rebuilt = String::new();
    rebuilt.push_str(&xml[..open]);
    rebuilt.push_str(&root);
    let mut edits = 0;
    let mut rest = body;
    while let Some(position) = rest.find("<w:p>") {
        rebuilt.push_str(&rest[..position]);
        rebuilt.push_str(&format!(
            r#"<w:p w14:paraId="{:08X}" w14:textId="{:08X}" w:rsidR="00AB12CD" w:rsidRDefault="00AB12CD">"#,
            edits + 1,
            0x7777_0000_u32 + u32::try_from(edits)?
        ));
        rest = &rest[position + "<w:p>".len()..];
        edits += 1;
    }
    rebuilt.push_str(rest);
    Ok((rebuilt, edits))
}

// ---------------------------------------------------------------------------
// PPTX generation
// ---------------------------------------------------------------------------

/// Build the producer-shape PPTX corpus.
///
/// Every slide gains one `mc:AlternateContent` wrapper of the kind PowerPoint
/// emits for a chart or 3D-text shape: an `mc:Choice Requires="a14"` branch
/// and an `mc:Fallback` branch, both built from a shape the production writer
/// itself authored, so the slide stays exactly as valid as before.
pub(crate) fn build_pptx_producer_corpus() -> Result<ProducerCorpus, Box<dyn Error>> {
    let shape = SemanticShape::Medium;
    let skeleton = semantic_pptx_bytes(shape)?;
    let reader = ArchiveReader::new(&skeleton)?;
    let names = reader.file_names().map(str::to_owned).collect::<Vec<_>>();
    let mut writer = StreamingArchiveWriter::new();
    let mut wrapped = 0;
    for name in &names {
        let bytes = reader.read(name.as_str())?;
        if name.starts_with("ppt/slides/slide") && name.ends_with(".xml") {
            let xml = std::str::from_utf8(&bytes)?;
            writer.write_deflated(name, producer_slide_part(xml, wrapped)?.as_bytes())?;
            wrapped += 1;
        } else {
            writer.write_deflated(name, &bytes)?;
        }
    }
    if wrapped == 0 {
        return Err("producer PPTX rewrite found no slide part".into());
    }
    let archive = writer.finish_to_bytes()?;

    let presentation = litchi_pptx::SourceBackedPresentation::from_read_at(Arc::new(
        OwnedSource::new(archive.clone()),
    ))?;
    let slide_count = presentation.slide_count();
    if slide_count != wrapped {
        return Err("producer PPTX slide count differs from the rewritten part count".into());
    }
    let target_index = slide_count / 2;
    let (target_text, _name) = presentation
        .slide(target_index)
        .ok_or("producer PPTX target slide is missing")?
        .text_and_name()?;
    drop(presentation);

    let reader = ArchiveReader::new(&archive)?;
    let archive_member_count = reader.file_names().count();
    let member = format!("ppt/slides/slide{}.xml", target_index + 1);
    let slide = reader.read(member.as_str())?;
    let parts = vec![census_of(
        member.as_str(),
        Some("producer-alternate-content"),
        &slide,
        None,
        Vec::new(),
    )];
    let target_payload = target_text.clone().into_bytes();
    Ok(ProducerCorpus {
        evidence: ProducerEvidence {
            generator: PPTX_PRODUCER_SHAPE_GENERATOR,
            shape: "producer-medium",
            shape_size: Some("producer-medium"),
            variant: None,
            sheet_roles: Vec::new(),
            shared_string_unique_count: None,
            shared_string_cell_percent: None,
            shared_string_part: None,
            parts,
            planning_sheet: 0,
            selected_sheet: target_index,
            selected_target: format!("slide:{target_index}"),
            selected_target_text: target_text,
            proven_refusals: Vec::new(),
            real_file: None,
        },
        corpus: Corpus {
            manifest: CorpusManifest {
                name: "pptx-producer-medium".to_owned(),
                generator: PPTX_PRODUCER_SHAPE_GENERATOR,
                package_format: "PPTX/OPC/ZIP",
                shape: "producer-medium",
                payload_kind: "producer-shape-alternate-content-slides",
                compression: "deflate",
                entry_count: slide_count,
                archive_member_count,
                entry_bytes: target_payload.len(),
                uncompressed_payload_bytes: slide.len(),
                archive_bytes: archive.len(),
                archive_sha256: sha256_hex(&archive),
                target_entry: format!("slide:{target_index}"),
                target_payload_bytes: target_payload.len(),
                target_payload_sha256: sha256_hex(&target_payload),
                rtf_variant: None,
                xlsx: None,
            },
            archive,
            target_name: format!("slide:{target_index}"),
            target_payload,
            xlsx: None,
        },
    })
}

fn producer_slide_part(xml: &str, slide: usize) -> Result<String, Box<dyn Error>> {
    let close = xml
        .find("</p:spTree>")
        .ok_or("producer PPTX slide has no shape tree")?;
    let open = xml
        .find("<p:sp>")
        .ok_or("producer PPTX slide has no ordinary shape")?;
    let end = xml[open..]
        .find("</p:sp>")
        .ok_or("producer PPTX shape is unterminated")?
        + open
        + "</p:sp>".len();
    let template = &xml[open..end];
    let choice = retarget_shape(template, 9_000 + slide, "choice")?;
    let fallback = retarget_shape(template, 9_500 + slide, "fallback")?;
    let block = format!(
        concat!(
            r#"<mc:AlternateContent xmlns:mc="{mce}">"#,
            r#"<mc:Choice xmlns:a14="{a14}" Requires="a14">{choice}</mc:Choice>"#,
            r#"<mc:Fallback>{fallback}</mc:Fallback>"#,
            r#"</mc:AlternateContent>"#
        ),
        mce = MCE_NAMESPACE,
        a14 = A14_NAMESPACE,
        choice = choice,
        fallback = fallback,
    );
    Ok(format!("{}{}{}", &xml[..close], block, &xml[close..]))
}

/// Give a cloned shape a distinct non-visual id and name so the slide keeps a
/// unique shape catalog.
fn retarget_shape(
    template: &str,
    identifier: usize,
    branch: &str,
) -> Result<String, Box<dyn Error>> {
    let marker = "<p:cNvPr ";
    let open = template
        .find(marker)
        .ok_or("producer PPTX shape has no non-visual properties")?;
    let close = template[open..]
        .find('>')
        .ok_or("producer PPTX non-visual properties are unterminated")?
        + open;
    let replacement =
        format!(r#"<p:cNvPr id="{identifier}" name="litchi-perf-producer-{branch}-{identifier}""#);
    let original = &template[open..=close];
    let tail = original
        .find("/>")
        .map_or_else(|| ">".to_owned(), |_| "/>".to_owned());
    Ok(format!(
        "{}{}{}{}",
        &template[..open],
        replacement,
        tail,
        &template[close + 1..]
    ))
}

// ---------------------------------------------------------------------------
// Scenario runners
// ---------------------------------------------------------------------------

/// Scenario the producer-shape selectors measure.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Scenario {
    /// Open the source-backed owner; nothing is materialized.
    Open,
    /// One selected cell of the shared-string worksheet.
    SelectedCell,
    /// Value-only multi-sheet planning, exactly the interval the XLSX planning
    /// guard measures.
    Planning,
    /// One-cell edit, commit and publication to a bounded counting sink.
    OneEditSave,
    /// One selected paragraph of the producer-shape DOCX.
    SelectedParagraph,
    /// One selected slide of the producer-shape PPTX.
    SelectedSlide,
}

pub(crate) fn run_case(
    case: Case,
    scenario: Scenario,
    corpus: &ProducerCorpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    match scenario {
        Scenario::Open => run_xlsx_open(case, corpus, warmup_iterations, samples),
        Scenario::SelectedCell => run_xlsx_selected_cell(case, corpus, warmup_iterations, samples),
        Scenario::Planning => run_xlsx_planning(case, corpus, warmup_iterations, samples),
        Scenario::OneEditSave => run_xlsx_one_edit_save(case, corpus, warmup_iterations, samples),
        Scenario::SelectedParagraph => {
            run_docx_selected_paragraph(case, corpus, warmup_iterations, samples)
        },
        Scenario::SelectedSlide => {
            run_pptx_selected_slide(case, corpus, warmup_iterations, samples)
        },
    }
}

fn source_of(corpus: &ProducerCorpus) -> Arc<dyn ReadAt> {
    Arc::new(OwnedSource::new(corpus.corpus.archive.clone()))
}

fn plain_result(case: Case, corpus: &ProducerCorpus, elapsed: Vec<u64>) -> CaseResult {
    CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: None,
        source: None,
        execution: None,
        output_sha256: None,
        operation_metrics: None,
    }
}

fn run_xlsx_open(
    case: Case,
    corpus: &ProducerCorpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let expected_sheets = corpus
        .corpus
        .xlsx
        .as_ref()
        .ok_or("producer XLSX corpus has no specification")?
        .sheet_count;
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let source = source_of(corpus);
        let started = Instant::now();
        let workbook = litchi_xlsx::SourceBackedWorkbook::from_read_at(source)?;
        let duration = started.elapsed();
        if workbook.len() != expected_sheets {
            return Err("producer XLSX open sheet count differs from the corpus".into());
        }
        std::hint::black_box(&workbook);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(plain_result(case, corpus, elapsed))
}

fn selected_cell_target(corpus: &ProducerCorpus) -> Result<(String, String), Box<dyn Error>> {
    let (sheet, address) = corpus
        .evidence
        .selected_target
        .split_once('!')
        .ok_or("producer XLSX selected target is not sheet!address")?;
    Ok((sheet.to_owned(), address.to_owned()))
}

fn run_xlsx_selected_cell(
    case: Case,
    corpus: &ProducerCorpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let (sheet_name, address) = selected_cell_target(corpus)?;
    let expected = corpus.evidence.selected_target_text.clone();
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let workbook = litchi_xlsx::SourceBackedWorkbook::from_read_at(source_of(corpus))?;
        let sheet = workbook
            .sheet(sheet_name.as_str())?
            .ok_or("producer XLSX selected worksheet is missing")?;
        let started = Instant::now();
        let cell = sheet.cell(address.as_str())?;
        let duration = started.elapsed();
        if describe_cell(cell.stored()) != expected {
            return Err("producer XLSX selected cell differs from the corpus".into());
        }
        std::hint::black_box(&cell);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(plain_result(case, corpus, elapsed))
}

fn run_xlsx_planning(
    case: Case,
    corpus: &ProducerCorpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let selectors = [litchi_xlsx::Selector::from(PLANNING_SHEET)];
    let before = sha256_hex(&corpus.corpus.archive);
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        // Opening is deliberately outside the planning interval, exactly as
        // the XLSX planning guard measures it.
        let editor = litchi_xlsx::cell_values::SourceBackedEditor::from_read_at(source_of(corpus))?;
        let started = Instant::now();
        let planned = editor.edit_sheets(selectors.iter().cloned());
        let duration = started.elapsed();
        let transaction = planned?;
        if transaction.worksheet_count() != 1 {
            return Err("producer XLSX planning selected an unexpected worksheet count".into());
        }
        let commit = transaction.commit()?;
        if commit.changed() || !commit.patch().is_empty() {
            return Err("producer XLSX empty planning commit was not an exact no-op".into());
        }
        if sha256_hex(&corpus.corpus.archive) != before {
            return Err("producer XLSX planning changed source bytes".into());
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(plain_result(case, corpus, elapsed))
}

fn run_xlsx_one_edit_save(
    case: Case,
    corpus: &ProducerCorpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let spec = corpus
        .corpus
        .xlsx
        .as_ref()
        .ok_or("producer XLSX corpus has no specification")?;
    let row = u32::try_from(spec.row_count / 2)?;
    let column = u32::try_from(spec.column_count / 2)?;
    let address = litchi_xlsx::Address::at(row, column)?;
    let replacement = numeric_value(PLANNING_SHEET, spec.row_count / 2, spec.column_count / 2) + 1;
    let maximum = u64::try_from(corpus.corpus.archive.len().saturating_mul(4).max(64 * 1024))?;
    let before = sha256_hex(&corpus.corpus.archive);
    let mut elapsed = Vec::with_capacity(samples);
    let mut sink_summaries = Vec::with_capacity(samples);
    let mut digests = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let mut sink = CountingSink::bounded(maximum, 64 * 1024);
        sink.reserve_budget()?;
        let editor = litchi_xlsx::cell_values::SourceBackedEditor::from_read_at(source_of(corpus))?;
        let mut duration = Duration::ZERO;
        let started = Instant::now();
        let mut edit = editor.edit_sheets([litchi_xlsx::Selector::from(PLANNING_SHEET)])?;
        edit.set(PLANNING_SHEET, address, i32::try_from(replacement)?)?;
        let commit = edit.commit()?;
        if !commit.changed() {
            return Err("producer XLSX one-cell edit produced no change".into());
        }
        editor.publish_multi_commit_to_stream(&mut sink, &commit)?;
        duration += started.elapsed();
        if sha256_hex(&corpus.corpus.archive) != before {
            return Err("producer XLSX edit/save changed source bytes".into());
        }
        let digest = sha256_hex(&sink.bytes);
        std::hint::black_box(&sink.bytes);
        if iteration >= warmup_iterations {
            sink_summaries.push(sink.summary());
            digests.push(digest);
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    let expected = digests
        .first()
        .cloned()
        .ok_or("producer XLSX edit/save retained no sample")?;
    if digests.iter().any(|digest| digest != &expected) {
        return Err("producer XLSX edit/save output digests are not stable".into());
    }
    let sink = crate::deterministic_sink_summary(&sink_summaries, "producer XLSX edit/save")?;
    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: Some(sink),
        source: None,
        execution: None,
        output_sha256: Some(expected),
        operation_metrics: None,
    })
}

fn run_docx_selected_paragraph(
    case: Case,
    corpus: &ProducerCorpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let index = corpus.evidence.selected_sheet;
    let expected = corpus.evidence.selected_target_text.clone();
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let package = litchi_docx::source_backed::Package::from_read_at(source_of(corpus))?;
        let document = package.document()?;
        let started = Instant::now();
        let text = document
            .paragraph(index)?
            .ok_or("producer DOCX selected paragraph is missing")?
            .text()?;
        let duration = started.elapsed();
        if text != expected {
            return Err("producer DOCX selected paragraph differs from the corpus".into());
        }
        std::hint::black_box(&text);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(plain_result(case, corpus, elapsed))
}

fn run_pptx_selected_slide(
    case: Case,
    corpus: &ProducerCorpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    let index = corpus.evidence.selected_sheet;
    let expected = corpus.evidence.selected_target_text.clone();
    let mut elapsed = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let presentation = litchi_pptx::SourceBackedPresentation::from_read_at(source_of(corpus))?;
        let slide = presentation
            .slide(index)
            .ok_or("producer PPTX selected slide is missing")?;
        let started = Instant::now();
        let (text, _name) = slide.text_and_name()?;
        let duration = started.elapsed();
        if text != expected {
            return Err("producer PPTX selected slide differs from the corpus".into());
        }
        std::hint::black_box(&text);
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    Ok(plain_result(case, corpus, elapsed))
}

/// Write the collected producer-shape evidence blocks for the retained packet.
///
/// This is the first-class marker and refusal census the 0587 survey's gap 3
/// asks for: instead of each record re-deriving "41 of 60 sheets declare mc"
/// by hand, a run states, per corpus and per part, which producer markers the
/// part carries and which library gate each one trips.
pub(crate) fn write_evidence(
    path: &Path,
    evidence: &[ProducerEvidence],
) -> Result<(), Box<dyn Error>> {
    let document = serde_json::json!({
        "schema": "litchi.perf-baseline.producer-shape-evidence.v1",
        "corpora": evidence,
    });
    let mut bytes = serde_json::to_vec_pretty(&document)?;
    bytes.push(b'\n');
    fs::write(path, bytes).map_err(|source| {
        format!(
            "producer evidence {} is unwritable: {source}",
            path.display()
        )
    })?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn xlsx_medium(variant: XlsxProducerVariant) -> ProducerCorpus {
        build_xlsx_producer_corpus(XlsxProducerShape::Medium, variant)
            .expect("producer XLSX medium corpus builds")
    }

    #[test]
    fn producer_generation_is_byte_identical_across_runs() {
        // The dense shape is proven deterministic across two independent
        // release-mode processes in the retained packet; repeating it here
        // would make the debug-profile gate minutes long for no extra fact.
        for variant in XlsxProducerVariant::ALL {
            let first = build_xlsx_producer_corpus(XlsxProducerShape::Medium, variant)
                .expect("first build");
            let second = build_xlsx_producer_corpus(XlsxProducerShape::Medium, variant)
                .expect("second build");
            assert_eq!(
                first.corpus.manifest.archive_sha256,
                second.corpus.manifest.archive_sha256,
                "{} archive is not deterministic",
                variant.suffix()
            );
            assert_eq!(first.corpus.archive, second.corpus.archive);
        }
        let docx_first = build_docx_producer_corpus().expect("first DOCX build");
        let docx_second = build_docx_producer_corpus().expect("second DOCX build");
        assert_eq!(docx_first.corpus.archive, docx_second.corpus.archive);
        let pptx_first = build_pptx_producer_corpus().expect("first PPTX build");
        let pptx_second = build_pptx_producer_corpus().expect("second PPTX build");
        assert_eq!(pptx_first.corpus.archive, pptx_second.corpus.archive);
    }

    #[test]
    fn producer_read_and_edit_variants_are_distinct_corpora() {
        let read = xlsx_medium(XlsxProducerVariant::Read);
        let edit = xlsx_medium(XlsxProducerVariant::Edit);
        assert_ne!(
            read.corpus.manifest.archive_sha256,
            edit.corpus.manifest.archive_sha256
        );
        assert_eq!(read.corpus.manifest.shape, "producer-medium-read");
        assert_eq!(edit.corpus.manifest.shape, "producer-medium-edit");
        assert!(read.evidence.shared_string_part.is_some());
        assert!(edit.evidence.shared_string_part.is_some());
    }

    #[test]
    fn xlsx_producer_worksheets_carry_the_real_producer_signature() {
        for variant in XlsxProducerVariant::ALL {
            let corpus = xlsx_medium(variant);
            let worksheets = corpus
                .evidence
                .parts
                .iter()
                .filter(|part| part.part.starts_with("xl/worksheets/"))
                .collect::<Vec<_>>();
            assert_eq!(worksheets.len(), 4);
            if variant == XlsxProducerVariant::Control {
                for part in &worksheets {
                    // The control is the marker-free counterpart: it is what
                    // the existing corpora already are.
                    assert!(!part.markup_compatibility_namespace, "{}", part.part);
                    assert!(!part.x14ac_namespace);
                    assert!(!part.cols_block);
                    assert!(part.source_stream_eligible, "{}", part.part);
                    assert!(part.source_stream_ineligible_reasons.is_empty());
                }
                continue;
            }
            for part in &worksheets {
                // The namespace declarations and the `<cols>` block are on
                // both producer variants: they are what defeats 0546's fused traversal,
                // the MCE presence scan and the selected-cell stream.
                assert!(part.markup_compatibility_namespace, "{}", part.part);
                assert!(part.x14ac_namespace);
                assert!(part.cols_block);
                if variant == XlsxProducerVariant::Read {
                    assert_eq!(part.mc_ignorable.as_deref(), Some("x14ac xr xr2 xr3"));
                    assert!(part.dy_descent_occurrences > 32, "{}", part.part);
                } else {
                    assert_eq!(part.mc_ignorable, None);
                    assert_eq!(part.dy_descent_occurrences, 0);
                }
                // The library's fused traversal and its selected-cell stream
                // are both unavailable on this shape, exactly as on real
                // files.
                assert!(!part.source_stream_eligible, "{}", part.part);
                assert!(part.selected_cell_stream_ineligible_on_cols);
                assert!(
                    part.source_stream_ineligible_reasons
                        .contains(&"markup-compatibility-namespace")
                );
                assert!(
                    part.source_stream_ineligible_reasons
                        .contains(&"x14ac-namespace")
                );
                if variant == XlsxProducerVariant::Read {
                    assert!(
                        part.source_stream_ineligible_reasons
                            .contains(&"dy-descent")
                    );
                }
            }
        }
    }

    #[test]
    fn xlsx_producer_read_variant_carries_shared_strings_and_relationships() {
        let corpus = xlsx_medium(XlsxProducerVariant::Read);
        let worksheets = corpus
            .evidence
            .parts
            .iter()
            .filter(|part| part.part.starts_with("xl/worksheets/"))
            .collect::<Vec<_>>();
        assert_eq!(worksheets[0].shared_string_cells, 0);
        assert!(worksheets[1].shared_string_cells > 0);
        assert_eq!(worksheets[2].shared_string_cells, 0);
        assert_eq!(
            worksheets[2].worksheet_relationship_part.as_deref(),
            Some("xl/worksheets/_rels/sheet3.xml.rels")
        );
        assert_eq!(
            worksheets[2].worksheet_relationship_targets,
            vec!["../printerSettings/printerSettings1.bin".to_owned()]
        );
        assert_eq!(
            corpus.evidence.shared_string_part.as_deref(),
            Some("xl/sharedStrings.xml")
        );
        assert_eq!(corpus.evidence.shared_string_cell_percent, Some(40));
        assert_eq!(corpus.evidence.shared_string_unique_count, Some(64));
        let workbook = corpus
            .evidence
            .parts
            .iter()
            .find(|part| part.part == "xl/workbook.xml")
            .expect("workbook part is censused");
        assert_eq!(
            workbook.mc_ignorable.as_deref(),
            Some("x15 xr xr6 xr10 xr2")
        );
        // The selected-cell target is a shared string, which is the fact the
        // survey's XLSX-2 turns on.
        assert!(
            corpus.evidence.selected_target_text.starts_with("text:"),
            "selected target was {}",
            corpus.evidence.selected_target_text
        );
    }

    #[test]
    fn xlsx_producer_shared_string_share_is_forty_percent() {
        let corpus = xlsx_medium(XlsxProducerVariant::Read);
        let sheet = corpus
            .evidence
            .parts
            .iter()
            .find(|part| part.part == "xl/worksheets/sheet2.xml")
            .expect("shared-string worksheet is present");
        let total = sheet.shared_string_cells + sheet.numeric_cells;
        assert_eq!(total, 32 * 32);
        assert_eq!(sheet.shared_string_cells * 100 / total, 40);
    }

    #[test]
    fn xlsx_producer_read_variant_records_the_value_editor_verdict_on_every_fact() {
        let corpus = xlsx_medium(XlsxProducerVariant::Read);
        let refusals = &corpus.evidence.proven_refusals;
        assert_eq!(refusals.len(), 6);
        // Changes 0657 and 0667 replaced the value-only editor's allow-lists
        // with a dependency rule and a retained shared-string table. Each
        // isolated producer fact, and the complete signature when planning
        // the numeric worksheet, is therefore admitted.
        assert_eq!(refusals[0].role, "producer-worksheet-mc-attributes");
        assert_eq!(refusals[0].message, ADMITTED_VERDICT);
        assert_eq!(refusals[1].role, "producer-page-setup");
        assert_eq!(refusals[1].message, ADMITTED_VERDICT);
        assert_eq!(refusals[2].role, "producer-shared-strings");
        assert_eq!(refusals[2].message, ADMITTED_VERDICT);
        assert_eq!(refusals[3].role, "producer-worksheet-relationship");
        assert_eq!(refusals[3].message, ADMITTED_VERDICT);
        assert_eq!(refusals[4].role, "producer-workbook-root");
        assert_eq!(refusals[4].message, ADMITTED_VERDICT);
        // The complete signature is admitted for the selected numeric
        // worksheet; its other worksheets remain independently selectable.
        assert_eq!(refusals[5].role, "producer-complete");
        assert_eq!(refusals[5].message, ADMITTED_VERDICT);
    }

    #[test]
    fn xlsx_producer_edit_variant_is_admitted_and_proves_nothing_else() {
        let corpus = xlsx_medium(XlsxProducerVariant::Edit);
        assert!(corpus.evidence.proven_refusals.is_empty());
        assert!(
            corpus
                .evidence
                .parts
                .iter()
                .all(|part| part.worksheet_relationship_part.is_none())
        );
    }

    #[test]
    fn docx_producer_document_carries_the_word_2013_namespace_set() {
        let corpus = build_docx_producer_corpus().expect("producer DOCX corpus builds");
        let part = &corpus.evidence.parts[0];
        assert_eq!(part.part, "word/document.xml");
        assert!(part.markup_compatibility_namespace);
        assert_eq!(part.mc_ignorable.as_deref(), Some("w14 w15 wp14"));
        assert!(!part.source_stream_eligible);
        let xml = String::from_utf8(
            ArchiveReader::new(&corpus.corpus.archive)
                .expect("archive reads")
                .read("word/document.xml")
                .expect("document part reads"),
        )
        .expect("document part is UTF-8");
        assert!(
            xml.contains(r#"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#)
        );
        assert!(
            xml.contains(r#"xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml""#)
        );
        assert!(xml.contains("w14:paraId="));
    }

    #[test]
    fn pptx_producer_slides_carry_alternate_content() {
        let corpus = build_pptx_producer_corpus().expect("producer PPTX corpus builds");
        let part = &corpus.evidence.parts[0];
        assert!(part.part.starts_with("ppt/slides/slide"));
        assert_eq!(part.alternate_content_occurrences, 1);
        assert!(part.markup_compatibility_namespace);
        assert!(
            part.source_stream_ineligible_reasons
                .contains(&"alternate-content")
        );
    }

    #[test]
    fn real_file_corpus_refuses_a_missing_path() {
        let error = build_xlsx_real_file_corpus(Path::new(
            "/nonexistent/litchi-perf-baseline/real-file.xlsx",
        ))
        .expect_err("a missing --real-file is refused");
        assert!(error.to_string().contains("--real-file"));
    }
}
