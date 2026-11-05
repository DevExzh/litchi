//! Test-only runtime refusal evidence for the generated producer worksheet.
//!
//! These tests call the same public raw selected-range scanner used by the
//! source-backed worksheet route.  They do not observe the surrounding route,
//! time a read, or claim that a selected result was published.  The generated
//! archive stays the reviewed four-sheet producer fixture; the XML controls
//! are bounded copies that remove one earlier marker at a time so the next
//! scanner refusal is identified by an actual event.

#![allow(clippy::unwrap_used, reason = "focused producer refusal assertions")]

use std::{
    io::{BufReader, Cursor},
    string::String,
};

use litchi_ooxml_common::mce::{Capabilities, Error as MceError, StreamError, StreamLimits};
use litchi_xlsx::Rect;
use litchi_xlsx::raw::selected_worksheet::{
    NotEligibleReason, RangeScanOutcome, SelectedPayload, scan_range,
};
use soapberry_zip::office::ArchiveReader;

use super::*;

const REVIEWED_ARCHIVE_BYTES: usize = 22_739;
const REVIEWED_ARCHIVE_SHA256: &str =
    "55e0901c093c27d80bcd4e5fef888f00798654cdef8c0d79d130a2ce7cbe0ae8";
const REVIEWED_SHEETS: usize = 4;
const REVIEWED_ROWS: usize = 32;
const REVIEWED_COLUMNS: usize = 32;
const REVIEWED_CELL_COUNT: usize = 4_096;
const REVIEWED_NUMERIC_COUNT: usize = 3_686;
const REVIEWED_TEXT_COUNT: usize = 410;

/// One bounded XML control over a generated producer worksheet.
///
/// A read scan (the source-backed worksheet route uses `Rect::ALL`) admits
/// the producer's `sheetViews`, regular `sheetFormatPr`, row `spans` hint and
/// `pageMargins` (litchi-xlsx `raw/worksheet/selected.rs`, "Inert worksheet
/// children", since `5fe77e0bb6`).  The producer's refusing markers are
/// therefore `cols` (styles) and, on the relationship-bearing sheet only,
/// `pageSetup` (structure).  The controls remove those markers one at a time
/// to reveal each downstream refusal, remove the admitted markers to prove
/// they never decide the verdict, and break one admitted marker to prove the
/// admission is checked rather than blanket.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Control {
    CompleteProducer,
    WithoutSheetViews,
    WithoutSheetViewsAndFormat,
    WithoutCols,
    WithoutColsIrregularSheetFormat,
    WithoutColsAndPageSetup,
    WithoutSheetViewsFormatAndCols,
    WithoutMetadataThroughRowSpans,
    WithoutPageMargins,
    WithoutPageMarginsAndPageSetup,
}

impl Control {
    const fn name(self) -> &'static str {
        match self {
            Self::CompleteProducer => "producer-complete-first-refusal-cols-styles",
            Self::WithoutSheetViews => "remove-admitted-sheetViews-still-cols-styles",
            Self::WithoutSheetViewsAndFormat => {
                "remove-admitted-sheetViews-and-sheetFormatPr-still-cols-styles"
            },
            Self::WithoutCols => "remove-cols-next-pageSetup-or-eligible",
            Self::WithoutColsIrregularSheetFormat => {
                "remove-cols-and-defaultRowHeight-irregular-sheetFormatPr-structure"
            },
            Self::WithoutColsAndPageSetup => {
                "remove-cols-and-pageSetup-eligible-with-admitted-metadata"
            },
            Self::WithoutSheetViewsFormatAndCols => {
                "remove-sheetViews-format-and-cols-next-pageSetup-or-eligible"
            },
            Self::WithoutMetadataThroughRowSpans => {
                "remove-sheetViews-format-cols-and-row-spans-next-pageSetup-or-eligible"
            },
            Self::WithoutPageMargins => "remove-through-pageMargins-next-pageSetup-or-eligible",
            Self::WithoutPageMarginsAndPageSetup => "remove-all-producer-metadata-eligible",
        }
    }
}

fn reviewed_read_corpus() -> ProducerCorpus {
    let corpus = build_xlsx_producer_corpus(XlsxProducerShape::Medium, XlsxProducerVariant::Read)
        .expect("reviewed producer XLSX corpus builds");
    assert_eq!(corpus.corpus.archive.len(), REVIEWED_ARCHIVE_BYTES);
    assert_eq!(
        corpus.corpus.manifest.archive_sha256,
        REVIEWED_ARCHIVE_SHA256
    );
    assert_eq!(corpus.corpus.manifest.archive_member_count, 12);
    assert_eq!(corpus.corpus.manifest.entry_count, REVIEWED_CELL_COUNT);
    let spec = corpus
        .corpus
        .xlsx
        .as_ref()
        .expect("XLSX shape specification");
    assert_eq!(spec.sheet_count, REVIEWED_SHEETS);
    assert_eq!(spec.row_count, REVIEWED_ROWS);
    assert_eq!(spec.column_count, REVIEWED_COLUMNS);

    let worksheets = corpus
        .evidence
        .parts
        .iter()
        .filter(|part| part.part.starts_with("xl/worksheets/sheet"))
        .collect::<Vec<_>>();
    assert_eq!(worksheets.len(), REVIEWED_SHEETS);
    assert_eq!(
        worksheets
            .iter()
            .map(|part| part.numeric_cells)
            .sum::<usize>(),
        REVIEWED_NUMERIC_COUNT
    );
    assert_eq!(
        worksheets
            .iter()
            .map(|part| part.shared_string_cells)
            .sum::<usize>(),
        REVIEWED_TEXT_COUNT
    );
    corpus
}

fn worksheet_xmls(archive: &[u8]) -> Vec<Vec<u8>> {
    let reader = ArchiveReader::new(archive, soapberry_zip::office::ArchiveLimits::default())
        .expect("producer archive opens");
    let mut worksheets = Vec::with_capacity(REVIEWED_SHEETS);
    for sheet in 1..=REVIEWED_SHEETS {
        let name = format!("xl/worksheets/sheet{sheet}.xml");
        worksheets.push(reader.read(&name).expect("producer worksheet exists"));
    }
    worksheets
}

/// Remove one complete XML element from the generated worksheet copy.
///
/// The generated fixture uses only ASCII element names and deterministic
/// self-closing or paired markup.  The helper still checks the boundary after
/// the name and handles both forms so a generator change cannot silently make
/// a control remove a substring from a different element.
fn remove_element(xml: &str, name: &str) -> String {
    let open = format!("<{name}");
    let bytes = xml.as_bytes();
    let mut search_from = 0;
    let start = loop {
        let Some(relative) = xml[search_from..].find(&open) else {
            panic!("control {} did not find <{name}>", name);
        };
        let candidate = search_from + relative;
        let after_name = candidate + open.len();
        if matches!(bytes.get(after_name), Some(b'>') | Some(b' ') | Some(b'/')) {
            break candidate;
        }
        search_from = after_name;
    };
    let open_end = start
        + xml[start..]
            .find('>')
            .expect("control element has a start-tag closure");
    let end = if bytes.get(open_end.saturating_sub(1)) == Some(&b'/') {
        open_end + 1
    } else {
        let close = format!("</{name}>");
        open_end
            + 1
            + xml[open_end + 1..]
                .find(&close)
                .expect("control element has an end-tag closure")
            + close.len()
    };
    let mut result = String::with_capacity(xml.len().saturating_sub(end - start));
    result.push_str(&xml[..start]);
    result.push_str(&xml[end..]);
    result
}

fn remove_row_spans(xml: &str) -> String {
    let marker = format!(r#" spans="1:{REVIEWED_COLUMNS}""#);
    assert_eq!(xml.matches(&marker).count(), REVIEWED_ROWS);
    xml.replace(&marker, "")
}

fn remove_default_row_height(xml: &str) -> String {
    let marker = r#" defaultRowHeight="15""#;
    assert_eq!(xml.matches(marker).count(), 1);
    xml.replacen(marker, "", 1)
}

fn controlled_xml(xml: &[u8], control: Control) -> String {
    let original = String::from_utf8(xml.to_vec()).expect("producer worksheet is UTF-8");
    let mut result = original;
    let strips_admitted_prefix = matches!(
        control,
        Control::WithoutSheetViews
            | Control::WithoutSheetViewsAndFormat
            | Control::WithoutSheetViewsFormatAndCols
            | Control::WithoutMetadataThroughRowSpans
            | Control::WithoutPageMargins
            | Control::WithoutPageMarginsAndPageSetup
    );
    if strips_admitted_prefix {
        result = remove_element(&result, "sheetViews");
    }
    if strips_admitted_prefix && !matches!(control, Control::WithoutSheetViews) {
        result = remove_element(&result, "sheetFormatPr");
    }
    if matches!(
        control,
        Control::WithoutCols
            | Control::WithoutColsIrregularSheetFormat
            | Control::WithoutColsAndPageSetup
            | Control::WithoutSheetViewsFormatAndCols
            | Control::WithoutMetadataThroughRowSpans
            | Control::WithoutPageMargins
            | Control::WithoutPageMarginsAndPageSetup
    ) {
        result = remove_element(&result, "cols");
    }
    if matches!(control, Control::WithoutColsIrregularSheetFormat) {
        result = remove_default_row_height(&result);
    }
    if matches!(
        control,
        Control::WithoutMetadataThroughRowSpans
            | Control::WithoutPageMargins
            | Control::WithoutPageMarginsAndPageSetup
    ) {
        result = remove_row_spans(&result);
    }
    if matches!(
        control,
        Control::WithoutPageMargins | Control::WithoutPageMarginsAndPageSetup
    ) {
        result = remove_element(&result, "pageMargins");
    }
    if matches!(
        control,
        Control::WithoutColsAndPageSetup | Control::WithoutPageMarginsAndPageSetup
    ) && result.contains("<pageSetup")
    {
        result = remove_element(&result, "pageSetup");
    }
    result
}

#[allow(
    clippy::result_large_err,
    reason = "the test helper returns the leaf scanner's own result type unchanged"
)]
fn scan_xml(xml: &str) -> litchi_xlsx::raw::selected_worksheet::StreamResult<RangeScanOutcome> {
    let mut limits = StreamLimits::default();
    limits.processing.max_input_bytes = xml.len();
    limits.processing.max_output_bytes = xml.len();
    let mut input = BufReader::new(Cursor::new(xml.as_bytes()));
    scan_range(&mut input, &Capabilities::default(), &limits, Rect::ALL)
}

fn assert_reason(xml: &str, expected: NotEligibleReason, control: Control, sheet: usize) {
    match scan_xml(xml) {
        Ok(RangeScanOutcome::NotEligible(actual)) => assert_eq!(
            actual,
            expected,
            "{} disagreed for sheet {}",
            control.name(),
            sheet + 1
        ),
        other => panic!(
            "{} expected {:?} for sheet {}, got {other:?}",
            control.name(),
            expected,
            sheet + 1
        ),
    }
}

/// The eligible selected payload of one controlled sheet: its numeric and
/// text record counts, and a debug rendering of every retained record and
/// its dependency metadata, so two controls can be compared exactly.
struct EligibleSheet {
    numeric: usize,
    text: usize,
    payload: String,
}

fn assert_eligible_producer_sheet(xml: &str, control: Control, sheet: usize) -> EligibleSheet {
    let selected = match scan_xml(xml) {
        Ok(RangeScanOutcome::Eligible(selected)) => selected,
        other => panic!(
            "{} expected sheet {} to become eligible, got {other:?}",
            control.name(),
            sheet + 1
        ),
    };
    assert_eq!(selected.cells.len(), REVIEWED_ROWS * REVIEWED_COLUMNS);
    let numeric = selected
        .cells
        .iter()
        .filter(|record| {
            matches!(
                &record.payload,
                SelectedPayload::Cell(litchi_xlsx::Cell::Value(litchi_xlsx::Value::Number(_)))
            )
        })
        .count();
    let text = selected
        .cells
        .iter()
        .filter(|record| matches!(&record.payload, SelectedPayload::SharedString(_)))
        .count();
    assert_eq!(numeric + text, REVIEWED_ROWS * REVIEWED_COLUMNS);
    EligibleSheet {
        numeric,
        text,
        payload: format!("{selected:?}"),
    }
}

/// Assert the post-`cols` split: the relationship-bearing sheet reaches its
/// `pageSetup` and refuses on structure; the other sheets are eligible.
/// Returns the eligible sheets' payloads (`None` for the refusing sheet).
fn assert_page_setup_split(worksheets: &[Vec<u8>], control: Control) -> Vec<Option<EligibleSheet>> {
    let mut numeric = 0;
    let mut text = 0;
    let mut sheets = Vec::with_capacity(worksheets.len());
    for (sheet, xml) in worksheets.iter().enumerate() {
        let controlled = controlled_xml(xml, control);
        if sheet == RELATIONSHIP_SHEET {
            assert_reason(
                &controlled,
                NotEligibleReason::UnsupportedStructure,
                control,
                sheet,
            );
            sheets.push(None);
        } else {
            let eligible = assert_eligible_producer_sheet(&controlled, control, sheet);
            numeric += eligible.numeric;
            text += eligible.text;
            sheets.push(Some(eligible));
        }
    }
    // The relationship sheet is the all-numeric one; the shared strings live
    // on an eligible sheet.
    assert_eq!(
        numeric,
        REVIEWED_NUMERIC_COUNT - REVIEWED_ROWS * REVIEWED_COLUMNS,
        "{}",
        control.name()
    );
    assert_eq!(text, REVIEWED_TEXT_COUNT, "{}", control.name());
    sheets
}

/// Assert every sheet is eligible and the aggregate is the reviewed
/// 4,096 / 3,686 / 410 contract; return the per-sheet payloads.
fn assert_all_eligible(worksheets: &[Vec<u8>], control: Control) -> Vec<EligibleSheet> {
    let sheets = worksheets
        .iter()
        .enumerate()
        .map(|(sheet, xml)| {
            assert_eligible_producer_sheet(&controlled_xml(xml, control), control, sheet)
        })
        .collect::<Vec<_>>();
    assert_eq!(
        sheets.iter().map(|sheet| sheet.numeric).sum::<usize>(),
        REVIEWED_NUMERIC_COUNT,
        "{}",
        control.name()
    );
    assert_eq!(
        sheets.iter().map(|sheet| sheet.text).sum::<usize>(),
        REVIEWED_TEXT_COUNT,
        "{}",
        control.name()
    );
    sheets
}

#[test]
fn producer_medium_read_reports_the_first_selected_scanner_refusal_per_sheet() {
    let corpus = reviewed_read_corpus();
    let worksheets = worksheet_xmls(&corpus.corpus.archive);
    assert_eq!(worksheets.len(), REVIEWED_SHEETS);

    // The complete generated producer worksheet opens with <dimension>,
    // <sheetViews> and a regular <sheetFormatPr>; a read scan admits the
    // latter two, so the first refusal is the producer's <cols> width
    // metadata, which the scanner defers to the styled materialized read.
    // The scanner's public result is the exact current first refusal, while
    // the archive identity and semantic counts above bind this observation
    // to the reviewed 4 x 32 x 32 fixture.
    for (sheet, xml) in worksheets.iter().enumerate() {
        assert_reason(
            &controlled_xml(xml, Control::CompleteProducer),
            NotEligibleReason::Styles,
            Control::CompleteProducer,
            sheet,
        );
    }
}

#[test]
fn producer_medium_read_marker_removal_controls_reveal_downstream_refusals() {
    let corpus = reviewed_read_corpus();
    let worksheets = worksheet_xmls(&corpus.corpus.archive);

    // Removing the admitted prefix markers does not move the first refusal:
    // <cols> still decides it.
    let staged = [
        (Control::CompleteProducer, NotEligibleReason::Styles),
        (Control::WithoutSheetViews, NotEligibleReason::Styles),
        (
            Control::WithoutSheetViewsAndFormat,
            NotEligibleReason::Styles,
        ),
    ];
    for (control, expected) in staged {
        for (sheet, xml) in worksheets.iter().enumerate() {
            assert_reason(&controlled_xml(xml, control), expected, control, sheet);
        }
    }

    // Removing only <cols> reveals the next refusal with every admitted
    // marker (sheetViews, sheetFormatPr, row spans, pageMargins) still in
    // place: the relationship-bearing third sheet reaches its later
    // <pageSetup> child and refuses on structure, and the three sheets
    // without a pageSetup become eligible.  Stripping the admitted markers
    // as well gives the same split, stage by stage, with identical selected
    // payloads.  This distinguishes the post-sheetData control without
    // claiming anything about the enclosing SourceWorksheet route.
    let with_admitted = assert_page_setup_split(&worksheets, Control::WithoutCols);
    for control in [
        Control::WithoutSheetViewsFormatAndCols,
        Control::WithoutMetadataThroughRowSpans,
        Control::WithoutPageMargins,
    ] {
        let stripped = assert_page_setup_split(&worksheets, control);
        for (sheet, (left, right)) in with_admitted.iter().zip(&stripped).enumerate() {
            assert_eq!(
                left.as_ref().map(|eligible| &eligible.payload),
                right.as_ref().map(|eligible| &eligible.payload),
                "{} changed the selected payload of sheet {}",
                control.name(),
                sheet + 1
            );
        }
    }

    // The admission is checked, not blanket: a <sheetFormatPr> without its
    // defaultRowHeight is one the materialized parser would repair, so the
    // read scan defers it on structure before any sheet reaches sheetData.
    for (sheet, xml) in worksheets.iter().enumerate() {
        assert_reason(
            &controlled_xml(xml, Control::WithoutColsIrregularSheetFormat),
            NotEligibleReason::UnsupportedStructure,
            Control::WithoutColsIrregularSheetFormat,
            sheet,
        );
    }

    // Removing the later pageSetup marker from its one relationship-bearing
    // sheet reaches the same eligible selected payload as the other three
    // sheets, with or without the admitted producer metadata.  The aggregate
    // vectors must still be the reviewed 4,096 / 3,686 / 410 contract, so
    // this remains a corpus control rather than a timing route claim.
    let admitted = assert_all_eligible(&worksheets, Control::WithoutColsAndPageSetup);
    let stripped = assert_all_eligible(&worksheets, Control::WithoutPageMarginsAndPageSetup);
    for (sheet, (left, right)) in admitted.iter().zip(&stripped).enumerate() {
        assert_eq!(
            left.payload,
            right.payload,
            "admitted producer metadata changed the selected payload of sheet {}",
            sheet + 1
        );
    }
}

#[test]
fn producer_malformed_prefix_remains_a_typed_scan_error() {
    let corpus = reviewed_read_corpus();
    let worksheets = worksheet_xmls(&corpus.corpus.archive);
    let xml = String::from_utf8(worksheets[0].clone()).expect("producer worksheet is UTF-8");
    // Corrupt the root start tag before <sheetViews>.  The MCE/XML layer must
    // preserve this typed input failure; the selected scanner must not turn a
    // malformed prefix into a capability result.
    let malformed = xml.replacen("<worksheet ", "<worksheet <", 1);
    assert_ne!(
        malformed, xml,
        "malformed-prefix control did not alter fixture"
    );
    match scan_xml(&malformed) {
        Err(error) => match &error {
            StreamError::Mce {
                error: MceError::Xml(_) | MceError::NonConformant(_),
                ..
            } => {},
            other => panic!("expected typed MCE/XML error, got {other:?}"),
        },
        other => panic!("malformed prefix produced a selected result: {other:?}"),
    }
}
