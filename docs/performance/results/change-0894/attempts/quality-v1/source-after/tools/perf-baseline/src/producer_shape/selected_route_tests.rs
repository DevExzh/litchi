//! Test-only runtime refusal evidence for the generated producer worksheet.
//!
//! These tests call the same public raw selected-range scanner used by the
//! source-backed worksheet route.  They do not observe the surrounding route,
//! time a read, or claim that a selected result was published.  The generated
//! archive stays the reviewed four-sheet producer fixture; the XML controls
//! are bounded copies that remove one admitted metadata family at a time so
//! the selected cell oracle is checked with and without each producer marker.

#![allow(clippy::unwrap_used, reason = "focused producer refusal assertions")]

use std::{
    io::{BufReader, Cursor},
    string::String,
};

use litchi_ooxml_common::mce::{Capabilities, Error as MceError, StreamError, StreamLimits};
use litchi_xlsx::Rect;
use litchi_xlsx::raw::selected_worksheet::{RangeScanOutcome, SelectedPayload, scan_range};
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

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Control {
    CompleteProducer,
    WithoutSheetViews,
    WithoutSheetViewsAndFormat,
    WithoutSheetViewsFormatAndCols,
    WithoutMetadataThroughRowSpans,
    WithoutPageMargins,
    WithoutPageMarginsAndPageSetup,
}

impl Control {
    const fn name(self) -> &'static str {
        match self {
            Self::CompleteProducer => "producer-complete-metadata",
            Self::WithoutSheetViews => "without-sheetViews",
            Self::WithoutSheetViewsAndFormat => "without-sheetViews-and-sheetFormatPr",
            Self::WithoutSheetViewsFormatAndCols => "without-sheetViews-format-and-cols",
            Self::WithoutMetadataThroughRowSpans => "without-metadata-through-row-spans",
            Self::WithoutPageMargins => "without-pageMargins",
            Self::WithoutPageMarginsAndPageSetup => "without-pageMargins-and-pageSetup",
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
    let reader = ArchiveReader::new(archive).expect("producer archive opens");
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

fn controlled_xml(xml: &[u8], control: Control) -> String {
    let original = String::from_utf8(xml.to_vec()).expect("producer worksheet is UTF-8");
    let mut result = original;
    if !matches!(control, Control::CompleteProducer) {
        result = remove_element(&result, "sheetViews");
    }
    if matches!(
        control,
        Control::WithoutSheetViewsAndFormat
            | Control::WithoutSheetViewsFormatAndCols
            | Control::WithoutMetadataThroughRowSpans
            | Control::WithoutPageMargins
            | Control::WithoutPageMarginsAndPageSetup
    ) {
        result = remove_element(&result, "sheetFormatPr");
    }
    if matches!(
        control,
        Control::WithoutSheetViewsFormatAndCols
            | Control::WithoutMetadataThroughRowSpans
            | Control::WithoutPageMargins
            | Control::WithoutPageMarginsAndPageSetup
    ) {
        result = remove_element(&result, "cols");
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
    if matches!(control, Control::WithoutPageMarginsAndPageSetup) && result.contains("<pageSetup") {
        result = remove_element(&result, "pageSetup");
    }
    result
}

fn scan_xml(xml: &str) -> litchi_xlsx::raw::selected_worksheet::StreamResult<RangeScanOutcome> {
    let mut limits = StreamLimits::default();
    limits.processing.max_input_bytes = xml.len();
    limits.processing.max_output_bytes = xml.len();
    let mut input = BufReader::new(Cursor::new(xml.as_bytes()));
    scan_range(&mut input, &Capabilities::default(), &limits, Rect::ALL)
}

fn assert_eligible_producer_sheet(xml: &str, sheet: usize) -> (usize, usize) {
    let Ok(RangeScanOutcome::Eligible(selected)) = scan_xml(xml) else {
        panic!(
            "sheet {} did not become eligible after removing its blockers",
            sheet + 1
        );
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
    (numeric, text)
}

#[test]
fn producer_medium_read_admits_complete_metadata_and_preserves_the_cell_oracle() {
    let corpus = reviewed_read_corpus();
    let worksheets = worksheet_xmls(&corpus.corpus.archive);
    assert_eq!(worksheets.len(), REVIEWED_SHEETS);

    // The candidate metadata profile admits every producer worksheet, including
    // the shared-string and worksheet-relationship roles.  These direct scan
    // results remain separate from the enclosing SourceWorksheet route: the
    // archive identity and semantic counts bind the observation to the reviewed
    // 4 x 32 x 32 fixture.
    let mut numeric = 0;
    let mut text = 0;
    for (sheet, xml) in worksheets.iter().enumerate() {
        let (sheet_numeric, sheet_text) =
            assert_eligible_producer_sheet(&controlled_xml(xml, Control::CompleteProducer), sheet);
        numeric += sheet_numeric;
        text += sheet_text;
    }
    assert_eq!(numeric, REVIEWED_NUMERIC_COUNT);
    assert_eq!(text, REVIEWED_TEXT_COUNT);
}

#[test]
fn producer_medium_read_metadata_removal_controls_preserve_eligibility() {
    let corpus = reviewed_read_corpus();
    let worksheets = worksheet_xmls(&corpus.corpus.archive);

    // Each control removes one admitted metadata family.  The candidate must
    // preserve the selected physical records regardless of which optional
    // producer metadata is absent; the same exact aggregate cell oracle is
    // asserted for every control rather than only the fully-populated shape.
    let controls = [
        Control::CompleteProducer,
        Control::WithoutSheetViews,
        Control::WithoutSheetViewsAndFormat,
        Control::WithoutSheetViewsFormatAndCols,
        Control::WithoutMetadataThroughRowSpans,
        Control::WithoutPageMargins,
        Control::WithoutPageMarginsAndPageSetup,
    ];
    for control in controls {
        let mut numeric = 0;
        let mut text = 0;
        for (sheet, xml) in worksheets.iter().enumerate() {
            let (sheet_numeric, sheet_text) =
                assert_eligible_producer_sheet(&controlled_xml(xml, control), sheet);
            numeric += sheet_numeric;
            text += sheet_text;
        }
        assert_eq!(
            numeric,
            REVIEWED_NUMERIC_COUNT,
            "{} numeric",
            control.name()
        );
        assert_eq!(text, REVIEWED_TEXT_COUNT, "{} text", control.name());
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
