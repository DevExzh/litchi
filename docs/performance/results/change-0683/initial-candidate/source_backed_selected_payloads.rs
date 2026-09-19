#![allow(clippy::unwrap_used, reason = "focused selected-scanner assertions")]

//! Public source-backed witnesses for the compact selected-record payload.
//!
//! The selected scanner keeps one of two mutually exclusive payloads for each
//! physical record.  These tests intentionally exercise every public semantic
//! form before the scanner's final publication fence, including repeated and
//! unselected shared-string dependencies.  They stay at the public
//! source-backed boundary so a representation change cannot accidentally alter
//! the values or the refusal-before-callback contract.

use std::io::{self, BufReader, Cursor};
use std::sync::Arc;
use std::sync::atomic::{AtomicUsize, Ordering};

use litchi_core::{ReadAt, SourceVersion};
use litchi_ooxml_common::mce::{Capabilities, StreamLimits};
use litchi_opc::constants::content_type as ct;
use litchi_xlsx::raw::selected_worksheet::{RangeScanOutcome, SelectedPayload, scan_range};
use litchi_xlsx::{Cell, Error, ErrorValue, Number, Rect, SourceBackedWorkbook, SourceCell, Value};
use soapberry_zip::office::StreamingArchiveWriter;

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const PACKAGE_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WORKSHEET_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const SHARED_STRINGS_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const STYLES_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

fn dependency_xlsx(worksheet: &str) -> Vec<u8> {
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored(
            "[Content_Types].xml",
            format!(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="{}"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="{}"/><Override PartName="/xl/sharedStrings.xml" ContentType="{}"/><Override PartName="/xl/styles.xml" ContentType="{}"/></Types>"#,
                ct::SML_SHEET_MAIN,
                ct::SML_WORKSHEET,
                ct::SML_SHARED_STRINGS,
                ct::SML_STYLES,
            )
            .as_bytes(),
        )
        .unwrap();
    writer
        .write_stored(
            "_rels/.rels",
            format!(
                r#"<Relationships xmlns="{PACKAGE_REL}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
            )
            .as_bytes(),
        )
        .unwrap();
    writer
        .write_stored(
            "xl/workbook.xml",
            format!(
                r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#
            )
            .as_bytes(),
        )
        .unwrap();
    writer
        .write_stored(
            "xl/_rels/workbook.xml.rels",
            format!(
                r#"<Relationships xmlns="{PACKAGE_REL}"><Relationship Id="rId1" Type="{WORKSHEET_REL}" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{SHARED_STRINGS_REL}" Target="sharedStrings.xml"/><Relationship Id="rId3" Type="{STYLES_REL}" Target="styles.xml"/></Relationships>"#
            )
            .as_bytes(),
        )
        .unwrap();
    writer
        .write_stored("xl/worksheets/sheet1.xml", worksheet.as_bytes())
        .unwrap();
    writer
        .write_stored(
            "xl/sharedStrings.xml",
            format!(
                r#"<sst xmlns="{SML}" count="3" uniqueCount="3"><si><t>unused-zero</t></si><si><t>selected-shared</t></si><si><t>unselected-max</t></si></sst>"#
            )
            .as_bytes(),
        )
        .unwrap();
    writer
        .write_stored(
            "xl/styles.xml",
            format!(r#"<styleSheet xmlns="{SML}"><cellXfs count="2"><xf/><xf numFmtId="1"/></cellXfs></styleSheet>"#)
                .as_bytes(),
        )
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

fn semantic_fixture() -> String {
    format!(
        r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1" s="1"><v>42</v></c><c r="B1" t="b"><v>1</v></c><c r="C1" t="e"><v>#N/A</v></c><c r="D1" t="inlineStr"><is><t>inline text</t></is></c><c r="E1"><f>A1+1</f><v>43</v></c><c r="F1" t="s"><v>1</v></c><c r="G1" t="s"><v>1</v></c><c r="H1"/></row><row r="2"><c r="I2" t="s"><v>2</v></c></row></sheetData></worksheet>"#
    )
}

fn assert_semantic_variants(cells: &[SourceCell]) {
    assert_eq!(
        cells
            .iter()
            .map(|cell| cell.address.a1())
            .collect::<Vec<_>>(),
        ["A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1"]
    );
    assert!(matches!(
        &cells[0].cell,
        Cell::Value(Value::Number(number)) if number.as_str() == "42"
    ));
    assert!(matches!(&cells[1].cell, Cell::Value(Value::Bool(true))));
    assert!(matches!(
        &cells[2].cell,
        Cell::Value(Value::Error(ErrorValue::NotAvailable))
    ));
    assert!(matches!(
        &cells[3].cell,
        Cell::Value(Value::Text(text)) if text.as_str() == "inline text"
    ));
    let Cell::Formula(formula) = &cells[4].cell else {
        panic!("formula payload was not retained as a formula")
    };
    assert_eq!(formula.text(), "A1+1");
    assert_eq!(
        formula.cached().map(|cache| cache.value()),
        Some(&Value::Number(Number::new("43").unwrap()))
    );
    for cell in &cells[5..=6] {
        assert!(matches!(
            &cell.cell,
            Cell::Value(Value::Text(text)) if text.as_str() == "selected-shared"
        ));
    }
    assert!(matches!(&cells[7].cell, Cell::Empty));
}

#[test]
fn raw_selected_payload_tags_cover_values_formula_strings_and_dependencies() {
    let worksheet = semantic_fixture();
    let mut input = BufReader::new(Cursor::new(worksheet.as_bytes()));
    let outcome = scan_range(
        &mut input,
        &Capabilities::default(),
        &StreamLimits::default(),
        Rect::from_a1("A1:H1").unwrap(),
    )
    .unwrap();
    let RangeScanOutcome::Eligible(selected) = outcome else {
        panic!("semantic fixture unexpectedly fell back from the selected scanner")
    };

    assert_eq!(selected.cells.len(), 8);
    assert!(matches!(
        &selected.cells[0].payload,
        SelectedPayload::Cell(Cell::Value(Value::Number(number))) if number.as_str() == "42"
    ));
    assert!(matches!(
        &selected.cells[1].payload,
        SelectedPayload::Cell(Cell::Value(Value::Bool(true)))
    ));
    assert!(matches!(
        &selected.cells[2].payload,
        SelectedPayload::Cell(Cell::Value(Value::Error(ErrorValue::NotAvailable)))
    ));
    assert!(matches!(
        &selected.cells[3].payload,
        SelectedPayload::Cell(Cell::Value(Value::Text(text))) if text.as_str() == "inline text"
    ));
    assert!(matches!(
        &selected.cells[4].payload,
        SelectedPayload::Cell(Cell::Formula(formula)) if formula.text() == "A1+1"
    ));
    assert!(matches!(
        &selected.cells[5].payload,
        SelectedPayload::SharedString(1)
    ));
    assert!(matches!(
        &selected.cells[6].payload,
        SelectedPayload::SharedString(1)
    ));
    assert!(matches!(
        &selected.cells[7].payload,
        SelectedPayload::Cell(Cell::Empty)
    ));
    assert_eq!(selected.dependencies.max_shared_string_index, Some(2));
    assert_eq!(selected.dependencies.max_direct_style_index, Some(1));
    assert_eq!(selected.dependencies.target_shared_string_index, None);
}

#[test]
fn selected_tagged_payloads_match_cold_warm_and_visit_routes() {
    let source = Arc::new(ProbeSource::new(dependency_xlsx(&semantic_fixture())));
    let workbook = SourceBackedWorkbook::from_read_at(source).unwrap();
    let sheet = workbook.sheet("Sheet1").unwrap().unwrap();

    let cold = sheet.cells("A1:H1").unwrap();
    assert_semantic_variants(&cold);

    // A second selected read keeps the worksheet store cold but exercises the
    // warmed source/dependency path.  It must agree byte-for-byte at the
    // public semantic boundary.
    let warm = sheet.cells("A1:H1").unwrap();
    assert_eq!(warm, cold);

    let mut visited = Vec::new();
    let count = sheet
        .visit_cells("A1:H1", |address, cell| {
            visited.push(SourceCell {
                address,
                cell: cell.clone(),
            });
            Ok(())
        })
        .unwrap();
    assert_eq!(count, cold.len());
    assert_eq!(visited, cold);
    assert_semantic_variants(&visited);
}

#[test]
fn repeated_shared_strings_validate_unselected_dependency_maximum() {
    let worksheet = format!(
        r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1" t="s"><v>1</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row></sheetData></worksheet>"#
    );
    let workbook =
        SourceBackedWorkbook::from_read_at(Arc::new(ProbeSource::new(dependency_xlsx(&worksheet))))
            .unwrap();
    let sheet = workbook.sheet("Sheet1").unwrap().unwrap();

    let selected = sheet.cells("A1:B1").unwrap();
    assert_eq!(selected.len(), 2);
    for cell in selected {
        assert!(matches!(
            cell.cell,
            Cell::Value(Value::Text(ref text)) if text.as_str() == "selected-shared"
        ));
    }

    // C1 is outside the requested rectangle but still contributes to the
    // scanner's dependency maximum.  The selected records must continue to
    // use their narrow physical indexes and must not accidentally resolve the
    // unselected table entry.
    let selected_again = sheet.cells("A1:B1").unwrap();
    assert_eq!(selected_again[0].cell, selected_again[1].cell);
    assert_ne!(
        selected_again[0].cell,
        Cell::Value(Value::Text("unselected-max".into()))
    );
}

#[test]
fn late_refusal_after_every_payload_arm_produces_zero_callbacks() {
    let worksheet = format!(
        r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1"><v>42</v></c><c r="B1" t="b"><v>1</v></c><c r="C1" t="e"><v>#N/A</v></c><c r="D1" t="inlineStr"><is><t>inline text</t></is></c><c r="E1"><f>A1+1</f><v>43</v></c><c r="F1" t="s"><v>1</v></c><c r="G1"/></row><row r="9"><c r="A9" t="inlineStr"><is><t>late</t></is><v>3</v></c></row></sheetData></worksheet>"#
    );
    let bytes = dependency_xlsx(&worksheet);

    let cells_workbook =
        SourceBackedWorkbook::from_read_at(Arc::new(ProbeSource::new(bytes.clone()))).unwrap();
    let cells_error = cells_workbook
        .sheet("Sheet1")
        .unwrap()
        .unwrap()
        .cells("A1:G1")
        .unwrap_err();
    assert!(matches!(cells_error, Error::Invalid(_)));
    let cells_error_text = cells_error.to_string();
    assert!(cells_error_text.contains("both inline text and a value"));

    let visit_workbook =
        SourceBackedWorkbook::from_read_at(Arc::new(ProbeSource::new(bytes))).unwrap();
    let sheet = visit_workbook.sheet("Sheet1").unwrap().unwrap();
    let callbacks = Arc::new(AtomicUsize::new(0));
    let callback_counter = Arc::clone(&callbacks);
    let visit_error = sheet
        .visit_cells("A1:G1", |_address, _cell| {
            callback_counter.fetch_add(1, Ordering::SeqCst);
            Ok(())
        })
        .unwrap_err();
    assert!(matches!(visit_error, Error::Invalid(_)));
    assert_eq!(callbacks.load(Ordering::SeqCst), 0);
    assert_eq!(visit_error.to_string(), cells_error_text);
}

#[test]
fn selected_callback_can_reenter_the_workbook_after_scan_publication() {
    let source = Arc::new(ProbeSource::new(dependency_xlsx(&format!(
        r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>"#
    ))));
    let reader: Arc<dyn ReadAt> = source.clone();
    let workbook = SourceBackedWorkbook::from_read_at(reader).unwrap();
    let sheet = workbook.sheet("Sheet1").unwrap().unwrap();
    let reads_before = source.reads();
    let mut reentered = false;

    let result = sheet.visit_cells("A1:B1", |_address, _cell| {
        source.begin_callback();
        source.begin_nested_query();
        let nested = sheet.cell("B1");
        source.end_nested_query();
        let nested = nested?;
        assert!(matches!(
            nested,
            litchi_xlsx::SourceCellView::Stored(Cell::Value(Value::Number(ref value)))
                if value.as_str() == "2"
        ));
        reentered = true;
        Ok(())
    });
    source.end_callback();
    let count = result.unwrap();

    assert_eq!(count, 2);
    assert!(reentered);
    assert_eq!(
        source.reads_during_callback(),
        0,
        "the worksheet reader must be gone before a callback can re-enter"
    );
    assert!(
        source.reads() > reads_before,
        "the nested query should be able to use the positional source"
    );
}

#[derive(Debug)]
struct ProbeSource {
    bytes: Vec<u8>,
    reads: AtomicUsize,
    callback_active: AtomicUsize,
    nested_query: AtomicUsize,
    reads_during_callback: AtomicUsize,
}

impl ProbeSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            reads: AtomicUsize::new(0),
            callback_active: AtomicUsize::new(0),
            nested_query: AtomicUsize::new(0),
            reads_during_callback: AtomicUsize::new(0),
        }
    }

    fn reads(&self) -> usize {
        self.reads.load(Ordering::SeqCst)
    }

    fn begin_callback(&self) {
        self.callback_active.store(1, Ordering::SeqCst);
    }

    fn end_callback(&self) {
        self.callback_active.store(0, Ordering::SeqCst);
        self.nested_query.store(0, Ordering::SeqCst);
    }

    fn begin_nested_query(&self) {
        self.nested_query.store(1, Ordering::SeqCst);
    }

    fn end_nested_query(&self) {
        self.nested_query.store(0, Ordering::SeqCst);
    }

    fn reads_during_callback(&self) -> usize {
        self.reads_during_callback.load(Ordering::SeqCst)
    }
}

impl ReadAt for ProbeSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.reads.fetch_add(1, Ordering::SeqCst);
        if self.callback_active.load(Ordering::SeqCst) != 0
            && self.nested_query.load(Ordering::SeqCst) == 0
        {
            self.reads_during_callback.fetch_add(1, Ordering::SeqCst);
        }
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(6_830, 0))
    }
}
