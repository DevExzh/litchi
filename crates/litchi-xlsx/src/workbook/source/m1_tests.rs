//! M1 leaf requests L-XLSX-R1, R2 and R6 on the source-backed owner: the
//! retained parsed grid, its delegates, the blocks outside `sheetData`, and
//! the early-stop range visitor.

use std::fmt::Debug;
use std::io;
use std::ops::ControlFlow;
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering};

use litchi_core::{ReadAt, SourceVersion};
use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcError, SourceLimits};
use soapberry_zip::office::StreamingArchiveWriter;

use super::values::limit_retained_weight;
use super::{SourceBackedWorkbook, SourceCellView};
use crate::{Cell, Error, Value};

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

/// A positional source that counts physical reads and can change revision.
struct ProbeSource {
    bytes: Vec<u8>,
    reads: AtomicUsize,
    revision: AtomicU64,
}

impl ProbeSource {
    fn new(bytes: Vec<u8>) -> Arc<Self> {
        Arc::new(Self {
            bytes,
            reads: AtomicUsize::new(0),
            revision: AtomicU64::new(0),
        })
    }

    fn reads(&self) -> usize {
        self.reads.load(Ordering::SeqCst)
    }

    fn changed(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }
}

impl ReadAt for ProbeSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.reads.fetch_add(1, Ordering::SeqCst);
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
        Ok(SourceVersion::new(
            10_011,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

/// One worksheet of a generated workbook.
struct Sheet {
    name: &'static str,
    xml: String,
    relationships: Option<String>,
}

fn sheet(name: &'static str, xml: String) -> Sheet {
    Sheet {
        name,
        xml,
        relationships: None,
    }
}

const SHARED_STRINGS: &str = r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2"><si><t>shared zero</t></si><si><t>shared one</t></si></sst>"#;
const STYLES: &str = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cellXfs count="2"><xf/><xf numFmtId="2"/></cellXfs><dxfs count="1"><dxf><font><b/></font></dxf></dxfs></styleSheet>"#;

fn workbook_bytes(sheets: &[Sheet], defined_names: &str) -> Vec<u8> {
    let mut overrides = String::new();
    let mut entries = String::new();
    let mut relationships = String::new();
    for (index, sheet) in sheets.iter().enumerate() {
        let number = index + 1;
        overrides.push_str(&format!(
            r#"<Override PartName="/xl/worksheets/sheet{number}.xml" ContentType="{}"/>"#,
            ct::SML_WORKSHEET
        ));
        entries.push_str(&format!(
            r#"<sheet name="{}" sheetId="{number}" r:id="rId{number}"/>"#,
            sheet.name
        ));
        relationships.push_str(&format!(
            r#"<Relationship Id="rId{number}" Type="{REL}/worksheet" Target="worksheets/sheet{number}.xml"/>"#
        ));
    }
    let strings_id = sheets.len() + 1;
    let styles_id = sheets.len() + 2;
    relationships.push_str(&format!(
        r#"<Relationship Id="rId{strings_id}" Type="{REL}/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId{styles_id}" Type="{REL}/styles" Target="styles.xml"/>"#
    ));

    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    writer
        .write_stored(
            "[Content_Types].xml",
            format!(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="{}"/>{overrides}<Override PartName="/xl/sharedStrings.xml" ContentType="{}"/><Override PartName="/xl/styles.xml" ContentType="{}"/></Types>"#,
                ct::SML_SHEET_MAIN,
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
                r#"<workbook xmlns="{SML}" xmlns:r="{REL}"><sheets>{entries}</sheets>{defined_names}</workbook>"#
            )
            .as_bytes(),
        )
        .unwrap();
    writer
        .write_stored(
            "xl/_rels/workbook.xml.rels",
            format!(r#"<Relationships xmlns="{PACKAGE_REL}">{relationships}</Relationships>"#)
                .as_bytes(),
        )
        .unwrap();
    for (index, sheet) in sheets.iter().enumerate() {
        let number = index + 1;
        writer
            .write_deflated(
                &format!("xl/worksheets/sheet{number}.xml"),
                sheet.xml.as_bytes(),
            )
            .unwrap();
        if let Some(relationships) = &sheet.relationships {
            writer
                .write_stored(
                    &format!("xl/worksheets/_rels/sheet{number}.xml.rels"),
                    relationships.as_bytes(),
                )
                .unwrap();
        }
    }
    writer
        .write_stored("xl/sharedStrings.xml", SHARED_STRINGS.as_bytes())
        .unwrap();
    writer
        .write_stored("xl/styles.xml", STYLES.as_bytes())
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

/// A worksheet with every modeled grid record and every block read.
fn rich_sheet() -> Sheet {
    Sheet {
        name: "Data",
        xml: format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="{SML}" xmlns:r="{REL}"><dimension ref="A1:D6"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="B3" sqref="B3"/></sheetView></sheetViews><sheetFormatPr baseColWidth="10" defaultRowHeight="15"/><cols><col min="2" max="3" width="12.5" customWidth="1" style="1"/></cols><sheetData><row r="1" spans="1:4" ht="20" customHeight="1"><c r="A1" t="s"><v>0</v></c><c r="B1" s="1"><v>2.5</v></c><c r="C1"><f>B1*2</f><v>5</v></c></row><row r="3" hidden="1"><c r="A3" t="inlineStr"><is><t>inline</t></is></c><c r="D3" t="b"><v>1</v></c></row><row r="6"><c r="B6" t="s"><v>1</v></c></row></sheetData><autoFilter ref="A1:D6"/><mergeCells count="1"><mergeCell ref="B3:C4"/></mergeCells><conditionalFormatting sqref="B1:B6"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation type="whole" allowBlank="1" showErrorMessage="1" sqref="C1:C6"><formula1>0</formula1><formula2>10</formula2></dataValidation></dataValidations><hyperlinks><hyperlink ref="A1" r:id="rId1"/><hyperlink ref="B6" location="'Other'!A1" display="jump"/></hyperlinks><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>"#
        ),
        relationships: Some(format!(
            r#"<Relationships xmlns="{PACKAGE_REL}"><Relationship Id="rId1" Type="{REL}/hyperlink" Target="https://example.com/" TargetMode="External"/></Relationships>"#
        )),
    }
}

/// A plain numeric worksheet of `rows` by `columns` cells.
fn numeric_sheet(name: &'static str, rows: usize, columns: usize) -> Sheet {
    let mut xml = format!(r#"<worksheet xmlns="{SML}"><sheetData>"#);
    for row in 1..=rows {
        xml.push_str(&format!(r#"<row r="{row}">"#));
        for column in 0..columns {
            let letter = char::from(b'A' + u8::try_from(column).unwrap());
            xml.push_str(&format!(
                r#"<c r="{letter}{row}"><v>{}</v></c>"#,
                row * 100 + column
            ));
        }
        xml.push_str("</row>");
    }
    xml.push_str("</sheetData></worksheet>");
    sheet(name, xml)
}

fn rich_workbook() -> Vec<u8> {
    workbook_bytes(
        &[rich_sheet(), numeric_sheet("Other", 4, 4)],
        r#"<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Data!$A$1:$D$6</definedName><definedName name="Total" comment="sum">Data!$C$1</definedName></definedNames>"#,
    )
}

fn open(bytes: &[u8]) -> (Arc<ProbeSource>, SourceBackedWorkbook) {
    let source = ProbeSource::new(bytes.to_vec());
    let reader: Arc<dyn ReadAt> = source.clone();
    (source, SourceBackedWorkbook::from_read_at(reader).unwrap())
}

fn debug<T: Debug>(value: T) -> String {
    format!("{value:?}")
}

fn number(view: &SourceCellView) -> &str {
    match view {
        SourceCellView::Stored(Cell::Value(Value::Number(value))) => value.as_str(),
        other => panic!("expected a stored number, got {other:?}"),
    }
}

#[test]
fn first_touch_parses_once_and_later_reads_are_lookups() {
    let (source, workbook) = open(&rich_workbook());
    let sheet = workbook.sheet("Other").unwrap().unwrap();
    assert!(!sheet.store_resident());
    sheet.grid().unwrap();

    assert_eq!(number(&sheet.cell("B2").unwrap()), "201");
    assert!(sheet.store_resident());
    let after_first = source.reads();
    assert_eq!(workbook.semantic_cache_diagnostics().builds, 1);

    // Repeated cell, range, visit, extent and row reads are lookups in the
    // pinned grid: no further physical read and no second parse.
    for _ in 0..3 {
        assert_eq!(number(&sheet.cell("D4").unwrap()), "403");
        assert_eq!(sheet.cells("A1:D4").unwrap().len(), 16);
        assert_eq!(sheet.visit_cells("B1:B4", |_, _| Ok(())).unwrap(), 4);
        assert_eq!(
            sheet.stored_extent().unwrap().map(|extent| extent.a1()),
            Some("A1:D4".to_owned())
        );
        assert!(sheet.row(2).unwrap().stored());
    }
    assert_eq!(source.reads(), after_first);
    let diagnostics = workbook.semantic_cache_diagnostics();
    assert_eq!(diagnostics.builds, 1);
    assert_eq!(diagnostics.retained_entries, 1);
    assert_eq!(
        diagnostics.clean_entries, 0,
        "the live handle pins the grid"
    );
}

#[test]
fn fresh_handles_share_one_retained_grid() {
    let (source, workbook) = open(&rich_workbook());
    let first = workbook.sheet("Other").unwrap().unwrap().grid().unwrap();
    let after_first = source.reads();

    // A new handle per read, as a resolver building one node per lookup does.
    for _ in 0..3 {
        let sheet = workbook.sheet("Other").unwrap().unwrap();
        assert_eq!(number(&sheet.cell("A1").unwrap()), "100");
        assert!(sheet.grid().unwrap().same_grid(&first));
    }
    assert_eq!(source.reads(), after_first);
    let diagnostics = workbook.semantic_cache_diagnostics();
    assert_eq!(diagnostics.builds, 1);
    assert!(diagnostics.hits >= 3);
    assert_eq!(diagnostics.misses, 1);
}

#[test]
fn concurrent_first_reads_share_one_build() {
    let (_source, workbook) = open(&rich_workbook());
    std::thread::scope(|scope| {
        for _ in 0..8 {
            scope.spawn(|| {
                let sheet = workbook.sheet("Other").unwrap().unwrap();
                sheet.grid().unwrap();
                assert_eq!(number(&sheet.cell("C3").unwrap()), "302");
            });
        }
    });
    let diagnostics = workbook.semantic_cache_diagnostics();
    assert_eq!(diagnostics.builds, 1);
    assert_eq!(diagnostics.retained_entries, 1);
    assert_eq!(diagnostics.clean_entries, 1, "every handle was dropped");
}

#[test]
fn trimming_releases_only_clean_values() {
    let (source, workbook) = open(&rich_workbook());
    let pinned = workbook.sheet("Data").unwrap().unwrap();
    pinned.extents().unwrap();
    workbook.sheet("Other").unwrap().unwrap().grid().unwrap();
    let before = workbook.semantic_cache_diagnostics();
    assert_eq!(before.retained_entries, 2);
    assert_eq!(before.clean_entries, 1);

    let released = workbook.trim_semantic_cache();
    assert_eq!(released, before.clean_weight);
    let after = workbook.semantic_cache_diagnostics();
    assert_eq!(after.retained_entries, 1);
    assert_eq!(after.clean_entries, 0);
    assert_eq!(after.clean_evictions, 1);

    // The pinned grid still answers without a read; the released one is
    // parsed again from the same snapshot, with the same value.
    let reads = source.reads();
    assert!(pinned.merges().unwrap().next().is_some());
    assert_eq!(source.reads(), reads);
    assert_eq!(workbook.semantic_cache_diagnostics().builds, 2);
    let other = workbook.sheet("Other").unwrap().unwrap();
    other.grid().unwrap();
    assert_eq!(number(&other.cell("A1").unwrap()), "100");
    assert_eq!(workbook.semantic_cache_diagnostics().builds, 3);

    drop(pinned);
    drop(other);
    assert!(workbook.trim_semantic_cache() > 0);
    assert_eq!(workbook.semantic_cache_diagnostics().retained_entries, 0);
}

#[test]
fn least_recently_used_clean_grid_is_evicted_first() {
    let bytes = workbook_bytes(
        &[
            numeric_sheet("A", 8, 4),
            numeric_sheet("B", 8, 4),
            numeric_sheet("C", 8, 4),
        ],
        "",
    );
    let weight = {
        let (_source, workbook) = open(&bytes);
        let grid = workbook.sheet("A").unwrap().unwrap().grid().unwrap();
        grid.retained_weight()
    };
    // Room for exactly one grid and its modeled entry overhead.
    let _limit = limit_retained_weight(weight + 64);
    let (_source, workbook) = open(&bytes);

    workbook.sheet("A").unwrap().unwrap().grid().unwrap();
    workbook.sheet("B").unwrap().unwrap().grid().unwrap();
    let diagnostics = workbook.semantic_cache_diagnostics();
    assert_eq!(diagnostics.retained_entries, 1);
    assert_eq!(diagnostics.clean_evictions, 1);
    assert_eq!(diagnostics.pinned_evictions, 0);
    assert!(diagnostics.retained_weight <= diagnostics.max_weight);

    // When every retained grid is pinned, the least-recently-used one is
    // released from the cache while its handle keeps it.
    let pinned = workbook.sheet("B").unwrap().unwrap();
    pinned.cell("A1").unwrap();
    workbook.sheet("C").unwrap().unwrap().grid().unwrap();
    let diagnostics = workbook.semantic_cache_diagnostics();
    assert_eq!(diagnostics.retained_entries, 1);
    assert_eq!(diagnostics.pinned_evictions, 1);
    assert_eq!(number(&pinned.cell("D8").unwrap()), "803");
}

#[test]
fn oversized_grid_is_pinned_by_its_handle_but_not_retained() {
    let bytes = workbook_bytes(&[numeric_sheet("Big", 20, 5)], "");
    let _limit = limit_retained_weight(1024);
    let (source, workbook) = open(&bytes);

    let sheet = workbook.sheet("Big").unwrap().unwrap();
    sheet.grid().unwrap();
    assert_eq!(number(&sheet.cell("E20").unwrap()), "2004");
    let diagnostics = workbook.semantic_cache_diagnostics();
    assert_eq!(diagnostics.oversized_bypasses, 1);
    assert_eq!(diagnostics.retained_entries, 0);
    assert!(sheet.grid().unwrap().retained_weight() > 1024);

    // The reading handle keeps its grid; another handle parses again.
    let reads = source.reads();
    assert_eq!(number(&sheet.cell("A1").unwrap()), "100");
    assert_eq!(source.reads(), reads);
    assert_eq!(workbook.semantic_cache_diagnostics().builds, 1);
    let fresh = workbook.sheet("Big").unwrap().unwrap();
    fresh.grid().unwrap();
    assert_eq!(number(&fresh.cell("A1").unwrap()), "100");
    assert_eq!(workbook.semantic_cache_diagnostics().builds, 2);
    assert!(!fresh.grid().unwrap().same_grid(&sheet.grid().unwrap()));
}

#[test]
fn failed_build_publishes_nothing_and_is_retried() {
    let bytes = workbook_bytes(
        &[sheet(
            "Broken",
            format!(r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1"><v>7</v></c>"#),
        )],
        "",
    );
    let (_source, workbook) = open(&bytes);
    let sheet = workbook.sheet("Broken").unwrap().unwrap();
    let first = sheet.grid().unwrap_err().to_string();
    assert!(!sheet.store_resident());
    // The refusal is recomputed from the snapshot on every read; it is never
    // retained as a value.
    assert_eq!(sheet.grid().unwrap_err().to_string(), first);
    assert_eq!(
        workbook
            .sheet("Broken")
            .unwrap()
            .unwrap()
            .grid()
            .unwrap_err()
            .to_string(),
        first
    );
    assert!(!sheet.store_resident());
    let diagnostics = workbook.semantic_cache_diagnostics();
    assert_eq!(diagnostics.builds, 0);
    assert_eq!(diagnostics.retained_entries, 0);
    assert!(workbook.tolerance_report().is_clean());
}

#[test]
fn retained_grid_still_observes_source_change_and_cancellation() {
    let (source, workbook) = open(&rich_workbook());
    workbook.sheet("Other").unwrap().unwrap().grid().unwrap();
    let pinned = workbook.sheet("Other").unwrap().unwrap();
    pinned.cell("A1").unwrap();
    source.changed();

    for result in [
        pinned.cell("A1").map(|_| ()),
        workbook
            .sheet("Other")
            .unwrap()
            .unwrap()
            .cell("A1")
            .map(|_| ()),
        pinned.grid().map(|_| ()),
        pinned.rows().map(|_| ()),
        pinned.data_validations().map(|_| ()),
    ] {
        assert!(
            matches!(result, Err(Error::Package(OpcError::SourceChanged { .. }))),
            "{result:?}"
        );
    }
}

#[test]
fn tolerance_report_survives_eviction() {
    // Style 9 does not exist; the read repairs it to the default format.
    let bytes = workbook_bytes(
        &[sheet(
            "Stale",
            format!(
                r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1" s="9"><v>1</v></c></row></sheetData></worksheet>"#
            ),
        )],
        "",
    );
    let (_source, workbook) = open(&bytes);
    let sheet = workbook.sheet("Stale").unwrap().unwrap();
    assert_eq!(sheet.tolerance_report().unwrap().total(), 1);
    drop(sheet);
    assert!(workbook.trim_semantic_cache() > 0);
    assert_eq!(workbook.semantic_cache_diagnostics().retained_entries, 0);
    let report = workbook.tolerance_report();
    assert_eq!(report.total(), 1);
    assert_eq!(report.defects()[0].sheet, Some(0));
}

#[test]
fn grid_delegates_match_the_eager_owner() {
    let bytes = rich_workbook();
    let eager = crate::Package::from_bytes(&bytes).unwrap();
    let expected = eager.sheet("Data").unwrap().unwrap();
    let (_source, workbook) = open(&bytes);
    let sheet = workbook.sheet("Data").unwrap().unwrap();

    assert_eq!(
        debug(sheet.rows().unwrap().collect::<Vec<_>>()),
        debug(expected.rows().unwrap().collect::<Vec<_>>())
    );
    for row in [0u32, 2, 4, 5] {
        assert_eq!(
            debug(sheet.row(row).unwrap()),
            debug(expected.row(row).unwrap())
        );
    }
    assert_eq!(
        debug(sheet.columns().unwrap().collect::<Vec<_>>()),
        debug(expected.columns().unwrap().collect::<Vec<_>>())
    );
    for column in ["A", "B", "C", "D"] {
        assert_eq!(
            debug(sheet.column(column).unwrap()),
            debug(expected.column(column).unwrap())
        );
    }
    assert_eq!(
        sheet.merges().unwrap().collect::<Vec<_>>(),
        expected.merges().unwrap().collect::<Vec<_>>()
    );
    assert_eq!(
        debug(sheet.defaults().unwrap()),
        debug(expected.defaults().unwrap())
    );
    assert_eq!(sheet.extents().unwrap(), expected.extents().unwrap());

    let grid = sheet.grid().unwrap();
    assert_eq!(grid.extents(), expected.extents().unwrap());
    assert_eq!(grid.stored_cell_count(), 6);
    assert_eq!(
        debug(grid.cells("A1:D6").unwrap().collect::<Vec<_>>()),
        debug(expected.cells("A1:D6").unwrap().collect::<Vec<_>>())
    );
    for at in ["A1", "B1", "C1", "C3", "B4", "Z99"] {
        assert_eq!(
            debug(grid.cell(at).unwrap()),
            debug(expected.cell(at).unwrap())
        );
    }
    assert_eq!(debug(grid.row(0).unwrap()), debug(expected.row(0).unwrap()));
    assert_eq!(
        debug(grid.column("B").unwrap()),
        debug(expected.column("B").unwrap())
    );
    assert_eq!(debug(grid.defaults()), debug(expected.defaults().unwrap()));
    assert!(grid.tolerance_report().is_clean());
}

#[test]
fn grid_outlives_its_worksheet_and_workbook() {
    let (_source, workbook) = open(&rich_workbook());
    let grid = workbook.sheet("Other").unwrap().unwrap().grid().unwrap();
    drop(workbook);
    assert!(matches!(
        grid.cell("C2").unwrap(),
        crate::cell::View::Stored(Cell::Value(Value::Number(value))) if value.as_str() == "202"
    ));
    assert_eq!(grid.rows().len(), 4);
    assert!(format!("{grid:?}").starts_with("Grid"));
}

#[test]
fn defined_names_match_the_eager_owner() {
    let bytes = rich_workbook();
    let eager = crate::Package::from_bytes(&bytes).unwrap();
    let (source, workbook) = open(&bytes);
    let reads = source.reads();
    assert_eq!(workbook.defined_names(), eager.defined_names());
    assert_eq!(workbook.defined_names().len(), 2);
    assert_eq!(workbook.defined_names()[1].name, "Total");
    assert_eq!(source.reads(), reads, "names are retained from open");
}

#[test]
fn deferred_views_reuse_the_lineage_catalog_policy_and_reject_mismatch() {
    let bytes = rich_workbook();
    let (_, workbook) = open(&bytes);
    let proof = workbook.shared_catalog_proof();

    let mut edit = workbook.edit_tab_state().unwrap();
    assert!(edit.hide("Other").unwrap());
    let first = edit.commit().unwrap();
    let first_view = first.package().view().unwrap();
    assert!(Arc::ptr_eq(&proof, &first_view.inner.catalog_proof()));
    assert!(Arc::ptr_eq(
        &proof.limits,
        &first_view.inner.source_limits()
    ));

    // Materializing the next deferred owner must keep the same proof and
    // policy handles rather than allocating one OnceLock value per owner.
    let mut edit = first_view.edit_tab_state().unwrap();
    assert!(edit.show("Other").unwrap());
    let second = edit.commit().unwrap();
    let second_view = second.package().view().unwrap();
    assert!(Arc::ptr_eq(&proof, &second_view.inner.catalog_proof()));
    assert!(Arc::ptr_eq(
        &proof.limits,
        &second_view.inner.source_limits()
    ));
    // Every proof-backed owner borrows the proof's name storage; no reopen
    // retains a duplicate catalogue (pointer identity of the names and of
    // each name's string backing).
    assert!(first_view.inner.defined_names.is_none());
    assert!(second_view.inner.defined_names.is_none());
    for (index, name) in proof.defined_names.iter().enumerate() {
        assert_eq!(
            first_view.defined_names()[index].name.as_ptr(),
            name.name.as_ptr()
        );
        assert_eq!(
            second_view.defined_names()[index].name.as_ptr(),
            name.name.as_ptr()
        );
    }
    assert_eq!(
        first_view.defined_names().as_ptr(),
        proof.defined_names.as_ptr()
    );
    assert_eq!(
        second_view.defined_names().as_ptr(),
        proof.defined_names.as_ptr()
    );
    let first_owner_options = first_view.inner.owner_options();
    let second_owner_options = second_view.inner.owner_options();
    assert!(Arc::ptr_eq(&first_owner_options, &second_owner_options));

    let mismatch = SourceBackedWorkbook::from_read_at_with_options_and_shared_catalog(
        ProbeSource::new(bytes),
        crate::OpenOptions::new().source_limits(
            SourceLimits::builder()
                .max_part_bytes(1_000_000)
                .unwrap()
                .build()
                .unwrap(),
        ),
        Some(proof),
        None,
    );
    assert!(mismatch.is_err());
}

#[test]
fn block_reads_match_the_eager_owner_and_scan_once() {
    let bytes = rich_workbook();
    let eager = crate::Package::from_bytes(&bytes).unwrap();
    let expected = eager.sheet("Data").unwrap().unwrap();
    let (source, workbook) = open(&bytes);
    let sheet = workbook.sheet("Data").unwrap().unwrap();

    let validations = sheet.data_validations().unwrap();
    let reads = source.reads();
    assert!(sheet.blocks_elided_sheet_data().unwrap());
    for _ in 0..2 {
        assert_eq!(
            debug(sheet.data_validations().unwrap()),
            debug(expected.data_validations().unwrap())
        );
        assert_eq!(
            debug(sheet.conditional_formattings().unwrap()),
            debug(expected.conditional_formattings().unwrap())
        );
        assert_eq!(
            debug(sheet.auto_filter().unwrap()),
            debug(expected.auto_filter().unwrap())
        );
        assert_eq!(
            debug(sheet.hyperlinks().unwrap()),
            debug(expected.hyperlinks().unwrap())
        );
        assert_eq!(
            debug(sheet.views().unwrap()),
            debug(expected.views().unwrap())
        );
    }
    assert_eq!(validations.len(), 1);
    assert_eq!(sheet.hyperlinks().unwrap().len(), 2);
    assert!(sheet.auto_filter().unwrap().is_some());
    assert!(sheet.views().unwrap().is_some());
    assert_eq!(sheet.conditional_formattings().unwrap().len(), 1);

    // One projection serves every block read after the first: the styles
    // part was read once for the differential-format count.
    let styles_and_blocks = source.reads();
    assert!(styles_and_blocks >= reads);
    for _ in 0..3 {
        sheet.data_validations().unwrap();
        sheet.conditional_formattings().unwrap();
        sheet.hyperlinks().unwrap();
    }
    assert_eq!(source.reads(), styles_and_blocks);
    assert!(
        !sheet.store_resident(),
        "block reads do not materialize the cell grid"
    );

    // A sheet without any of these blocks answers empty, as the eager owner.
    let other = workbook.sheet("Other").unwrap().unwrap();
    assert!(other.data_validations().unwrap().is_empty());
    assert!(other.conditional_formattings().unwrap().is_empty());
    assert!(other.auto_filter().unwrap().is_none());
    assert!(other.hyperlinks().unwrap().is_empty());
    assert!(other.views().unwrap().is_none());
}

#[test]
fn block_reads_do_not_count_cell_events() {
    // 210,000 cells give the worksheet more than 1,000,000 XML events, the
    // ceiling each block parser charges.
    let mut xml = format!(r#"<worksheet xmlns="{SML}"><sheetData>"#);
    for row in 1..=21_000 {
        xml.push_str(&format!(r#"<row r="{row}">"#));
        for column in b'A'..=b'J' {
            xml.push_str(&format!(
                r#"<c r="{}{row}"><v>1</v></c>"#,
                char::from(column)
            ));
        }
        xml.push_str("</row>");
    }
    xml.push_str(r#"</sheetData><autoFilter ref="A1:J21000"/></worksheet>"#);
    let bytes = workbook_bytes(&[sheet("Large", xml)], "");

    let eager = crate::Package::from_bytes(&bytes).unwrap();
    let error = eager
        .sheet("Large")
        .unwrap()
        .unwrap()
        .auto_filter()
        .unwrap_err();
    assert!(error.to_string().contains("event limit"), "{error}");

    // The physical scan still counts every event under the caller's budget.
    let limits = SourceLimits::builder()
        .max_events(2_000_000)
        .unwrap()
        .build()
        .unwrap();
    let workbook = SourceBackedWorkbook::from_read_at_with_options(
        ProbeSource::new(bytes),
        crate::OpenOptions::new().source_limits(limits),
    )
    .unwrap();
    let sheet = workbook.sheet("Large").unwrap().unwrap();
    let filter = sheet.auto_filter().unwrap().expect("auto-filter");
    assert!(debug(filter).contains("A1:J21000"));
    assert!(sheet.views().unwrap().is_none());
    assert!(sheet.blocks_elided_sheet_data().unwrap());
}

#[test]
fn visit_cells_until_stops_at_the_break() {
    let (_source, workbook) = open(&rich_workbook());
    let sheet = workbook.sheet("Other").unwrap().unwrap();
    sheet.grid().unwrap();

    let mut seen = Vec::new();
    let flow = sheet
        .visit_cells_until("A1:D4", |address, cell| {
            seen.push(address.a1());
            if seen.len() == 3 {
                ControlFlow::Break(cell.clone())
            } else {
                ControlFlow::Continue(())
            }
        })
        .unwrap();
    assert_eq!(seen, ["A1", "B1", "C1"]);
    assert!(matches!(
        flow,
        ControlFlow::Break(Cell::Value(Value::Number(ref value))) if value.as_str() == "102"
    ));

    let flow = sheet
        .visit_cells_until("B2:C3", |_, _| ControlFlow::<()>::Continue(()))
        .unwrap();
    assert_eq!(flow, ControlFlow::Continue(4));
    let flow = sheet
        .visit_cells_until("F1:G9", |_, _| ControlFlow::<()>::Continue(()))
        .unwrap();
    assert_eq!(flow, ControlFlow::Continue(0));
}

#[test]
fn visit_cells_until_fences_outrank_the_break() {
    let (source, workbook) = open(&rich_workbook());
    let sheet = workbook.sheet("Other").unwrap().unwrap();
    let result = sheet.visit_cells_until("A1:D4", |_, _| {
        source.changed();
        ControlFlow::Break(())
    });
    assert!(matches!(
        result,
        Err(Error::Package(OpcError::SourceChanged { .. }))
    ));
}

#[test]
fn visit_cells_until_on_the_stream_route_refuses_before_any_callback() {
    let bytes = workbook_bytes(
        &[sheet(
            "Late",
            format!(
                r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="9"><c r="A9" t="inlineStr"><is><t>late</t></is><v>3</v></c></row></sheetData></worksheet>"#
            ),
        )],
        "",
    );
    let (_source, workbook) = open(&bytes);
    let sheet = workbook.sheet("Late").unwrap().unwrap();
    let mut callbacks = 0usize;
    let error = sheet
        .visit_cells_until("A1:B1", |_, _| {
            callbacks += 1;
            ControlFlow::Break(())
        })
        .unwrap_err();
    assert_eq!(callbacks, 0);
    assert!(error.to_string().contains("both inline text and a value"));
}

#[test]
fn stream_route_admits_inert_children_without_materializing() {
    let xml = format!(
        r#"<worksheet xmlns="{SML}"><dimension ref="A1:B2"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="B2" sqref="B2"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData><row r="1" spans="1:2"><c r="A1"><v>1</v></c><c r="B1" t="s"><v>1</v></c></row><row r="2" spans="1:2"><c r="B2"><v>4</v></c></row></sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>"#
    );
    let bytes = workbook_bytes(&[sheet("Plain", xml)], "");
    let eager = crate::Package::from_bytes(&bytes).unwrap();
    let expected = eager.sheet("Plain").unwrap().unwrap();
    let (_source, workbook) = open(&bytes);
    let sheet = workbook.sheet("Plain").unwrap().unwrap();

    for at in ["A1", "B1", "B2", "A2"] {
        let view = sheet.cell(at).unwrap();
        match expected.cell(at).unwrap() {
            crate::cell::View::Stored(cell) => {
                assert_eq!(view, SourceCellView::Stored(cell.clone()));
            },
            crate::cell::View::Missing => assert_eq!(view, SourceCellView::Missing),
            other => panic!("unexpected eager view {other:?}"),
        }
    }
    assert_eq!(sheet.cells("A1:B2").unwrap().len(), 3);
    assert!(
        !sheet.store_resident(),
        "an eligible streamed read retains no grid"
    );
    assert_eq!(workbook.semantic_cache_diagnostics().builds, 0);
}

#[test]
fn non_grid_handles_keep_typed_refusals() {
    let (_source, workbook) = open(&rich_workbook());
    let sheet = workbook.sheet("Data").unwrap().unwrap();
    assert!(matches!(sheet.row(1_048_576), Err(Error::Coordinate(_))));
    assert!(sheet.column("XFE").is_err());
    assert!(sheet.cell("A0").is_err());
}

#[test]
fn cold_cell_reads_stream_until_structural_reads_pin_a_grid() {
    let (_source, workbook) = open(&rich_workbook());
    let sheet = workbook.sheet("Other").unwrap().unwrap();
    assert_eq!(number(&sheet.cell("A1").unwrap()), "100");
    assert!(!sheet.store_resident());
    assert_eq!(workbook.semantic_cache_diagnostics().builds, 0);
    sheet.grid().unwrap();
    assert_eq!(number(&sheet.cell("A1").unwrap()), "100");
    assert_eq!(workbook.semantic_cache_diagnostics().builds, 1);
}

#[test]
fn block_reads_honor_the_open_profile_and_never_cache_limit_failures() {
    let limits = SourceLimits::builder()
        .max_events(32)
        .unwrap()
        .build()
        .unwrap();
    let source = ProbeSource::new(rich_workbook());
    let workbook = SourceBackedWorkbook::from_read_at_with_options(
        source.clone(),
        crate::OpenOptions::new().source_limits(limits),
    )
    .unwrap();
    let sheet = workbook.sheet("Other").unwrap().unwrap();
    for _ in 0..2 {
        assert!(matches!(
            sheet.auto_filter(),
            Err(Error::Package(OpcError::Limit {
                kind: litchi_opc::SourceLimitKind::Events,
                maximum: 32,
                ..
            }))
        ));
        assert_eq!(workbook.semantic_cache_diagnostics().builds, 0);
    }
}

#[test]
fn block_projection_preserves_mce_parser_results() {
    for tail in [
        r#"<autoFilter ref="A1:B2"/>"#,
        r#"<mc:AlternateContent><mc:Choice Requires="x14"><autoFilter ref="A1:B2"/></mc:Choice><mc:Fallback><autoFilter ref="A1:C3"/></mc:Fallback></mc:AlternateContent>"#,
    ] {
        let xml = format!(
            r#"<worksheet xmlns="{SML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" mc:Ignorable="x14"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>{tail}</worksheet>"#
        );
        let bytes = workbook_bytes(&[sheet("Data", xml)], "");
        let eager = crate::Package::from_bytes(&bytes).unwrap();
        let (_, source) = open(&bytes);
        let eager = eager.sheet("Data").unwrap().unwrap();
        let source = source.sheet("Data").unwrap().unwrap();
        assert_eq!(debug(source.auto_filter()), debug(eager.auto_filter()));
        assert_eq!(
            debug(source.data_validations()),
            debug(eager.data_validations())
        );
        assert_eq!(debug(source.views()), debug(eager.views()));
    }
}
