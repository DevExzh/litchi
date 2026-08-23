use std::{
    io::{self, Cursor, Write},
    ops::Range,
    sync::{
        Arc, Mutex,
        atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering},
    },
};

use litchi_core::{CancellationSource, OwnedSource, ReadAt, SourceVersion};
use litchi_odf_common::{
    constants,
    core::{
        PackageWriter, SourceContentPublicationError, SourceContentPublicationOptions,
        SourceContentPublicationProgress,
    },
    package::raw_identical_members,
};
use litchi_ods::{
    Cell, CellValue, CellView, SourceBackedSpreadsheet, Spreadsheet,
    worksheet::{CellChange, MAX_CELL_CHANGES},
};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

fn content(table_attributes: &str, rows: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}" office:version="1.4"><office:body><office:spreadsheet><table:table table:name="Data"{table_attributes}>{rows}</table:table></office:spreadsheet></office:body></office:document-content>"#,
    )
}

fn ordinary_content() -> String {
    content(
        "",
        concat!(
            r#"<table:table-row><table:table-cell office:value-type="string"><text:p>alpha</text:p></table:table-cell><table:table-cell office:value-type="string" table:number-columns-repeated="3"><text:p>same</text:p></table:table-cell></table:table-row>"#,
            r#"<table:table-row><table:table-cell office:value-type="float" office:value="7"><text:p>7</text:p></table:table-cell></table:table-row>"#,
        ),
    )
}

fn package(content_xml: &str, signed: bool) -> litchi_core::Result<Vec<u8>> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(constants::ODF_SPREADSHEET)?;
    writer.add_file("content.xml", content_xml.as_bytes())?;
    writer.add_file_with_media_type(
        "Pictures/opaque.bin",
        &vec![0x5a; 128 * 1024],
        "application/octet-stream",
    )?;
    if signed {
        writer.add_file_with_media_type(
            "META-INF/documentsignatures.xml",
            br#"<signatures xmlns="urn:test"/>"#,
            "application/vnd.oasis.opendocument.digital-signature",
        )?;
    }
    writer.finish_to_bytes()
}

fn text(value: &str) -> Cell {
    Cell::new(CellValue::Text(value.to_string()), value)
}

fn cell_text(spreadsheet: &Spreadsheet, row: usize, column: usize) -> Option<&str> {
    spreadsheet
        .cell("Data", row, column)
        .and_then(|cell| match cell {
            CellView::Stored(cell) => Some(cell.text.as_str()),
            CellView::Missing => None,
        })
}

#[test]
fn source_cell_commit_streams_one_and_repeated_run_edits_with_exact_patch() {
    let source = package(&ordinary_content(), false).unwrap();
    let owner =
        SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(source.clone()))).unwrap();
    let before = owner.cell_snapshot().unwrap();
    let mut edit = before.edit().unwrap();
    assert_eq!(
        edit.set_cell("Data", 0, 0, text("omega")).unwrap(),
        Some(true)
    );
    assert_eq!(
        edit.set_cells(
            "Data",
            vec![
                CellChange::new(0, 1, text("left")),
                CellChange::new(0, 3, text("right")),
            ],
        )
        .unwrap(),
        Some(2)
    );
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(commit.changed_cells(), 3);
    let applied = commit.patch().apply(&before).unwrap();
    assert_eq!(applied.content_xml(), commit.snapshot().content_xml());
    let restored = commit.patch().inverse().apply(&applied).unwrap();
    assert_eq!(restored.content_xml(), before.content_xml());
    assert!(commit.patch().apply(&applied).is_err());

    let mut output = Vec::new();
    let report = commit.write_to(&mut output).unwrap();
    assert_eq!(report.changed_cells(), 3);
    assert!(!report.is_no_op());
    assert_eq!(report.bytes(), output.len() as u64);
    let reopened = Spreadsheet::from_bytes(output.clone()).unwrap();
    assert_eq!(cell_text(&reopened, 0, 0), Some("omega"));
    assert_eq!(cell_text(&reopened, 0, 1), Some("left"));
    assert_eq!(cell_text(&reopened, 0, 2), Some("same"));
    assert_eq!(cell_text(&reopened, 0, 3), Some("right"));
    let identical = raw_identical_members(&source, &output).unwrap();
    assert!(identical.contains("mimetype"));
    assert!(identical.contains("META-INF/manifest.xml"));
    assert!(identical.contains("Pictures/opaque.bin"));
    assert!(!identical.contains("content.xml"));
}

#[test]
fn chained_source_cell_commits_do_not_reuse_offsets_from_the_original_layout() {
    let source = package(&ordinary_content(), false).unwrap();
    let owner =
        SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(source))).unwrap();

    // The first replacement changes the row length, so every later row moves
    // in the derived content.xml.  The owner's cached layout still describes
    // the original source and must not be used for a commit based on the
    // first commit's semantic snapshot.
    let first_value = "expanded-cell-value-".repeat(32);
    let first_cell = Cell::new(CellValue::Text(first_value.clone()), first_value.clone());
    let mut first_edit = owner.edit_cells().unwrap();
    assert_eq!(
        first_edit.set_cell("Data", 0, 0, first_cell).unwrap(),
        Some(true)
    );
    let first_commit = first_edit.commit().unwrap();

    let mut second_edit = first_commit.snapshot().edit().unwrap();
    assert_eq!(
        second_edit.set_cell("Data", 1, 0, text("second-edit")).unwrap(),
        Some(true)
    );
    let second_commit = second_edit.commit().unwrap();
    let mut output = Vec::new();
    second_commit.write_to(&mut output).unwrap();

    let reopened = Spreadsheet::from_bytes(output).unwrap();
    assert_eq!(cell_text(&reopened, 0, 0), Some(first_value.as_str()));
    assert_eq!(cell_text(&reopened, 1, 0), Some("second-edit"));
}

#[test]
fn changed_source_cell_commit_skips_second_content_payload_read() {
    let source_bytes = package(&ordinary_content(), false).unwrap();
    let content = zip_payload_range(&source_bytes, "content.xml");
    let source = Arc::new(ContentProbeSource::new(source_bytes));
    let owner = SourceBackedSpreadsheet::from_read_at(source.clone()).unwrap();
    source.forbid_range_until_output(content);

    let mut edit = owner.edit_cells().unwrap();
    assert_eq!(
        edit.set_cell("Data", 0, 0, text("omega")).unwrap(),
        Some(true)
    );
    let commit = edit.commit().unwrap();
    assert!(commit.changed());

    let mut sink = ProbeSink {
        bytes: Vec::new(),
        source: Arc::clone(&source),
    };
    let report = commit.write_to(&mut sink).unwrap();
    assert!(!report.is_no_op());
    assert_eq!(source.forbidden_read_count(), 0);

    let reopened = Spreadsheet::from_bytes(sink.bytes).unwrap();
    assert_eq!(cell_text(&reopened, 0, 0), Some("omega"));
}

#[test]
fn source_cell_commit_preserves_standard_table_children_and_refuses_paragraph_flattening() {
    const COLUMN: &str =
        r#"<table:table-column table:style-name="co1" table:number-columns-repeated="4"/>"#;
    let with_column = format!("{COLUMN}{}", ordinary_rows());
    let source = package(&content("", &with_column), false).unwrap();
    let owner = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(source))).unwrap();
    let mut edit = owner.edit_cells().unwrap();
    assert_eq!(
        edit.set_cell("Data", 0, 0, text("omega")).unwrap(),
        Some(true)
    );
    let commit = edit.commit().unwrap();
    let mut output = Vec::new();
    commit.write_to(&mut output).unwrap();
    let reopened =
        SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(output))).unwrap();
    assert!(reopened.content_xml().unwrap().contains(COLUMN));

    for unsafe_cell in [
        r#"<table:table-cell office:value-type="string"><text:p>first</text:p><text:p>second</text:p></table:table-cell>"#,
        r#"<table:table-cell office:value-type="string">outside<text:p>inside</text:p></table:table-cell>"#,
        r#"<table:table-cell office:value-type="string"><text:p><![CDATA[opaque]]></text:p></table:table-cell>"#,
        r#"<table:table-cell office:value-type="string"><text:p>a&amp;b</text:p></table:table-cell>"#,
        r#"<table:table-cell office:value-type="string"><?xml version="1.0"?><text:p>opaque</text:p></table:table-cell>"#,
    ] {
        let row = format!(
            r#"<table:table-row>{unsafe_cell}<table:table-cell office:value-type="string"><text:p>editable</text:p></table:table-cell></table:table-row>"#,
        );
        let source = package(&content("", &row), false).unwrap();
        let owner =
            SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(source))).unwrap();
        let mut edit = owner.edit_cells().unwrap();
        assert_eq!(
            edit.set_cell("Data", 0, 1, text("replacement")).unwrap(),
            Some(true)
        );
        assert!(edit.commit().is_err());
    }
}

#[test]
fn source_cell_edit_is_failure_atomic_and_refuses_unsafe_owners() {
    let source = package(&ordinary_content(), false).unwrap();
    let owner = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(source))).unwrap();
    let mut edit = owner.edit_cells().unwrap();
    assert!(edit.set_cell("Data", 9, 9, text("missing")).is_err());
    assert_eq!(edit.changed_cells(), 0);
    assert!(
        edit.set_cells(
            "Data",
            vec![
                CellChange::new(0, 0, text("one")),
                CellChange::new(0, 0, text("two")),
            ],
        )
        .is_err()
    );
    assert_eq!(edit.changed_cells(), 0);

    let mut formula = text("formula");
    formula.set_formula("of:=1+1").unwrap();
    assert!(edit.set_cell("Data", 0, 0, formula).is_err());
    assert_eq!(edit.changed_cells(), 0);

    let above = (0..=MAX_CELL_CHANGES)
        .map(|column| CellChange::new(0, column, text("x")))
        .collect();
    assert!(edit.set_cells("Data", above).is_err());
    assert_eq!(edit.changed_cells(), 0);

    let repeated_rows = content(
        "",
        r#"<table:table-row table:number-rows-repeated="2"><table:table-cell office:value-type="string"><text:p>same</text:p></table:table-cell></table:table-row>"#,
    );
    let repeated = package(&repeated_rows, false).unwrap();
    let repeated =
        SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(repeated))).unwrap();
    assert!(
        repeated
            .edit_cells()
            .unwrap()
            .set_cell("Data", 1, 0, text("changed"))
            .is_err()
    );

    let protected_source = package(
        &content(" table:protected=\"true\"", &ordinary_rows()),
        false,
    )
    .unwrap();
    let protected =
        SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(protected_source.clone())))
            .unwrap();
    let mut no_op = protected.edit_cells().unwrap();
    assert_eq!(
        no_op.set_cell("Data", 0, 0, text("alpha")).unwrap(),
        Some(false)
    );
    let mut exact = Vec::new();
    assert!(
        no_op
            .commit()
            .unwrap()
            .write_to(&mut exact)
            .unwrap()
            .is_no_op()
    );
    assert_eq!(exact, protected_source);
    assert!(
        protected
            .edit_cells()
            .unwrap()
            .set_cell("Data", 0, 0, text("changed"))
            .is_err()
    );

    let opaque_row = content(
        "",
        r#"<table:table-row><table:table-cell office:value-type="string"><text:p>old</text:p><office:annotation office:name="note"><text:p>opaque</text:p></office:annotation></table:table-cell></table:table-row>"#,
    );
    let opaque = package(&opaque_row, false).unwrap();
    let opaque = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(opaque))).unwrap();
    let mut edit = opaque.edit_cells().unwrap();
    assert_eq!(
        edit.set_cell("Data", 0, 0, text("changed")).unwrap(),
        Some(true)
    );
    assert!(edit.commit().is_err());

    let unknown_neighbor = content(
        "",
        r#"<table:table-row><table:table-cell office:value-type="vendor" office:boolean-value="true"><text:p>opaque</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>editable</text:p></table:table-cell></table:table-row>"#,
    );
    let source = package(&unknown_neighbor, false).unwrap();
    let owner = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(source))).unwrap();
    let mut edit = owner.edit_cells().unwrap();
    assert!(edit.set_cell("Data", 0, 1, text("changed")).is_err());
    assert_eq!(edit.changed_cells(), 0);

    let unknown_between = content(
        "",
        concat!(
            r#"<table:table-row><table:table-cell office:value-type="string"><text:p>first</text:p></table:table-cell></table:table-row>"#,
            r#"<table:table-row><table:table-cell office:value-type="vendor" office:boolean-value="true"><text:p>opaque</text:p></table:table-cell></table:table-row>"#,
            r#"<table:table-row><table:table-cell office:value-type="string"><text:p>last</text:p></table:table-cell></table:table-row>"#,
        ),
    );
    let source = package(&unknown_between, false).unwrap();
    let owner = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(source))).unwrap();
    let mut edit = owner.edit_cells().unwrap();
    assert_eq!(
        edit.set_cell("Data", 0, 0, text("changed-first")).unwrap(),
        Some(true)
    );
    assert_eq!(
        edit.set_cell("Data", 2, 0, text("changed-last")).unwrap(),
        Some(true)
    );
    assert!(edit.commit().is_err());
}

#[test]
fn source_cell_commit_validates_multi_row_windows_in_order() {
    // A changed window spanning several rows, including an unchanged row and
    // an empty row between the edited ones, commits exactly.
    let rows = concat!(
        r#"<table:table-row><table:table-cell office:value-type="string"><text:p>alpha</text:p></table:table-cell></table:table-row>"#,
        r#"<table:table-row/>"#,
        r#"<table:table-row><table:table-cell office:value-type="string"><text:p>middle</text:p></table:table-cell></table:table-row>"#,
        r#"<table:table-row><table:table-cell office:value-type="float" office:value="7"><text:p>7</text:p></table:table-cell></table:table-row>"#,
    );
    let source = package(&content("", rows), false).unwrap();
    let owner = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(source))).unwrap();
    let mut edit = owner.edit_cells().unwrap();
    assert_eq!(
        edit.set_cell("Data", 0, 0, text("one")).unwrap(),
        Some(true)
    );
    assert_eq!(
        edit.set_cell("Data", 3, 0, text("four")).unwrap(),
        Some(true)
    );
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(commit.changed_cells(), 2);
    let mut output = Vec::new();
    commit.write_to(&mut output).unwrap();
    let reopened = Spreadsheet::from_bytes(output).unwrap();
    assert_eq!(cell_text(&reopened, 0, 0), Some("one"));
    assert_eq!(cell_text(&reopened, 2, 0), Some("middle"));
    assert_eq!(cell_text(&reopened, 3, 0), Some("four"));
}

#[test]
fn source_cell_commit_reports_the_first_failing_row_in_a_window() {
    // Row 0 is clean and validates first; the refusal must still name the
    // problem in the later row, whether it is found by the descendant scan
    // or by the synthetic-document reparse.
    let cases = [
        (
            concat!(
                r#"<table:table-row><table:table-cell office:value-type="string"><text:p>alpha</text:p></table:table-cell></table:table-row>"#,
                r#"<table:table-row><table:table-cell office:value-type="string"><text:p>first</text:p><text:p>second</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>beta</text:p></table:table-cell></table:table-row>"#,
            ),
            "flat ODS edit requires at most one direct text paragraph per cell",
        ),
        (
            concat!(
                r#"<table:table-row><table:table-cell office:value-type="string"><text:p>alpha</text:p></table:table-cell></table:table-row>"#,
                r#"<table:table-row><table:table-cell office:value-type="string">loose<text:p>inside</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>beta</text:p></table:table-cell></table:table-row>"#,
            ),
            "flat ODS edit would discard text outside a cell paragraph",
        ),
        (
            concat!(
                r#"<table:table-row><table:table-cell office:value-type="string"><text:p>alpha</text:p></table:table-cell></table:table-row>"#,
                r#"<table:table-row><table:table-cell office:value-type="string"><text:p><![CDATA[opaque]]></text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>beta</text:p></table:table-cell></table:table-row>"#,
            ),
            "flat ODS edit would discard unsupported cell text markup",
        ),
        (
            concat!(
                r#"<table:table-row><table:table-cell office:value-type="string"><text:p>alpha</text:p></table:table-cell></table:table-row>"#,
                r#"<table:table-row table:unmodeled="urn:test"><table:table-cell office:value-type="string"><text:p>filler</text:p></table:table-cell><table:table-cell office:value-type="string"><text:p>beta</text:p></table:table-cell></table:table-row>"#,
            ),
            "flat ODS edit would discard unmodeled attribute 'unmodeled'",
        ),
    ];
    for (rows, expected) in cases {
        let source = package(&content("", rows), false).unwrap();
        let owner =
            SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(source))).unwrap();
        let mut edit = owner.edit_cells().unwrap();
        assert_eq!(
            edit.set_cell("Data", 0, 0, text("changed")).unwrap(),
            Some(true)
        );
        assert_eq!(
            edit.set_cell("Data", 1, 1, text("changed")).unwrap(),
            Some(true)
        );
        let error = edit.commit().unwrap_err();
        assert!(
            error.to_string().contains(expected),
            "expected '{expected}', got '{error}'",
        );
    }
}

fn ordinary_rows() -> String {
    ordinary_content()
        .split_once("<table:table table:name=\"Data\">")
        .and_then(|(_, tail)| tail.split_once("</table:table>").map(|(rows, _)| rows))
        .unwrap()
        .to_string()
}

fn zip_payload_range(bytes: &[u8], name: &str) -> Range<u64> {
    let mut archive = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let file = archive.by_name(name).unwrap();
    let start = file.data_start().unwrap();
    start..start + file.compressed_size()
}

#[test]
fn no_op_is_exact_and_changed_signed_source_is_refused_before_output() {
    let unsigned = package(&ordinary_content(), false).unwrap();
    let owner = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(unsigned.clone())))
        .unwrap();
    let mut edit = owner.edit_cells().unwrap();
    assert_eq!(
        edit.set_cell("Data", 0, 0, text("alpha")).unwrap(),
        Some(false)
    );
    let commit = edit.commit().unwrap();
    let mut output = Vec::new();
    let report = commit.write_to(&mut output).unwrap();
    assert!(report.is_no_op());
    assert_eq!(report.changed_cells(), 0);
    assert_eq!(output, unsigned);

    let signed = package(&ordinary_content(), true).unwrap();
    let signed_owner =
        SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(signed.clone()))).unwrap();
    let mut no_op = signed_owner.edit_cells().unwrap();
    assert_eq!(
        no_op.set_cell("Data", 0, 0, text("alpha")).unwrap(),
        Some(false)
    );
    let mut exact = Vec::new();
    assert!(
        no_op
            .commit()
            .unwrap()
            .write_to(&mut exact)
            .unwrap()
            .is_no_op()
    );
    assert_eq!(exact, signed);

    let mut edit = signed_owner.edit_cells().unwrap();
    assert_eq!(
        edit.set_cell("Data", 0, 0, text("omega")).unwrap(),
        Some(true)
    );
    let commit = edit.commit().unwrap();
    let mut output = Vec::new();
    let error = commit.write_to(&mut output).unwrap_err();
    assert!(matches!(
        error,
        SourceContentPublicationError::Unsupported { .. }
    ));
    assert!(output.is_empty());
}

#[test]
fn source_cell_edit_accepts_the_exact_transaction_bound() {
    let rows = format!(
        r#"<table:table-row><table:table-cell office:value-type="string" table:number-columns-repeated="{MAX_CELL_CHANGES}"><text:p>old</text:p></table:table-cell></table:table-row>"#,
    );
    let source = package(&content("", &rows), false).unwrap();
    let owner = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(source))).unwrap();
    let mut edit = owner.edit_cells().unwrap();
    let changes = (0..MAX_CELL_CHANGES)
        .map(|column| CellChange::new(0, column, text("new")))
        .collect();
    assert_eq!(
        edit.set_cells("Data", changes).unwrap(),
        Some(MAX_CELL_CHANGES)
    );
    let commit = edit.commit().unwrap();
    assert_eq!(commit.changed_cells(), MAX_CELL_CHANGES);
    let mut output = Vec::new();
    commit.write_to(&mut output).unwrap();
    let reopened = Spreadsheet::from_bytes(output).unwrap();
    assert_eq!(cell_text(&reopened, 0, 0), Some("new"));
    assert_eq!(cell_text(&reopened, 0, MAX_CELL_CHANGES - 1), Some("new"));
}

#[test]
fn publication_propagates_stale_limit_cancellation_and_partial_sink_state() {
    let source = Arc::new(MutableSource::new(
        package(&ordinary_content(), false).unwrap(),
    ));
    let owner = SourceBackedSpreadsheet::from_read_at(source.clone()).unwrap();
    let mut edit = owner.edit_cells().unwrap();
    edit.set_cell("Data", 0, 0, text("omega")).unwrap();
    let commit = edit.commit().unwrap();

    let replacement_limit = commit.snapshot().content_xml().len() as u64 - 1;
    let mut output = Vec::new();
    let error = commit
        .write_to_with_options(
            &mut output,
            SourceContentPublicationOptions::new().with_max_replacement_bytes(replacement_limit),
        )
        .unwrap_err();
    assert!(matches!(
        error,
        SourceContentPublicationError::LimitExceeded {
            progress: SourceContentPublicationProgress::Untouched,
            ..
        }
    ));
    assert!(output.is_empty());

    let (cancellation, token) = CancellationSource::pair();
    cancellation.cancel();
    let error = commit
        .write_to_with_options(
            Vec::new(),
            SourceContentPublicationOptions::new().with_cancellation(token),
        )
        .unwrap_err();
    assert!(matches!(
        error,
        SourceContentPublicationError::Cancelled {
            progress: SourceContentPublicationProgress::Untouched
        }
    ));

    let mut partial = PartialErrorSink::default();
    let error = commit.write_to(&mut partial).unwrap_err();
    assert!(matches!(
        error,
        SourceContentPublicationError::Sink {
            progress: SourceContentPublicationProgress::Indeterminate { accepted_before: 7 },
            ..
        }
    ));
    assert_eq!(partial.bytes.len(), 7);

    source.bump();
    let mut output = Vec::new();
    let error = commit.write_to(&mut output).unwrap_err();
    assert!(matches!(
        error,
        SourceContentPublicationError::SourceChanged {
            progress: SourceContentPublicationProgress::Untouched,
            ..
        }
    ));
    assert!(output.is_empty());
}

struct MutableSource {
    bytes: Arc<Vec<u8>>,
    revision: AtomicU64,
}

struct ContentProbeSource {
    bytes: Arc<Vec<u8>>,
    revision: AtomicU64,
    forbidden_range: Mutex<Option<Range<u64>>>,
    output_started: AtomicBool,
    forbidden_reads: AtomicUsize,
}

impl ContentProbeSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Arc::new(bytes),
            revision: AtomicU64::new(0),
            forbidden_range: Mutex::new(None),
            output_started: AtomicBool::new(false),
            forbidden_reads: AtomicUsize::new(0),
        }
    }

    fn forbid_range_until_output(&self, range: Range<u64>) {
        *self.forbidden_range.lock().unwrap() = Some(range);
    }

    fn forbidden_read_count(&self) -> usize {
        self.forbidden_reads.load(Ordering::Relaxed)
    }
}

impl ReadAt for ContentProbeSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len()).map_err(|_| io::Error::other("source too large"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let start = usize::try_from(offset).unwrap_or(usize::MAX);
        let Some(bytes) = self.bytes.get(start..) else {
            return Ok(0);
        };
        let length = bytes.len().min(output.len());
        let end = offset.saturating_add(length as u64);
        if !self.output_started.load(Ordering::Relaxed)
            && self
                .forbidden_range
                .lock()
                .unwrap()
                .as_ref()
                .is_some_and(|range| offset < range.end && range.start < end)
        {
            self.forbidden_reads.fetch_add(1, Ordering::Relaxed);
            return Err(io::Error::other(
                "content payload was read before publication output",
            ));
        }
        output[..length].copy_from_slice(&bytes[..length]);
        Ok(length)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x4f44_5350,
            self.revision.load(Ordering::Relaxed),
        ))
    }
}

struct ProbeSink {
    bytes: Vec<u8>,
    source: Arc<ContentProbeSource>,
}

impl Write for ProbeSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.source.output_started.store(true, Ordering::Relaxed);
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

impl MutableSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Arc::new(bytes),
            revision: AtomicU64::new(0),
        }
    }

    fn bump(&self) {
        self.revision.fetch_add(1, Ordering::Relaxed);
    }
}

impl ReadAt for MutableSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len()).map_err(|_| io::Error::other("source too large"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let start = usize::try_from(offset).unwrap_or(usize::MAX);
        let Some(bytes) = self.bytes.get(start..) else {
            return Ok(0);
        };
        let length = bytes.len().min(output.len());
        output[..length].copy_from_slice(&bytes[..length]);
        Ok(length)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x4f44_5353,
            self.revision.load(Ordering::Relaxed),
        ))
    }
}

#[derive(Default)]
struct PartialErrorSink {
    bytes: Vec<u8>,
    wrote_once: bool,
}

impl Write for PartialErrorSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.wrote_once {
            return Err(io::Error::other("injected sink failure"));
        }
        self.wrote_once = true;
        let accepted = bytes.len().min(7);
        self.bytes.extend_from_slice(&bytes[..accepted]);
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}
