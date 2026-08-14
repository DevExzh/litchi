#![cfg(any(unix, windows))]
#![allow(clippy::unwrap_used, reason = "focused filesystem-source assertions")]

use std::fs;
use std::io;
use std::num::{NonZeroU64, NonZeroUsize};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, Resource,
};
use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcError, OpcPackage, PackURI, ReadLimits, SourceCacheLimits};
use litchi_xlsx::cell_values::SourceBackedEditor;
use litchi_xlsx::workbook::{SourceBackedWorkbook, SourceCellView};
use litchi_xlsx::{Address, Cell, Error, Value};
use soapberry_zip::office::StreamingArchiveWriter;
use tempfile::NamedTempFile;

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const FIRST: &str = "/xl/worksheets/sheet1.xml";
const SECOND: &str = "/xl/worksheets/sheet2.xml";
const UNUSED: &str = "/xl/media/untouched.bin";

fn archive_with_second(include_second: bool) -> Vec<u8> {
    let second_override = if include_second {
        format!(
            r#"<Override PartName="{SECOND}" ContentType="{}"/>"#,
            ct::SML_WORKSHEET
        )
    } else {
        String::new()
    };
    let workbook_xml = if include_second {
        r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="First" sheetId="1" r:id="rIdFirst"/><sheet name="Second" sheetId="2" r:id="rIdSecond"/></sheets></workbook>"#
    } else {
        r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="First" sheetId="1" r:id="rIdFirst"/></sheets></workbook>"#
    };
    let workbook_relationships = if include_second {
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdFirst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rIdSecond" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>"#
    } else {
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdFirst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#
    };
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored(
            "[Content_Types].xml",
            format!(
                r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/><Override PartName="/xl/workbook.xml" ContentType="{}"/><Override PartName="{}" ContentType="{}"/>{}</Types>"#,
                ct::SML_SHEET_MAIN,
                FIRST,
                ct::SML_WORKSHEET,
                second_override,
            )
            .as_bytes(),
        )
        .unwrap();
    writer
        .write_stored(
            "_rels/.rels",
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdRoot" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#,
        )
        .unwrap();
    writer
        .write_stored("xl/workbook.xml", workbook_xml.as_bytes())
        .unwrap();
    writer
        .write_stored(
            "xl/_rels/workbook.xml.rels",
            workbook_relationships.as_bytes(),
        )
        .unwrap();
    writer
        .write_stored(
            "xl/worksheets/sheet1.xml",
            format!(
                r#"<worksheet xmlns="{SML}"><sheetData><row r="1"><c r="A1"><v>7</v></c><c r="B1"><v>8</v></c></row></sheetData></worksheet>"#
            )
            .as_bytes(),
        )
        .unwrap();
    if include_second {
        // The catalog and graph can be opened without touching this malformed
        // payload. A selected read of First must continue to succeed.
        writer
            .write_stored(SECOND.trim_start_matches('/'), b"<worksheet")
            .unwrap();
    }
    writer
        .write_stored(
            UNUSED.trim_start_matches('/'),
            b"opaque untouched member with producer-specific bytes\0\xff",
        )
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

fn archive() -> Vec<u8> {
    archive_with_second(true)
}

fn editor_archive() -> Vec<u8> {
    archive_with_second(false)
}

fn write_fixture() -> (NamedTempFile, Vec<u8>) {
    let bytes = archive();
    let file = NamedTempFile::new().unwrap();
    fs::write(file.path(), &bytes).unwrap();
    (file, bytes)
}

fn write_editor_fixture() -> (NamedTempFile, Vec<u8>) {
    let bytes = editor_archive();
    let file = NamedTempFile::new().unwrap();
    fs::write(file.path(), &bytes).unwrap();
    (file, bytes)
}

fn part_bytes(bytes: &[u8], name: &str) -> Vec<u8> {
    OpcPackage::from_bytes(bytes)
        .unwrap()
        .get_part(&PackURI::new(name).unwrap())
        .unwrap()
        .blob()
        .to_vec()
}

fn address(value: &str) -> Address {
    Address::from_a1(value).unwrap()
}

fn managed_context(memory: u64) -> (Budget, ExecutionContext) {
    let budget = Budget::root(
        "xlsx-file-source-test",
        Limits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (_cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(memory.max(1)).unwrap(),
        0,
    )
    .unwrap();
    (
        budget.clone(),
        ExecutionContext::new(budget, cancellation, execution_limits),
    )
}

#[test]
fn path_catalog_and_selected_reads_are_lazy_over_unrequested_payloads() {
    let (file, _bytes) = write_fixture();
    let workbook = SourceBackedWorkbook::from_path(file.path()).unwrap();
    assert_eq!(
        workbook
            .sheets()
            .map(|sheet| sheet.name().to_owned())
            .collect::<Vec<_>>(),
        ["First", "Second"]
    );

    let first = workbook.sheet("First").unwrap().unwrap();
    assert!(matches!(
        first.cell("A1").unwrap(),
        SourceCellView::Stored(Cell::Value(Value::Number(ref value)))
            if value.as_str() == "7"
    ));
    // The malformed Second payload would fail if catalog/list or the selected
    // First read eagerly extracted and parsed every worksheet.
}

#[test]
fn path_editor_noop_is_byte_exact_and_one_edit_keeps_untouched_members() {
    let (file, bytes) = write_editor_fixture();
    let editor = SourceBackedEditor::from_path(file.path()).unwrap();
    let noop = editor.edit("First").unwrap().commit().unwrap();
    let mut output = Vec::new();
    editor.publish_commit_to_stream(&mut output, &noop).unwrap();
    assert_eq!(output, bytes);

    let editor = SourceBackedEditor::from_path(file.path()).unwrap();
    let mut edit = editor.edit("First").unwrap();
    edit.set(address("A1"), 42_u32).unwrap();
    let commit = edit.commit().unwrap();
    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_ne!(output, bytes);
    assert_eq!(part_bytes(&output, UNUSED), part_bytes(&bytes, UNUSED));
}

#[test]
fn path_managed_editor_releases_payload_budget_after_drop() {
    let (file, bytes) = write_editor_fixture();
    let workbook_bytes = part_bytes(&bytes, "/xl/workbook.xml").len() as u64;
    let sheet_bytes = part_bytes(&bytes, FIRST).len() as u64;
    let (budget, context) = managed_context(workbook_bytes + sheet_bytes + 1024);
    let editor = SourceBackedEditor::from_path_with_execution_context(
        file.path(),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    assert_eq!(budget.used(Resource::Memory), 0);
    let commit = editor.edit("First").unwrap().commit().unwrap();
    assert!(budget.used(Resource::Memory) >= workbook_bytes + sheet_bytes);
    drop(commit);
    drop(editor);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn path_replacement_and_limits_are_typed_errors() {
    let (file, _bytes) = write_fixture();
    let workbook = SourceBackedWorkbook::from_path(file.path()).unwrap();
    let first = workbook.sheet("First").unwrap().unwrap();
    fs::write(file.path(), b"replaced source with a different length").unwrap();
    assert!(matches!(
        first.cell("A1"),
        Err(Error::Package(OpcError::SourceChanged { .. }))
    ));

    let (file, _bytes) = write_editor_fixture();
    let editor = SourceBackedEditor::from_path(file.path()).unwrap();
    let commit = editor.edit("First").unwrap().commit().unwrap();
    fs::write(file.path(), b"replacement source for editor").unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        editor.publish_commit_to_stream(&mut output, &commit),
        Err(Error::Package(OpcError::SourceChanged { .. }))
    ));
    assert!(output.is_empty());

    let (file, bytes) = write_fixture();
    let limits = ReadLimits::builder()
        .max_input_bytes((bytes.len() as u64).saturating_sub(1))
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        SourceBackedWorkbook::from_path_with_limits(file.path(), limits),
        Err(Error::Package(OpcError::ReadLimit { .. }))
    ));

    let missing = file.path().with_extension("missing.xlsx");
    assert!(matches!(
        SourceBackedEditor::from_path(missing),
        Err(Error::Package(OpcError::IoError(_)))
    ));

    let directory = tempfile::tempdir().unwrap();
    assert!(matches!(
        SourceBackedWorkbook::from_path(directory.path()),
        Err(Error::Package(OpcError::IoError(error)))
            if matches!(
                error.kind(),
                io::ErrorKind::InvalidInput
                    | io::ErrorKind::PermissionDenied
                    | io::ErrorKind::IsADirectory
            )
    ));
}

#[test]
fn path_policy_variants_accept_explicit_cache_and_execution_options() {
    let (file, _bytes) = write_fixture();
    let cache = SourceCacheLimits::new(1 << 20, 8).unwrap();
    SourceBackedWorkbook::from_path(file.path()).unwrap();
    SourceBackedWorkbook::from_path_with_limits(file.path(), ReadLimits::default()).unwrap();
    SourceBackedWorkbook::from_path_with_cache_limits(file.path(), cache).unwrap();
    SourceBackedWorkbook::from_path_with_limits_and_cache_limits(
        file.path(),
        ReadLimits::default(),
        cache,
    )
    .unwrap();

    SourceBackedWorkbook::open(file.path()).unwrap();
    SourceBackedWorkbook::open_with_limits(file.path(), ReadLimits::default()).unwrap();
    SourceBackedWorkbook::open_with_cache_limits(file.path(), cache).unwrap();
    SourceBackedWorkbook::open_with_limits_and_cache_limits(
        file.path(),
        ReadLimits::default(),
        cache,
    )
    .unwrap();

    let (_budget, context) = managed_context(1 << 20);
    SourceBackedWorkbook::from_path_with_execution_context(
        file.path(),
        ReadLimits::default(),
        context,
    )
    .unwrap();

    let (_budget, context) = managed_context(1 << 20);
    SourceBackedWorkbook::from_path_with_limits_and_execution_context(
        file.path(),
        ReadLimits::default(),
        context,
    )
    .unwrap();

    let (_budget, context) = managed_context(1 << 20);
    SourceBackedWorkbook::open_with_limits_and_cache_limits_and_execution_context(
        file.path(),
        ReadLimits::default(),
        cache,
        context,
    )
    .unwrap();

    SourceBackedEditor::open(file.path()).unwrap();
    SourceBackedEditor::open_with_limits(file.path(), ReadLimits::default()).unwrap();
    SourceBackedEditor::open_with_cache_limits(file.path(), cache).unwrap();
    SourceBackedEditor::open_with_limits_and_cache_limits(
        file.path(),
        ReadLimits::default(),
        cache,
    )
    .unwrap();

    let (_budget, context) = managed_context(1 << 20);
    SourceBackedEditor::from_path_with_limits_and_execution_context(
        file.path(),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    let (_budget, context) = managed_context(1 << 20);
    SourceBackedEditor::open_with_limits_and_cache_limits_and_execution_context(
        file.path(),
        ReadLimits::default(),
        cache,
        context,
    )
    .unwrap();
}
