#![allow(
    clippy::unwrap_used,
    reason = "test assertions use panic-on-failure extraction by design"
)]

use std::io::{self, Cursor};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering};

use litchi_core::sheet::{Cell, CellValue};
use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as BudgetLimits,
    OwnedSource, ReadAt, SourceVersion, TextOutputError, TextOutputLimitKind, TextOutputOptions,
};
use litchi_opc::constants::relationship_type;
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, SourceBackedPackage, SourceCacheLimits};
use litchi_xlsb::package::PackageError;
use litchi_xlsb::raw::{Header, Limits as RawLimits, Records, Writer, kind};
use litchi_xlsb::{ReadLimits, SourceBackedWorkbook};

fn fixture() -> Vec<u8> {
    std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ooxml/xlsb/Simple.xlsb"
    ))
    .unwrap()
}

fn rewrite_part(partname: &str, blob: Vec<u8>) -> Vec<u8> {
    let mut package = OpcPackage::from_reader(Cursor::new(fixture())).unwrap();
    package
        .get_part_mut(&PackURI::new(partname).unwrap())
        .unwrap()
        .set_blob(blob);
    let mut output = Vec::new();
    package.to_stream(&mut output).unwrap();
    output
}

fn remove_part(partname: &str) -> Vec<u8> {
    let mut package = OpcPackage::from_reader(Cursor::new(fixture())).unwrap();
    assert!(package.remove_part(&PackURI::new(partname).unwrap()));
    let mut output = Vec::new();
    package.to_stream(&mut output).unwrap();
    output
}

fn rewrite_content_type(partname: &str, content_type: &str) -> Vec<u8> {
    let mut package = OpcPackage::from_reader(Cursor::new(fixture())).unwrap();
    package
        .get_part_mut(&PackURI::new(partname).unwrap())
        .unwrap()
        .set_content_type(content_type.to_string())
        .unwrap();
    let mut output = Vec::new();
    package.to_stream(&mut output).unwrap();
    output
}

fn add_comments_part(content_type: &str) -> Vec<u8> {
    let mut package = OpcPackage::from_reader(Cursor::new(fixture())).unwrap();
    let sheet_uri = PackURI::new("/xl/worksheets/sheet1.bin").unwrap();
    let comments_uri = PackURI::new("/xl/comments1.bin").unwrap();
    let mut comments = Vec::new();
    litchi_xlsb::comments::write(&mut Writer::new(&mut comments), &[]).unwrap();
    package.add_part(Box::new(BlobPart::new(
        comments_uri,
        content_type.to_owned(),
        comments,
    )));
    package
        .get_part_mut(&sheet_uri)
        .unwrap()
        .rels_mut()
        .add_relationship(
            relationship_type::COMMENTS.to_owned(),
            "../comments1.bin".to_owned(),
            "rIdComments".to_owned(),
            false,
        );
    let mut output = Vec::new();
    package.to_stream(&mut output).unwrap();
    output
}

fn rewrite_sheet_relationship(relationship_type: &str, content_type: &str) -> Vec<u8> {
    let mut package = OpcPackage::from_reader(Cursor::new(fixture())).unwrap();
    let workbook_uri = PackURI::new("/xl/workbook.bin").unwrap();
    let sheet_uri = PackURI::new("/xl/worksheets/sheet1.bin").unwrap();
    let (r_id, target_ref, is_external) = package
        .get_part(&workbook_uri)
        .unwrap()
        .rels()
        .iter()
        .find_map(|relationship| {
            let target = relationship.target_partname().ok()?;
            (target == sheet_uri).then(|| {
                (
                    relationship.r_id().to_owned(),
                    relationship.target_ref().to_owned(),
                    relationship.is_external(),
                )
            })
        })
        .unwrap();
    {
        let workbook = package.get_part_mut(&workbook_uri).unwrap();
        workbook.rels_mut().remove(&r_id).unwrap();
        workbook.rels_mut().add_relationship(
            relationship_type.to_owned(),
            target_ref,
            r_id,
            is_external,
        );
    }
    package
        .get_part_mut(&sheet_uri)
        .unwrap()
        .set_content_type(content_type.to_owned())
        .unwrap();
    let mut output = Vec::new();
    package.to_stream(&mut output).unwrap();
    output
}

fn rewrite_sheet_target(target_partname: &str, target_ref: &str) -> Vec<u8> {
    let mut package = OpcPackage::from_reader(Cursor::new(fixture())).unwrap();
    let workbook_uri = PackURI::new("/xl/workbook.bin").unwrap();
    let target_partname = PackURI::new(target_partname).unwrap();
    let (r_id, rel_type, is_external) = package
        .get_part(&workbook_uri)
        .unwrap()
        .rels()
        .iter()
        .find_map(|relationship| {
            let target = relationship.target_partname().ok()?;
            (target == target_partname).then(|| {
                (
                    relationship.r_id().to_owned(),
                    relationship.reltype().to_owned(),
                    relationship.is_external(),
                )
            })
        })
        .unwrap();
    let workbook = package.get_part_mut(&workbook_uri).unwrap();
    workbook.rels_mut().remove(&r_id).unwrap();
    workbook
        .rels_mut()
        .add_relationship(rel_type, target_ref.to_owned(), r_id, is_external);
    let mut output = Vec::new();
    package.to_stream(&mut output).unwrap();
    output
}

fn set_date1904() -> Vec<u8> {
    let mut package = OpcPackage::from_reader(Cursor::new(fixture())).unwrap();
    let workbook_uri = PackURI::new("/xl/workbook.bin").unwrap();
    let mut workbook_blob = package.get_part(&workbook_uri).unwrap().blob().to_vec();
    let payload_offset = Records::new(&workbook_blob)
        .find_map(|record| {
            let record = record.unwrap();
            if record.kind() != kind::WORKBOOK_PROP || record.payload().is_empty() {
                return None;
            }
            let (_, header_len) =
                Header::parse(&workbook_blob[record.offset()..], RawLimits::DEFAULT).unwrap();
            Some(record.offset() + header_len)
        })
        .unwrap();
    workbook_blob[payload_offset] |= 1;
    package
        .get_part_mut(&workbook_uri)
        .unwrap()
        .set_blob(workbook_blob);
    let mut output = Vec::new();
    package.to_stream(&mut output).unwrap();
    output
}

struct VersionedSource {
    bytes: Vec<u8>,
    revision: AtomicU64,
}

struct CancelOnArmedRead {
    bytes: Vec<u8>,
    cancellation_source: CancellationSource,
    armed: AtomicBool,
    triggered: AtomicBool,
}

impl CancelOnArmedRead {
    fn new(bytes: Vec<u8>, cancellation_source: CancellationSource) -> Self {
        Self {
            bytes,
            cancellation_source,
            armed: AtomicBool::new(false),
            triggered: AtomicBool::new(false),
        }
    }

    fn arm(&self) {
        self.armed.store(true, Ordering::SeqCst);
    }

    fn triggered(&self) -> bool {
        self.triggered.load(Ordering::SeqCst)
    }
}

impl ReadAt for CancelOnArmedRead {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        if self.armed.swap(false, Ordering::SeqCst) {
            self.triggered.store(true, Ordering::SeqCst);
            self.cancellation_source.cancel();
        }
        let offset = usize::try_from(offset)
            .map_err(|_error| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let end = offset.saturating_add(output.len()).min(self.bytes.len());
        output[..end - offset].copy_from_slice(&self.bytes[offset..end]);
        Ok(end - offset)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(0x4341_4e43, 0))
    }
}

impl VersionedSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            revision: AtomicU64::new(0),
        }
    }

    fn change(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }
}

fn managed_context() -> (CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "xlsb-source-backed-test",
        BudgetLimits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(u64::MAX).unwrap(),
        0,
    )
    .unwrap();
    (
        cancellation_source,
        ExecutionContext::new(budget, cancellation, execution_limits),
    )
}

fn assert_cancelled<T>(result: Result<T, PackageError>) {
    assert!(matches!(
        result,
        Err(PackageError::Opc(OpcError::Cancelled))
    ));
}

impl ReadAt for VersionedSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset)
            .map_err(|_error| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let end = offset.saturating_add(output.len()).min(self.bytes.len());
        output[..end - offset].copy_from_slice(&self.bytes[offset..end]);
        Ok(end - offset)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x584c_5342,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

#[test]
fn catalog_defers_sheet_payloads_and_materializes_one_selection() {
    let workbook =
        SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(fixture()))).unwrap();
    assert!(workbook.worksheet_count().unwrap() > 0);
    let names = workbook.worksheet_names().unwrap();
    assert_eq!(names.len(), workbook.worksheet_count().unwrap());

    let catalog_diagnostics = workbook.cache_diagnostics();
    assert_eq!(catalog_diagnostics.cold_loads, 1);
    let selected = workbook.worksheet_by_index(0).unwrap().unwrap();
    assert_eq!(selected.name().unwrap(), names[0].as_str());
    assert_eq!(selected.workbook_position().unwrap(), 0);
    selected.materialize().unwrap();

    let first_materialization = workbook.cache_diagnostics();
    assert!(first_materialization.cold_loads > catalog_diagnostics.cold_loads);
    selected.materialize().unwrap();
    let repeated_materialization = workbook.cache_diagnostics();
    assert_eq!(
        repeated_materialization.cold_loads,
        first_materialization.cold_loads
    );
    assert!(repeated_materialization.hits > first_materialization.hits);
}

#[test]
fn selectors_remain_metadata_only_and_source_checked() {
    let source = Arc::new(VersionedSource::new(fixture()));
    let workbook = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let name = workbook.worksheet_names().unwrap().remove(0);
    let selected = workbook.worksheet_by_name(&name).unwrap().unwrap();
    selected.materialize().unwrap();
    assert!(workbook.worksheet_by_name(&name).unwrap().is_some());
    assert!(workbook.worksheet_by_name("missing").unwrap().is_none());

    source.change();
    assert!(workbook.worksheet_count().is_err());
    assert!(workbook.worksheet_by_name(&name).is_err());
    assert!(selected.materialize().is_err());
}

#[test]
fn managed_cancellation_is_checked_for_catalog_selectors_state_and_handles() {
    let source = Arc::new(VersionedSource::new(fixture()));
    let (cancellation_source, context) = managed_context();
    let workbook = SourceBackedWorkbook::from_read_at_with_execution_context(
        source,
        ReadLimits::default(),
        context,
    )
    .unwrap();
    let sheet_name = workbook.sheet_names().unwrap().remove(0);
    let worksheet_name = workbook.worksheet_names().unwrap().remove(0);
    let selected_sheet = workbook.sheet_by_index(0).unwrap().unwrap();
    let selected_worksheet = workbook
        .worksheet_by_name(&worksheet_name)
        .unwrap()
        .unwrap();
    selected_worksheet.materialize().unwrap();

    cancellation_source.cancel();

    assert_cancelled(workbook.sheet_count());
    assert_cancelled(workbook.sheet_names());
    assert_cancelled(workbook.sheets());
    assert_cancelled(workbook.sheet_by_index(0));
    assert_cancelled(workbook.sheet_by_index(usize::MAX));
    assert_cancelled(workbook.sheet_by_name(&sheet_name));
    assert_cancelled(workbook.sheet_by_name("missing"));
    assert_cancelled(workbook.worksheet_count());
    assert_cancelled(workbook.worksheet_names());
    assert_cancelled(workbook.worksheets());
    assert_cancelled(workbook.worksheet_by_index(0));
    assert_cancelled(workbook.worksheet_by_index(usize::MAX));
    assert_cancelled(workbook.worksheet_by_name(&worksheet_name));
    assert_cancelled(workbook.worksheet_by_name("missing"));
    assert_cancelled(workbook.active_catalog_position());
    assert_cancelled(workbook.active_worksheet_index());
    assert_cancelled(workbook.source_version());
    assert_cancelled(workbook.is_1904_date_system());
    assert_cancelled(selected_sheet.name());
    assert_cancelled(selected_sheet.workbook_position());
    assert_cancelled(selected_worksheet.materialize());
}

#[test]
fn managed_cancellation_is_checked_before_catalog_construction() {
    let (cancellation_source, context) = managed_context();
    cancellation_source.cancel();
    assert!(matches!(
        SourceBackedWorkbook::from_read_at_with_execution_context(
            Arc::new(OwnedSource::new(fixture())),
            ReadLimits::default(),
            context,
        ),
        Err(PackageError::Opc(OpcError::Cancelled))
    ));
}

#[test]
fn managed_cancellation_precedes_stale_source_preflight() {
    let source = Arc::new(VersionedSource::new(fixture()));
    let (cancellation_source, context) = managed_context();
    let workbook = SourceBackedWorkbook::from_read_at_with_execution_context(
        source.clone(),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    source.change();
    cancellation_source.cancel();

    assert_cancelled(workbook.sheet_count());
}

#[test]
fn stale_source_preflight_remains_source_changed_without_cancellation() {
    let source = Arc::new(VersionedSource::new(fixture()));
    let (_cancellation_source, context) = managed_context();
    let workbook = SourceBackedWorkbook::from_read_at_with_execution_context(
        source.clone(),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    source.change();

    assert!(matches!(
        workbook.sheet_count(),
        Err(PackageError::Opc(OpcError::SourceChanged { .. }))
    ));
}

#[test]
fn cancellation_during_selected_materialization_is_typed_and_not_cached() {
    let (cancellation_source, context) = managed_context();
    let source = Arc::new(CancelOnArmedRead::new(fixture(), cancellation_source));
    let workbook = SourceBackedWorkbook::from_read_at_with_execution_context(
        source.clone(),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    let selected = workbook.worksheet_by_index(0).unwrap().unwrap();
    let before = workbook.cache_diagnostics();

    source.arm();
    assert!(matches!(
        selected.materialize(),
        Err(PackageError::Opc(OpcError::Cancelled))
    ));
    assert!(source.triggered());

    let after = workbook.cache_diagnostics();
    assert_eq!(after.retained_entries, before.retained_entries);
}

#[test]
fn case_variant_duplicate_sheet_targets_are_rejected() {
    let source = rewrite_sheet_target("/xl/worksheets/sheet2.bin", "worksheets/SHEET1.bin");
    assert!(matches!(
        SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(source))),
        Err(PackageError::InvalidRelationship(_))
    ));
}

#[test]
fn unique_case_variant_sheet_target_opens_and_materializes() {
    let source = rewrite_sheet_target("/xl/worksheets/sheet2.bin", "worksheets/SHEET2.bin");
    let workbook = SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(source))).unwrap();
    let selected = workbook.worksheet_by_index(1).unwrap().unwrap();
    assert_eq!(
        selected.name().unwrap(),
        workbook.worksheet_names().unwrap()[1]
    );
    selected.materialize().unwrap();
}

#[test]
fn managed_package_handoff_checks_cancellation_before_catalog_construction() {
    let (cancellation_source, context) = managed_context();
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(fixture())),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    cancellation_source.cancel();

    assert!(matches!(
        SourceBackedWorkbook::from_source_backed_package(package),
        Err(PackageError::Opc(OpcError::Cancelled))
    ));
}

#[test]
fn unselected_malformed_sheet_remains_deferred() {
    let source = rewrite_part("/xl/worksheets/sheet2.bin", vec![0xff; 8]);
    let workbook = SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(source))).unwrap();
    assert!(
        workbook
            .worksheet_by_index(0)
            .unwrap()
            .unwrap()
            .materialize()
            .is_ok()
    );
    assert!(
        workbook
            .worksheet_by_index(1)
            .unwrap()
            .unwrap()
            .materialize()
            .is_err()
    );
}

#[test]
fn declared_part_read_limit_is_forwarded_at_catalog_open() {
    let package = OpcPackage::from_reader(Cursor::new(fixture())).unwrap();
    let dependency_limit = [
        "/xl/workbook.bin",
        "/xl/styles.bin",
        "/xl/sharedStrings.bin",
        "/xl/worksheets/sheet1.bin",
    ]
    .iter()
    .filter_map(|partname| package.get_part(&PackURI::new(*partname).unwrap()).ok())
    .map(|part| part.blob().len())
    .max()
    .unwrap();
    let oversized = vec![0xff; dependency_limit.saturating_add(1)];
    let source = rewrite_part("/xl/worksheets/sheet2.bin", oversized);
    let limits = ReadLimits::builder()
        .max_part_bytes(u64::try_from(dependency_limit).unwrap())
        .unwrap()
        .build()
        .unwrap();
    assert!(
        SourceBackedWorkbook::from_read_at_with_limits(Arc::new(OwnedSource::new(source)), limits,)
            .is_err()
    );
}

#[test]
fn missing_related_parts_are_rejected_during_catalog_construction() {
    for partname in ["/xl/worksheets/sheet1.bin", "/xl/styles.bin"] {
        let source = remove_part(partname);
        assert!(
            SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(source))).is_err(),
            "accepted missing related part {partname}"
        );
    }
}

#[test]
fn mismatched_binary_part_content_types_are_rejected() {
    for partname in [
        "/xl/worksheets/sheet1.bin",
        "/xl/styles.bin",
        "/xl/sharedStrings.bin",
    ] {
        let source = rewrite_content_type(partname, "application/octet-stream");
        assert!(
            SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(source))).is_err(),
            "accepted mismatched content type for {partname}"
        );
    }
}

#[test]
fn comments_content_type_is_checked_before_materialization() {
    let source = add_comments_part("application/octet-stream");
    let workbook = SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(source))).unwrap();
    assert!(
        workbook
            .worksheet_by_index(0)
            .unwrap()
            .unwrap()
            .materialize()
            .is_err()
    );
}

#[test]
fn recognized_nonworksheet_sheet_relationships_are_exact_and_typed() {
    const RELATIONSHIPS: &[(&str, &str)] = &[
        (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
            "application/vnd.ms-excel.chartsheet",
        ),
        (
            "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet",
            "application/vnd.ms-excel.chartsheet",
        ),
        (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet",
            "application/vnd.ms-excel.dialogsheet",
        ),
        (
            "http://purl.oclc.org/ooxml/officeDocument/relationships/dialogsheet",
            "application/vnd.ms-excel.dialogsheet",
        ),
        (
            "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet",
            "application/vnd.ms-excel.macrosheet",
        ),
        (
            "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet",
            "application/vnd.ms-excel.intlmacrosheet",
        ),
    ];
    for &(relationship_type, content_type) in RELATIONSHIPS {
        let source = rewrite_sheet_relationship(relationship_type, content_type);
        assert!(
            SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(source))).is_ok(),
            "rejected recognized sheet relationship {relationship_type}"
        );
    }
}

#[test]
fn recognized_nonworksheet_sheet_relationships_require_binary_content_types() {
    const RELATIONSHIPS: &[&str] = &[
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
        "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet",
        "http://purl.oclc.org/ooxml/officeDocument/relationships/dialogsheet",
        "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet",
        "http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet",
    ];
    for &relationship_type in RELATIONSHIPS {
        let source = rewrite_sheet_relationship(relationship_type, "application/octet-stream");
        assert!(
            SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(source))).is_err(),
            "accepted mismatched content type for {relationship_type}"
        );
    }
}

#[test]
fn full_sheet_catalog_preserves_nonworksheet_tabs_and_worksheet_selectors() {
    const RELATIONSHIPS: &[(&str, &str)] = &[
        (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
            "application/vnd.ms-excel.chartsheet",
        ),
        (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet",
            "application/vnd.ms-excel.dialogsheet",
        ),
        (
            "http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet",
            "application/vnd.ms-excel.macrosheet",
        ),
    ];
    for &(relationship_type, content_type) in RELATIONSHIPS {
        let source = rewrite_sheet_relationship(relationship_type, content_type);
        let workbook =
            SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(source))).unwrap();
        let names = workbook.sheet_names().unwrap();
        assert_eq!(workbook.sheet_count().unwrap(), names.len());
        assert!(names.len() >= 2);

        let sheets = workbook.sheets().unwrap();
        assert_eq!(sheets.len(), names.len());
        for (position, sheet) in sheets.iter().enumerate() {
            assert_eq!(sheet.name().unwrap(), names[position].as_str());
            assert_eq!(sheet.workbook_position().unwrap(), position);
        }
        assert!(matches!(
            sheets[0].materialize(),
            Err(PackageError::UnsupportedFeature(_))
        ));

        assert_eq!(workbook.worksheet_count().unwrap(), names.len() - 1);
        assert_eq!(workbook.worksheet_names().unwrap(), names[1..].to_vec());
        assert_eq!(
            workbook
                .worksheet_by_index(0)
                .unwrap()
                .unwrap()
                .workbook_position()
                .unwrap(),
            1
        );
        assert!(workbook.sheet_by_name(&names[0]).unwrap().is_some());
        assert!(workbook.sheet_by_index(names.len()).unwrap().is_none());
    }
}

#[test]
fn source_text_matches_legacy_terminal_newline_and_reports_progress() {
    let workbook =
        SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(fixture()))).unwrap();
    let expected = workbook.text().unwrap();
    let mut output = Vec::new();
    let report = workbook
        .write_text_to(&mut output, TextOutputOptions::default())
        .unwrap();

    let mut expected_without_terminal = expected.into_bytes();
    expected_without_terminal.pop();
    assert_eq!(output, expected_without_terminal);
    assert!(report.objects_written() > 0);
    assert_eq!(report.bytes_written(), output.len() as u64);
}

#[test]
fn source_text_respects_output_limit() {
    let workbook =
        SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(fixture()))).unwrap();
    let mut output = Vec::new();
    let error = workbook
        .write_text_to(&mut output, TextOutputOptions::new("\n", "\n\n", 1, 1))
        .unwrap_err();

    assert!(matches!(
        &error,
        TextOutputError::Limit { limit, progress }
            if limit.kind() == TextOutputLimitKind::OutputBytes
                && limit.limit() == 1
                && limit.observed() > 1
                && progress.bytes_written() == 0
                && progress.objects_written() == 0
    ));
    assert!(output.len() <= 1);
}

#[test]
fn source_text_rejects_stale_source_before_output() {
    let source = Arc::new(VersionedSource::new(fixture()));
    let workbook = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    source.change();

    let text_error = workbook.text().unwrap_err();
    assert!(matches!(
        text_error,
        PackageError::Opc(OpcError::SourceChanged { .. })
    ));

    let mut output = Vec::new();
    let sink_error = workbook
        .write_text_to(&mut output, TextOutputOptions::default())
        .unwrap_err();
    assert!(matches!(
        sink_error,
        TextOutputError::Document {
            source: PackageError::Opc(OpcError::SourceChanged { .. }),
            progress,
        } if progress.bytes_written() == 0 && progress.objects_written() == 0
    ));
    assert!(output.is_empty());
}

#[test]
fn source_text_rejects_nonworksheet_with_typed_error() {
    let source = rewrite_sheet_relationship(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet",
        "application/vnd.ms-excel.chartsheet",
    );
    let workbook = SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(source))).unwrap();
    let mut output = Vec::new();
    let error = workbook
        .write_text_to(&mut output, TextOutputOptions::default())
        .unwrap_err();

    assert!(matches!(
        error,
        TextOutputError::Document {
            source: PackageError::UnsupportedFeature(_),
            ..
        }
    ));
    assert!(output.is_empty());
    assert!(matches!(
        workbook.text(),
        Err(PackageError::UnsupportedFeature(_))
    ));
}

#[test]
fn date1904_flag_is_retained_in_source_catalog() {
    let source = set_date1904();
    let workbook = SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(source))).unwrap();
    assert!(workbook.is_1904_date_system().unwrap());
}

#[test]
fn unknown_sheet_relationship_suffix_is_rejected() {
    let source = rewrite_sheet_relationship(
        "https://attacker.invalid/relationships/chartsheet",
        "application/vnd.ms-excel.chartsheet",
    );
    assert!(
        SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(source))).is_err(),
        "accepted an unknown chartsheet relationship authority"
    );
}

#[test]
fn finite_cache_policy_is_forwarded_to_deferred_parts() {
    let cache_limits = SourceCacheLimits::new(1, 1).unwrap();
    let workbook = SourceBackedWorkbook::from_read_at_with_cache_limits(
        Arc::new(OwnedSource::new(fixture())),
        cache_limits,
    )
    .unwrap();
    let selected = workbook.worksheet_by_index(0).unwrap().unwrap();
    let before = workbook.cache_diagnostics().cold_loads;
    selected.materialize().unwrap();
    let first = workbook.cache_diagnostics().cold_loads;
    selected.materialize().unwrap();
    let second = workbook.cache_diagnostics().cold_loads;
    assert!(first > before);
    assert!(second > first);
    assert_eq!(workbook.cache_diagnostics().retained_entries, 0);
}

fn wide_string(value: &str) -> Vec<u8> {
    let mut data = (value.encode_utf16().count() as u32).to_le_bytes().to_vec();
    for unit in value.encode_utf16() {
        data.extend_from_slice(&unit.to_le_bytes());
    }
    data
}

fn nullable_wide_string(value: Option<&str>) -> Vec<u8> {
    value.map_or_else(|| u32::MAX.to_le_bytes().to_vec(), wide_string)
}

fn table_header_payload(id: u32, display_name: &str) -> Vec<u8> {
    let mut data = Vec::new();
    for value in [0_u32, 1, 0, 0, 0, id, 1, 0, 0] {
        data.extend_from_slice(&value.to_le_bytes());
    }
    for _ in 0..6 {
        data.extend_from_slice(&u32::MAX.to_le_bytes());
    }
    data.extend_from_slice(&0_u32.to_le_bytes());
    data.extend_from_slice(&nullable_wide_string(Some(display_name)));
    data.extend_from_slice(&nullable_wide_string(Some(display_name)));
    for _ in 0..4 {
        data.extend_from_slice(&nullable_wide_string(None));
    }
    data
}

fn table_column_payload(caption: &str) -> Vec<u8> {
    let mut data = Vec::new();
    for value in [1_u32, 0, u32::MAX, u32::MAX, u32::MAX, 0] {
        data.extend_from_slice(&value.to_le_bytes());
    }
    data.extend_from_slice(&nullable_wide_string(None));
    data.extend_from_slice(&nullable_wide_string(Some(caption)));
    for _ in 0..4 {
        data.extend_from_slice(&nullable_wide_string(None));
    }
    data
}

fn table_part_payload(id: u32, display_name: &str) -> Vec<u8> {
    let mut data = Vec::new();
    let mut writer = Writer::new(&mut data);
    writer
        .write_record(kind::BEGIN_LIST, &table_header_payload(id, display_name))
        .unwrap();
    writer
        .write_record(kind::BEGIN_LIST_COLS, &1_u32.to_le_bytes())
        .unwrap();
    writer
        .write_record(kind::BEGIN_LIST_COL, &table_column_payload("Amount"))
        .unwrap();
    writer.write_record(kind::END_LIST_COL, &[]).unwrap();
    writer.write_record(kind::END_LIST_COLS, &[]).unwrap();
    writer.write_record(kind::END_LIST, &[]).unwrap();
    data
}

fn malformed_table_part_payload() -> Vec<u8> {
    let mut data = Vec::new();
    let mut writer = Writer::new(&mut data);
    writer.write_record(kind::BEGIN_LIST, &[0; 8]).unwrap();
    data
}

fn resident_table_formula(table_id: u32) -> (Vec<u8>, Vec<u8>) {
    let mut rgce = vec![0x18, 0x19, 0, 0, 0x19, 0];
    rgce.extend_from_slice(&table_id.to_le_bytes());
    rgce.extend_from_slice(&0_u16.to_le_bytes());
    rgce.extend_from_slice(&0_u16.to_le_bytes());
    let mut formula = (u32::try_from(rgce.len()).unwrap()).to_le_bytes().to_vec();
    formula.extend_from_slice(&rgce);
    formula.extend_from_slice(&0_u32.to_le_bytes());
    (formula, rgce)
}

fn worksheet_with_table(table_relationship_id: Option<&str>, formula: Option<&[u8]>) -> Vec<u8> {
    let mut data = Vec::new();
    let mut writer = Writer::new(&mut data);
    let mut dimensions = Vec::new();
    for value in [0_u32, 1, 0, 0] {
        dimensions.extend_from_slice(&value.to_le_bytes());
    }
    writer.write_record(kind::WS_DIM, &dimensions).unwrap();
    if let Some(relationship_id) = table_relationship_id {
        writer
            .write_record(kind::BEGIN_LIST_PARTS, &1_u32.to_le_bytes())
            .unwrap();
        writer
            .write_record(kind::LIST_PART, &wide_string(relationship_id))
            .unwrap();
        writer.write_record(kind::END_LIST_PARTS, &[]).unwrap();
    }
    writer.write_record(kind::BEGIN_SHEET_DATA, &[]).unwrap();
    if let Some(formula) = formula {
        let mut row = 0_u32.to_le_bytes().to_vec();
        row.extend_from_slice(&0_u32.to_le_bytes());
        row.extend_from_slice(&0_u16.to_le_bytes());
        row.extend_from_slice(&[0, 0, 0]);
        row.extend_from_slice(&0_u32.to_le_bytes());
        writer.write_record(kind::ROW_HDR, &row).unwrap();

        let mut cell = 0_u32.to_le_bytes().to_vec();
        cell.extend_from_slice(&[0, 0, 0, 0]);
        cell.extend_from_slice(&42_f64.to_le_bytes());
        cell.extend_from_slice(&0_u16.to_le_bytes());
        cell.extend_from_slice(formula);
        writer.write_record(kind::FMLA_NUM, &cell).unwrap();
    }
    writer.write_record(kind::END_SHEET_DATA, &[]).unwrap();
    writer.write_record(kind::END_SHEET, &[]).unwrap();
    data
}

fn bundle_sheet_payload(id: u32, relationship_id: &str, name: &str) -> Vec<u8> {
    let mut data = 0_u32.to_le_bytes().to_vec();
    data.extend_from_slice(&id.to_le_bytes());
    data.extend_from_slice(&wide_string(relationship_id));
    data.extend_from_slice(&wide_string(name));
    data
}

fn two_sheet_workbook_payload() -> Vec<u8> {
    let mut data = Vec::new();
    let mut writer = Writer::new(&mut data);
    writer
        .write_record(kind::BUNDLE_SH, &bundle_sheet_payload(1, "rId1", "Sheet1"))
        .unwrap();
    writer
        .write_record(kind::BUNDLE_SH, &bundle_sheet_payload(2, "rId2", "Sheet2"))
        .unwrap();
    writer.write_record(kind::SUP_SELF, &[]).unwrap();
    let mut extern_sheet = 1_u32.to_le_bytes().to_vec();
    extern_sheet.extend_from_slice(&0_u32.to_le_bytes());
    extern_sheet.extend_from_slice(&1_u32.to_le_bytes());
    extern_sheet.extend_from_slice(&1_u32.to_le_bytes());
    writer
        .write_record(kind::EXTERN_SHEET, &extern_sheet)
        .unwrap();
    data
}

fn table_workbook_source(
    first: Option<(String, Vec<u8>, bool)>,
    second: Option<(String, Vec<u8>, bool)>,
    formula: Option<Vec<u8>>,
) -> Vec<u8> {
    let mut package = OpcPackage::from_reader(Cursor::new(fixture())).unwrap();
    package
        .get_part_mut(&PackURI::new("/xl/workbook.bin").unwrap())
        .unwrap()
        .set_blob(two_sheet_workbook_payload());

    let first_relationship_id = first.as_ref().map(|_| "rIdTable1");
    let second_relationship_id = second.as_ref().map(|_| "rIdTable2");
    package
        .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.bin").unwrap())
        .unwrap()
        .set_blob(worksheet_with_table(
            first_relationship_id,
            formula.as_deref(),
        ));
    package
        .get_part_mut(&PackURI::new("/xl/worksheets/sheet2.bin").unwrap())
        .unwrap()
        .set_blob(worksheet_with_table(second_relationship_id, None));

    if let Some((content_type, payload, external)) = first {
        package
            .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.bin").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                relationship_type::TABLE.to_owned(),
                if external {
                    "https://example.invalid/table1.bin".to_owned()
                } else {
                    "../tables/table1.bin".to_owned()
                },
                "rIdTable1".to_owned(),
                external,
            );
        if !external {
            package.add_part(Box::new(BlobPart::new(
                PackURI::new("/xl/tables/table1.bin").unwrap(),
                content_type,
                payload,
            )));
        }
    }
    if let Some((content_type, payload, external)) = second {
        package
            .get_part_mut(&PackURI::new("/xl/worksheets/sheet2.bin").unwrap())
            .unwrap()
            .rels_mut()
            .add_relationship(
                relationship_type::TABLE.to_owned(),
                if external {
                    "https://example.invalid/table2.bin".to_owned()
                } else {
                    "../tables/table2.bin".to_owned()
                },
                "rIdTable2".to_owned(),
                external,
            );
        if !external {
            package.add_part(Box::new(BlobPart::new(
                PackURI::new("/xl/tables/table2.bin").unwrap(),
                content_type,
                payload,
            )));
        }
    }

    let mut output = Vec::new();
    package.to_stream(&mut output).unwrap();
    output
}

fn add_pivot_relationship(source: Vec<u8>) -> Vec<u8> {
    let mut package = OpcPackage::from_reader(Cursor::new(source)).unwrap();
    package
        .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.bin").unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            relationship_type::PIVOT_TABLE.to_owned(),
            "../pivotTables/pivotTable1.bin".to_owned(),
            "rIdPivot".to_owned(),
            true,
        );
    let mut output = Vec::new();
    package.to_stream(&mut output).unwrap();
    output
}

fn zip_member_payload_range(bytes: &[u8], member: &str) -> (usize, usize) {
    const CENTRAL_HEADER_LEN: usize = 46;
    const LOCAL_HEADER_LEN: usize = 30;

    for offset in 0..=bytes.len().saturating_sub(CENTRAL_HEADER_LEN) {
        if bytes.get(offset..offset + 4) != Some(b"PK\x01\x02") {
            continue;
        }
        let central_name_length = usize::from(u16::from_le_bytes(
            bytes[offset + 28..offset + 30].try_into().unwrap(),
        ));
        let central_extra_length = usize::from(u16::from_le_bytes(
            bytes[offset + 30..offset + 32].try_into().unwrap(),
        ));
        let central_comment_length = usize::from(u16::from_le_bytes(
            bytes[offset + 32..offset + 34].try_into().unwrap(),
        ));
        let central_name_start = offset.checked_add(CENTRAL_HEADER_LEN).unwrap_or_else(|| {
            panic!("ZIP member {member:?} central-directory name offset overflows")
        });
        let central_name_end = central_name_start
            .checked_add(central_name_length)
            .unwrap_or_else(|| {
                panic!("ZIP member {member:?} central-directory name length overflows")
            });
        let central_end = central_name_end
            .checked_add(central_extra_length)
            .and_then(|end| end.checked_add(central_comment_length))
            .unwrap_or_else(|| {
                panic!("ZIP member {member:?} central-directory entry length overflows")
            });
        if central_end > bytes.len() {
            panic!("ZIP member {member:?} has a truncated central-directory entry");
        }
        if bytes.get(central_name_start..central_name_end) != Some(member.as_bytes()) {
            continue;
        }

        let compressed_size = usize::try_from(u32::from_le_bytes(
            bytes[offset + 20..offset + 24].try_into().unwrap(),
        ))
        .unwrap_or_else(|_| panic!("ZIP member {member:?} compressed size does not fit usize"));
        let local_header_offset = usize::try_from(u32::from_le_bytes(
            bytes[offset + 42..offset + 46].try_into().unwrap(),
        ))
        .unwrap_or_else(|_| panic!("ZIP member {member:?} local-header offset does not fit usize"));
        let local_header_end = local_header_offset
            .checked_add(LOCAL_HEADER_LEN)
            .unwrap_or_else(|| panic!("ZIP member {member:?} local-header offset overflows"));
        let local_header = bytes
            .get(local_header_offset..local_header_end)
            .unwrap_or_else(|| panic!("ZIP member {member:?} has a truncated local header"));
        if local_header.get(..4) != Some(b"PK\x03\x04") {
            panic!("ZIP member {member:?} central entry points to a non-local header");
        }
        let local_name_length =
            usize::from(u16::from_le_bytes(local_header[26..28].try_into().unwrap()));
        let local_extra_length =
            usize::from(u16::from_le_bytes(local_header[28..30].try_into().unwrap()));
        let payload_start = local_header_end
            .checked_add(local_name_length)
            .and_then(|start| start.checked_add(local_extra_length))
            .unwrap_or_else(|| panic!("ZIP member {member:?} payload offset overflows"));
        let payload_end = payload_start
            .checked_add(compressed_size)
            .unwrap_or_else(|| panic!("ZIP member {member:?} payload length overflows"));
        if payload_end > bytes.len() {
            panic!("ZIP member {member:?} has a truncated compressed payload");
        }
        return (payload_start, payload_end);
    }
    panic!("ZIP member {member:?} has no central-directory entry");
}

struct TablePayloadSource {
    bytes: Vec<u8>,
    table_ranges: Vec<(usize, usize)>,
    table_payload_reads: Vec<AtomicUsize>,
    cancellation_source: Option<CancellationSource>,
    cancel_table_index: AtomicUsize,
    mutate_table_index: AtomicUsize,
    revision: AtomicU64,
}

impl TablePayloadSource {
    fn new(
        bytes: Vec<u8>,
        table_members: &[&str],
        cancellation_source: Option<CancellationSource>,
    ) -> Self {
        let table_ranges = table_members
            .iter()
            .map(|member| zip_member_payload_range(&bytes, member))
            .collect();
        let table_payload_reads = table_members.iter().map(|_| AtomicUsize::new(0)).collect();
        Self {
            bytes,
            table_ranges,
            table_payload_reads,
            cancellation_source,
            cancel_table_index: AtomicUsize::new(usize::MAX),
            mutate_table_index: AtomicUsize::new(usize::MAX),
            revision: AtomicU64::new(0),
        }
    }

    fn table_payload_reads(&self) -> usize {
        self.table_payload_reads
            .iter()
            .map(|counter| counter.load(Ordering::SeqCst))
            .sum()
    }

    fn table_payload_reads_for(&self, index: usize) -> usize {
        self.table_payload_reads
            .get(index)
            .map_or(0, |counter| counter.load(Ordering::SeqCst))
    }

    fn arm_cancel_on_table(&self, index: usize) {
        self.cancel_table_index.store(index, Ordering::SeqCst);
    }

    fn arm_mutation_on_table(&self, index: usize) {
        self.mutate_table_index.store(index, Ordering::SeqCst);
    }

    fn reset_revision(&self) {
        self.revision.store(0, Ordering::SeqCst);
    }
}

impl ReadAt for TablePayloadSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset)
            .map_err(|_error| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let end = offset.saturating_add(output.len()).min(self.bytes.len());
        let table_index = self
            .table_ranges
            .iter()
            .position(|(start, finish)| offset < *finish && *start < end);
        if let Some(table_index) = table_index {
            self.table_payload_reads[table_index].fetch_add(1, Ordering::SeqCst);
            if self
                .cancel_table_index
                .compare_exchange(table_index, usize::MAX, Ordering::SeqCst, Ordering::SeqCst)
                .is_ok()
                && let Some(cancellation_source) = &self.cancellation_source
            {
                cancellation_source.cancel();
            }
            if self
                .mutate_table_index
                .compare_exchange(table_index, usize::MAX, Ordering::SeqCst, Ordering::SeqCst)
                .is_ok()
            {
                self.revision.fetch_add(1, Ordering::SeqCst);
            }
        }
        output[..end - offset].copy_from_slice(&self.bytes[offset..end]);
        Ok(end - offset)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x5442_4c45,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

#[test]
fn source_catalog_defers_table_payload_and_materializes_cross_sheet_reference() {
    let (formula, rgce) = resident_table_formula(8);
    let bytes = table_workbook_source(
        Some((
            "application/vnd.ms-excel.table".to_owned(),
            table_part_payload(7, "LocalTable"),
            false,
        )),
        Some((
            "application/vnd.ms-excel.table".to_owned(),
            table_part_payload(8, "SalesTable"),
            false,
        )),
        Some(formula),
    );
    let source = Arc::new(TablePayloadSource::new(
        bytes,
        &["xl/tables/table1.bin", "xl/tables/table2.bin"],
        None,
    ));
    let cache_limits = SourceCacheLimits::new(1, 1).unwrap();
    let workbook =
        SourceBackedWorkbook::from_read_at_with_cache_limits(source.clone(), cache_limits).unwrap();
    assert_eq!(source.table_payload_reads(), 0);
    let catalog = workbook.cache_diagnostics();
    assert_eq!(catalog.cold_loads, 1);

    let selected = workbook.worksheet_by_index(0).unwrap().unwrap();
    let worksheet = selected.materialize().unwrap();
    assert!(source.table_payload_reads() > 0);
    let cell = worksheet.get_cell(0, 0).unwrap();
    assert!(matches!(
        cell.value(),
        CellValue::Formula { formula, .. } if formula == "SalesTable[[#Headers],[#Data],[Amount]]"
    ));
    assert_eq!(cell.cached_value(), Some(&CellValue::Float(42.0)));
    assert_eq!(cell.formula_bytes(), Some(rgce.as_slice()));
    assert_eq!(cell.raw_formula_bytes(), None);

    let first = source.table_payload_reads();
    let first_diagnostics = workbook.cache_diagnostics();
    selected.materialize().unwrap();
    assert_eq!(source.table_payload_reads(), first);
    let second_diagnostics = workbook.cache_diagnostics();
    assert!(second_diagnostics.cold_loads > first_diagnostics.cold_loads);
    assert_eq!(second_diagnostics.retained_entries, 0);
}

#[test]
fn source_table_wrong_content_type_is_rejected_during_catalog_open() {
    let source = table_workbook_source(
        Some((
            "application/octet-stream".to_owned(),
            table_part_payload(7, "SalesTable"),
            false,
        )),
        None,
        None,
    );
    assert!(matches!(
        SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(source))),
        Err(PackageError::InvalidContentType { .. })
    ));
}

#[test]
fn source_malformed_and_duplicate_tables_fail_during_materialization() {
    let malformed = table_workbook_source(
        Some((
            "application/vnd.ms-excel.table".to_owned(),
            malformed_table_part_payload(),
            false,
        )),
        None,
        None,
    );
    let malformed_workbook =
        SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(malformed))).unwrap();
    let error = malformed_workbook
        .worksheet_by_index(0)
        .unwrap()
        .unwrap()
        .materialize()
        .unwrap_err();
    assert!(
        matches!(error, PackageError::InvalidLength { .. }),
        "unexpected malformed-table error: {error:?}"
    );

    for (first_id, first_name, second_id, second_name) in [
        (7, "SalesTable", 7, "OtherTable"),
        (7, "SalesTable", 8, "SalesTable"),
    ] {
        let duplicate = table_workbook_source(
            Some((
                "application/vnd.ms-excel.table".to_owned(),
                table_part_payload(first_id, first_name),
                false,
            )),
            Some((
                "application/vnd.ms-excel.table".to_owned(),
                table_part_payload(second_id, second_name),
                false,
            )),
            None,
        );
        let duplicate_workbook =
            SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(duplicate))).unwrap();
        assert!(matches!(
            duplicate_workbook
                .worksheet_by_index(0)
                .unwrap()
                .unwrap()
                .materialize(),
            Err(PackageError::InvalidFormula(_))
        ));
    }
}

#[test]
fn source_table_cancellation_does_not_publish_partial_cache() {
    let bytes = table_workbook_source(
        Some((
            "application/vnd.ms-excel.table".to_owned(),
            table_part_payload(7, "LocalTable"),
            false,
        )),
        Some((
            "application/vnd.ms-excel.table".to_owned(),
            table_part_payload(8, "SalesTable"),
            false,
        )),
        None,
    );
    let (cancellation_source, context) = managed_context();
    let source = Arc::new(TablePayloadSource::new(
        bytes,
        &["xl/tables/table1.bin", "xl/tables/table2.bin"],
        Some(cancellation_source),
    ));
    let workbook = SourceBackedWorkbook::from_read_at_with_execution_context(
        source.clone(),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    let selected = workbook.worksheet_by_index(0).unwrap().unwrap();
    source.arm_cancel_on_table(1);
    assert_cancelled(selected.materialize());
    assert!(source.table_payload_reads_for(0) > 0);
    assert!(source.table_payload_reads_for(1) > 0);
}

#[test]
fn source_table_mutation_does_not_publish_partial_cache_and_retry_succeeds() {
    let (formula, _rgce) = resident_table_formula(8);
    let bytes = table_workbook_source(
        Some((
            "application/vnd.ms-excel.table".to_owned(),
            table_part_payload(7, "LocalTable"),
            false,
        )),
        Some((
            "application/vnd.ms-excel.table".to_owned(),
            table_part_payload(8, "SalesTable"),
            false,
        )),
        Some(formula),
    );
    let source = Arc::new(TablePayloadSource::new(
        bytes,
        &["xl/tables/table1.bin", "xl/tables/table2.bin"],
        None,
    ));
    let workbook = SourceBackedWorkbook::from_read_at(source.clone()).unwrap();
    let selected = workbook.worksheet_by_index(0).unwrap().unwrap();
    source.arm_mutation_on_table(1);
    assert!(matches!(
        selected.materialize(),
        Err(PackageError::Opc(OpcError::SourceChanged { .. }))
    ));
    let failed_reads = source.table_payload_reads();
    assert!(source.table_payload_reads_for(0) > 0);
    assert!(source.table_payload_reads_for(1) > 0);
    let failed_second_reads = source.table_payload_reads_for(1);

    source.reset_revision();
    let worksheet = selected.materialize().unwrap();
    assert!(source.table_payload_reads() > failed_reads);
    assert!(source.table_payload_reads_for(1) > failed_second_reads);
    assert!(worksheet.get_cell(0, 0).unwrap().is_formula());
}

#[test]
fn source_external_and_pivot_table_dependencies_remain_typed_refusals() {
    let external = table_workbook_source(
        Some((
            "application/vnd.ms-excel.table".to_owned(),
            table_part_payload(7, "SalesTable"),
            true,
        )),
        None,
        None,
    );
    assert!(matches!(
        SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(external))),
        Err(PackageError::InvalidRelationship(_))
    ));

    let pivot = add_pivot_relationship(table_workbook_source(
        Some((
            "application/vnd.ms-excel.table".to_owned(),
            table_part_payload(7, "SalesTable"),
            false,
        )),
        None,
        None,
    ));
    let pivot_workbook =
        SourceBackedWorkbook::from_read_at(Arc::new(OwnedSource::new(pivot))).unwrap();
    assert!(matches!(
        pivot_workbook
            .worksheet_by_index(0)
            .unwrap()
            .unwrap()
            .materialize(),
        Err(PackageError::UnsupportedFeature(_))
    ));
}
