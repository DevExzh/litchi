#![allow(
    clippy::unwrap_used,
    reason = "test assertions use panic-on-failure extraction by design"
)]

use std::io::{self, Cursor};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::{OwnedSource, ReadAt, SourceVersion};
use litchi_opc::constants::relationship_type;
use litchi_opc::{BlobPart, OpcPackage, PackURI, SourceCacheLimits};
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
    litchi_xlsb::comments::write(&mut litchi_xlsb::raw::Writer::new(&mut comments), &[]).unwrap();
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

struct VersionedSource {
    bytes: Vec<u8>,
    revision: AtomicU64,
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
