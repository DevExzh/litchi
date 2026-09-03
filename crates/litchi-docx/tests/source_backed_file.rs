#![cfg(any(unix, windows))]

//! Filesystem-backed source tests for the lazy DOCX facade.

use std::fs::{self, OpenOptions};
use std::io::{Seek, SeekFrom, Write};
use std::path::PathBuf;

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Resource,
};
use litchi_docx::{Error, Package, ReadLimits, source_backed};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, Part, SourceCacheLimits, TargetMode,
};
use tempfile::TempDir;

const MAIN: &str = "word/document.xml";
const HEADER: &str = "word/header1.xml";
const UNUSED: &str = "word/opaque.bin";
const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn fixture(text: &str) -> Vec<u8> {
    let document = format!(
        r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body><w:p><w:r><w:t>{text}</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rHeader"/></w:sectPr></w:body></w:document>"#
    )
    .into_bytes();

    let mut package = OpcPackage::new();
    let mut main = BlobPart::new(
        PackURI::new(format!("/{MAIN}")).unwrap(),
        ct::WML_DOCUMENT_MAIN.to_owned(),
        document,
    );
    main.rels_mut()
        .try_add_relationship(
            rt::HEADER.to_owned(),
            "header1.xml".to_owned(),
            "rHeader".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    package.try_add_part(Box::new(main)).unwrap();
    // The source-backed catalog and section inventory must not inspect this
    // valid but unselected story payload.
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/{HEADER}")).unwrap(),
            ct::WML_HEADER.to_owned(),
            format!(
                r#"<w:hdr xmlns:w="{W}"><w:p><w:r><w:t>inert header</w:t></w:r></w:p></w:hdr>"#
            )
            .into_bytes(),
        )))
        .unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(format!("/{UNUSED}")).unwrap(),
            "application/octet-stream".to_owned(),
            (0_u8..=u8::MAX)
                .cycle()
                .take(64 * 1024)
                .map(|index| index.wrapping_mul(17))
                .collect(),
        )))
        .unwrap();
    package.relate_to(MAIN, rt::OFFICE_DOCUMENT);
    PackageWriter::to_bytes(&package).unwrap()
}

fn write_fixture(directory: &TempDir, text: &str) -> (PathBuf, Vec<u8>) {
    let path = directory.path().join("source.docx");
    let bytes = fixture(text);
    fs::write(&path, &bytes).unwrap();
    (path, bytes)
}

fn package_part_payload<'a>(zip: &'a [u8], name: &str) -> &'a [u8] {
    let name = name.as_bytes();
    for offset in zip
        .windows(4)
        .enumerate()
        .filter_map(|(offset, signature)| (signature == b"PK\x01\x02").then_some(offset))
    {
        if offset + 46 > zip.len() {
            continue;
        }
        let compressed_size =
            u32::from_le_bytes(zip[offset + 20..offset + 24].try_into().unwrap()) as usize;
        let name_length =
            u16::from_le_bytes(zip[offset + 28..offset + 30].try_into().unwrap()) as usize;
        let extra_length =
            u16::from_le_bytes(zip[offset + 30..offset + 32].try_into().unwrap()) as usize;
        if offset + 46 + name_length + extra_length > zip.len()
            || &zip[offset + 46..offset + 46 + name_length] != name
        {
            continue;
        }
        let local_offset =
            u32::from_le_bytes(zip[offset + 42..offset + 46].try_into().unwrap()) as usize;
        let local_name_length = u16::from_le_bytes(
            zip[local_offset + 26..local_offset + 28]
                .try_into()
                .unwrap(),
        ) as usize;
        let local_extra_length = u16::from_le_bytes(
            zip[local_offset + 28..local_offset + 30]
                .try_into()
                .unwrap(),
        ) as usize;
        let data_start = local_offset + 30 + local_name_length + local_extra_length;
        return &zip[data_start..data_start + compressed_size];
    }
    panic!("ZIP member {name:?} was not found");
}

fn managed_context(memory: u64) -> (Budget, ExecutionContext) {
    let budget = Budget::root(
        "docx-file-source-test",
        CoreLimits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (_cancellation_source, cancellation) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        std::num::NonZeroUsize::MIN,
        std::num::NonZeroUsize::MIN,
        std::num::NonZeroU64::new(memory.max(1)).unwrap(),
        0,
    )
    .unwrap();
    (
        budget.clone(),
        ExecutionContext::new(budget, cancellation, limits),
    )
}

#[test]
fn open_and_catalog_leave_payloads_cold_until_selected_queries() {
    let directory = tempfile::tempdir().unwrap();
    let (path, _bytes) = write_fixture(&directory, "filesystem");

    let package = source_backed::Package::open(&path).unwrap();
    let opened = package.cache_diagnostics();
    assert_eq!(opened.cold_loads, 0);
    assert_eq!(opened.successful_loads, 0);
    assert_eq!(opened.retained_entries, 0);

    let sections = package.section_inventory_snapshot().unwrap();
    assert_eq!(sections.inventory().sections().len(), 1);
    assert_eq!(sections.inventory().sections()[0].headers().len(), 1);
    let after_sections = package.cache_diagnostics();
    assert_eq!(after_sections.cold_loads, 1);
    assert_eq!(after_sections.successful_loads, 1);
    assert_eq!(after_sections.retained_entries, 1);

    let document = package.document().unwrap();
    assert_eq!(document.extract_text().unwrap(), "filesystem");
    let after_document = package.cache_diagnostics();
    assert_eq!(after_document.cold_loads, 1);
    assert_eq!(after_document.successful_loads, 1);
    assert_eq!(after_document.hits, after_sections.hits + 1);
}

#[test]
fn from_path_variants_forward_limits_and_managed_cache_policy() {
    let directory = tempfile::tempdir().unwrap();
    let (path, bytes) = write_fixture(&directory, "limits");

    let limits = ReadLimits::builder()
        .max_input_bytes((bytes.len() - 1) as u64)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        source_backed::Package::from_path_with_limits(&path, limits),
        Err(Error::Opc(OpcError::ReadLimit { .. }))
    ));

    let cache_limits = SourceCacheLimits::new(1 << 20, 8).unwrap();
    let package = source_backed::Package::from_path_with_limits_and_cache_limits(
        &path,
        ReadLimits::default(),
        cache_limits,
    )
    .unwrap();
    assert!(!package.cache_diagnostics().budget_managed);
    assert_eq!(
        package.document().unwrap().extract_text().unwrap(),
        "limits"
    );

    let (budget, context) = managed_context(1 << 20);
    let managed =
        source_backed::Package::from_path_with_limits_and_cache_limits_and_execution_context(
            &path,
            ReadLimits::default(),
            cache_limits,
            context,
        )
        .unwrap();
    assert!(managed.cache_diagnostics().budget_managed);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(
        managed.document().unwrap().extract_text().unwrap(),
        "limits"
    );
    let diagnostics = managed.cache_diagnostics();
    assert_eq!(diagnostics.cold_loads, 1);
    assert_eq!(diagnostics.successful_loads, 1);
    assert!(diagnostics.budget_cache_reserved_bytes > 0);
    assert!(budget.used(Resource::Memory) > 0);
    drop(managed);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn exact_noop_and_changed_document_overlay_preserve_raw_unselected_members() {
    let directory = tempfile::tempdir().unwrap();
    let (path, source_bytes) = write_fixture(&directory, "before");

    let package = source_backed::Package::from_path(&path).unwrap();
    let noop = package.edit_document().unwrap().commit().unwrap();
    let mut output = Vec::new();
    package
        .publish_document_commit_to_stream(&mut output, &noop)
        .unwrap();
    assert_eq!(output, source_bytes);

    let package = source_backed::Package::from_path(&path).unwrap();
    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(litchi_core::Position::new(0), "after")
        .unwrap();
    let commit = edit.commit().unwrap();
    let mut changed = Vec::new();
    package
        .publish_document_commit_to_stream(&mut changed, &commit)
        .unwrap();
    assert_ne!(changed, source_bytes);
    assert_eq!(
        package_part_payload(&changed, UNUSED),
        package_part_payload(&source_bytes, UNUSED)
    );
    let reopened = Package::from_reader(std::io::Cursor::new(changed)).unwrap();
    assert_eq!(
        reopened
            .document_snapshot()
            .unwrap()
            .paragraph(litchi_core::Position::new(0))
            .unwrap()
            .text()
            .unwrap(),
        "after"
    );
}

#[test]
fn replacing_the_path_keeps_the_open_source_pinned() {
    let directory = tempfile::tempdir().unwrap();
    let (path, original) = write_fixture(&directory, "original");
    let replacement_path = directory.path().join("replacement.docx");
    let replacement = fixture("replacement");
    fs::write(&replacement_path, &replacement).unwrap();

    let package = source_backed::Package::from_path(&path).unwrap();
    let version = package.source_version().unwrap();
    fs::remove_file(&path).unwrap();
    fs::rename(&replacement_path, &path).unwrap();

    assert_eq!(package.source_version().unwrap(), version);
    assert_eq!(
        package.document().unwrap().extract_text().unwrap(),
        "original"
    );
    assert_ne!(fs::read(&path).unwrap(), original);
    assert_eq!(
        source_backed::Package::from_path(&path)
            .unwrap()
            .document()
            .unwrap()
            .extract_text()
            .unwrap(),
        "replacement"
    );
}

#[test]
fn in_place_file_change_is_reported_as_source_changed() {
    let directory = tempfile::tempdir().unwrap();
    let (path, bytes) = write_fixture(&directory, "changed");
    let package = source_backed::Package::from_path(&path).unwrap();
    let original_version = package.source_version().unwrap();

    let mut file = OpenOptions::new().write(true).open(&path).unwrap();
    file.seek(SeekFrom::Start(bytes.len() as u64)).unwrap();
    file.write_all(b"x").unwrap();
    file.flush().unwrap();

    assert!(matches!(
        package.source_version(),
        Err(Error::Opc(OpcError::SourceChanged { expected, actual }))
            if expected == original_version && actual.revision() > expected.revision()
    ));
    assert!(matches!(
        package.document(),
        Err(Error::Opc(OpcError::SourceChanged { .. }))
    ));
}

#[test]
fn filesystem_path_errors_are_typed() {
    let directory = tempfile::tempdir().unwrap();
    let missing = directory.path().join("missing.docx");
    assert!(matches!(
        source_backed::Package::from_path(&missing),
        Err(Error::Io(error)) if error.kind() == std::io::ErrorKind::NotFound
    ));
    assert!(matches!(
        source_backed::Package::from_path(directory.path()),
        Err(Error::Io(error))
            if matches!(
                error.kind(),
                std::io::ErrorKind::InvalidInput
                    | std::io::ErrorKind::PermissionDenied
                    | std::io::ErrorKind::IsADirectory
            )
    ));
}
