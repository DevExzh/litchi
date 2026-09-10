use std::io;
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;
use std::sync::atomic::{AtomicUsize, Ordering};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits,
    OwnedSource, Position, ReadAt, Resource, SourceVersion,
};
use litchi_docx::{Error, ReadLimits, source_backed};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{
    BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, Part, SourceCacheLimits, TargetMode,
};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const FINITE_INPUT_BYTES: u64 = 64 * 1024 * 1024;
const FINITE_OUTPUT_BYTES: u64 = 64 * 1024 * 1024;
const FINITE_OBJECTS: u64 = 1_000_000;
const FINITE_DEPTH: u64 = 1024;
const FINITE_WORK: u64 = 1 << 30;
const SOURCE_DOCUMENT_SCAN_WORKSPACE_BASE: u64 = 131_072;
const SOURCE_DOCUMENT_SCAN_WORKSPACE_PER_BYTE: u64 = 32;

fn source_document_scan_workspace(xml_len: usize) -> u64 {
    (xml_len as u64)
        .saturating_mul(SOURCE_DOCUMENT_SCAN_WORKSPACE_PER_BYTE)
        .saturating_add(SOURCE_DOCUMENT_SCAN_WORKSPACE_BASE)
}

fn source_document_index_admission(xml_len: usize) -> u64 {
    (xml_len as u64 / 4 + 1)
        .saturating_mul(24)
        .saturating_add(1024)
}

fn fixture() -> Vec<u8> {
    let document = format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>managed</w:t></w:r></w:p><w:p><w:r><w:t>read</w:t></w:r></w:p></w:body></w:document>"#
    )
    .into_bytes();
    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/word/document.xml").unwrap(),
            ct::WML_DOCUMENT_MAIN.to_owned(),
            document,
        )))
        .unwrap();
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    PackageWriter::to_bytes(&package).unwrap()
}

fn source() -> Arc<dyn ReadAt> {
    Arc::new(OwnedSource::new(fixture()))
}

fn settings_fixture() -> Vec<u8> {
    let document = format!(r#"<w:document xmlns:w="{W}"><w:body/></w:document>"#).into_bytes();
    let settings = format!(r#"<w:settings xmlns:w="{W}"/>"#).into_bytes();
    let mut package = OpcPackage::new();
    let mut main = BlobPart::new(
        PackURI::new("/word/document.xml").unwrap(),
        ct::WML_DOCUMENT_MAIN.to_owned(),
        document,
    );
    main.rels_mut()
        .try_add_relationship(
            rt::SETTINGS.to_owned(),
            "settings.xml".to_owned(),
            "rSettings".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    package.try_add_part(Box::new(main)).unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/word/settings.xml").unwrap(),
            ct::WML_SETTINGS.to_owned(),
            settings,
        )))
        .unwrap();
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    PackageWriter::to_bytes(&package).unwrap()
}

struct CountingSource {
    bytes: Vec<u8>,
    main_payload: std::ops::Range<usize>,
    payload_reads: AtomicUsize,
}

impl CountingSource {
    fn new(bytes: Vec<u8>) -> Self {
        let main_payload = payload_range(&bytes, "word/document.xml");
        Self {
            bytes,
            main_payload,
            payload_reads: AtomicUsize::new(0),
        }
    }

    fn payload_reads(&self) -> usize {
        self.payload_reads.load(Ordering::SeqCst)
    }
}

impl ReadAt for CountingSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let end = offset.saturating_add(output.len()).min(self.bytes.len());
        let requested = offset..end;
        if requested.start < self.main_payload.end && self.main_payload.start < requested.end {
            self.payload_reads.fetch_add(1, Ordering::SeqCst);
        }
        output[..end - offset].copy_from_slice(&self.bytes[offset..end]);
        Ok(end - offset)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(23, 0))
    }
}

fn payload_range(zip: &[u8], name: &str) -> std::ops::Range<usize> {
    let name = name.as_bytes();
    for (offset, _) in zip
        .windows(4)
        .enumerate()
        .filter(|(_, signature)| *signature == b"PK\x01\x02")
    {
        if offset + 46 > zip.len() {
            continue;
        }
        let compressed =
            u32::from_le_bytes(zip[offset + 20..offset + 24].try_into().unwrap()) as usize;
        let name_len =
            u16::from_le_bytes(zip[offset + 28..offset + 30].try_into().unwrap()) as usize;
        if offset + 46 + name_len > zip.len() || &zip[offset + 46..offset + 46 + name_len] != name {
            continue;
        }
        let local = u32::from_le_bytes(zip[offset + 42..offset + 46].try_into().unwrap()) as usize;
        let local_name =
            u16::from_le_bytes(zip[local + 26..local + 28].try_into().unwrap()) as usize;
        let local_extra =
            u16::from_le_bytes(zip[local + 28..local + 30].try_into().unwrap()) as usize;
        let start = local + 30 + local_name + local_extra;
        return start..start + compressed;
    }
    panic!("ZIP member was not found: {name:?}");
}

fn context(memory: u64) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "docx-managed-source-test",
        Limits::new(
            memory,
            FINITE_INPUT_BYTES,
            FINITE_OUTPUT_BYTES,
            FINITE_OBJECTS,
            FINITE_DEPTH,
            FINITE_WORK,
        ),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::MIN,
        NonZeroUsize::MIN,
        NonZeroU64::new(memory.max(1)).unwrap(),
        0,
    )
    .unwrap();
    (
        budget.clone(),
        cancellation_source,
        ExecutionContext::new(budget, cancellation, execution_limits),
    )
}

fn managed(memory: u64) -> (Budget, CancellationSource, source_backed::Package) {
    let (budget, cancellation_source, context) = context(memory);
    let package =
        source_backed::Package::from_read_at_with_limits_and_cache_limits_and_execution_context(
            source(),
            ReadLimits::default(),
            SourceCacheLimits::new(1 << 20, 8).unwrap(),
            context,
        )
        .unwrap();
    (budget, cancellation_source, package)
}

fn managed_bytes(bytes: Vec<u8>) -> (Budget, CancellationSource, source_backed::Package) {
    let (budget, cancellation_source, context) = context(bytes.len() as u64);
    let package = source_backed::Package::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(bytes)),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    (budget, cancellation_source, package)
}

#[test]
fn managed_document_read_retains_budgeted_part_data_and_supports_selective_queries() {
    let (budget, cancellation_source, package) = managed(1 << 20);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert!(package.cache_diagnostics().budget_managed);

    let document = package.document().unwrap();
    assert_eq!(document.extract_text().unwrap(), "managedread");
    assert_eq!(document.paragraph_count().unwrap(), 2);
    assert_eq!(document.paragraph_text(1).unwrap().as_deref(), Some("read"));
    assert_eq!(
        document.source_version().id(),
        package.source_version().unwrap().id()
    );
    assert_eq!(document.source_version().revision(), 0);
    assert!(budget.used(Resource::Memory) > 0);

    let error = document.paragraphs().unwrap_err();
    assert!(matches!(
        error,
        Error::UnsafeEdit {
            operation: "document paragraphs",
            ..
        }
    ));
    cancellation_source.cancel();
    assert!(matches!(
        document.extract_text(),
        Err(Error::Opc(OpcError::Cancelled))
    ));
    drop(document);
    // The package-owned clean cache entry retains its reservation until the
    // package itself is dropped.
    assert!(budget.used(Resource::Memory) > 0);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_exact_memory_budget_succeeds_and_one_under_fails_before_publication() {
    let document_bytes = format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>budget</w:t></w:r></w:p></w:body></w:document>"#
    )
    .into_bytes();
    let mut package = OpcPackage::new();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/word/document.xml").unwrap(),
            ct::WML_DOCUMENT_MAIN.to_owned(),
            document_bytes.clone(),
        )))
        .unwrap();
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    let bytes = PackageWriter::to_bytes(&package).unwrap();

    let retained_payload_bytes = document_bytes.len() as u64;
    let parser_workspace = source_document_scan_workspace(document_bytes.len());
    let index_admission = source_document_index_admission(document_bytes.len());
    let retained_memory = retained_payload_bytes.saturating_add(index_admission);
    let exact_limit = retained_memory.saturating_add(parser_workspace);
    let (budget, _source, exact_context) = context(exact_limit);
    let exact_source = Arc::new(CountingSource::new(bytes.clone()));
    let exact = source_backed::Package::from_read_at_with_execution_context(
        exact_source.clone(),
        ReadLimits::default(),
        exact_context,
    )
    .unwrap();
    let document = exact.document().unwrap();
    assert_eq!(document.extract_text().unwrap(), "budget");
    assert_eq!(budget.used(Resource::Memory), retained_memory);
    assert!(exact_source.payload_reads() > 0);
    let document_clone = document.clone();
    drop(exact);
    assert_eq!(budget.used(Resource::Memory), retained_memory);
    drop(document);
    assert_eq!(budget.used(Resource::Memory), retained_memory);
    drop(document_clone);
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _source, one_under_context) = context(exact_limit - 1);
    let one_under_source = Arc::new(CountingSource::new(bytes));
    let one_under = source_backed::Package::from_read_at_with_execution_context(
        one_under_source.clone(),
        ReadLimits::default(),
        one_under_context,
    )
    .unwrap();
    let payload_reads_before = one_under_source.payload_reads();
    assert!(matches!(
        one_under.document(),
        Err(Error::Opc(OpcError::Execution(
            ExecutionError::ResourceLimit(limit),
        ))) if limit.resource == Resource::Memory
    ));
    assert!(
        one_under_source.payload_reads() > payload_reads_before,
        "the one-byte-short limit reaches the retained payload before parser workspace admission"
    );
    // The failed lazy parse leaves the successfully read main Part cached by
    // the package. Its payload remains charged until that package is dropped.
    assert_eq!(budget.used(Resource::Memory), retained_payload_bytes);
    drop(one_under);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_cancellation_is_checked_at_open_and_before_lazy_read() {
    let (budget, cancellation_source, open_context) = context(fixture().len() as u64);
    cancellation_source.cancel();
    assert!(matches!(
        source_backed::Package::from_read_at_with_execution_context(
            source(),
            ReadLimits::default(),
            open_context,
        ),
        Err(Error::Opc(OpcError::Cancelled))
    ));
    assert_eq!(budget.used(Resource::Memory), 0);

    let (_budget, cancellation_source, context) = context(fixture().len() as u64);
    let package = source_backed::Package::from_read_at_with_execution_context(
        source(),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    cancellation_source.cancel();
    assert!(matches!(
        package.document(),
        Err(Error::Opc(OpcError::Cancelled))
    ));
}

#[test]
fn managed_document_variables_check_cancellation_before_relationship_metadata() {
    let bytes = settings_fixture();
    let unmanaged =
        source_backed::Package::from_read_at(Arc::new(OwnedSource::new(bytes.clone()))).unwrap();
    let commit = unmanaged
        .edit_document_variables()
        .unwrap()
        .commit()
        .unwrap();

    let (_budget, cancellation_source, package) = managed_bytes(bytes.clone());
    cancellation_source.cancel();
    assert!(matches!(
        package.document_variables_snapshot(),
        Err(Error::Opc(OpcError::Cancelled))
    ));

    let (_budget, cancellation_source, package) = managed_bytes(bytes.clone());
    cancellation_source.cancel();
    assert!(matches!(
        package.edit_document_variables(),
        Err(Error::Opc(OpcError::Cancelled))
    ));

    let (_budget, cancellation_source, package) = managed_bytes(bytes);
    cancellation_source.cancel();
    let mut output = Vec::new();
    assert!(matches!(
        package.publish_document_variables_commit_to_stream(&mut output, &commit),
        Err(Error::Opc(OpcError::Cancelled))
    ));
    assert!(output.is_empty());
}

#[test]
fn managed_main_document_edits_retain_source_and_candidate_reservations() {
    let bytes = fixture();
    let source = Arc::new(CountingSource::new(bytes.clone()));
    let (budget, _cancellation_source, context) = context(1 << 20);
    let package = source_backed::Package::from_read_at_with_execution_context(
        source.clone(),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    let payload_reads_before = source.payload_reads();
    assert_eq!(budget.used(Resource::Memory), 0);

    let snapshot = package.document_snapshot().unwrap();
    assert!(source.payload_reads() > payload_reads_before);
    let source_reserved = budget.used(Resource::Memory);
    assert!(source_reserved > 0);
    let snapshot_clone = snapshot.clone();
    assert_eq!(budget.used(Resource::Memory), source_reserved);

    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(0), "managed replacement")
        .unwrap();
    let candidate_reserved = budget.used(Resource::Memory);
    assert!(candidate_reserved >= source_reserved);
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());

    let mut output = Vec::new();
    package
        .publish_document_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert!(!output.is_empty());
    drop(commit);
    drop(snapshot_clone);
    drop(snapshot);
    assert_eq!(budget.used(Resource::Memory), 0);
}
