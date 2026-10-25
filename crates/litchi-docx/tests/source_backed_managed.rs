use std::io;
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;
use std::sync::atomic::{AtomicUsize, Ordering};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits,
    OwnedSource, Position, ReadAt, Resource, SourceVersion,
};
use litchi_docx::{Error, SourceLimits, source_backed};
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
    let package = source_backed::Package::from_read_at_with_options(
        source(),
        litchi_docx::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .cache_limits(SourceCacheLimits::new(1 << 20, 8).unwrap())
            .execution_context(context),
    )
    .unwrap();
    (budget, cancellation_source, package)
}

fn managed_bytes(bytes: Vec<u8>) -> (Budget, CancellationSource, source_backed::Package) {
    let (budget, cancellation_source, context) = context(bytes.len() as u64);
    let package = source_backed::Package::from_read_at_with_options(
        Arc::new(OwnedSource::new(bytes)),
        litchi_docx::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(context),
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
    assert!(document.revisions().unwrap().is_empty());
    assert!(matches!(
        package
            .with_semantic_limits(litchi_docx::SemanticLimits::default().revisions(
                litchi_docx::revision::Limits {
                    max_source_bytes: 0,
                    ..litchi_docx::revision::Limits::default()
                }
            ))
            .document()
            .unwrap()
            .revisions(),
        Err(Error::RevisionLimit {
            kind: litchi_docx::revision::LimitKind::SourceBytes,
            ..
        })
    ));
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
            Error::Refused(refusal) if refusal.operation() == "document paragraphs"
    ));
    cancellation_source.cancel();
    assert!(matches!(
        document.extract_text(),
        Err(Error::Opc(OpcError::Cancelled))
    ));
    assert!(matches!(
        document.revisions(),
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
    let exact = source_backed::Package::from_read_at_with_options(
        exact_source.clone(),
        litchi_docx::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(exact_context),
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
    let one_under = source_backed::Package::from_read_at_with_options(
        one_under_source.clone(),
        litchi_docx::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(one_under_context),
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

fn is_memory_limit(error: &Error) -> bool {
    matches!(
        error,
        Error::Opc(OpcError::Execution(ExecutionError::ResourceLimit(limit)))
            if limit.resource == Resource::Memory
    )
}

// Review of the 0759 merge: the incoming `revisions()` parsed the whole
// managed document without the parser admission that `extract_text` and the
// other document queries reserve, so it succeeded where they were refused.
#[test]
fn managed_revisions_query_is_admitted_and_charged_like_extract_text() {
    let mut body = String::new();
    for index in 0..8 {
        body.push_str(&format!(
            r#"<w:p><w:r><w:t>paragraph {index}</w:t></w:r><w:ins w:id="{index}" w:author="A"><w:r><w:t>x</w:t></w:r></w:ins></w:p>"#
        ));
    }
    let document_bytes =
        format!(r#"<w:document xmlns:w="{W}"><w:body>{body}</w:body></w:document>"#).into_bytes();
    let mut opc = OpcPackage::new();
    opc.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/word/document.xml").unwrap(),
        ct::WML_DOCUMENT_MAIN.to_owned(),
        document_bytes.clone(),
    )))
    .unwrap();
    opc.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    let (budget, _cancellation_source, context) = context(1 << 20);
    let package = source_backed::Package::from_read_at_with_options(
        Arc::new(OwnedSource::new(PackageWriter::to_bytes(&opc).unwrap())),
        litchi_docx::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(context),
    )
    .unwrap();
    let document = package.document().unwrap();

    // Leave exactly one parser workspace free: both queries are admitted and
    // release their admission when they return.
    let workspace = source_document_scan_workspace(document_bytes.len());
    let free = budget
        .limit(Resource::Memory)
        .saturating_sub(budget.used(Resource::Memory));
    assert!(free > workspace);
    let hold = budget.reserve(Resource::Memory, free - workspace).unwrap();
    let held = budget.used(Resource::Memory);
    assert_eq!(document.extract_text().unwrap().matches('x').count(), 8);
    assert_eq!(budget.used(Resource::Memory), held);
    assert_eq!(document.revisions().unwrap().len(), 8);
    assert_eq!(budget.used(Resource::Memory), held);

    // One byte short of that workspace: both are refused, identically.
    let short = budget.reserve(Resource::Memory, 1).unwrap();
    let text = document.extract_text().unwrap_err();
    assert!(is_memory_limit(&text), "{text:?}");
    let revisions = match document.revisions() {
        Ok(revisions) => panic!(
            "revisions() parsed {} revisions without the parser admission extract_text was refused",
            revisions.len()
        ),
        Err(error) => error,
    };
    assert!(is_memory_limit(&revisions), "{revisions:?}");
    assert_eq!(revisions.to_string(), text.to_string());

    drop(short);
    assert_eq!(document.revisions().unwrap().len(), 8);
    drop(hold);
    drop(document);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_cancellation_is_checked_at_open_and_before_lazy_read() {
    let (budget, cancellation_source, open_context) = context(fixture().len() as u64);
    cancellation_source.cancel();
    assert!(matches!(
        source_backed::Package::from_read_at_with_options(
            source(),
            litchi_docx::OpenOptions::default()
                .source_limits(SourceLimits::default())
                .execution_context(open_context)
        ),
        Err(Error::Opc(OpcError::Cancelled))
    ));
    assert_eq!(budget.used(Resource::Memory), 0);

    let (_budget, cancellation_source, context) = context(fixture().len() as u64);
    let package = source_backed::Package::from_read_at_with_options(
        source(),
        litchi_docx::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(context),
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
        package
            .apply_document_variables(commit.patch())
            .and_then(|committed| committed.package().write_to(&mut output)),
        Err(Error::Opc(OpcError::Cancelled))
    ));
    assert!(output.is_empty());
}

#[test]
fn managed_main_document_edits_retain_source_and_candidate_reservations() {
    let bytes = fixture();
    let source = Arc::new(CountingSource::new(bytes.clone()));
    let (budget, _cancellation_source, context) = context(1 << 20);
    let package = source_backed::Package::from_read_at_with_options(
        source.clone(),
        litchi_docx::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(context),
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

    let mut edit = package.edit().unwrap();
    edit.replace_paragraph_text(Position::new(0), "managed replacement")
        .unwrap();
    let candidate_reserved = budget.used(Resource::Memory);
    assert!(candidate_reserved >= source_reserved);
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());

    let mut output = Vec::new();
    package
        .apply(commit.patch())
        .and_then(|committed| {
            let committed = committed.into_package();
            committed.write_to(&mut output)?;
            committed.document_snapshot()
        })
        .unwrap();
    assert!(!output.is_empty());
    drop(commit);
    drop(snapshot_clone);
    drop(snapshot);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn cloned_open_options_share_budget_and_cancellation() {
    let (budget, cancellation, context) = context(4 << 20);
    let options = litchi_docx::OpenOptions::default()
        .cache_limits(SourceCacheLimits::new(1 << 20, 8).unwrap())
        .execution_context(context);
    let first =
        source_backed::Package::from_read_at_with_options(source(), options.clone()).unwrap();
    let second = source_backed::Package::from_read_at_with_options(source(), options).unwrap();
    let first_document = first.document().unwrap();
    let first_memory = budget.used(Resource::Memory);
    assert!(first_memory > 0);
    let second_document = second.document().unwrap();
    assert!(budget.used(Resource::Memory) > first_memory);
    cancellation.cancel();
    assert!(first.document().is_err());
    assert!(second.document().is_err());
    drop((first, second, first_document, second_document));
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn owning_ingress_refuses_lazy_options_before_io() {
    use litchi_docx::{OpenOptions, OpenPolicy, Package};
    struct NeverRead;
    impl io::Read for NeverRead {
        fn read(&mut self, _: &mut [u8]) -> io::Result<usize> {
            panic!("unsupported policies must fail before input I/O");
        }
    }
    let (_, _, context) = context(1 << 20);
    let cases = [
        (
            OpenOptions::default().execution_context(context),
            OpenPolicy::ExecutionContext,
        ),
        (
            OpenOptions::default().cache_limits(SourceCacheLimits::default()),
            OpenPolicy::CacheLimits,
        ),
        (
            OpenOptions::default().source_read_policy(litchi_opc::SourceReadPolicy::exact()),
            OpenPolicy::SourceReadPolicy,
        ),
    ];
    let directory = tempfile::tempdir().unwrap();
    for (options, expected) in cases {
        let assert_policy = |result| {
            assert!(matches!(result,
            Err(Error::UnsupportedOpenPolicy { policy }) if policy == expected))
        };
        assert_policy(Package::from_reader_with_options(
            NeverRead,
            options.clone(),
        ));
        assert_policy(Package::from_bytes_with_options(&[], options.clone()));
        assert_policy(Package::open_with_options(
            directory.path(),
            options.clone(),
        ));
        #[cfg(feature = "encryption")]
        {
            let mut clear = Vec::new();
            Package::new().unwrap().write_to(&mut clear).unwrap();
            let decrypted = litchi_docx::encryption::open(
                clear,
                "pw",
                &litchi_docx::encryption::Limits::default(),
            )
            .unwrap();
            assert_policy(Package::from_decrypted(decrypted, options.clone()));
            assert_policy(Package::from_reader_with_password_and_options(
                NeverRead,
                "pw",
                options.clone(),
            ));
            assert_policy(Package::from_bytes_with_password_and_options(
                &[],
                "pw",
                options.clone(),
            ));
            assert_policy(Package::open_with_password_and_options(
                directory.path(),
                "pw",
                options.clone(),
            ));
        }
    }
}

#[cfg(any(unix, windows))]
#[test]
fn cancelled_path_options_fail_before_opening_the_file() {
    let (_, cancellation, context) = context(1 << 20);
    cancellation.cancel();
    let directory = tempfile::tempdir().unwrap();
    let error = source_backed::Package::open_with_options(
        directory.path().join("missing.docx"),
        litchi_docx::OpenOptions::default().execution_context(context),
    )
    .err()
    .unwrap();
    assert!(matches!(error, Error::Opc(OpcError::Cancelled)));
}

#[test]
fn sequential_ingress_retains_input_charges_with_source_artifacts_and_clones() {
    let bytes = fixture();
    let (budget, _, context) = context(4 << 20);
    let options = litchi_docx::OpenOptions::default()
        .source_limits(
            SourceLimits::builder()
                .max_input_bytes(bytes.len() as u64)
                .unwrap()
                .build()
                .unwrap(),
        )
        .cache_limits(SourceCacheLimits::new(1 << 20, 8).unwrap())
        .execution_context(context);
    let package = source_backed::Package::from_reader_with_options(&bytes[..], options).unwrap();
    let input_memory = budget.used(Resource::Memory);
    assert!(input_memory >= bytes.len() as u64);
    // Resource::InputBytes counts external ingestion and later OPC positional
    // reads; SourceLimits::InputBytes above still bounds the archive length.
    assert!(budget.used(Resource::InputBytes) > bytes.len() as u64);
    assert_eq!(package.cache_diagnostics().cold_loads, 0);
    let clone = package.clone();
    let artifact = package.source_backed_package().unwrap().source_artifact();
    assert_eq!(
        package.document().unwrap().extract_text().unwrap(),
        "managedread"
    );
    drop((package, clone));
    assert_eq!(budget.used(Resource::Memory), input_memory);
    let mut output = Vec::new();
    artifact.write_to(&mut output).unwrap();
    assert_eq!(output, bytes);
    drop(artifact);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn sequential_options_preserve_read_ahead_and_publication_release() {
    let bytes = fixture();
    let (budget, _, context) = context(4 << 20);
    let package = source_backed::Package::from_bytes_with_options(
        &bytes,
        litchi_docx::OpenOptions::default()
            .source_read_policy(litchi_opc::SourceReadPolicy::forward_start(4096).unwrap())
            .execution_context(context),
    )
    .unwrap();
    assert!(package.source_read_diagnostics().unwrap().unwrap().enabled);
    let mut output = Vec::new();
    package.write_to(&mut output).unwrap();
    assert_eq!(output, bytes);
    assert!(!package.source_read_diagnostics().unwrap().unwrap().enabled);
    assert!(budget.used(Resource::Memory) >= bytes.len() as u64);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn sequential_ingress_failure_releases_retained_input() {
    let (budget, _, context) = context(4 << 20);
    assert!(
        source_backed::Package::from_bytes_with_options(
            b"not a zip",
            litchi_docx::OpenOptions::default().execution_context(context),
        )
        .is_err()
    );
    assert_eq!(budget.used(Resource::Memory), 0);
    assert!(budget.used(Resource::InputBytes) >= 9);
}

#[test]
fn sequential_default_and_cache_options_preserve_archive_limits_and_bytes() {
    let bytes = fixture();
    source_backed::Package::from_bytes(&bytes).unwrap();
    for options in [
        litchi_docx::OpenOptions::default(),
        litchi_docx::OpenOptions::default().cache_limits(SourceCacheLimits::new(1024, 1).unwrap()),
    ] {
        let limits = |maximum| {
            SourceLimits::builder()
                .max_input_bytes(maximum)
                .unwrap()
                .build()
                .unwrap()
        };
        let package = source_backed::Package::from_bytes_with_options(
            &bytes,
            options.clone().source_limits(limits(bytes.len() as u64)),
        )
        .unwrap();
        let mut output = Vec::new();
        package.write_to(&mut output).unwrap();
        assert_eq!(output, bytes);
        assert!(
            matches!(source_backed::Package::from_reader_with_options(&bytes[..],
            options.source_limits(limits(bytes.len() as u64 - 1))),
            Err(Error::Opc(OpcError::Limit { kind: litchi_opc::SourceLimitKind::InputBytes, observed, maximum }))
                if observed == bytes.len() as u64 && maximum == bytes.len() as u64 - 1)
        );
    }
}
