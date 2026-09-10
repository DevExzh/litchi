#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "the focused fixture uses infallible test setup"
)]

//! Focused DOCX coverage for the opt-in source read-ahead policy.
//!
//! The test keeps the semantic read path managed, then publishes one bounded
//! paragraph append. Publication must permanently return the archive adapter
//! to exact reads and release its retained read-ahead reservation.

use std::io::{self, Cursor};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::{
    Arc,
    atomic::{AtomicUsize, Ordering},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits,
    OwnedSource, ReadAt, Resource, SourceVersion,
};
use litchi_docx::{Package, ReadLimits, source_backed};
use litchi_opc::SourceCacheLimits;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use soapberry_zip::office::StreamingArchiveWriter;

const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const RELATIONSHIPS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPES: &str = "http://schemas.openxmlformats.org/package/2006/content-types";

fn fixture() -> Vec<u8> {
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="{WORD}"><w:body><w:p><w:r><w:t>seed</w:t></w:r></w:p></w:body></w:document>"#
    );
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{}"/></Types>"#,
        ct::WML_DOCUMENT_MAIN
    );
    let relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS}"><Relationship Id="rIdDocument" Type="{}" Target="word/document.xml"/></Relationships>"#,
        rt::OFFICE_DOCUMENT
    );

    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .expect("content types fixture must be writable");
    writer
        .write_stored("_rels/.rels", relationships.as_bytes())
        .expect("package relationships fixture must be writable");
    // Keep the main part outside the initial 4 KiB window so the semantic
    // read exercises the policy rather than succeeding only from the index.
    writer
        .write_stored("word/opaque.bin", &[0xA5; 8192])
        .expect("opaque fixture must be writable");
    writer
        .write_deflated_sized("word/document.xml", document.as_bytes())
        .expect("document fixture must be writable");
    writer
        .finish_to_bytes()
        .expect("fixture archive must finish")
}

#[derive(Debug)]
struct CountingSource {
    inner: OwnedSource,
    calls: AtomicUsize,
}

impl CountingSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            inner: OwnedSource::new(bytes),
            calls: AtomicUsize::new(0),
        }
    }

    fn calls(&self) -> usize {
        self.calls.load(Ordering::SeqCst)
    }
}

impl ReadAt for CountingSource {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.calls.fetch_add(1, Ordering::SeqCst);
        self.inner.read_at(offset, output)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.inner.version()
    }
}

fn managed_context(name: &'static str) -> (Budget, CancellationSource, ExecutionContext) {
    let memory = 64 * 1024 * 1024;
    let budget = Budget::root(
        name,
        CoreLimits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::MIN,
        NonZeroUsize::MIN,
        NonZeroU64::new(memory).expect("managed memory limit must be nonzero"),
        0,
    )
    .expect("managed execution limits must be valid");
    (
        budget.clone(),
        cancellation_source,
        ExecutionContext::new(budget, cancellation, execution_limits),
    )
}

fn append_limits() -> source_backed::TailAppendLimits {
    source_backed::TailAppendLimits {
        max_source_xml_bytes: 2 * 1024 * 1024,
        max_text_bytes: 64 * 1024,
        max_fragment_bytes: 65_536,
        max_candidate_xml_bytes: 2 * 1024 * 1024,
        max_events: 4096,
        max_depth: 32,
        max_paragraphs: 65_536,
        max_settings_xml_bytes: 4096,
        max_workspace_bytes: 2 * 1024 * 1024,
        max_output_bytes: 2 * 1024 * 1024,
        max_token_bytes: 4096,
    }
}

#[test]
fn managed_read_ahead_supports_semantic_read_then_releases_before_docx_publication() {
    let archive = fixture();

    // Establish a same-fixture exact-read control for the semantic payload
    // read. The comparison is scoped to the lazy document operation, after
    // each package has completed its mandatory source indexing.
    let (exact_budget, _exact_cancellation, exact_context) = managed_context("docx-exact-control");
    let exact_source = Arc::new(CountingSource::new(archive.clone()));
    let exact_package =
        source_backed::Package::from_read_at_with_limits_and_cache_limits_and_execution_context(
            exact_source.clone(),
            ReadLimits::default(),
            SourceCacheLimits::default(),
            exact_context,
        )
        .expect("exact managed DOCX fixture must open");
    let exact_before = exact_source.calls();
    assert_eq!(
        exact_package
            .document()
            .expect("exact semantic read")
            .extract_text()
            .expect("exact text"),
        "seed"
    );
    let exact_document_calls = exact_source.calls() - exact_before;
    assert!(
        exact_document_calls > 0,
        "the exact control must read the payload"
    );
    drop(exact_package);
    assert_eq!(exact_budget.used(Resource::Memory), 0);

    let (budget, _cancellation, context) = managed_context("docx-read-ahead-publication");
    let candidate_source = Arc::new(CountingSource::new(archive));
    let source_read_policy = source_backed::SourceReadPolicy::forward_start(4096)
        .expect("focused read-ahead policy must be valid");
    let package = source_backed::Package::from_read_at_with_limits_and_cache_limits_and_source_read_policy_and_execution_context(
        candidate_source.clone(),
        ReadLimits::default(),
        SourceCacheLimits::default(),
        source_read_policy,
        context,
    )
    .expect("managed read-ahead DOCX fixture must open");

    let opened = package
        .source_read_diagnostics()
        .expect("read-ahead diagnostics")
        .expect("opt-in policy diagnostics");
    assert!(opened.enabled);
    assert_eq!(opened.configured_window_bytes, 4096);
    assert_eq!(opened.retained_window_bytes, 4096);

    let document_before = candidate_source.calls();
    let document = package.document().expect("semantic read with read-ahead");
    assert_eq!(document.extract_text().expect("semantic text"), "seed");
    assert_eq!(document.paragraph_count().expect("paragraph count"), 1);
    drop(document);
    let candidate_document_calls = candidate_source.calls() - document_before;
    assert!(
        candidate_document_calls <= exact_document_calls,
        "read-ahead should not require more physical calls for this range: candidate={candidate_document_calls}, exact={exact_document_calls}"
    );
    assert!(
        package
            .source_read_diagnostics()
            .expect("diagnostics after semantic read")
            .expect("opt-in diagnostics")
            .enabled,
        "semantic reads must leave the explicit policy enabled until publication"
    );

    let memory_before_prepare = budget.used(Resource::Memory);
    let retained_window = opened.retained_window_bytes as u64;
    assert!(
        memory_before_prepare >= retained_window,
        "the managed budget must include the retained read-ahead window"
    );
    let plan = package
        .tail_append_plain_paragraph("tail")
        .with_limits(append_limits())
        .prepare()
        .expect("simple DOCX append must prepare");
    assert!(!plan.is_noop(), "the append must be a real edit");
    let prepared = package
        .source_read_diagnostics()
        .expect("diagnostics after preparation")
        .expect("opt-in diagnostics");
    assert!(
        !prepared.enabled,
        "preparing a publication plan must permanently select exact reads"
    );
    assert_eq!(prepared.retained_window_bytes, 0);
    let mut output = Vec::new();
    let publication = plan
        .write_to_stream(&mut output)
        .expect("DOCX publication must succeed");
    assert!(!publication.is_noop());
    drop(publication);

    let after = package
        .source_read_diagnostics()
        .expect("publication diagnostics")
        .expect("opt-in diagnostics");
    assert!(
        !after.enabled,
        "publication must permanently select exact reads"
    );
    assert_eq!(after.retained_window_bytes, 0);

    let reopened = Package::from_reader(Cursor::new(output)).expect("published DOCX must reopen");
    assert_eq!(
        reopened
            .document()
            .expect("published main document")
            .text()
            .expect("published text"),
        "seedtail"
    );
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}
