use std::io::{self, Write};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, OwnedSource, ReadAt,
    SourceVersion, TextOutputError, TextOutputOptions,
};
use litchi_docx::source_backed::Package as SourceBackedPackage;
use litchi_docx::{Package, ReadLimits};
use soapberry_zip::office::StreamingArchiveWriter;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_W: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn document_xml(body: &str) -> Vec<u8> {
    format!(
        r#"<?xml version="1.0"?><w:document xmlns:w="{W}"><w:body>{body}<w:sectPr/></w:body></w:document>"#
    )
    .into_bytes()
}

fn strict_document_xml(body: &str) -> Vec<u8> {
    format!(r#"<w:document xmlns:w="{STRICT_W}"><w:body>{body}</w:body></w:document>"#).into_bytes()
}

fn package_bytes(document: &[u8], with_media: bool) -> Vec<u8> {
    let media_override = if with_media {
        r#"<Override PartName="/word/media/image1.bin" ContentType="application/octet-stream"/>"#
    } else {
        ""
    };
    let content_types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>{media_override}</Types>"#
    );
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .expect("content types");
    writer
        .write_stored(
            "_rels/.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#
                .to_string()
                .as_bytes(),
        )
        .expect("package relationships");
    writer
        .write_stored("word/document.xml", document)
        .expect("main document");
    if with_media {
        writer
            .write_stored("word/media/image1.bin", b"cold media")
            .expect("media");
    }
    writer.finish_to_bytes().expect("finish DOCX archive")
}

fn eager_package(bytes: impl AsRef<[u8]>) -> Package {
    Package::from_reader(io::Cursor::new(bytes.as_ref().to_vec())).expect("eager package")
}

fn eager_text(
    bytes: &[u8],
    options: TextOutputOptions<'_>,
) -> (Vec<u8>, litchi_core::TextOutputReport) {
    let package = eager_package(bytes);
    let document = package.document().expect("eager document");
    let mut output = Vec::new();
    let report = document
        .write_text_to(&mut output, options)
        .expect("eager semantic text");
    (output, report)
}

fn source_text(
    bytes: Vec<u8>,
    options: TextOutputOptions<'_>,
) -> (Vec<u8>, litchi_core::TextOutputReport) {
    let package = SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(bytes)))
        .expect("source-backed package");
    let mut output = Vec::new();
    let report = package
        .write_text_to(&mut output, options)
        .expect("source semantic text");
    (output, report)
}

fn parity_options(include_empty: bool) -> TextOutputOptions<'static> {
    TextOutputOptions::new("\n", "unused", u64::MAX, u64::MAX).with_empty_objects(include_empty)
}

#[test]
fn eager_and_source_paragraph_sinks_have_parity() {
    let bytes = package_bytes(
        &document_xml(
            r#"<w:p><w:r><w:t>one &amp; 世界</w:t><w:tab/><w:t>two</w:t><w:br/><w:t><![CDATA[three]]></w:t></w:r><w:hyperlink><w:r><w:t>four</w:t></w:r></w:hyperlink></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:fldSimple><w:r><w:t>field</w:t></w:r></w:fldSimple><w:t>&#x1F600;</w:t><w:noBreakHyphen/><w:softHyphen/></w:r></w:p>"#,
        ),
        false,
    );
    let expected = "one & 世界\ttwo\nthreefour\ncell\nfield😀‑\u{ad}";
    let (eager, eager_report) = eager_text(&bytes, parity_options(false));
    let (source, source_report) = source_text(bytes, parity_options(false));
    assert_eq!(eager, expected.as_bytes());
    assert_eq!(source, eager);
    assert_eq!(eager_report.bytes_written(), expected.len() as u64);
    assert_eq!(eager_report.objects_written(), 3);
    assert_eq!(source_report, eager_report);
}

#[test]
fn separators_and_empty_paragraph_policy_are_explicit() {
    let bytes = package_bytes(
        &document_xml(
            r#"<w:p><w:r><w:t>A</w:t></w:r></w:p><w:p/><w:p><w:r><w:t>B</w:t></w:r></w:p>"#,
        ),
        false,
    );
    let options = TextOutputOptions::new("::", "unused", u64::MAX, u64::MAX);
    let (included, included_report) = eager_text(&bytes, options);
    assert_eq!(included, b"A::::B");
    assert_eq!(included_report.objects_written(), 3);

    let options =
        TextOutputOptions::new("::", "unused", u64::MAX, u64::MAX).with_empty_objects(false);
    let (excluded, excluded_report) = source_text(bytes, options);
    assert_eq!(excluded, b"A::B");
    assert_eq!(excluded_report.objects_written(), 2);
}

#[test]
fn nested_empty_paragraph_is_rejected_without_output() {
    let bytes = package_bytes(
        &document_xml(r#"<w:p><w:custom><w:p/></w:custom></w:p>"#),
        false,
    );
    let package = eager_package(&bytes);
    let document = package.document().expect("document");
    let mut output = Vec::new();

    let result = document.write_text_to(&mut output, TextOutputOptions::default());

    assert!(matches!(
        result,
        Err(TextOutputError::Document { progress, .. })
            if progress.bytes_written() == 0 && progress.objects_written() == 0
    ));
    assert!(output.is_empty());
}

#[test]
fn strict_namespaces_mce_fallback_and_foreign_text_are_checked() {
    let strict = package_bytes(
        &strict_document_xml(r#"<w:p><w:r><w:t>strict</w:t></w:r></w:p>"#),
        false,
    );
    assert_eq!(eager_text(&strict, parity_options(false)).0, b"strict");

    let mce = format!(
        r#"<w:document xmlns:w="{W}" xmlns:mc="{MCE}" xmlns:x="urn:future" mc:Ignorable="x"><w:body><mc:AlternateContent><mc:Choice Requires="x"><w:p><w:r><w:t>ignored</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>fallback</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent></w:body></w:document>"#
    );
    assert_eq!(
        eager_text(&package_bytes(mce.as_bytes(), false), parity_options(false)).0,
        b"fallback"
    );

    let foreign = package_bytes(
        &document_xml(r#"<w:p><w:r><x:t xmlns:x="urn:foreign">bad</x:t></w:r></w:p>"#),
        false,
    );
    let package = eager_package(&foreign);
    let document = package.document().expect("foreign document");
    let mut output = Vec::new();
    assert!(matches!(
        document.write_text_to(&mut output, parity_options(true)),
        Err(TextOutputError::Document { .. })
    ));
    assert!(output.is_empty());
}

#[test]
fn oversized_namespace_event_is_rejected_before_resolution() {
    let huge_namespace = "x".repeat(1024 * 1024 + 1);
    let xml = format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="{huge_namespace}"><w:body/></w:document>"#
    );
    let bytes = package_bytes(xml.as_bytes(), false);
    let package = eager_package(&bytes);
    let document = package.document().expect("document");
    let mut output = Vec::new();

    let result = document.write_text_to(&mut output, TextOutputOptions::default());

    assert!(matches!(
        result,
        Err(TextOutputError::Document { progress, .. })
            if progress.bytes_written() == 0 && progress.objects_written() == 0
    ));
    assert!(output.is_empty());
}

#[test]
fn malformed_later_paragraph_preserves_prior_progress() {
    let bytes = package_bytes(
        &document_xml(
            r#"<w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><x:t xmlns:x="urn:foreign">bad</x:t></w:r></w:p>"#,
        ),
        false,
    );
    let package = eager_package(&bytes);
    let document = package.document().expect("malformed semantic document");
    let mut output = Vec::new();
    let error = document
        .write_text_to(&mut output, parity_options(false))
        .expect_err("foreign later paragraph must fail");
    assert_eq!(output, b"first");
    assert!(matches!(
        error,
        TextOutputError::Document { progress, .. }
            if progress.bytes_written() == 5 && progress.objects_written() == 1
    ));
}

#[test]
fn output_and_object_limits_report_exact_progress() {
    let bytes = package_bytes(
        &document_xml(
            r#"<w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p>"#,
        ),
        false,
    );
    let package = eager_package(&bytes);
    let document = package.document().expect("limit document");
    let mut output = Vec::new();
    let error = document
        .write_text_to(
            &mut output,
            TextOutputOptions::new("\n", "unused", 3, 2).with_empty_objects(false),
        )
        .expect_err("output limit");
    assert!(matches!(
        error,
        TextOutputError::Limit { progress, .. }
            if progress.bytes_written() == 3 && progress.objects_written() == 1
    ));
    assert_eq!(output, b"one");

    let mut output = Vec::new();
    let error = document
        .write_text_to(
            &mut output,
            TextOutputOptions::new("\n", "unused", u64::MAX, 1).with_empty_objects(false),
        )
        .expect_err("object limit");
    assert!(matches!(
        error,
        TextOutputError::Limit { progress, .. }
            if progress.bytes_written() == 3 && progress.objects_written() == 1
    ));
    assert_eq!(output, b"one");
}

#[derive(Default)]
struct ShortWriter {
    bytes: Vec<u8>,
    chunk: usize,
}

impl Write for ShortWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let written = bytes.len().min(self.chunk.max(1));
        self.bytes.extend_from_slice(&bytes[..written]);
        Ok(written)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct ZeroWriter;

impl Write for ZeroWriter {
    fn write(&mut self, _bytes: &[u8]) -> io::Result<usize> {
        Ok(0)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct OverreportingWriter;

impl Write for OverreportingWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        Ok(bytes.len().saturating_add(1))
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct InterruptedOnce {
    bytes: Vec<u8>,
    interrupted: bool,
}

impl Write for InterruptedOnce {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if !self.interrupted {
            self.interrupted = true;
            return Err(io::Error::from(io::ErrorKind::Interrupted));
        }
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct FailAfter {
    bytes: Vec<u8>,
    remaining: usize,
}

impl Write for FailAfter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.remaining == 0 {
            return Err(io::Error::other("injected sink failure"));
        }
        let written = bytes.len().min(self.remaining);
        self.bytes.extend_from_slice(&bytes[..written]);
        self.remaining -= written;
        Ok(written)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn sink_progress_handles_short_interrupted_zero_overreport_and_prefix_failure() {
    let bytes = package_bytes(
        &document_xml(r#"<w:p><w:r><w:t>é!</w:t></w:r></w:p>"#),
        false,
    );
    let package = eager_package(&bytes);
    let document = package.document().expect("sink document");

    let mut short = ShortWriter {
        chunk: 1,
        ..ShortWriter::default()
    };
    let report = document
        .write_text_to(&mut short, parity_options(true))
        .expect("short writes retry");
    assert_eq!(short.bytes, "é!".as_bytes());
    assert_eq!(report.bytes_written(), 3);

    let mut interrupted = InterruptedOnce {
        bytes: Vec::new(),
        interrupted: false,
    };
    document
        .write_text_to(&mut interrupted, parity_options(true))
        .expect("interrupted writes retry");
    assert_eq!(interrupted.bytes, "é!".as_bytes());

    let mut zero = ZeroWriter;
    assert!(matches!(
        document.write_text_to(&mut zero, parity_options(true)),
        Err(TextOutputError::Sink { progress, .. }) if progress.bytes_written() == 0
    ));

    let mut overreporting = OverreportingWriter;
    assert!(matches!(
        document.write_text_to(&mut overreporting, parity_options(true)),
        Err(TextOutputError::Sink { progress, .. }) if progress.bytes_written() == 0
    ));

    let mut failing = FailAfter {
        bytes: Vec::new(),
        remaining: 2,
    };
    assert!(matches!(
        document.write_text_to(&mut failing, parity_options(true)),
        Err(TextOutputError::Sink { progress, .. }) if progress.bytes_written() == 2
    ));
    assert_eq!(failing.bytes, "é".as_bytes());
}

#[test]
fn parser_depth_decoded_and_illegal_character_limits_are_bounded() {
    let mut deep = String::new();
    for _ in 0..129 {
        deep.push_str("<w:sdt>");
    }
    deep.push_str("<w:r><w:t>deep</w:t></w:r>");
    for _ in 0..129 {
        deep.push_str("</w:sdt>");
    }
    let package = eager_package(package_bytes(
        &document_xml(&format!("<w:p>{deep}</w:p>")),
        false,
    ));
    let document = package.document().expect("deep document");
    let mut output = Vec::new();
    assert!(matches!(
        document.write_text_to(&mut output, parity_options(true)),
        Err(TextOutputError::Document { .. })
    ));
    assert!(output.is_empty());

    let run = "x".repeat(1024 * 1024);
    let mut runs = String::new();
    for _ in 0..17 {
        runs.push_str("<w:r><w:t>");
        runs.push_str(&run);
        runs.push_str("</w:t></w:r>");
    }
    let package = eager_package(package_bytes(
        &document_xml(&format!("<w:p>{runs}</w:p>")),
        false,
    ));
    let document = package.document().expect("decoded-limit document");
    let mut output = Vec::new();
    assert!(matches!(
        document.write_text_to(&mut output, parity_options(true)),
        Err(TextOutputError::Document { .. })
    ));
    assert!(output.is_empty());

    let illegal = document_xml("<w:p><w:r><w:t>bad\u{1}</w:t></w:r></w:p>");
    let package = eager_package(package_bytes(&illegal, false));
    let document = package.document().expect("illegal document");
    let mut output = Vec::new();
    assert!(matches!(
        document.write_text_to(&mut output, parity_options(true)),
        Err(TextOutputError::Document { .. })
    ));
}

struct VersionedSource {
    bytes: Vec<u8>,
    revision: AtomicU64,
    reads: AtomicUsize,
    change_on_read: AtomicBool,
}

impl VersionedSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            revision: AtomicU64::new(0),
            reads: AtomicUsize::new(0),
            change_on_read: AtomicBool::new(false),
        }
    }

    fn changed(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }

    fn arm_change_on_read(&self) {
        self.change_on_read.store(true, Ordering::SeqCst);
    }
}

impl ReadAt for VersionedSource {
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
        if self.change_on_read.swap(false, Ordering::SeqCst) {
            self.changed();
        }
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            901,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

struct ChangeOnWrite {
    source: Arc<VersionedSource>,
    bytes: Vec<u8>,
    changed: bool,
}

struct ChangeAndFail {
    source: Arc<VersionedSource>,
    changed: bool,
}

impl Write for ChangeAndFail {
    fn write(&mut self, _bytes: &[u8]) -> io::Result<usize> {
        if !self.changed {
            self.changed = true;
            self.source.changed();
        }
        Err(io::Error::other("injected sink failure"))
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

impl Write for ChangeOnWrite {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if !self.changed {
            self.changed = true;
            self.source.changed();
        }
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn source_stale_before_during_and_after_output_takes_precedence() {
    let source = Arc::new(VersionedSource::new(package_bytes(
        &document_xml(r#"<w:p><w:r><w:t>stale</w:t></w:r></w:p>"#),
        false,
    )));
    let package = SourceBackedPackage::from_read_at(source.clone()).expect("source package");
    source.changed();
    let mut output = Vec::new();
    assert!(matches!(
        package.write_text_to(&mut output, parity_options(true)),
        Err(TextOutputError::Document { progress, .. }) if progress.bytes_written() == 0
    ));

    let source = Arc::new(VersionedSource::new(package_bytes(
        &document_xml(r#"<w:p><w:r><w:t>during</w:t></w:r></w:p>"#),
        false,
    )));
    let package = SourceBackedPackage::from_read_at(source.clone()).expect("source package");
    source.arm_change_on_read();
    let mut output = Vec::new();
    assert!(matches!(
        package.write_text_to(&mut output, parity_options(true)),
        Err(TextOutputError::Document { progress, .. }) if progress.bytes_written() == 0
    ));
    assert!(output.is_empty());

    let source = Arc::new(VersionedSource::new(package_bytes(
        &document_xml(r#"<w:p><w:r><w:t>after</w:t></w:r></w:p>"#),
        false,
    )));
    let package = SourceBackedPackage::from_read_at(source.clone()).expect("source package");
    let mut sink = ChangeOnWrite {
        source,
        bytes: Vec::new(),
        changed: false,
    };
    let error = package
        .write_text_to(&mut sink, parity_options(true))
        .expect_err("source change after output");
    assert!(matches!(
        error,
        TextOutputError::Document { progress, .. }
            if progress.bytes_written() == 5 && progress.objects_written() == 1
    ));
    assert_eq!(sink.bytes, b"after");
}

#[test]
fn source_change_overrides_simultaneous_sink_failure() {
    let source = Arc::new(VersionedSource::new(package_bytes(
        &document_xml(r#"<w:p><w:r><w:t>sink</w:t></w:r></w:p>"#),
        false,
    )));
    let package = SourceBackedPackage::from_read_at(source.clone()).expect("source package");
    let mut sink = ChangeAndFail {
        source,
        changed: false,
    };
    let error = package
        .write_text_to(&mut sink, parity_options(true))
        .expect_err("source change must take precedence");
    assert!(matches!(
        error,
        TextOutputError::Document { progress, .. }
            if progress.bytes_written() == 0 && progress.objects_written() == 0
    ));
}

fn execution_context() -> (CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "docx-sequential-text-test",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (source, cancellation) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("worker limit"),
        NonZeroUsize::new(1).expect("task limit"),
        NonZeroU64::new(u64::MAX).expect("memory limit"),
        0,
    )
    .expect("execution limits");
    (source, ExecutionContext::new(budget, cancellation, limits))
}

struct CancelOnWrite {
    source: CancellationSource,
    bytes: Vec<u8>,
    cancelled: bool,
}

impl Write for CancelOnWrite {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if !self.cancelled {
            self.cancelled = true;
            self.source.cancel();
        }
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn source_cancellation_before_and_after_output_is_document_progress() {
    let (cancellation, context) = execution_context();
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(package_bytes(
            &document_xml(r#"<w:p><w:r><w:t>cancel</w:t></w:r></w:p>"#),
            false,
        ))),
        ReadLimits::default(),
        context,
    )
    .expect("managed source package");
    cancellation.cancel();
    let mut output = Vec::new();
    assert!(matches!(
        package.write_text_to(&mut output, parity_options(true)),
        Err(TextOutputError::Document { progress, .. }) if progress.bytes_written() == 0
    ));

    let (cancellation, context) = execution_context();
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(package_bytes(
            &document_xml(r#"<w:p><w:r><w:t>cancel</w:t></w:r></w:p>"#),
            false,
        ))),
        ReadLimits::default(),
        context,
    )
    .expect("managed source package");
    let mut sink = CancelOnWrite {
        source: cancellation,
        bytes: Vec::new(),
        cancelled: false,
    };
    let error = package
        .write_text_to(&mut sink, parity_options(true))
        .expect_err("cancellation after output");
    assert!(matches!(
        error,
        TextOutputError::Document { progress, .. }
            if progress.bytes_written() == 6 && progress.objects_written() == 1
    ));
    assert_eq!(sink.bytes, b"cancel");
}

fn oversized_declaration(mut bytes: Vec<u8>) -> Vec<u8> {
    let target = b"word/document.xml";
    for offset in 0..bytes.len().saturating_sub(46) {
        if &bytes[offset..offset + 4] != b"PK\x01\x02" {
            continue;
        }
        let name_len = u16::from_le_bytes([bytes[offset + 28], bytes[offset + 29]]) as usize;
        let start = offset + 46;
        let end = start.saturating_add(name_len);
        if end <= bytes.len() && &bytes[start..end] == target {
            let declared = (64_u64 * 1024 * 1024 + 1) as u32;
            bytes[offset + 24..offset + 28].copy_from_slice(&declared.to_le_bytes());
            return bytes;
        }
    }
    panic!("main document central directory entry not found");
}

#[test]
fn declared_size_is_rejected_before_payload_and_media_stays_cold() {
    let source = Arc::new(VersionedSource::new(oversized_declaration(package_bytes(
        &document_xml(r#"<w:p><w:r><w:t>tiny</w:t></w:r></w:p>"#),
        true,
    ))));
    let package = SourceBackedPackage::from_read_at(source.clone()).expect("oversized package");
    let reads_before_sink = source.reads.load(Ordering::SeqCst);
    let mut output = Vec::new();
    let error = package
        .write_text_to(&mut output, parity_options(true))
        .expect_err("declared semantic limit");
    assert!(matches!(
        error,
        TextOutputError::Document { progress, .. }
            if progress.bytes_written() == 0 && progress.objects_written() == 0
    ));
    assert!(output.is_empty());
    assert_eq!(source.reads.load(Ordering::SeqCst), reads_before_sink);
}
