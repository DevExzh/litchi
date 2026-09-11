#![allow(
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

mod support;

use std::io::{self, Write};
use std::sync::{
    Arc,
    atomic::{AtomicU64, Ordering},
};

use litchi_core::{
    Error, ReadAt, SourceVersion, TextOutputError, TextOutputLimitKind, TextOutputOptions,
};
use litchi_odt::{Document, SourceBackedDocument};

const MIMETYPE: &str = "application/vnd.oasis.opendocument.text";
const OFFICE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const TABLE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const DRAW_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";

fn content(body: &str) -> Vec<u8> {
    format!(
        r#"<office:document-content xmlns:office="{OFFICE_NS}" xmlns:text="{TEXT_NS}" xmlns:table="{TABLE_NS}" xmlns:draw="{DRAW_NS}" office:version="1.3"><office:body><office:text>{body}</office:text></office:body></office:document-content>"#
    )
    .into_bytes()
}

fn package(body: &str) -> Vec<u8> {
    let content = content(body);
    raw_package(content.as_slice())
}

fn raw_package(content_xml: &[u8]) -> Vec<u8> {
    // Keep producer bytes, including XML line-ending spelling, in the source
    // archive so the owned and positional facades consume the same payload.
    support::package(MIMETYPE, &[("content.xml", content_xml)])
}

fn write_owned(document: &Document, options: TextOutputOptions<'_>) -> (Vec<u8>, u64, u64) {
    let mut output = Vec::new();
    let report = document.write_text_to(&mut output, options).unwrap();
    (output, report.bytes_written(), report.objects_written())
}

fn normalized_line_ending_fixture() -> (&'static str, &'static str) {
    (
        concat!(
            "<text:p>literal\rCR</text:p>",
            "<text:p>CRLF\r\nline</text:p>",
            "<text:p>repeat\r\rCR</text:p>",
            "<text:p>UTF8尾<![CDATA[cdata\r\n尾]]></text:p>",
        ),
        "literal\nCR|CRLF\nline|repeat\n\nCR|UTF8尾cdata\n尾",
    )
}

#[test]
fn mixed_document_writes_exact_text_in_semantic_start_order() {
    let bytes = package(concat!(
        "<text:h text:outline-level=\"1\">Heading</text:h>",
        "<text:p>before</text:p>",
        "<table:table table:name=\"Table1\"><table:table-row><table:table-cell>",
        "<text:p>cell</text:p></table:table-cell></table:table-row></table:table>",
        "<text:p>anchor<draw:frame><draw:text-box><text:p>frame</text:p>",
        "</draw:text-box></draw:frame>after</text:p>",
        "<text:p/>",
    ));
    let document = Document::from_bytes(bytes).unwrap();
    let expected = document.text().unwrap();
    let (output, bytes_written, objects_written) =
        write_owned(&document, TextOutputOptions::default());

    assert_eq!(output, expected.as_bytes());
    assert_eq!(bytes_written, expected.len() as u64);
    assert_eq!(
        objects_written, 6,
        "the heading, paragraphs, table-cell paragraph, nested frame paragraph, and empty block are objects"
    );

    let order = ["Heading", "before", "cell", "anchorafter", "frame"];
    let mut previous = 0;
    for marker in order {
        let position = expected.find(marker).unwrap();
        assert!(position >= previous, "{marker} was emitted out of order");
        previous = position;
    }
}

#[test]
fn varied_block_sizes_and_nested_blocks_match_owned_and_source_backed_output() {
    let oversized = "x".repeat(4097);
    let body = format!(
        concat!(
            "<text:p>small</text:p>",
            "<text:p/>",
            "<text:p>outer-before<draw:frame><draw:text-box>",
            "<text:p>nested</text:p>",
            "</draw:text-box></draw:frame>outer-after</text:p>",
            "<text:p>{oversized}</text:p>",
            "<text:p>tail</text:p>",
        ),
        oversized = oversized,
    );
    let bytes = package(&body);
    let eager = Document::from_bytes(bytes.clone()).unwrap();
    let options = TextOutputOptions::new("|", "unused", 16 * 1024, 16);
    let expected = format!("small||outer-beforeouter-after|nested|{oversized}|tail");

    let (owned_output, owned_bytes, owned_objects) = write_owned(&eager, options);
    assert_eq!(owned_output, expected.as_bytes());
    assert_eq!(owned_bytes, expected.len() as u64);
    assert_eq!(owned_objects, 6);

    let (source, _) = MemorySource::new(bytes);
    let source_document = SourceBackedDocument::from_read_at(source).unwrap();
    let mut source_output = Vec::new();
    let source_report = source_document
        .write_text_to(&mut source_output, options)
        .unwrap();
    assert_eq!(source_output, expected.as_bytes());
    assert_eq!(source_output, owned_output);
    assert_eq!(source_report.bytes_written(), owned_bytes);
    assert_eq!(source_report.objects_written(), owned_objects);
}

#[test]
fn whitespace_entities_cdata_expansions_and_utf8_match_text() {
    let document = Document::from_bytes(package(
        r#"<text:p>  lead &amp; <![CDATA[<cdata>]]><text:s text:c="2"/><text:tab/><text:line-break/>尾</text:p>"#,
    ))
    .unwrap();
    let expected = document.text().unwrap();
    let (output, bytes_written, objects_written) =
        write_owned(&document, TextOutputOptions::default());

    assert_eq!(output, expected.as_bytes());
    assert_eq!(bytes_written, expected.len() as u64);
    assert_eq!(objects_written, 1);
    assert!(expected.contains("lead & <cdata>"));
    assert!(expected.contains('\t'), "text:tab was not preserved");
    assert!(expected.contains('\n'), "text:line-break was not preserved");
    assert!(expected.contains('尾'), "UTF-8 text was not preserved");
}

#[test]
fn suppressed_stored_content_matches_document_text() {
    let document = Document::from_bytes(package(concat!(
        "<text:tracked-changes><text:changed-region text:id=\"change-1\">",
        "<text:deletion><text:p>stored-change</text:p></text:deletion>",
        "</text:changed-region></text:tracked-changes>",
        "<text:p>visible <text:note text:note-class=\"footnote\">",
        "<text:note-citation>1</text:note-citation><text:note-body>",
        "<text:p>stored-note</text:p></text:note-body></text:note>",
        "<text:ruby><text:ruby-base>base</text:ruby-base>",
        "<text:ruby-text>pronunciation</text:ruby-text></text:ruby></text:p>",
    )))
    .unwrap();
    let expected = document.text().unwrap();
    let (output, bytes_written, objects_written) =
        write_owned(&document, TextOutputOptions::default());

    assert_eq!(output, expected.as_bytes());
    assert_eq!(bytes_written, expected.len() as u64);
    assert_eq!(objects_written, 1);
    assert!(expected.contains("visible"));
    assert!(expected.contains("base"));
    assert!(!expected.contains("stored-change"));
    assert!(!expected.contains("stored-note"));
    assert!(!expected.contains("pronunciation"));
}

#[test]
fn empty_object_policy_and_custom_separator_are_applied() {
    let document = Document::from_bytes(package(
        "<text:p>one</text:p><text:p/><text:p>three</text:p>",
    ))
    .unwrap();

    let (included, included_bytes, included_objects) =
        write_owned(&document, TextOutputOptions::new("::", "unused", 64, 8));
    assert_eq!(included, b"one::::three");
    assert_eq!(included_bytes, 12);
    assert_eq!(included_objects, 3);

    let (omitted, omitted_bytes, omitted_objects) = write_owned(
        &document,
        TextOutputOptions::new("::", "unused", 64, 8).with_empty_objects(false),
    );
    assert_eq!(omitted, b"one::three");
    assert_eq!(omitted_bytes, 10);
    assert_eq!(omitted_objects, 2);
}

#[test]
fn object_limit_reports_the_rejected_object_without_partial_output() {
    let document =
        Document::from_bytes(package("<text:p>one</text:p><text:p>two</text:p>")).unwrap();
    let mut output = Vec::new();
    let error = document
        .write_text_to(&mut output, TextOutputOptions::new("|", "unused", 64, 1))
        .unwrap_err();

    assert_eq!(output, b"one");
    assert_eq!(error.progress().bytes_written(), 3);
    assert_eq!(error.progress().objects_written(), 1);
    let limit = error.limit().unwrap();
    assert_eq!(limit.kind(), TextOutputLimitKind::Objects);
    assert_eq!(limit.observed(), 2);
    assert_eq!(limit.limit(), 1);
}

#[test]
fn output_byte_limit_reports_separator_and_next_object_extent() {
    let document =
        Document::from_bytes(package("<text:p>one</text:p><text:p>two</text:p>")).unwrap();
    let mut output = Vec::new();
    let error = document
        .write_text_to(&mut output, TextOutputOptions::new("|", "unused", 6, 8))
        .unwrap_err();

    assert_eq!(output, b"one");
    assert_eq!(error.progress().bytes_written(), 3);
    assert_eq!(error.progress().objects_written(), 1);
    let limit = error.limit().unwrap();
    assert_eq!(limit.kind(), TextOutputLimitKind::OutputBytes);
    assert_eq!(limit.observed(), 7);
    assert_eq!(limit.limit(), 6);
}

#[derive(Default)]
struct PrefixThenFail {
    accepted: Vec<u8>,
    prefix_limit: usize,
}

impl Write for PrefixThenFail {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.accepted.len() >= self.prefix_limit {
            return Err(io::Error::other("injected sink failure"));
        }
        let remaining = self.prefix_limit - self.accepted.len();
        let amount = remaining.min(bytes.len());
        self.accepted.extend_from_slice(&bytes[..amount]);
        Ok(amount)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn sink_failure_reports_accepted_multibyte_prefix_exactly() {
    let document = Document::from_bytes(package("<text:p>é文</text:p>")).unwrap();
    let mut sink = PrefixThenFail {
        accepted: Vec::new(),
        prefix_limit: 3,
    };
    let error = document
        .write_text_to(&mut sink, TextOutputOptions::new("|", "unused", 64, 8))
        .unwrap_err();

    assert_eq!(sink.accepted, "é文".as_bytes()[..3]);
    assert_eq!(error.progress().bytes_written(), 3);
    assert_eq!(error.progress().objects_written(), 0);
    assert!(matches!(&error, TextOutputError::Sink { .. }));
}

#[test]
fn sink_failure_after_emitted_block_preserves_progress_and_text() {
    let document =
        Document::from_bytes(package("<text:p>one</text:p><text:p>two</text:p>")).unwrap();
    let mut sink = PrefixThenFail {
        accepted: Vec::new(),
        prefix_limit: 5,
    };
    let error = document
        .write_text_to(&mut sink, TextOutputOptions::new("|", "unused", 64, 8))
        .unwrap_err();

    assert_eq!(sink.accepted, b"one|t");
    assert_eq!(error.progress().bytes_written(), 5);
    assert_eq!(error.progress().objects_written(), 1);
    assert!(matches!(&error, TextOutputError::Sink { .. }));
}

struct RevisionThenFail {
    accepted: Vec<u8>,
    first_object_bytes: usize,
    revision: Arc<AtomicU64>,
    revised: bool,
}

impl RevisionThenFail {
    fn new(revision: Arc<AtomicU64>, first_object_bytes: usize) -> Self {
        Self {
            accepted: Vec::new(),
            first_object_bytes,
            revision,
            revised: false,
        }
    }
}

impl Write for RevisionThenFail {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.accepted.len() < self.first_object_bytes {
            let remaining = self.first_object_bytes - self.accepted.len();
            let amount = remaining.min(bytes.len());
            self.accepted.extend_from_slice(&bytes[..amount]);
            return Ok(amount);
        }
        if !self.revised {
            self.revised = true;
            self.revision.fetch_add(1, Ordering::Relaxed);
        }
        Err(io::Error::other("injected separator sink failure"))
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn stale_source_precedes_separator_sink_failure_with_truthful_progress() {
    let bytes = package("<text:p>one</text:p><text:p>two</text:p>");
    let (source, revision) = MemorySource::new(bytes);
    let document = SourceBackedDocument::from_read_at(source).unwrap();
    let mut sink = RevisionThenFail::new(revision, 3);
    let error = document
        .write_text_to(&mut sink, TextOutputOptions::new("|", "unused", 64, 8))
        .unwrap_err();

    assert_eq!(sink.accepted, b"one");
    assert!(matches!(
        &error,
        TextOutputError::Document {
            source: Error::SourceChanged { .. },
            ..
        }
    ));
    assert!(!matches!(&error, TextOutputError::Sink { .. }));
    assert_eq!(error.progress().bytes_written(), 3);
    assert_eq!(error.progress().objects_written(), 1);
}

#[test]
fn malformed_tail_returns_document_error_after_valid_emitted_block() {
    let document = Document::from_bytes(package(
        r#"<text:p>valid</text:p><text:p>broken<text:s text:c="1000001"/></text:p>"#,
    ))
    .unwrap();
    let mut output = Vec::new();
    let error = document
        .write_text_to(&mut output, TextOutputOptions::default())
        .unwrap_err();

    assert_eq!(output, b"valid");
    assert!(matches!(&error, TextOutputError::Document { .. }));
    assert_eq!(error.progress().bytes_written(), 5);
    assert_eq!(error.progress().objects_written(), 1);
}

struct MemorySource {
    bytes: Arc<Vec<u8>>,
    revision: Arc<AtomicU64>,
}

impl MemorySource {
    fn new(bytes: Vec<u8>) -> (Arc<Self>, Arc<AtomicU64>) {
        let revision = Arc::new(AtomicU64::new(0));
        let source = Arc::new(Self {
            bytes: Arc::new(bytes),
            revision: Arc::clone(&revision),
        });
        (source, revision)
    }
}

impl ReadAt for MemorySource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_| io::Error::other("test source length does not fit u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let start = usize::try_from(offset)
            .map_err(|_| io::Error::other("test source offset does not fit usize"))?;
        let input = self
            .bytes
            .get(start..)
            .ok_or_else(|| io::Error::other("test source offset is out of bounds"))?;
        let amount = input.len().min(output.len());
        output[..amount].copy_from_slice(&input[..amount]);
        Ok(amount)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x4f44_5401,
            self.revision.load(Ordering::Relaxed),
        ))
    }
}

#[test]
fn source_backed_writer_matches_owned_writer_and_rejects_stale_source_first() {
    let bytes = package(concat!(
        "<text:h text:outline-level=\"1\">Heading</text:h>",
        "<text:p>body</text:p><text:p>尾</text:p>",
    ));
    let eager = Document::from_bytes(bytes.clone()).unwrap();
    let (source, revision) = MemorySource::new(bytes);
    let source_document = SourceBackedDocument::from_read_at(source).unwrap();
    let options = TextOutputOptions::new("|", "unused", 64, 8);

    let (owned_output, owned_bytes, owned_objects) = write_owned(&eager, options);
    let mut source_output = Vec::new();
    let source_report = source_document
        .write_text_to(&mut source_output, options)
        .unwrap();
    assert_eq!(source_output, owned_output);
    assert_eq!(source_report.bytes_written(), owned_bytes);
    assert_eq!(source_report.objects_written(), owned_objects);

    revision.fetch_add(1, Ordering::Relaxed);
    let mut stale_output = Vec::new();
    let error = source_document
        .write_text_to(&mut stale_output, TextOutputOptions::default())
        .unwrap_err();
    assert!(stale_output.is_empty());
    assert!(matches!(
        &error,
        TextOutputError::Document {
            source: Error::SourceChanged { .. },
            ..
        }
    ));
    assert_eq!(error.progress().bytes_written(), 0);
    assert_eq!(error.progress().objects_written(), 0);
}

#[test]
fn owned_and_source_backed_exports_match_normalized_line_endings_across_blocks() {
    let (body, expected) = normalized_line_ending_fixture();
    let content_xml = content(body);
    let bytes = raw_package(content_xml.as_slice());
    let eager = Document::from_bytes(bytes.clone()).unwrap();
    let (source, _) = MemorySource::new(bytes);
    let source_document = SourceBackedDocument::from_read_at(source).unwrap();

    assert_eq!(
        source_document.content_xml().unwrap().as_bytes(),
        content_xml.as_slice()
    );

    let options = TextOutputOptions::new("|", "unused", 64, 8);
    let (owned_output, owned_bytes, owned_objects) = write_owned(&eager, options);
    assert_eq!(owned_output, expected.as_bytes());
    assert_eq!(owned_bytes, 48);
    assert_eq!(owned_objects, 4);

    let mut source_output = Vec::new();
    let source_report = source_document
        .write_text_to(&mut source_output, options)
        .unwrap();
    assert_eq!(source_output, expected.as_bytes());
    assert_eq!(source_report.bytes_written(), 48);
    assert_eq!(source_report.objects_written(), 4);
}

#[test]
fn normalized_output_limit_accepts_exact_length_and_refuses_one_byte_lower() {
    let (body, expected) = normalized_line_ending_fixture();
    let document = Document::from_bytes(raw_package(content(body).as_slice())).unwrap();

    let exact = TextOutputOptions::new("|", "unused", expected.len() as u64, 8);
    let (output, bytes_written, objects_written) = write_owned(&document, exact);
    assert_eq!(output, expected.as_bytes());
    assert_eq!(bytes_written, 48);
    assert_eq!(objects_written, 4);

    let mut limited_output = Vec::new();
    let error = document
        .write_text_to(
            &mut limited_output,
            TextOutputOptions::new("|", "unused", 47, 8),
        )
        .unwrap_err();
    assert_eq!(limited_output, b"literal\nCR|CRLF\nline|repeat\n\nCR");
    assert_eq!(error.progress().bytes_written(), 31);
    assert_eq!(error.progress().objects_written(), 3);
    let limit = error.limit().unwrap();
    assert_eq!(limit.kind(), TextOutputLimitKind::OutputBytes);
    assert_eq!(limit.observed(), 48);
    assert_eq!(limit.limit(), 47);
}
