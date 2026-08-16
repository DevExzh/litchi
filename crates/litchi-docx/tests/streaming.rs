//! Public contract tests for bounded, fresh DOCX paragraph/run creation.

use std::io::{self, Cursor, Write};
use std::num::{NonZeroU64, NonZeroUsize};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, Resource,
};
use litchi_docx::{
    Package, StreamingDocumentError, StreamingDocumentLimits, StreamingDocumentWriter,
};
use litchi_opc::PackURI;
use sha2::{Digest, Sha256};

fn context_for(budget: Budget) -> (CancellationSource, ExecutionContext) {
    let (source, token) = CancellationSource::pair();
    let execution = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("one worker"),
        NonZeroUsize::new(1).expect("one task"),
        NonZeroU64::new(1024 * 1024).expect("nonzero in-flight bytes"),
        1,
    )
    .expect("valid execution limits");
    (source, ExecutionContext::new(budget, token, execution))
}

fn context() -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "docx-stream-integration",
        CoreLimits::new(
            1024 * 1024,
            1024 * 1024,
            1024 * 1024,
            10_000,
            64,
            10_000_000,
        ),
    );
    let (source, execution) = context_for(budget.clone());
    (budget, source, execution)
}

fn limits() -> StreamingDocumentLimits {
    StreamingDocumentLimits::default()
}

fn emit_sample<W: Write>(writer: &mut StreamingDocumentWriter<W>) {
    writer.start_paragraph().expect("paragraph");
    writer.start_run().expect("first run");
    writer.write_text(" leading & <é> ").expect("first text");
    writer.finish_run().expect("first run finish");
    writer.start_run().expect("second run");
    writer.write_text("tail").expect("second text");
    writer.finish_run().expect("second run finish");
    writer.finish_paragraph().expect("paragraph finish");
    writer.start_paragraph().expect("empty paragraph");
    writer.finish_paragraph().expect("empty paragraph finish");
}

fn render_sample() -> Vec<u8> {
    let (_budget, _source, execution) = context();
    let mut writer =
        StreamingDocumentWriter::new(Vec::new(), execution, limits()).expect("streaming writer");
    emit_sample(&mut writer);
    writer.finish().expect("package finish")
}

#[test]
fn exported_writer_is_deterministic_three_member_and_reopens_with_run_parity() {
    let first = render_sample();
    let second = render_sample();
    assert_eq!(first, second);
    let first_hash = Sha256::digest(&first);
    assert_eq!(first_hash, Sha256::digest(&second));

    let physical = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(first.clone())
        .expect("physical package");
    assert_eq!(
        physical.member_names().expect("member names"),
        vec![
            "[Content_Types].xml".to_owned(),
            "_rels/.rels".to_owned(),
            "word/document.xml".to_owned(),
        ]
    );
    let document_xml = physical
        .blob_for(&PackURI::new("/word/document.xml").expect("document URI"))
        .expect("document XML");
    assert!(document_xml.starts_with(b"<?xml version=\"1.0\""));
    assert!(document_xml.windows(2).all(|pair| pair != b"\r\n"));
    assert!(!document_xml.windows(2).any(|pair| pair == b"< "));
    assert!(document_xml.windows(5).any(|window| window == b"&amp;"));
    assert!(document_xml.windows(4).any(|window| window == b"&lt;"));
    assert!(document_xml.windows(4).any(|window| window == b"&gt;"));

    let package = Package::from_reader(Cursor::new(first)).expect("reopen DOCX");
    let document = package.document().expect("main document");
    assert_eq!(
        document.text().expect("document text"),
        " leading & <é> tail"
    );
    assert_eq!(document.paragraph_count().expect("paragraph count"), 2);
    let paragraphs = document.paragraphs().expect("paragraphs");
    assert_eq!(paragraphs.len(), 2);
    assert_eq!(
        paragraphs[0].text().expect("first paragraph text"),
        " leading & <é> tail"
    );
    assert_eq!(paragraphs[1].text().expect("empty paragraph text"), "");
    let runs = paragraphs[0].runs().expect("runs");
    assert_eq!(runs.len(), 2);
    assert_eq!(runs[0].text().expect("first run text"), " leading & <é> ");
    assert_eq!(runs[1].text().expect("second run text"), "tail");
}

#[derive(Debug, Default)]
struct ShortSink {
    bytes: Vec<u8>,
    max_chunk: usize,
}

impl Write for ShortSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let count = bytes.len().min(self.max_chunk.max(1));
        self.bytes.extend_from_slice(&bytes[..count]);
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug, Default)]
struct InterruptedSink {
    bytes: Vec<u8>,
    interrupted: bool,
}

impl Write for InterruptedSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if !self.interrupted {
            self.interrupted = true;
            return Err(io::Error::new(io::ErrorKind::Interrupted, "retry"));
        }
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn exported_writer_accepts_non_seek_short_and_interrupted_sinks() {
    let (_budget, _source, execution) = context();
    let mut short = StreamingDocumentWriter::new(
        ShortSink {
            max_chunk: 3,
            ..ShortSink::default()
        },
        execution,
        limits(),
    )
    .expect("short sink writer");
    emit_sample(&mut short);
    let short_sink = short.finish().expect("short sink finish");
    assert!(!short_sink.bytes.is_empty());
    let short_package =
        Package::from_reader(Cursor::new(short_sink.bytes)).expect("reopen short-sink package");
    assert_eq!(
        short_package
            .document()
            .expect("short document")
            .text()
            .expect("short text"),
        " leading & <é> tail"
    );

    let (_budget, _source, execution) = context();
    let mut interrupted =
        StreamingDocumentWriter::new(InterruptedSink::default(), execution, limits())
            .expect("interrupted sink writer");
    emit_sample(&mut interrupted);
    let interrupted_sink = interrupted.finish().expect("interrupted sink finish");
    assert!(!interrupted_sink.bytes.is_empty());
}

#[test]
fn exported_writer_rejects_invalid_text_and_reports_progress_and_state() {
    let (_budget, _source, execution) = context();
    let mut writer =
        StreamingDocumentWriter::new(Vec::new(), execution, limits()).expect("streaming writer");
    assert!(matches!(
        writer.finish_run(),
        Err(StreamingDocumentError::InvalidInput { .. })
    ));
    assert!(!writer.is_poisoned());
    writer.start_paragraph().expect("paragraph");
    writer.start_run().expect("run");
    let prefix = writer.output_bytes();
    let error = writer
        .write_text("before\tafter")
        .expect_err("tab must be rejected");
    assert!(matches!(
        error,
        StreamingDocumentError::InvalidInput { written, .. } if written == prefix
    ));
    assert_eq!(writer.output_bytes(), prefix);
    assert_eq!(writer.input_bytes(), 0);
    assert!(writer.is_poisoned());
    let repeated = writer.finish_run().expect_err("poisoned writer");
    assert_eq!(repeated.written(), prefix);
}

#[test]
fn cancellation_and_hierarchical_scratch_release_are_observable() {
    let parent = Budget::root(
        "docx-stream-parent",
        CoreLimits::new(64, 1024 * 1024, 1024 * 1024, 10_000, 64, 10_000),
    );
    let child = parent.child(
        "docx-stream-child",
        CoreLimits::new(64, 1024 * 1024, 1024 * 1024, 10_000, 64, 10_000),
    );
    let (source, execution) = context_for(child.clone());
    let mut writer = StreamingDocumentWriter::new(Vec::new(), execution, limits())
        .expect("exact scratch reservation");
    assert_eq!(child.used(Resource::Memory), 64);
    assert_eq!(parent.used(Resource::Memory), 64);
    let written = writer.output_bytes();
    source.cancel();
    let error = writer.start_paragraph().expect_err("cancelled writer");
    assert!(matches!(
        error,
        StreamingDocumentError::Cancelled { written: value } if value == written
    ));
    assert!(writer.is_poisoned());
    drop(writer);
    assert_eq!(child.used(Resource::Memory), 0);
    assert_eq!(parent.used(Resource::Memory), 0);
}
