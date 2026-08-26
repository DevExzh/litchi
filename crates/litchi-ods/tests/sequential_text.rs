#![allow(
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use std::io::{self, Write};
use std::sync::{
    Arc,
    atomic::{AtomicU64, AtomicUsize, Ordering},
};
use std::thread;

use litchi_core::{
    Error, OwnedSource, ReadAt, SourceVersion, TextOutputError, TextOutputLimitKind,
    TextOutputOptions,
};
use litchi_ods::{Builder, SourceBackedSpreadsheet, Spreadsheet};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

fn content(body: &str) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}" office:version="1.3"><office:body><office:spreadsheet>{body}</office:spreadsheet></office:body></office:document-content>"#
    )
}

fn spreadsheet(body: &str) -> Spreadsheet {
    Builder::new()
        .content_xml(content(body))
        .build()
        .and_then(Spreadsheet::from_bytes)
        .unwrap()
}

fn row(cells: &str) -> String {
    format!(r#"<table:table-row>{cells}</table:table-row>"#)
}

fn cell(text: &str) -> String {
    format!(
        r#"<table:table-cell office:value-type="string"><text:p>{text}</text:p></table:table-cell>"#
    )
}

fn write_owned(document: &Spreadsheet, options: TextOutputOptions<'_>) -> (Vec<u8>, u64, u64) {
    let mut output = Vec::new();
    let report = document.write_text_to(&mut output, options).unwrap();
    (output, report.bytes_written(), report.objects_written())
}

#[test]
fn owned_sink_matches_text_for_rows_repetition_tabs_and_empty_boundaries() {
    let first_cells = format!("{}{}", cell("  left &amp; right  "), cell("尾"));
    let first_row = row(&first_cells);
    let repeated_cell = r#"<table:table-cell office:value-type="string" table:number-columns-repeated="2"><text:p>same</text:p></table:table-cell>"#;
    let body = format!(
        r#"<table:table table:name="First">{}<table:table-row table:number-rows-repeated="2">{}</table:table-row></table:table><table:table table:name="Empty"/>"#,
        first_row, repeated_cell,
    );
    let document = spreadsheet(&body);
    let expected = document.text().unwrap();
    let (output, bytes, objects) = write_owned(&document, TextOutputOptions::default());

    assert_eq!(output, expected.as_bytes());
    assert_eq!(bytes, expected.len() as u64);
    assert_eq!(objects, 4);
    assert_eq!(expected, "  left  right  \t尾\nsame\tsame\nsame\tsame\n");
}

#[test]
fn direct_repeated_cells_rows_and_combined_runs_preserve_logical_order() {
    let body = format!(
        r#"<table:table table:name="Sheet"><table:table-row>{}</table:table-row><table:table-row table:number-rows-repeated="2">{}</table:table-row><table:table-row table:number-rows-repeated="2">{}</table:table-row></table:table>"#,
        r#"<table:table-cell office:value-type="string" table:number-columns-repeated="3"><text:p>a</text:p></table:table-cell>"#,
        cell("b"),
        r#"<table:table-cell office:value-type="string" table:number-columns-repeated="2"><text:p>c</text:p></table:table-cell>"#,
    );
    let document = spreadsheet(&body);
    let expected = document.text().unwrap();
    let (output, bytes, objects) = write_owned(&document, TextOutputOptions::default());

    assert_eq!(output, expected.as_bytes());
    assert_eq!(expected, "a\ta\ta\nb\nb\nc\tc\nc\tc");
    assert_eq!(bytes, expected.len() as u64);
    assert_eq!(objects, 5);
}

#[test]
fn leading_middle_and_trailing_empty_sheets_keep_boundaries_and_can_be_omitted() {
    let body = format!(
        r#"<table:table table:name="Leading"/><table:table table:name="First">{}</table:table><table:table table:name="Middle"/><table:table table:name="Second">{}</table:table><table:table table:name="Trailing"/>"#,
        row(&cell("one")),
        row(&cell("two")),
    );
    let document = spreadsheet(&body);

    let (included, bytes, objects) =
        write_owned(&document, TextOutputOptions::new("::", "unused", 64, 8));
    assert_eq!(included, b"::one::::two::");
    assert_eq!(bytes, 14);
    assert_eq!(objects, 5);

    let (omitted, bytes, objects) = write_owned(
        &document,
        TextOutputOptions::new("::", "unused", 64, 8).with_empty_objects(false),
    );
    assert_eq!(omitted, b"one::two");
    assert_eq!(bytes, 8);
    assert_eq!(objects, 2);
    assert_eq!(document.text().unwrap(), "\none\n\ntwo\n");
}

#[test]
fn controls_and_utf8_cell_text_match_the_existing_projection() {
    let document = spreadsheet(&format!(
        r#"<table:table table:name="Sheet">{}</table:table>"#,
        row(&cell(
            "lead<text:s text:c=\"2\"/><text:tab/><text:line-break/>尾"
        )),
    ));
    let expected = document.text().unwrap();
    let (output, bytes, objects) = write_owned(&document, TextOutputOptions::default());

    assert_eq!(output, expected.as_bytes());
    assert_eq!(expected, "lead尾");
    assert_eq!(bytes, expected.len() as u64);
    assert_eq!(objects, 1);
}

#[test]
fn empty_rows_and_custom_separator_follow_empty_object_policy() {
    let body = format!(
        r#"<table:table table:name="Sheet">{}<table:table-row/><table:table-row>{}</table:table-row></table:table>"#,
        row(&cell("one")),
        cell("three"),
    );
    let document = spreadsheet(&body);

    let (included, bytes, objects) =
        write_owned(&document, TextOutputOptions::new("::", "unused", 64, 8));
    assert_eq!(included, b"one::::three");
    assert_eq!(bytes, 12);
    assert_eq!(objects, 3);

    let (omitted, bytes, objects) = write_owned(
        &document,
        TextOutputOptions::new("::", "unused", 64, 8).with_empty_objects(false),
    );
    assert_eq!(omitted, b"one::three");
    assert_eq!(bytes, 10);
    assert_eq!(objects, 2);
}

#[test]
fn object_and_output_limits_report_exact_prior_progress() {
    let body = format!(
        r#"<table:table table:name="Sheet">{}{}</table:table>"#,
        row(&cell("one")),
        row(&cell("two")),
    );
    let document = spreadsheet(&body);

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
    limit: usize,
}

impl Write for PrefixThenFail {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.accepted.len() >= self.limit {
            return Err(io::Error::other("injected sink failure"));
        }
        let amount = (self.limit - self.accepted.len()).min(bytes.len());
        self.accepted.extend_from_slice(&bytes[..amount]);
        Ok(amount)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn sink_failure_preserves_accepted_multibyte_prefix() {
    let document = spreadsheet(&format!(
        r#"<table:table table:name="Sheet">{}</table:table>"#,
        row(&cell("é文")),
    ));
    let mut sink = PrefixThenFail {
        limit: 3,
        ..PrefixThenFail::default()
    };
    let error = document
        .write_text_to(&mut sink, TextOutputOptions::default())
        .unwrap_err();

    assert_eq!(sink.accepted, "é文".as_bytes()[..3]);
    assert!(matches!(&error, TextOutputError::Sink { .. }));
    assert_eq!(error.progress().bytes_written(), 3);
    assert_eq!(error.progress().objects_written(), 0);
}

#[derive(Default)]
struct ChunkedWriter {
    accepted: Vec<u8>,
    chunk: usize,
}

impl Write for ChunkedWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let amount = self.chunk.min(bytes.len());
        self.accepted.extend_from_slice(&bytes[..amount]);
        Ok(amount)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct InterruptedOnce {
    accepted: Vec<u8>,
    interrupted: bool,
}

impl Write for InterruptedOnce {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if !self.interrupted {
            self.interrupted = true;
            return Err(io::Error::from(io::ErrorKind::Interrupted));
        }
        self.accepted.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct WriteZero;

impl Write for WriteZero {
    fn write(&mut self, _bytes: &[u8]) -> io::Result<usize> {
        Ok(0)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn eventually_successful_short_writes_are_retried_until_complete() {
    let document = spreadsheet(&format!(
        r#"<table:table table:name="Sheet">{}</table:table>"#,
        row(&cell("short writes")),
    ));
    let mut sink = ChunkedWriter {
        chunk: 2,
        ..ChunkedWriter::default()
    };
    let report = document
        .write_text_to(&mut sink, TextOutputOptions::default())
        .unwrap();

    assert_eq!(sink.accepted, b"short writes");
    assert_eq!(report.bytes_written(), 12);
    assert_eq!(report.objects_written(), 1);
}

#[test]
fn one_interrupted_write_is_retried_without_losing_output() {
    let document = spreadsheet(&format!(
        r#"<table:table table:name="Sheet">{}</table:table>"#,
        row(&cell("retry")),
    ));
    let mut sink = InterruptedOnce {
        accepted: Vec::new(),
        interrupted: false,
    };
    let report = document
        .write_text_to(&mut sink, TextOutputOptions::default())
        .unwrap();

    assert_eq!(sink.accepted, b"retry");
    assert_eq!(report.bytes_written(), 5);
    assert_eq!(report.objects_written(), 1);
}

#[test]
fn write_zero_is_reported_as_a_sink_failure_with_no_progress() {
    let document = spreadsheet(&format!(
        r#"<table:table table:name="Sheet">{}</table:table>"#,
        row(&cell("zero")),
    ));
    let error = document
        .write_text_to(&mut WriteZero, TextOutputOptions::default())
        .unwrap_err();

    assert!(matches!(
        &error,
        TextOutputError::Sink { source, .. }
            if source.kind() == io::ErrorKind::WriteZero
    ));
    assert_eq!(error.progress().bytes_written(), 0);
    assert_eq!(error.progress().objects_written(), 0);
}

struct MutableSource {
    bytes: Arc<Vec<u8>>,
    revision: Arc<AtomicU64>,
}

impl MutableSource {
    fn new(bytes: Vec<u8>) -> (Arc<Self>, Arc<AtomicU64>) {
        let revision = Arc::new(AtomicU64::new(0));
        let source = Arc::new(Self {
            bytes: Arc::new(bytes),
            revision: Arc::clone(&revision),
        });
        (source, revision)
    }
}

impl ReadAt for MutableSource {
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
            0x4f44_5301,
            self.revision.load(Ordering::Relaxed),
        ))
    }
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
            let amount = (self.first_object_bytes - self.accepted.len()).min(bytes.len());
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

struct RevisionAfterComplete {
    accepted: Vec<u8>,
    expected_bytes: usize,
    revision: Arc<AtomicU64>,
    revised: bool,
}

impl Write for RevisionAfterComplete {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.accepted.extend_from_slice(bytes);
        if !self.revised && self.accepted.len() == self.expected_bytes {
            self.revised = true;
            self.revision.fetch_add(1, Ordering::Relaxed);
        }
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct CountingSource {
    bytes: Arc<Vec<u8>>,
    reads: Arc<AtomicUsize>,
}

impl ReadAt for CountingSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_| io::Error::other("test source length does not fit u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.reads.fetch_add(1, Ordering::Relaxed);
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
        Ok(SourceVersion::new(0x4f44_5302, 0))
    }
}

#[test]
fn source_backed_sink_matches_owned_and_rejects_stale_source_before_work() {
    let body = format!(
        r#"<table:table table:name="Sheet">{}{}</table:table>"#,
        row(&cell("one")),
        row(&cell("two")),
    );
    let bytes = Builder::new().content_xml(content(&body)).build().unwrap();
    let eager = Spreadsheet::from_bytes(bytes.clone()).unwrap();
    let (source, revision) = MutableSource::new(bytes);
    let source_document = SourceBackedSpreadsheet::from_read_at(source).unwrap();
    let options = TextOutputOptions::new("|", "unused", 64, 8);
    let (expected, bytes_written, objects_written) = write_owned(&eager, options);

    let mut output = Vec::new();
    let report = source_document.write_text_to(&mut output, options).unwrap();
    assert_eq!(output, expected);
    assert_eq!(report.bytes_written(), bytes_written);
    assert_eq!(report.objects_written(), objects_written);

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
fn stale_source_precedes_separator_sink_failure_with_truthful_progress() {
    let body = format!(
        r#"<table:table table:name="Sheet">{}{}</table:table>"#,
        row(&cell("one")),
        row(&cell("two")),
    );
    let bytes = Builder::new().content_xml(content(&body)).build().unwrap();
    let (source, revision) = MutableSource::new(bytes);
    let document = SourceBackedSpreadsheet::from_read_at(source).unwrap();
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
    assert_eq!(error.progress().bytes_written(), 3);
    assert_eq!(error.progress().objects_written(), 1);
}

#[test]
fn stale_source_after_full_output_still_takes_precedence() {
    let body = format!(
        r#"<table:table table:name="Sheet">{}{}</table:table>"#,
        row(&cell("one")),
        row(&cell("two")),
    );
    let bytes = Builder::new().content_xml(content(&body)).build().unwrap();
    let (source, revision) = MutableSource::new(bytes);
    let document = SourceBackedSpreadsheet::from_read_at(source).unwrap();
    let mut sink = RevisionAfterComplete {
        accepted: Vec::new(),
        expected_bytes: 7,
        revision,
        revised: false,
    };
    let error = document
        .write_text_to(&mut sink, TextOutputOptions::new("|", "unused", 64, 8))
        .unwrap_err();

    assert_eq!(sink.accepted, b"one|two");
    assert!(matches!(
        &error,
        TextOutputError::Document {
            source: Error::SourceChanged { .. },
            ..
        }
    ));
    assert_eq!(error.progress().bytes_written(), 7);
    assert_eq!(error.progress().objects_written(), 2);
}

#[test]
fn stable_source_writer_reuses_retained_projection_without_new_reads() {
    let body = format!(
        r#"<table:table table:name="Sheet">{}</table:table>"#,
        row(&cell("stable")),
    );
    let bytes = Builder::new().content_xml(content(&body)).build().unwrap();
    let reads = Arc::new(AtomicUsize::new(0));
    let source = Arc::new(CountingSource {
        bytes: Arc::new(bytes),
        reads: Arc::clone(&reads),
    });
    let document = SourceBackedSpreadsheet::from_read_at(source).unwrap();
    let after_open = reads.load(Ordering::Relaxed);
    let mut output = Vec::new();
    document
        .write_text_to(&mut output, TextOutputOptions::default())
        .unwrap();
    assert_eq!(output, b"stable");
    assert_eq!(reads.load(Ordering::Relaxed), after_open);
}

#[test]
fn concurrent_arc_source_backed_writers_have_independent_progress() {
    let body = format!(
        r#"<table:table table:name="Sheet">{}{}</table:table>"#,
        row(&cell("one")),
        row(&cell("two")),
    );
    let bytes = Builder::new().content_xml(content(&body)).build().unwrap();
    let document =
        Arc::new(SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap());
    let handles = (0..4)
        .map(|_| {
            let document = Arc::clone(&document);
            thread::spawn(move || {
                let mut output = Vec::new();
                let report = document
                    .write_text_to(&mut output, TextOutputOptions::new("|", "unused", 64, 8))
                    .unwrap();
                (output, report.bytes_written(), report.objects_written())
            })
        })
        .collect::<Vec<_>>();

    for handle in handles {
        let (output, bytes, objects) = handle.join().unwrap();
        assert_eq!(output, b"one|two");
        assert_eq!(bytes, 7);
        assert_eq!(objects, 2);
    }
}

#[test]
fn source_backed_sink_does_not_change_existing_text_projection() {
    let body = format!(
        r#"<table:table table:name="Sheet">{}</table:table>"#,
        row(&cell("stable")),
    );
    let bytes = Builder::new().content_xml(content(&body)).build().unwrap();
    let source_document =
        SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();
    let before = source_document.text().unwrap();
    let mut output = Vec::new();
    source_document
        .write_text_to(&mut output, TextOutputOptions::default())
        .unwrap();
    assert_eq!(output, before.as_bytes());
    assert_eq!(source_document.text().unwrap(), before);
}
