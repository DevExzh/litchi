#![allow(
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use std::io::{self, Write};
use std::sync::{
    Arc,
    atomic::{AtomicU64, AtomicUsize, Ordering},
};

use litchi_core::{
    Error, ReadAt, SourceVersion, TextOutputError, TextOutputLimitKind, TextOutputOptions,
};
use litchi_odp::{Presentation, SourceBackedPresentation};

const MIME: &str = "application/vnd.oasis.opendocument.presentation";
const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const DRAW: &str = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const PRESENTATION: &str = "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

fn package(body: &str) -> Vec<u8> {
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:draw="{DRAW}" xmlns:presentation="{PRESENTATION}" xmlns:text="{TEXT}" office:version="1.3"><office:body><office:presentation>{body}</office:presentation></office:body></office:document-content>"#
    );
    let mut writer = litchi_odp::core::PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer.add_file("content.xml", content.as_bytes()).unwrap();
    writer.finish_to_bytes().unwrap()
}

fn representative_package() -> Vec<u8> {
    package(concat!(
        r#"<draw:page><draw:frame presentation:class="title"><draw:text-box><text:p>&#32;&#32;Title&#32;&#32;</text:p></draw:text-box></draw:frame>"#,
        r#"<draw:frame><draw:text-box><text:p>body one</text:p><text:p>body two</text:p></draw:text-box></draw:frame>"#,
        r#"<draw:rect><draw:text-box><text:p>outer shape</text:p></draw:text-box></draw:rect><draw:g><draw:rect><draw:text-box><text:p>inner shape</text:p></draw:text-box></draw:rect></draw:g></draw:page>"#,
        r#"<draw:page><draw:frame><draw:text-box><text:p>&#32;&#32;&#32;</text:p></draw:text-box></draw:frame><draw:g><draw:rect><draw:text-box><text:p>retained child</text:p></draw:text-box></draw:rect></draw:g></draw:page>"#,
        r#"<draw:page/>"#,
    ))
}

fn write_owned(presentation: &Presentation, options: TextOutputOptions<'_>) -> (Vec<u8>, u64, u64) {
    let mut output = Vec::new();
    let report = presentation.write_text_to(&mut output, options).unwrap();
    (output, report.bytes_written(), report.objects_written())
}

#[test]
fn semantic_slide_text_matches_all_text_and_preserves_empty_slide_objects() {
    let presentation = Presentation::from_bytes(representative_package()).unwrap();
    let slides = presentation.slides().unwrap();
    let expected = slides
        .iter()
        .map(litchi_odp::Slide::all_text)
        .collect::<Vec<_>>();
    assert_eq!(
        expected[0],
        "Title\nbody one\nbody two\nouter shape\ninner shape"
    );
    assert_eq!(expected[1], "retained child");
    assert_eq!(expected[2], "");

    let (output, bytes_written, objects_written) =
        write_owned(&presentation, TextOutputOptions::default());
    let expected_output = expected.join("\n\n");
    assert_eq!(output, expected_output.as_bytes());
    assert_eq!(bytes_written, expected_output.len() as u64);
    assert_eq!(objects_written, 3);

    let (without_empty, _, without_empty_objects) = write_owned(
        &presentation,
        TextOutputOptions::default().with_empty_objects(false),
    );
    assert_eq!(
        without_empty,
        b"Title\nbody one\nbody two\nouter shape\ninner shape\n\nretained child"
    );
    assert_eq!(without_empty_objects, 2);
}

#[test]
fn controls_entities_cdata_utf8_and_custom_separators_match_semantics() {
    let presentation = Presentation::from_bytes(package(
        concat!(
            r#"<draw:page><draw:frame><draw:text-box><text:p> lead &amp; &#x1f600; <![CDATA[<cdata>a"#,
            "\r\n",
            r#"b]]><text:s text:c="2"/><text:tab/>尾<text:line-break/>end</text:p></draw:text-box></draw:frame></draw:page>"#,
        ),
    ))
    .unwrap();
    let expected = presentation.slides().unwrap()[0].all_text();
    assert_eq!(expected, "lead & 😀 <cdata>a b  \t尾\nend");

    let (output, bytes_written, objects_written) =
        write_owned(&presentation, TextOutputOptions::new("::", "||", 1024, 4));
    assert_eq!(output, expected.as_bytes());
    assert_eq!(bytes_written, expected.len() as u64);
    assert_eq!(objects_written, 1);
}

#[test]
fn paired_text_controls_are_applied_on_start_events() {
    let presentation = Presentation::from_bytes(package(
        r#"<draw:page><draw:frame><draw:text-box><text:p>A<text:s text:c="2"></text:s><text:tab></text:tab>B<text:line-break></text:line-break>C</text:p></draw:text-box></draw:frame></draw:page>"#,
    ))
    .unwrap();
    let mut output = Vec::new();
    let report = presentation
        .write_text_to(&mut output, TextOutputOptions::default())
        .unwrap();

    assert_eq!(output, b"A  \tB\nC");
    assert_eq!(report.bytes_written(), 7);
    assert_eq!(report.objects_written(), 1);
}

#[test]
fn paragraph_and_slide_separators_are_applied_without_joining_the_document() {
    let presentation = Presentation::from_bytes(representative_package()).unwrap();
    let options = TextOutputOptions::new("::", "||", 1024, 4);
    let (output, bytes_written, objects_written) = write_owned(&presentation, options);

    assert_eq!(
        output,
        b"Title::body one::body two::outer shape::inner shape||retained child||"
    );
    assert_eq!(bytes_written, output.len() as u64);
    assert_eq!(objects_written, 3);
}

#[test]
fn object_limit_reports_rejected_slide_with_exact_progress() {
    let presentation = Presentation::from_bytes(representative_package()).unwrap();
    let mut output = Vec::new();
    let error = presentation
        .write_text_to(&mut output, TextOutputOptions::new("\n", "\n\n", 1024, 1))
        .unwrap_err();

    assert_eq!(
        output,
        b"Title\nbody one\nbody two\nouter shape\ninner shape"
    );
    assert_eq!(error.progress().bytes_written(), output.len() as u64);
    assert_eq!(error.progress().objects_written(), 1);
    let limit = error.limit().unwrap();
    assert_eq!(limit.kind(), TextOutputLimitKind::Objects);
    assert_eq!(limit.observed(), 2);
    assert_eq!(limit.limit(), 1);
}

#[test]
fn output_limit_accounts_for_slide_separator_and_next_slide() {
    let presentation = Presentation::from_bytes(representative_package()).unwrap();
    let first = presentation.slides().unwrap()[0].all_text();
    let mut output = Vec::new();
    let error = presentation
        .write_text_to(
            &mut output,
            TextOutputOptions::new("\n", "||", first.len() as u64, 4),
        )
        .unwrap_err();

    assert_eq!(output, first.as_bytes());
    assert_eq!(error.progress().bytes_written(), first.len() as u64);
    assert_eq!(error.progress().objects_written(), 1);
    let limit = error.limit().unwrap();
    assert_eq!(limit.kind(), TextOutputLimitKind::OutputBytes);
    assert_eq!(
        limit.observed(),
        (first.len() + 2 + "retained child".len()) as u64
    );
    assert_eq!(limit.limit(), first.len() as u64);
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
fn sink_failures_preserve_utf8_prefix_progress_and_write_zero() {
    let presentation = Presentation::from_bytes(package(
        r#"<draw:page><draw:frame><draw:text-box><text:p>αβγ</text:p></draw:text-box></draw:frame></draw:page>"#,
    ))
    .unwrap();
    let mut sink = PrefixThenFail {
        accepted: Vec::new(),
        prefix_limit: 2,
    };
    let error = presentation
        .write_text_to(&mut sink, TextOutputOptions::default())
        .unwrap_err();
    assert_eq!(sink.accepted, "α".as_bytes());
    assert_eq!(error.progress().bytes_written(), 2);
    assert_eq!(error.progress().objects_written(), 0);
    assert!(matches!(error, TextOutputError::Sink { .. }));

    let mut zero = ZeroWriter;
    let error = presentation
        .write_text_to(&mut zero, TextOutputOptions::default())
        .unwrap_err();
    assert_eq!(error.progress().bytes_written(), 0);
    assert!(matches!(error, TextOutputError::Sink { .. }));
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

struct ShortWriter {
    output: Vec<u8>,
    max_write: usize,
}

impl Write for ShortWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let amount = self.max_write.min(bytes.len());
        self.output.extend_from_slice(&bytes[..amount]);
        Ok(amount)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn eventually_successful_short_writes_preserve_complete_progress() {
    let presentation = Presentation::from_bytes(representative_package()).unwrap();
    let mut sink = ShortWriter {
        output: Vec::new(),
        max_write: 1,
    };
    let report = presentation
        .write_text_to(&mut sink, TextOutputOptions::default())
        .unwrap();

    assert_eq!(
        sink.output,
        b"Title\nbody one\nbody two\nouter shape\ninner shape\n\nretained child\n\n"
    );
    assert_eq!(report.bytes_written(), sink.output.len() as u64);
    assert_eq!(report.objects_written(), 3);
}

#[test]
fn malformed_later_slide_returns_document_error_after_prior_progress() {
    let bytes = package(
        r#"<draw:page><draw:frame><draw:text-box><text:p>first</text:p></draw:text-box></draw:frame></draw:page><draw:page><draw:frame><draw:text-box><text:p><text:p>broken</text:p></text:p></draw:text-box></draw:frame></draw:page>"#,
    );
    let presentation = Presentation::from_bytes(bytes).unwrap();
    let mut output = Vec::new();
    let error = presentation
        .write_text_to(&mut output, TextOutputOptions::default())
        .unwrap_err();

    assert_eq!(output, b"first");
    assert_eq!(error.progress().bytes_written(), 5);
    assert_eq!(error.progress().objects_written(), 1);
    assert!(matches!(error, TextOutputError::Document { .. }));
}

struct CountingSource {
    bytes: Vec<u8>,
    reads: AtomicUsize,
    revision: AtomicU64,
}

impl CountingSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            reads: AtomicUsize::new(0),
            revision: AtomicU64::new(0),
        }
    }

    fn bump_revision(&self) {
        self.revision.fetch_add(1, Ordering::Relaxed);
    }
}

impl ReadAt for CountingSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_| io::Error::other("test source length does not fit u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.reads.fetch_add(1, Ordering::Relaxed);
        let offset = usize::try_from(offset)
            .map_err(|error| io::Error::new(io::ErrorKind::InvalidInput, error))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let amount = output.len().min(self.bytes.len() - offset);
        output[..amount].copy_from_slice(&self.bytes[offset..offset + amount]);
        Ok(amount)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x4f44_5002,
            self.revision.load(Ordering::Relaxed),
        ))
    }
}

#[test]
fn source_backed_sink_matches_owned_and_does_not_use_text_cache() {
    let bytes = representative_package();
    let owned = Presentation::from_bytes(bytes.clone()).unwrap();
    let source = Arc::new(CountingSource::new(bytes));
    let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
    let reads_before = source.reads.load(Ordering::Relaxed);
    let mut output = Vec::new();
    let report = presentation
        .write_text_to(&mut output, TextOutputOptions::default())
        .unwrap();

    assert_eq!(
        output,
        b"Title\nbody one\nbody two\nouter shape\ninner shape\n\nretained child\n\n"
    );
    assert_eq!(report.objects_written(), 3);
    assert_eq!(source.reads.load(Ordering::Relaxed), reads_before);
    assert_eq!(presentation.text().unwrap(), owned.text().unwrap());
    assert_eq!(source.revision.load(Ordering::Relaxed), 0);
}

#[test]
fn source_backed_sink_rejects_oversized_text_space_attribute() {
    let count = "1".repeat(1_048_577);
    let content = format!(
        r#"<draw:page><draw:frame><draw:text-box><text:p><text:s text:c="{count}"/></text:p></draw:text-box></draw:frame></draw:page>"#
    );
    let source = Arc::new(CountingSource::new(package(&content)));
    let presentation = SourceBackedPresentation::from_read_at(source).unwrap();
    let mut output = Vec::new();
    let error = presentation
        .write_text_to(&mut output, TextOutputOptions::default())
        .unwrap_err();

    assert!(output.is_empty());
    assert!(matches!(error, TextOutputError::Document { .. }));
}

#[test]
fn source_backed_sink_rejects_oversized_element_name() {
    let name = "n".repeat(1_048_577);
    let content = format!(r#"<draw:page><{name}/></draw:page>"#);
    let source = Arc::new(CountingSource::new(package(&content)));
    let presentation = SourceBackedPresentation::from_read_at(source).unwrap();
    let mut output = Vec::new();
    let error = presentation
        .write_text_to(&mut output, TextOutputOptions::default())
        .unwrap_err();

    assert!(output.is_empty());
    assert!(matches!(error, TextOutputError::Document { .. }));
}

#[test]
fn source_backed_sink_does_not_echo_unsupported_entity_name() {
    let name = "e".repeat(1_048_577);
    let content = format!(
        r#"<draw:page><draw:frame><draw:text-box><text:p>&{name};</text:p></draw:text-box></draw:frame></draw:page>"#
    );
    let source = Arc::new(CountingSource::new(package(&content)));
    let presentation = SourceBackedPresentation::from_read_at(source).unwrap();
    let mut output = Vec::new();
    let error = presentation
        .write_text_to(&mut output, TextOutputOptions::default())
        .unwrap_err();

    assert!(output.is_empty());
    assert!(matches!(
        error,
        TextOutputError::Document {
            source: Error::InvalidFormat(message),
            ..
        } if message.len() < 128
    ));
}

#[test]
fn source_backed_sink_reports_stale_before_call_with_zero_progress() {
    let source = Arc::new(CountingSource::new(representative_package()));
    let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
    source.bump_revision();
    let mut output = Vec::new();
    let error = presentation
        .write_text_to(&mut output, TextOutputOptions::default())
        .unwrap_err();

    assert!(output.is_empty());
    assert_eq!(error.progress().bytes_written(), 0);
    assert!(matches!(
        error,
        TextOutputError::Document {
            source: Error::SourceChanged { .. },
            ..
        }
    ));
}

struct StaleAfterPrefix {
    source: Arc<CountingSource>,
    accepted: Vec<u8>,
    prefix_limit: usize,
}

impl Write for StaleAfterPrefix {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.accepted.len() >= self.prefix_limit {
            return Err(io::Error::other("injected sink failure"));
        }
        let amount = (self.prefix_limit - self.accepted.len()).min(bytes.len());
        self.accepted.extend_from_slice(&bytes[..amount]);
        self.source.bump_revision();
        Ok(amount)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn source_change_takes_precedence_over_sink_failure_with_progress() {
    let source = Arc::new(CountingSource::new(representative_package()));
    let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
    let mut sink = StaleAfterPrefix {
        source,
        accepted: Vec::new(),
        prefix_limit: 3,
    };
    let error = presentation
        .write_text_to(&mut sink, TextOutputOptions::default())
        .unwrap_err();

    assert_eq!(sink.accepted, b"Tit");
    assert_eq!(error.progress().bytes_written(), 3);
    assert_eq!(error.progress().objects_written(), 0);
    assert!(matches!(
        error,
        TextOutputError::Document {
            source: Error::SourceChanged { .. },
            ..
        }
    ));
}

struct StaleAfterSuccess {
    source: Arc<CountingSource>,
    output: Vec<u8>,
    changed: bool,
}

impl Write for StaleAfterSuccess {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.output.extend_from_slice(bytes);
        if !self.changed {
            self.changed = true;
            self.source.bump_revision();
        }
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn source_change_after_successful_sink_write_takes_precedence() {
    let source = Arc::new(CountingSource::new(representative_package()));
    let presentation = SourceBackedPresentation::from_read_at(source.clone()).unwrap();
    let mut sink = StaleAfterSuccess {
        source,
        output: Vec::new(),
        changed: false,
    };
    let error = presentation
        .write_text_to(&mut sink, TextOutputOptions::default())
        .unwrap_err();

    assert_eq!(
        sink.output,
        b"Title\nbody one\nbody two\nouter shape\ninner shape\n\nretained child\n\n"
    );
    assert_eq!(error.progress().bytes_written(), sink.output.len() as u64);
    assert_eq!(error.progress().objects_written(), 3);
    assert!(matches!(
        error,
        TextOutputError::Document {
            source: Error::SourceChanged { .. },
            ..
        }
    ));
}
