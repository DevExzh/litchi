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

/// Writes `paragraphs`, one run each, handing every run's text to
/// `write_text` in pieces of `piece(n)` bytes (rounded up to a character
/// boundary) for the n-th piece.
fn render_split(paragraphs: &[String], piece: &dyn Fn(usize) -> usize) -> Vec<u8> {
    let (_budget, _source, execution) = context();
    let mut writer =
        StreamingDocumentWriter::new(Vec::new(), execution, limits()).expect("streaming writer");
    for text in paragraphs {
        writer.start_paragraph().expect("paragraph");
        writer.start_run().expect("run");
        let mut rest = text.as_str();
        let mut number = 0;
        while !rest.is_empty() {
            let mut take = piece(number).clamp(1, rest.len());
            while !rest.is_char_boundary(take) {
                take += 1;
            }
            writer.write_text(&rest[..take]).expect("text piece");
            rest = &rest[take..];
            number += 1;
        }
        writer.finish_run().expect("run finish");
        writer.finish_paragraph().expect("paragraph finish");
    }
    writer.finish().expect("package finish")
}

/// Change 0762: the document member is compressed in fixed chunks cut at
/// absolute member offsets, so the package bytes are a function of the
/// document alone, not of how the caller split each run's text.
#[test]
fn package_bytes_do_not_depend_on_how_run_text_is_split() {
    let paragraphs: Vec<String> = (0..600)
        .map(|index| {
            format!(
                "paragraph {index:04}: café & <tags> {}",
                "lorem ipsum dolor ".repeat(index % 7)
            )
        })
        .collect();
    let whole = render_split(&paragraphs, &|_| usize::MAX);
    let physical = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(whole.clone())
        .expect("physical package");
    let document_xml = physical
        .blob_for(&PackURI::new("/word/document.xml").expect("document URI"))
        .expect("document XML");
    // The member spans several 16 KiB chunks.
    assert!(document_xml.len() > 3 * 16 * 1024);
    let pieces: [&dyn Fn(usize) -> usize; 4] = [
        &|_| 1,
        &|number| 1 + number % 5,
        &|number| 3 + (number * 7) % 11,
        &|number| if number % 2 == 0 { 2 } else { 40 },
    ];
    for piece in pieces {
        assert_eq!(render_split(&paragraphs, piece), whole);
    }
    let package = Package::from_reader(Cursor::new(whole)).expect("reopen DOCX");
    let document = package.document().expect("main document");
    assert_eq!(
        document.paragraph_count().expect("paragraph count"),
        paragraphs.len()
    );
    for (paragraph, expected) in document
        .paragraphs()
        .expect("paragraphs")
        .iter()
        .zip(&paragraphs)
    {
        assert_eq!(&paragraph.text().expect("paragraph text"), expected);
    }
}

// ---------------------------------------------------------------------------
// Change 0763: the writer charges Objects, Work and input bytes through rough
// budget leases. A sole holder of the budget must see exactly the refusals of
// exact accounting; the leases must be returned when the writer is poisoned
// and when it finishes.
// ---------------------------------------------------------------------------

/// One API call of the lease script.
#[derive(Debug, Clone, Copy)]
enum Call<'a> {
    StartParagraph,
    StartRun,
    Text(&'a str),
    FinishRun,
    FinishParagraph,
}

fn lease_script(texts: &[String]) -> Vec<Call<'_>> {
    let mut calls = Vec::new();
    for (index, text) in texts.iter().enumerate() {
        calls.push(Call::StartParagraph);
        calls.push(Call::StartRun);
        if index % 3 == 2 {
            // Two text calls in one run.
            let split = text
                .char_indices()
                .nth(text.chars().count() / 2)
                .map_or(0, |(at, _)| at);
            calls.push(Call::Text(&text[..split]));
            calls.push(Call::Text(&text[split..]));
        } else {
            calls.push(Call::Text(text));
        }
        calls.push(Call::FinishRun);
        calls.push(Call::FinishParagraph);
    }
    calls
}

/// The exact charges each call makes, in order: the writer's documented
/// accounting (Work in 64-character checkpoints, then the text's bytes).
fn call_charges(call: Call<'_>) -> Vec<(Resource, u64)> {
    match call {
        Call::StartParagraph | Call::StartRun => {
            vec![(Resource::Objects, 1), (Resource::Work, 1)]
        },
        Call::Text(text) => {
            let characters = u64::try_from(text.chars().count()).expect("characters");
            let mut charges =
                vec![(Resource::Work, 64); usize::try_from(characters / 64).expect("pieces")];
            if characters % 64 != 0 {
                charges.push((Resource::Work, characters % 64));
            }
            charges.push((
                Resource::InputBytes,
                u64::try_from(text.len()).expect("bytes"),
            ));
            charges
        },
        Call::FinishRun | Call::FinishParagraph => vec![(Resource::Work, 1)],
    }
}

fn apply_call<W: Write>(
    writer: &mut StreamingDocumentWriter<W>,
    call: Call<'_>,
) -> Result<(), StreamingDocumentError> {
    match call {
        Call::StartParagraph => writer.start_paragraph(),
        Call::StartRun => writer.start_run(),
        Call::Text(text) => writer.write_text(text),
        Call::FinishRun => writer.finish_run(),
        Call::FinishParagraph => writer.finish_paragraph(),
    }
}

fn budget_with(resource: Resource, limit: u64) -> Budget {
    let value = |candidate: Resource| {
        if candidate == resource {
            limit
        } else {
            1 << 40
        }
    };
    Budget::root(
        "docx-lease-limit",
        CoreLimits::new(
            1024 * 1024,
            value(Resource::InputBytes),
            1 << 40,
            value(Resource::Objects),
            64,
            value(Resource::Work),
        ),
    )
}

#[test]
fn a_sole_writer_is_refused_exactly_at_every_objects_work_and_input_limit() {
    let texts: Vec<String> = (0..9)
        .map(|index| format!("lease {index}: café & <b> {}", "é".repeat(index * 17)))
        .collect();
    let calls = lease_script(&texts);
    // `new` charges five Objects and two Work units exactly, before any
    // lease exists.
    for (resource, fixed, name) in [
        (Resource::Objects, 5_u64, "objects"),
        (Resource::Work, 2, "work"),
        (Resource::InputBytes, 0, "input bytes"),
    ] {
        let total: u64 = fixed
            + calls
                .iter()
                .flat_map(|&call| call_charges(call))
                .filter(|(charged, _)| *charged == resource)
                .map(|(_, amount)| amount)
                .sum::<u64>();
        for limit in fixed..=total + 1 {
            // The exact model: the first charge that would pass the limit.
            let mut used = fixed;
            let mut expected = None;
            'calls: for (index, &call) in calls.iter().enumerate() {
                for (charged, amount) in call_charges(call) {
                    if charged != resource {
                        continue;
                    }
                    if used + amount > limit {
                        expected = Some((index, used + amount));
                        break 'calls;
                    }
                    used += amount;
                }
            }
            let budget = budget_with(resource, limit);
            let (_source, execution) = context_for(budget.clone());
            let mut writer = StreamingDocumentWriter::new(Vec::new(), execution, limits())
                .expect("streaming writer");
            let mut refused = None;
            for (index, &call) in calls.iter().enumerate() {
                if let Err(error) = apply_call(&mut writer, call) {
                    refused = Some((index, error));
                    break;
                }
            }
            match (expected, refused) {
                (None, None) => {
                    writer.finish().expect("package finish");
                    assert_eq!(budget.used(resource), total, "{name} limit {limit}");
                },
                (Some((index, observed)), Some((actual, error))) => {
                    assert_eq!(actual, index, "{name} limit {limit}");
                    assert!(
                        matches!(
                            error,
                            StreamingDocumentError::LimitExceeded {
                                resource: refused_resource,
                                observed: refused_observed,
                                limit: refused_limit,
                                ..
                            } if refused_resource == name
                                && refused_observed == observed
                                && refused_limit == limit
                        ),
                        "{name} limit {limit}: {error:?}"
                    );
                    // The poisoned writer returned its leases: the budget
                    // shows exactly the charges made before the refusal.
                    assert!(writer.is_poisoned());
                    assert_eq!(budget.used(resource), used, "{name} limit {limit}");
                },
                (expected, refused) => {
                    panic!("{name} limit {limit}: expected {expected:?}, writer gave {refused:?}")
                },
            }
        }
    }
}

#[test]
fn a_sibling_sees_the_writers_pre_claim_until_the_writer_finishes() {
    let root = Budget::root(
        "docx-lease-root",
        CoreLimits::new(1024 * 1024, 1 << 30, 1 << 30, 1 << 20, 64, 100_000),
    );
    let child = |name: &str| {
        root.child(
            name.to_owned(),
            CoreLimits::new(1024 * 1024, 1 << 30, 1 << 30, 1 << 20, 64, 100_000),
        )
    };
    let (writer_budget, sibling) = (child("writer"), child("sibling"));
    let (_source, execution) = context_for(writer_budget.clone());
    let mut writer =
        StreamingDocumentWriter::new(Vec::new(), execution, limits()).expect("streaming writer");
    writer.start_paragraph().expect("paragraph");
    // The writer used 3 Work units and holds the rest of one 64 Ki chunk,
    // which the sibling observes as used.
    assert_eq!(writer_budget.used(Resource::Work), 64 * 1024 + 2);
    let refused = sibling
        .consume(Resource::Work, 100_000 - 3)
        .expect_err("exact accounting would admit this; the pre-claim does not");
    assert_eq!(refused.limit, 100_000);
    assert!(root.used(Resource::Work) <= 100_000);
    writer.start_run().expect("run");
    writer.write_text("sibling").expect("text");
    writer.finish_run().expect("run finish");
    writer.finish_paragraph().expect("paragraph finish");
    writer.finish().expect("package finish");
    // Finishing returned the lease: exactly the writer's charges remain.
    let used = 2 + 1 + 1 + 7 + 1 + 1;
    assert_eq!(writer_budget.used(Resource::Work), used);
    sibling
        .consume(Resource::Work, 100_000 - used)
        .expect("the rest of the root");
    assert_eq!(root.used(Resource::Work), 100_000);
}

#[test]
fn leases_are_returned_when_the_writer_is_poisoned_or_dropped() {
    let budget = Budget::root(
        "docx-lease-poison",
        CoreLimits::new(1024 * 1024, 1 << 30, 1 << 30, 1 << 20, 64, 1 << 30),
    );
    let (source, execution) = context_for(budget.clone());
    let mut writer =
        StreamingDocumentWriter::new(Vec::new(), execution, limits()).expect("streaming writer");
    writer.start_paragraph().expect("paragraph");
    writer.start_run().expect("run");
    writer.write_text("café").expect("text");
    assert!(
        budget.used(Resource::Work) > 8,
        "the Work lease holds a chunk"
    );
    source.cancel();
    let error = writer.finish_run().expect_err("cancelled");
    assert!(matches!(error, StreamingDocumentError::Cancelled { .. }));
    // Cancellation poisoned the writer, which returned every lease.
    assert_eq!(budget.used(Resource::Objects), 5 + 2);
    assert_eq!(budget.used(Resource::Work), 2 + 1 + 1 + 4);
    assert_eq!(budget.used(Resource::InputBytes), 5);
    drop(writer);
    assert_eq!(budget.used(Resource::Memory), 0);

    // Dropping a healthy writer returns its leases too.
    let budget = Budget::root(
        "docx-lease-drop",
        CoreLimits::new(1024 * 1024, 1 << 30, 1 << 30, 1 << 20, 64, 1 << 30),
    );
    let (_source, execution) = context_for(budget.clone());
    let mut writer =
        StreamingDocumentWriter::new(Vec::new(), execution, limits()).expect("streaming writer");
    writer.start_paragraph().expect("paragraph");
    assert!(budget.used(Resource::Objects) > 6);
    drop(writer);
    assert_eq!(budget.used(Resource::Objects), 6);
    assert_eq!(budget.used(Resource::Work), 3);
}
