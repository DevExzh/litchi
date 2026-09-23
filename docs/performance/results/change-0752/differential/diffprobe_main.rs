//! Differential probe for change 0752.
//!
//! Built twice from one source file, once against the base crates and once
//! against the candidate, it drives the streaming DOCX, XLSX and PPTX writers
//! through limit sweeps, cancellation points, failing sinks and adversarial
//! text, and prints a transcript of every observable: each call's result,
//! error, reported progress, every budget counter, and each output digest.
//! The two transcripts must be byte-identical.
use litchi_core::{Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, Resource};
use litchi_docx::{StreamingDocumentLimits, StreamingDocumentWriter};
use litchi_pptx::{
    StreamingPresentationLimits, StreamingPresentationOptions, StreamingPresentationWriter,
    TextBoxSpec,
};
use litchi_xlsx::{StreamingCell, StreamingCellValue, StreamingWorkbookLimits, StreamingWorkbookWriter};
use sha2::{Digest, Sha256};
use std::io::{self, Write};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Arc;

const RESOURCES: [Resource; 6] = [
    Resource::Memory,
    Resource::InputBytes,
    Resource::OutputBytes,
    Resource::Objects,
    Resource::Depth,
    Resource::Work,
];

fn hex(bytes: &[u8]) -> String {
    Sha256::digest(bytes).iter().map(|b| format!("{b:02x}")).collect()
}

fn used(budgets: &[Budget]) -> String {
    budgets
        .iter()
        .map(|b| {
            RESOURCES
                .iter()
                .map(|r| b.used(*r).to_string())
                .collect::<Vec<_>>()
                .join("/")
        })
        .collect::<Vec<_>>()
        .join("|")
}

/// A sink that accepts at most `capacity` bytes in total and at most `chunk`
/// bytes per call, then fails; it can also cancel an operation on a byte.
struct Sink {
    bytes: Vec<u8>,
    capacity: usize,
    chunk: usize,
    zero_after: bool,
    cancel_at: Option<(usize, CancellationSource)>,
}

impl Sink {
    fn new() -> Self {
        Self { bytes: Vec::new(), capacity: usize::MAX, chunk: usize::MAX, zero_after: false, cancel_at: None }
    }
}

impl Write for Sink {
    fn write(&mut self, buf: &[u8]) -> io::Result<usize> {
        if let Some((at, source)) = &self.cancel_at {
            if self.bytes.len() + buf.len() > *at {
                source.cancel();
            }
        }
        let room = self.capacity.saturating_sub(self.bytes.len());
        if room == 0 && !buf.is_empty() {
            if self.zero_after {
                return Ok(0);
            }
            return Err(io::Error::new(io::ErrorKind::Other, "probe sink is full"));
        }
        let n = buf.len().min(room).min(self.chunk);
        self.bytes.extend_from_slice(&buf[..n]);
        Ok(n)
    }
    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn context(budget: Budget) -> (CancellationSource, ExecutionContext) {
    let (source, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(1 << 20).unwrap(),
        1,
    )
    .unwrap();
    (source, ExecutionContext::new(budget, token, limits))
}

fn big() -> Limits {
    Limits::new(1 << 30, 1 << 30, 1 << 30, 1 << 30, 64, 1 << 40)
}

fn with(limits: Limits, resource: Resource, value: u64) -> Limits {
    let mut values = [0u64; 6];
    for (i, r) in RESOURCES.iter().enumerate() {
        values[i] = if *r == resource { value } else { limits.get(*r) };
    }
    Limits::new(values[0], values[1], values[2], values[3], values[4], values[5])
}

// ----------------------------------------------------------------- DOCX

const DOCX_TEXTS: &[&str] = &[
    "plain",
    "litchi-perf-docx-streaming-v1-000123-caf\u{e9}-<&>",
    "&&&&&&&&&&&&&&&&<<<<<<<<<<<<>>>>>>>>>>>>>>>\"'",
    "\u{1F600}\u{4E2D}\u{20AC}\u{E9}\u{7F}\u{FFFD}\u{10FFFF}",
];

fn docx_long_texts() -> Vec<String> {
    vec![
        format!("{}&\u{E9}<>{}", "a".repeat(58), "b".repeat(5)),
        format!("{}\u{E9}", "a".repeat(63)),
        format!("{}\u{1F600}{}", "x".repeat(61), "y".repeat(130)),
        "q".repeat(64 * 3 + 1),
        format!("{}&", "z".repeat(60)),
        "\u{E9}".repeat(100),
    ]
}

/// The scripted operations; `texts` are written one run per paragraph.
fn docx_script(texts: &[String]) -> Vec<(&'static str, Option<usize>)> {
    let mut ops = Vec::new();
    for index in 0..texts.len() {
        ops.push(("start_paragraph", None));
        ops.push(("start_run", None));
        ops.push(("write_text", Some(index)));
        if index % 2 == 0 {
            ops.push(("write_text", Some((index + 1) % texts.len())));
        }
        ops.push(("finish_run", None));
        ops.push(("finish_paragraph", None));
    }
    ops.push(("finish", None));
    ops
}

fn docx_run(
    label: &str,
    budgets: &[Budget],
    leaf: Budget,
    limits: StreamingDocumentLimits,
    texts: &[String],
    mut sink: Sink,
    cancel_before: Option<usize>,
    cancel_in_sink: Option<usize>,
    out: &mut Vec<String>,
) {
    let (source, context) = context(leaf);
    if let Some(at) = cancel_in_sink {
        sink.cancel_at = Some((at, source.clone()));
    }
    let mut line = format!("docx {label}:");
    let mut writer = match StreamingDocumentWriter::new(sink, context, limits) {
        Ok(writer) => writer,
        Err(error) => {
            out.push(format!("{line} new -> {error:?} written={} used={}", error.written(), used(budgets)));
            return;
        },
    };
    line.push_str(&format!(" new ok out={} used={};", writer.output_bytes(), used(budgets)));
    for (step, (op, text)) in docx_script(texts).into_iter().enumerate() {
        if cancel_before == Some(step) {
            source.cancel();
        }
        let result = match op {
            "start_paragraph" => writer.start_paragraph(),
            "start_run" => writer.start_run(),
            "write_text" => writer.write_text(&texts[text.unwrap()]),
            "finish_run" => writer.finish_run(),
            "finish_paragraph" => writer.finish_paragraph(),
            _ => {
                let result = writer.finish();
                match result {
                    Ok(sink) => line.push_str(&format!(
                        " finish ok bytes={} sha={} used={}",
                        sink.bytes.len(),
                        hex(&sink.bytes),
                        used(budgets)
                    )),
                    Err(error) => line.push_str(&format!(
                        " finish -> {error:?} written={} used={}",
                        error.written(),
                        used(budgets)
                    )),
                }
                out.push(line);
                return;
            },
        };
        match result {
            Ok(()) => line.push_str(&format!(
                " {op} ok out={} doc={} in={} used={};",
                writer.output_bytes(),
                writer.document_xml_bytes(),
                writer.input_bytes(),
                used(budgets)
            )),
            Err(error) => {
                line.push_str(&format!(
                    " {op} -> {error:?} written={} poisoned={} used={}",
                    error.written(),
                    writer.is_poisoned(),
                    used(budgets)
                ));
                // One more call shows the poisoned (or retryable) state.
                let again = writer.start_paragraph();
                line.push_str(&format!(" then start_paragraph -> {again:?}"));
                out.push(line);
                return;
            },
        }
    }
}

fn docx_limits() -> StreamingDocumentLimits {
    StreamingDocumentLimits::new(1 << 20, 1 << 20, 64, 64, 8, 1 << 20, 1 << 20, 1 << 20, 64)
}

fn docx_probe(out: &mut Vec<String>) {
    let mut texts: Vec<String> = DOCX_TEXTS.iter().map(|t| t.to_string()).collect();
    texts.extend(docx_long_texts());
    // Reference run with an unlimited budget measures each total to sweep.
    let root = Budget::root("docx-root", big());
    docx_run("reference", &[root.clone()], root.clone(), docx_limits(), &texts, Sink::new(), None, None, out);
    let totals: Vec<(Resource, u64)> = RESOURCES.iter().map(|r| (*r, root.used(*r))).collect();
    for (resource, total) in totals {
        if resource == Resource::Depth {
            continue;
        }
        // The 64-byte scratch reservation is released by `finish`, so the
        // reference run ends with no Memory in use; sweep past the reservation.
        let total = if resource == Resource::Memory { 65 } else { total };
        let values: Vec<u64> = if total <= 4096 {
            (0..=total + 1).collect()
        } else {
            let mut v: Vec<u64> = (0..=64).collect();
            let step = (total / 700).max(1);
            v.extend((64..=total + 1).step_by(step as usize));
            v.extend(total.saturating_sub(70)..=total + 1);
            v
        };
        for value in values {
            // Single-level budget.
            let leaf = Budget::root("docx-leaf", with(big(), resource, value));
            docx_run(
                &format!("{resource:?}={value}"),
                &[leaf.clone()],
                leaf.clone(),
                docx_limits(),
                &texts,
                Sink::new(),
                None,
                None,
                out,
            );
            // Hierarchy: the parent is the tight level.
            let parent = Budget::root("docx-parent", with(big(), resource, value));
            let child = parent.child("docx-child", big());
            docx_run(
                &format!("parent {resource:?}={value}"),
                &[parent.clone(), child.clone()],
                child,
                docx_limits(),
                &texts,
                Sink::new(),
                None,
                None,
                out,
            );
        }
    }
    // Writer-local byte limits.
    let reference_output = {
        let mut probe = Vec::new();
        let root = Budget::root("r", big());
        docx_run("measure", &[root.clone()], root, docx_limits(), &texts, Sink::new(), None, None, &mut probe);
        probe.pop().unwrap()
    };
    out.push(format!("docx measured {reference_output}"));
    for field in 0..6 {
        for value in (0..=2600u64).step_by(7).chain(2590..=2700) {
            let mut limits = docx_limits();
            match field {
                0 => limits.max_output_bytes = value.max(22),
                1 => limits.max_input_bytes = value,
                2 => limits.max_run_text_bytes = value,
                3 => limits.max_paragraph_xml_bytes = value.max(11),
                4 => limits.max_document_xml_bytes = value.max(160),
                _ => limits.max_runs_per_paragraph = value % 3,
            }
            let root = Budget::root("docx-local", big());
            docx_run(&format!("field{field}={value}"), &[root.clone()], root, limits, &texts, Sink::new(), None, None, out);
        }
    }
    // Cancellation before each call, and inside the sink at each byte.
    for step in 0..60 {
        let root = Budget::root("docx-cancel", big());
        docx_run(&format!("cancel-before {step}"), &[root.clone()], root, docx_limits(), &texts, Sink::new(), Some(step), None, out);
    }
    for at in (0..3000).step_by(13) {
        let root = Budget::root("docx-cancel-sink", big());
        docx_run(&format!("cancel-in-sink {at}"), &[root.clone()], root, docx_limits(), &texts, Sink::new(), None, Some(at), out);
    }
    // Failing, zero-progress and short-write sinks.
    for capacity in (0..3000).step_by(11) {
        for (chunk, zero) in [(usize::MAX, false), (7, false), (usize::MAX, true)] {
            let root = Budget::root("docx-sink", big());
            let mut sink = Sink::new();
            sink.capacity = capacity;
            sink.chunk = chunk;
            sink.zero_after = zero;
            docx_run(&format!("sink cap={capacity} chunk={chunk} zero={zero}"), &[root.clone()], root, docx_limits(), &texts, sink, None, None, out);
        }
    }
    // Invalid characters at every position of a text longer than a Work
    // checkpoint, under Work limits around each checkpoint.
    let base_text: Vec<char> = format!("{}\u{E9}{}", "a".repeat(70), "b".repeat(70)).chars().collect();
    for position in (0..base_text.len()).step_by(3) {
        for invalid in ['\t', '\u{FFFE}', '\u{1}'] {
            let mut chars = base_text.clone();
            chars[position] = invalid;
            let text: String = chars.into_iter().collect();
            for work in [7, 8, 64 + 7, 64 + 8, 128 + 7, 128 + 8, 1 << 20] {
                let root = Budget::root("docx-invalid", with(big(), Resource::Work, work));
                docx_run(&format!("invalid {position} {:?} work={work}", invalid), &[root.clone()], root, docx_limits(), &[text.clone()], Sink::new(), None, None, out);
            }
        }
    }
}

// ----------------------------------------------------------------- XLSX

fn xlsx_rows() -> Vec<Vec<StreamingCell<'static>>> {
    let texts = ["plain", "caf\u{e9} <&> \"'", "\u{1F600}\u{4E2D}", "a longer text value that spans a little"];
    let mut rows = Vec::new();
    for row in 0..12usize {
        let mut cells = Vec::new();
        for column in 0..(row % 4 + 1) {
            let value = match (row + column) % 5 {
                0 => StreamingCellValue::Text(texts[(row + column) % texts.len()]),
                1 => StreamingCellValue::Number(row as f64 * 1.5 - column as f64),
                2 => StreamingCellValue::Bool(row % 2 == 0),
                3 => StreamingCellValue::Blank,
                _ => StreamingCellValue::Text(texts[row % texts.len()]),
            };
            cells.push(StreamingCell::new(u32::try_from(column * 3 + 1).unwrap(), value));
        }
        rows.push(cells);
    }
    rows
}

fn xlsx_run(label: &str, budgets: &[Budget], leaf: Budget, limits: StreamingWorkbookLimits, sink: Sink, out: &mut Vec<String>) {
    let (_source, context) = context(leaf);
    let mut line = format!("xlsx {label}:");
    let mut writer = match StreamingWorkbookWriter::new(sink, context, limits) {
        Ok(writer) => writer,
        Err(error) => {
            out.push(format!("{line} new -> {error:?} used={}", used(budgets)));
            return;
        },
    };
    line.push_str(&format!(" new ok out={} used={};", writer.output_bytes(), used(budgets)));
    for (index, cells) in xlsx_rows().into_iter().enumerate() {
        let row = u32::try_from(index * 2 + 1).unwrap();
        match writer.write_row(row, cells) {
            Ok(()) => line.push_str(&format!(
                " row{row} ok out={} xml={} cells={} used={};",
                writer.output_bytes(),
                writer.worksheet_xml_bytes(),
                writer.cell_count(),
                used(budgets)
            )),
            Err(error) => {
                line.push_str(&format!(" row{row} -> {error:?} poisoned={} out={} used={}", writer.is_poisoned(), writer.output_bytes(), used(budgets)));
                out.push(line);
                return;
            },
        }
    }
    match writer.finish() {
        Ok(sink) => line.push_str(&format!(" finish ok bytes={} sha={} used={}", sink.bytes.len(), hex(&sink.bytes), used(budgets))),
        Err(error) => line.push_str(&format!(" finish -> {error:?} used={}", used(budgets))),
    }
    out.push(line);
}

fn xlsx_limits() -> StreamingWorkbookLimits {
    StreamingWorkbookLimits::new(1000, 1000, 1000, 4096, 1 << 20, 1 << 20)
}

fn xlsx_probe(out: &mut Vec<String>) {
    let root = Budget::root("xlsx-root", big());
    xlsx_run("reference", &[root.clone()], root.clone(), xlsx_limits(), Sink::new(), out);
    let totals: Vec<(Resource, u64)> = RESOURCES.iter().map(|r| (*r, root.used(*r))).collect();
    for (resource, total) in totals {
        if resource == Resource::Depth {
            continue;
        }
        let top = if resource == Resource::Memory { 4096 + 2 } else { total + 1 };
        let values: Vec<u64> = if top <= 6000 { (0..=top).collect() } else { (0..=top).step_by((top / 900) as usize).chain(top.saturating_sub(80)..=top).collect() };
        for value in values {
            let leaf = Budget::root("xlsx-leaf", with(big(), resource, value));
            xlsx_run(&format!("{resource:?}={value}"), &[leaf.clone()], leaf, xlsx_limits(), Sink::new(), out);
            let parent = Budget::root("xlsx-parent", with(big(), resource, value));
            let child = parent.child("xlsx-child", big());
            xlsx_run(&format!("parent {resource:?}={value}"), &[parent.clone(), child.clone()], child, xlsx_limits(), Sink::new(), out);
        }
    }
    for value in (0..=4000u64).step_by(9) {
        let mut limits = xlsx_limits();
        limits.max_output_bytes = value.max(22);
        let root = Budget::root("xlsx-output", big());
        xlsx_run(&format!("max_output={value}"), &[root.clone()], root, limits, Sink::new(), out);
    }
    for capacity in (0..4000).step_by(17) {
        for chunk in [usize::MAX, 5] {
            let root = Budget::root("xlsx-sink", big());
            let mut sink = Sink::new();
            sink.capacity = capacity;
            sink.chunk = chunk;
            xlsx_run(&format!("sink cap={capacity} chunk={chunk}"), &[root.clone()], root, xlsx_limits(), sink, out);
        }
    }
}

// ----------------------------------------------------------------- PPTX

fn pptx_run(label: &str, limits: StreamingPresentationLimits, sink: Sink, out: &mut Vec<String>) {
    let texts = ["Title & <subtitle>", "caf\u{e9} \"quoted\" 'single'", "\u{1F600}\u{4E2D}\u{20AC}", "plain text box"];
    let mut line = format!("pptx {label}:");
    let mut writer = match StreamingPresentationWriter::with_options(sink, 3, StreamingPresentationOptions::default(), limits) {
        Ok(writer) => writer,
        Err(error) => {
            out.push(format!("{line} new -> {error:?}"));
            return;
        },
    };
    line.push_str(&format!(" new ok out={};", writer.output_bytes()));
    for slide in 0..3usize {
        let title = if slide == 1 { None } else { Some(texts[slide]) };
        let mut slide_writer = match writer.start_slide(title) {
            Ok(slide_writer) => slide_writer,
            Err(error) => {
                out.push(format!("{line} start_slide{slide} -> {error:?}"));
                return;
            },
        };
        for box_index in 0..(slide + 2) {
            let spec = TextBoxSpec::new(texts[(slide + box_index) % texts.len()], 100 * box_index as i64, 200, 3000, 4000);
            match slide_writer.write_text_box(spec) {
                Ok(()) => line.push_str(&format!(" box ok out={} xml={};", slide_writer.output_bytes(), slide_writer.slide_xml_bytes())),
                Err(error) => {
                    out.push(format!("{line} box -> {error:?} out={}", slide_writer.output_bytes()));
                    return;
                },
            }
        }
        writer = match slide_writer.finish() {
            Ok(writer) => writer,
            Err(error) => {
                out.push(format!("{line} slide finish -> {error:?}"));
                return;
            },
        };
        line.push_str(&format!(" slide{slide} ok out={};", writer.output_bytes()));
    }
    match writer.finish() {
        Ok(sink) => line.push_str(&format!(" finish ok bytes={} sha={}", sink.bytes.len(), hex(&sink.bytes))),
        Err(error) => line.push_str(&format!(" finish -> {error:?}")),
    }
    out.push(line);
}

fn pptx_probe(out: &mut Vec<String>) {
    pptx_run("reference", StreamingPresentationLimits::default(), Sink::new(), out);
    for value in (0..=80_000u64).step_by(97) {
        let mut limits = StreamingPresentationLimits::default();
        limits.max_output_bytes = value.max(1);
        pptx_run(&format!("max_output={value}"), limits, Sink::new(), out);
    }
    for capacity in (0..80_000).step_by(173) {
        let mut sink = Sink::new();
        sink.capacity = capacity;
        pptx_run(&format!("sink cap={capacity}"), StreamingPresentationLimits::default(), sink, out);
    }
}

/// Writes the harness's large streaming DOCX corpus (131,072 one-run
/// paragraphs of `litchi-perf-docx-streaming-v1-{index:06}-café-<&>`) to
/// `path`, with limits that refuse nothing.
fn write_large_docx(path: &std::ffi::OsStr) {
    let root = Budget::root("dump", big());
    let (_source, context) = context(root);
    let limits = StreamingDocumentLimits::new(1 << 30, 1 << 29, 1 << 20, 1 << 20, 8, 1 << 20, 1 << 20, 1 << 29, 64);
    let mut writer = StreamingDocumentWriter::new(Sink::new(), context, limits).unwrap();
    for index in 0..131_072usize {
        writer.start_paragraph().unwrap();
        writer.start_run().unwrap();
        writer.write_text(&format!("litchi-perf-docx-streaming-v1-{index:06}-caf\u{e9}-<&>")).unwrap();
        writer.finish_run().unwrap();
        writer.finish_paragraph().unwrap();
    }
    let sink = writer.finish().unwrap();
    std::fs::write(path, &sink.bytes).unwrap();
    eprintln!("wrote {} bytes sha256 {}", sink.bytes.len(), hex(&sink.bytes));
}

fn main() {
    if let Some(path) = std::env::var_os("DOCX_LARGE_OUT") {
        write_large_docx(&path);
        return;
    }
    let mut out = Vec::new();
    docx_probe(&mut out);
    xlsx_probe(&mut out);
    pptx_probe(&mut out);
    let transcript = out.join("\n");
    let stdout = io::stdout();
    let mut handle = stdout.lock();
    writeln!(handle, "{transcript}").unwrap();
    eprintln!("lines {} transcript sha256 {}", out.len(), hex(transcript.as_bytes()));
    let _ = Arc::new(AtomicBool::new(false)).load(Ordering::Relaxed);
}
