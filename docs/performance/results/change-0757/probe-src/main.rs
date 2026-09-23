//! Probe for change 0757: the fresh XLS writer on worksheets with many
//! distinct strings, the path whose shared-string order the change fixes, and
//! the harness's `{xls,doc,ppt}_fresh_write_to` bodies (DOC and PPT for the
//! 0753 review follow-ups), timed per write under a counting global allocator.
//! Scratch tool, not production code.
//!
//! usage: probe ITERATIONS CASE
//! prints one JSON object: per-write nanoseconds and allocation counts.
use std::alloc::{GlobalAlloc, Layout, System};
use std::io::Cursor;
use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering::Relaxed};
use std::time::Instant;

struct Counting;
static ALLOCS: AtomicU64 = AtomicU64::new(0);
static REALLOCS: AtomicU64 = AtomicU64::new(0);
static BYTES: AtomicU64 = AtomicU64::new(0);
static LIVE: AtomicUsize = AtomicUsize::new(0);
static PEAK: AtomicUsize = AtomicUsize::new(0);

fn grow(by: usize) {
    let live = LIVE.fetch_add(by, Relaxed) + by;
    PEAK.fetch_max(live, Relaxed);
}

unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        ALLOCS.fetch_add(1, Relaxed);
        BYTES.fetch_add(layout.size() as u64, Relaxed);
        grow(layout.size());
        unsafe { System.alloc(layout) }
    }
    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        LIVE.fetch_sub(layout.size(), Relaxed);
        unsafe { System.dealloc(ptr, layout) }
    }
    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        REALLOCS.fetch_add(1, Relaxed);
        if new_size > layout.size() {
            BYTES.fetch_add((new_size - layout.size()) as u64, Relaxed);
            grow(new_size - layout.size());
        } else {
            LIVE.fetch_sub(layout.size() - new_size, Relaxed);
        }
        unsafe { System.realloc(ptr, layout, new_size) }
    }
}

#[global_allocator]
static GLOBAL: Counting = Counting;

fn writer_text(kind: &str, first: usize, second: usize, third: usize) -> String {
    format!("litchi-perf-baseline-{kind}-v1-{first:03}-{second:05}-{third:03} deterministic payload")
}

fn writer_payload_text(kind: &str, first: usize, second: usize, third: usize, length: usize) -> String {
    const REPEATED_TEXT: &str = "litchi-perf-baseline-payload-heavy-v1 ";
    let mut text = writer_text(kind, first, second, third);
    while text.len() < length {
        text.push_str(REPEATED_TEXT);
    }
    text.truncate(length);
    text
}

/// The harness's `write_fresh_xls` for its three shapes.
fn harness(shape: &str) -> litchi_xls::writer::Writer {
    let mut writer = litchi_xls::writer::Writer::new();
    match shape {
        "tiny" | "large" => {
            let (sheets, rows, columns) = if shape == "tiny" { (1, 4, 4) } else { (4, 128, 16) };
            for sheet in 0..sheets {
                let worksheet = writer.add_worksheet(&format!("Bench{sheet:02}")).unwrap();
                for row in 0..rows {
                    for column in 0..columns {
                        let value = (sheet * rows * columns + row * columns + column) as f64;
                        writer.write_number(worksheet, row as u32, column as u16, value).unwrap();
                    }
                }
            }
        },
        _ => {
            for sheet in 0..128 {
                let worksheet = writer.add_worksheet(&format!("Payload{sheet:03}")).unwrap();
                let text = writer_payload_text("xls", sheet, 0, 0, 32_700);
                writer.write_string(worksheet, 0, 0, &text).unwrap();
            }
        },
    }
    writer
}

/// Four worksheets of 2,500 rows by 8 columns of string cells: every string
/// distinct, or drawn from 64 labels.
fn multi_string(distinct: bool) -> litchi_xls::writer::Writer {
    let labels: Vec<String> = (0..64).map(|index| format!("label {index:02} ✓")).collect();
    let mut writer = litchi_xls::writer::Writer::new();
    for sheet in 0..4usize {
        let worksheet = writer.add_worksheet(&format!("Strings{sheet}")).unwrap();
        for row in 0..2_500u32 {
            for column in 0..8u16 {
                if distinct {
                    writer
                        .write_string(worksheet, row, column, &format!("s{sheet} r{row} c{column} text"))
                        .unwrap();
                } else {
                    let index = (row as usize * 8 + column as usize + sheet) % labels.len();
                    writer.write_string(worksheet, row, column, &labels[index]).unwrap();
                }
            }
        }
    }
    writer
}

/// The harness's `write_fresh_doc` body.
fn doc(shape: &str) -> litchi_doc::writer::Writer {
    let paragraphs = match shape { "tiny" => 3, "large" => 512, _ => 128 };
    let payload = (shape == "payload-heavy").then_some(20_000);
    let mut writer = litchi_doc::writer::Writer::new();
    for paragraph in 0..paragraphs {
        let text = payload.map_or_else(
            || writer_text("doc", 0, paragraph, 0),
            |n| writer_payload_text("doc", 0, paragraph, 0, n),
        );
        writer.add_paragraph(&text).unwrap();
    }
    writer
}

/// The harness's `write_fresh_ppt` body.
fn ppt(shape: &str) -> litchi_ppt::writer::Writer {
    let (slides, boxes) = match shape { "tiny" => (1, 2), "large" => (12, 12), _ => (16, 8) };
    let payload = (shape == "payload-heavy").then_some(40_000);
    let mut writer = litchi_ppt::writer::Writer::new();
    for slide_number in 0..slides {
        let slide = writer.add_slide().unwrap();
        for box_number in 0..boxes {
            let text = payload.map_or_else(
                || writer_text("ppt", slide_number, box_number, 0),
                |n| writer_payload_text("ppt", slide_number, box_number, 0, n),
            );
            let x = 36 + (box_number % 3) as i32 * 180;
            let y = 36 + (box_number / 3) as i32 * 90;
            writer.add_textbox(slide, x, y, 144, 54, &text).unwrap();
        }
    }
    writer
}

enum Probe {
    Xls(litchi_xls::writer::Writer),
    Doc(litchi_doc::writer::Writer),
    Ppt(litchi_ppt::writer::Writer),
}

impl Probe {
    fn write(&mut self) -> Vec<u8> {
        let mut output = Cursor::new(Vec::new());
        match self {
            Self::Xls(writer) => writer.write_to(&mut output).unwrap(),
            Self::Doc(writer) => writer.write_to(&mut output).unwrap(),
            Self::Ppt(writer) => writer.write_to(&mut output).unwrap(),
        }
        output.into_inner()
    }
}

fn probe(case: &str) -> Probe {
    match case.split_once('/') {
        Some(("doc_fresh_write_to", shape)) => Probe::Doc(doc(shape)),
        Some(("ppt_fresh_write_to", shape)) => Probe::Ppt(ppt(shape)),
        _ => Probe::Xls(build(case)),
    }
}

fn build(case: &str) -> litchi_xls::writer::Writer {
    match case {
        "xls_fresh_write_to/tiny" => harness("tiny"),
        "xls_fresh_write_to/large" => harness("large"),
        "xls_fresh_write_to/payload-heavy" => harness("payload-heavy"),
        "xls_multi_string/distinct" => multi_string(true),
        "xls_multi_string/repeated" => multi_string(false),
        other => panic!("unknown case {other}"),
    }
}

fn fnv1a(bytes: &[u8]) -> u64 {
    bytes.iter().fold(0xcbf2_9ce4_8422_2325, |hash, &byte| {
        (hash ^ u64::from(byte)).wrapping_mul(0x0000_0100_0000_01b3)
    })
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let iterations: usize = args[1].parse().unwrap();
    let case = &args[2];
    // Only the write is timed and counted; the writer is built once, as the
    // harness does not, so these figures isolate `write_to`.
    let mut writer = probe(case);
    let mut write = || writer.write();
    let warm = write();
    let (length, hash) = (warm.len(), fnv1a(&warm));
    drop(warm);
    let (allocs0, reallocs0, bytes0) = (ALLOCS.load(Relaxed), REALLOCS.load(Relaxed), BYTES.load(Relaxed));
    let live0 = LIVE.load(Relaxed);
    PEAK.store(live0, Relaxed);
    let mut nanos = Vec::with_capacity(iterations);
    for _ in 0..iterations {
        let start = Instant::now();
        let output = write();
        nanos.push(start.elapsed().as_nanos() as u64);
        assert_eq!(output.len(), length);
        std::hint::black_box(&output);
    }
    let n = iterations.max(1) as u64;
    println!(
        "{{\"case\": \"{case}\", \"iterations\": {iterations}, \"output_bytes\": {length}, \"output_fnv1a\": \"{hash:016x}\", \"allocations_per_write\": {}, \"reallocations_per_write\": {}, \"allocated_bytes_per_write\": {}, \"peak_live_bytes\": {}, \"ns\": {:?}}}",
        (ALLOCS.load(Relaxed) - allocs0) / n,
        (REALLOCS.load(Relaxed) - reallocs0) / n,
        (BYTES.load(Relaxed) - bytes0) / n,
        PEAK.load(Relaxed) - live0,
        nanos
    );
}
