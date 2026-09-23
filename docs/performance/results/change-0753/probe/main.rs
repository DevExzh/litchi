//! Allocation probe for change 0753: runs the harness's `write_fresh_{doc,ppt,xls}`
//! bodies under a counting global allocator. Scratch tool, not production code.
use std::alloc::{GlobalAlloc, Layout, System};
use std::io::Cursor;
use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering::Relaxed};

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

fn doc(shape: &str) -> Vec<u8> {
    let paragraphs = match shape { "tiny" => 3, "large" => 512, _ => 128 };
    let payload = (shape == "payload-heavy").then_some(20_000);
    let mut writer = litchi_doc::writer::Writer::new();
    for paragraph in 0..paragraphs {
        let text = payload.map_or_else(|| writer_text("doc", 0, paragraph, 0), |n| writer_payload_text("doc", 0, paragraph, 0, n));
        writer.add_paragraph(&text).unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn ppt(shape: &str) -> Vec<u8> {
    let (slides, boxes) = match shape { "tiny" => (1, 2), "large" => (12, 12), _ => (16, 8) };
    let payload = (shape == "payload-heavy").then_some(40_000);
    let mut writer = litchi_ppt::writer::Writer::new();
    for slide_number in 0..slides {
        let slide = writer.add_slide().unwrap();
        for box_number in 0..boxes {
            let text = payload.map_or_else(|| writer_text("ppt", slide_number, box_number, 0), |n| writer_payload_text("ppt", slide_number, box_number, 0, n));
            let x = 36 + (box_number % 3) as i32 * 180;
            let y = 36 + (box_number / 3) as i32 * 90;
            writer.add_textbox(slide, x, y, 144, 54, &text).unwrap();
        }
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn xls(shape: &str) -> Vec<u8> {
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
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let iterations: usize = args.get(1).map_or(1, |value| value.parse().unwrap());
    let only = args.get(2).cloned();
    println!("[");
    let mut first = true;
    for kind in ["doc", "ppt", "xls"] {
        for shape in ["tiny", "large", "payload-heavy"] {
            let name = format!("{kind}_fresh_write_to/{shape}");
            if only.as_ref().is_some_and(|value| value != &name) {
                continue;
            }
            let run = || match kind { "doc" => doc(shape), "ppt" => ppt(shape), _ => xls(shape) };
            let warm = run();
            let digest = warm.len();
            drop(warm);
            let (allocs0, reallocs0, bytes0) = (ALLOCS.load(Relaxed), REALLOCS.load(Relaxed), BYTES.load(Relaxed));
            let live0 = LIVE.load(Relaxed);
            PEAK.store(live0, Relaxed);
            for _ in 0..iterations {
                let output = run();
                assert_eq!(output.len(), digest);
                std::hint::black_box(&output);
            }
            let allocs = (ALLOCS.load(Relaxed) - allocs0) / iterations as u64;
            let reallocs = (REALLOCS.load(Relaxed) - reallocs0) / iterations as u64;
            let bytes = (BYTES.load(Relaxed) - bytes0) / iterations as u64;
            let peak = PEAK.load(Relaxed) - live0;
            if !first { println!(","); }
            first = false;
            print!("  {{\"case\": \"{name}\", \"allocations\": {allocs}, \"reallocations\": {reallocs}, \"allocated_bytes\": {bytes}, \"peak_live_bytes\": {peak}, \"output_bytes\": {digest}}}");
        }
    }
    println!("\n]");
}
