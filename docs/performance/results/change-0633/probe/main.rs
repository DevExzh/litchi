//! Attribution probe for the litchi-xls edit-and-save path (change 0620).
//!
//! This is scratch measurement code, not production code and not part of the
//! registered harness. It drives the public `litchi_xls::cell_values` editor
//! over a real XLS fixture so that one open, one staged edit and one commit can
//! be timed, counted and profiled separately on each of the three publication
//! paths (`commit`, `commit_source_backed`, `commit_source_backed_plan`).
//!
//! Output is one JSON object per run on stdout.

use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicU64, Ordering};
use std::time::Instant;

use litchi_xls::cell_values::{Reference, Selector, Snapshot, Storage, Value};

// ---------------------------------------------------------------------------
// Counting allocator
// ---------------------------------------------------------------------------

static ALLOCS: AtomicU64 = AtomicU64::new(0);
static ALLOC_BYTES: AtomicU64 = AtomicU64::new(0);
static LIVE_BYTES: AtomicU64 = AtomicU64::new(0);
static PEAK_BYTES: AtomicU64 = AtomicU64::new(0);
static COUNTING: AtomicU64 = AtomicU64::new(0);

struct Counting;

#[inline]
fn on(delta: i64, size: u64, counted: bool) {
    if !counted {
        return;
    }
    if delta > 0 {
        ALLOCS.fetch_add(1, Ordering::Relaxed);
        ALLOC_BYTES.fetch_add(size, Ordering::Relaxed);
        let live = LIVE_BYTES.fetch_add(size, Ordering::Relaxed) + size;
        PEAK_BYTES.fetch_max(live, Ordering::Relaxed);
    } else {
        LIVE_BYTES.fetch_sub(size, Ordering::Relaxed);
    }
}

unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let counted = COUNTING.load(Ordering::Relaxed) != 0;
        let pointer = unsafe { System.alloc(layout) };
        if !pointer.is_null() {
            on(1, layout.size() as u64, counted);
        }
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        let counted = COUNTING.load(Ordering::Relaxed) != 0;
        on(-1, layout.size() as u64, counted);
        unsafe { System.dealloc(pointer, layout) }
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        let counted = COUNTING.load(Ordering::Relaxed) != 0;
        let new_pointer = unsafe { System.realloc(pointer, layout, new_size) };
        if !new_pointer.is_null() {
            on(-1, layout.size() as u64, counted);
            on(1, new_size as u64, counted);
        }
        new_pointer
    }
}

#[global_allocator]
static GLOBAL: Counting = Counting;

fn reset_counters() {
    ALLOCS.store(0, Ordering::Relaxed);
    ALLOC_BYTES.store(0, Ordering::Relaxed);
    LIVE_BYTES.store(0, Ordering::Relaxed);
    PEAK_BYTES.store(0, Ordering::Relaxed);
}

// ---------------------------------------------------------------------------
// A sink that counts published bytes without retaining them
// ---------------------------------------------------------------------------

struct CountingSink {
    bytes: u64,
    writes: u64,
}

impl std::io::Write for CountingSink {
    fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
        self.bytes += buffer.len() as u64;
        self.writes += 1;
        Ok(buffer.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

// ---------------------------------------------------------------------------
// Target discovery
// ---------------------------------------------------------------------------

#[derive(Clone, Debug)]
struct NumericTarget {
    sheet: String,
    row: u32,
    column: u32,
    storage: Storage,
    before: f64,
}

#[derive(Clone, Debug)]
struct TextTarget {
    sheet: String,
    row: u32,
    column: u32,
    before: String,
    after: String,
}

fn find_numeric(snapshot: &Snapshot) -> Option<NumericTarget> {
    for worksheet in snapshot.worksheets() {
        let name = worksheet.name().to_string();
        for cell in worksheet.cells() {
            if !matches!(
                cell.storage(),
                Storage::Number | Storage::Rk | Storage::MulRk
            ) {
                continue;
            }
            let Value::Number(before) = cell.value() else {
                continue;
            };
            if !before.is_finite() {
                continue;
            }
            return Some(NumericTarget {
                sheet: name,
                row: u32::from(cell.reference().row()),
                column: u32::from(cell.reference().column()),
                storage: cell.storage(),
                before: *before,
            });
        }
    }
    None
}

/// Finds a `LabelSst` cell plus a *different* text that another `LabelSst`
/// cell already uses, so the replacement resolves through the existing SST and
/// stages no resource change.
fn find_text(snapshot: &Snapshot) -> Option<TextTarget> {
    let mut first: Option<(String, u32, u32, String)> = None;
    for worksheet in snapshot.worksheets() {
        let name = worksheet.name().to_string();
        for cell in worksheet.cells() {
            if cell.storage() != Storage::LabelSst {
                continue;
            }
            let Value::Text(text) = cell.value() else {
                continue;
            };
            match &first {
                None => {
                    first = Some((
                        name.clone(),
                        u32::from(cell.reference().row()),
                        u32::from(cell.reference().column()),
                        text.clone(),
                    ));
                },
                Some((sheet, row, column, before)) if before != text => {
                    return Some(TextTarget {
                        sheet: sheet.clone(),
                        row: *row,
                        column: *column,
                        before: before.clone(),
                        after: text.clone(),
                    });
                },
                Some(_) => {},
            }
        }
    }
    None
}

// ---------------------------------------------------------------------------
// Operations
// ---------------------------------------------------------------------------

#[derive(Clone, Copy, PartialEq, Eq)]
enum Operation {
    Inventory,
    Open,
    NumberPlan,
    NumberSourceBacked,
    NumberGeneric,
    StringGeneric,
    NoopGeneric,
}

impl Operation {
    fn parse(text: &str) -> Option<Self> {
        Some(match text {
            "inventory" => Self::Inventory,
            "open" => Self::Open,
            "number-plan" => Self::NumberPlan,
            "number-source-backed" => Self::NumberSourceBacked,
            "number-generic" => Self::NumberGeneric,
            "string-generic" => Self::StringGeneric,
            "noop-generic" => Self::NoopGeneric,
            _ => return None,
        })
    }

    fn name(self) -> &'static str {
        match self {
            Self::Inventory => "inventory",
            Self::Open => "open",
            Self::NumberPlan => "number-plan",
            Self::NumberSourceBacked => "number-source-backed",
            Self::NumberGeneric => "number-generic",
            Self::StringGeneric => "string-generic",
            Self::NoopGeneric => "noop-generic",
        }
    }
}

struct Sample {
    open_ns: u128,
    stage_ns: u128,
    commit_ns: u128,
    publish_ns: u128,
    published_bytes: u64,
}

fn render_diagnostics(diagnostics: litchi_xls::cell_values::SourceBackedDiagnostics) -> String {
    format!(
        "{{\"changed_cells\":{},\"touched_streams\":{},\"splice_count\":{},\"replacement_bytes\":{},\"changed_spans\":{},\"source_bytes\":{},\"source_workbook_bytes\":{},\"target_workbook_bytes\":{}}}",
        diagnostics.changed_cells(),
        diagnostics.touched_streams(),
        diagnostics.splice_count(),
        diagnostics.replacement_bytes(),
        diagnostics.changed_spans(),
        diagnostics.source_bytes(),
        diagnostics.source_workbook_bytes(),
        diagnostics.target_workbook_bytes()
    )
}

fn escape(text: &str) -> String {
    let mut out = String::new();
    for character in text.chars() {
        match character {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            c if (c as u32) < 0x20 => out.push_str(&format!("\\u{:04x}", c as u32)),
            c => out.push(c),
        }
    }
    out
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut input = String::new();
    let mut operation = Operation::Open;
    let mut warmups = 3_usize;
    let mut samples = 20_usize;
    let mut arguments = std::env::args().skip(1);
    while let Some(argument) = arguments.next() {
        match argument.as_str() {
            "--input" => input = arguments.next().ok_or("--input needs a path")?,
            "--operation" => {
                let value = arguments.next().ok_or("--operation needs a name")?;
                operation = Operation::parse(&value).ok_or("unknown --operation")?;
            },
            "--warmups" => warmups = arguments.next().ok_or("--warmups")?.parse()?,
            "--samples" => samples = arguments.next().ok_or("--samples")?.parse()?,
            other => return Err(format!("unknown argument {other}").into()),
        }
    }
    if input.is_empty() {
        return Err("--input is required".into());
    }
    let bytes = std::fs::read(&input)?;
    let source_bytes = bytes.len() as u64;

    if operation == Operation::Inventory {
        let snapshot = Snapshot::from_bytes(bytes.clone())?;
        let numeric = find_numeric(&snapshot);
        let text = find_text(&snapshot);
        let mut sheets = Vec::new();
        for worksheet in snapshot.worksheets() {
            let mut number = 0_u64;
            let mut rk = 0_u64;
            let mut mulrk = 0_u64;
            let mut label = 0_u64;
            let mut formula = 0_u64;
            let mut other = 0_u64;
            for cell in worksheet.cells() {
                match cell.storage() {
                    Storage::Number => number += 1,
                    Storage::Rk => rk += 1,
                    Storage::MulRk => mulrk += 1,
                    Storage::LabelSst => label += 1,
                    Storage::Formula => formula += 1,
                    _ => other += 1,
                }
            }
            sheets.push(format!(
                "{{\"name\":\"{}\",\"number\":{number},\"rk\":{rk},\"mulrk\":{mulrk},\"labelsst\":{label},\"formula\":{formula},\"other\":{other}}}",
                escape(worksheet.name())
            ));
        }
        println!(
            "{{\"input\":\"{}\",\"source_bytes\":{source_bytes},\"workbook_stream_bytes\":{},\"worksheets\":{},\"sheets\":[{}],\"numeric_target\":{},\"text_target\":{}}}",
            escape(&input),
            snapshot.workbook_stream().len(),
            snapshot.worksheet_count(),
            sheets.join(","),
            match &numeric {
                Some(target) => format!(
                    "{{\"sheet\":\"{}\",\"row\":{},\"column\":{},\"storage\":\"{:?}\",\"before\":{}}}",
                    escape(&target.sheet),
                    target.row,
                    target.column,
                    target.storage,
                    target.before
                ),
                None => "null".to_string(),
            },
            match &text {
                Some(target) => format!(
                    "{{\"sheet\":\"{}\",\"row\":{},\"column\":{},\"before\":\"{}\",\"after\":\"{}\"}}",
                    escape(&target.sheet),
                    target.row,
                    target.column,
                    escape(&target.before),
                    escape(&target.after)
                ),
                None => "null".to_string(),
            }
        );
        return Ok(());
    }

    // Discover the targets once, outside every measured region.
    let probe = Snapshot::from_bytes(bytes.clone())?;
    let numeric = find_numeric(&probe);
    let text = find_text(&probe);
    let workbook_stream_bytes = probe.workbook_stream().len() as u64;
    drop(probe);

    let mut collected: Vec<Sample> = Vec::with_capacity(samples);
    let mut digest: Option<u64> = None;
    let total_iterations = warmups + samples;
    reset_counters();
    let mut counted_allocs = 0_u64;
    let mut counted_alloc_bytes = 0_u64;
    let mut counted_peak = 0_u64;
    let mut open_allocs = 0_u64;
    let mut open_alloc_bytes = 0_u64;
    let mut open_peak = 0_u64;
    let mut diagnostics_json = String::from("null");

    for iteration in 0..total_iterations {
        let measured = iteration >= warmups;
        // Every iteration works on its own copy so no state is carried over.
        let owned = bytes.clone();
        if measured && iteration == warmups {
            reset_counters();
            COUNTING.store(1, Ordering::Relaxed);
        }
        let open_started = Instant::now();
        let snapshot = Snapshot::from_bytes(owned)?;
        let open_ns = open_started.elapsed().as_nanos();
        if measured && iteration == warmups {
            open_allocs = ALLOCS.load(Ordering::Relaxed);
            open_alloc_bytes = ALLOC_BYTES.load(Ordering::Relaxed);
            open_peak = PEAK_BYTES.load(Ordering::Relaxed);
        }

        if operation == Operation::Open {
            if measured && iteration == warmups {
                COUNTING.store(0, Ordering::Relaxed);
                counted_allocs = ALLOCS.load(Ordering::Relaxed);
                counted_alloc_bytes = ALLOC_BYTES.load(Ordering::Relaxed);
                counted_peak = PEAK_BYTES.load(Ordering::Relaxed);
            }
            if measured {
                collected.push(Sample {
                    open_ns,
                    stage_ns: 0,
                    commit_ns: 0,
                    publish_ns: 0,
                    published_bytes: 0,
                });
            }
            continue;
        }

        let mut sink = CountingSink {
            bytes: 0,
            writes: 0,
        };
        let stage_started = Instant::now();
        let mut transaction = snapshot.edit();
        match operation {
            Operation::NumberPlan | Operation::NumberSourceBacked | Operation::NumberGeneric => {
                let target = numeric.as_ref().ok_or("fixture has no editable numeric cell")?;
                transaction.set_numeric(
                    Selector::Name(&target.sheet),
                    Reference::new(target.row, target.column)?,
                    target.before + 1.0,
                )?;
            },
            Operation::StringGeneric => {
                let target = text.as_ref().ok_or("fixture has no two distinct SST cells")?;
                transaction.set_value(
                    Selector::Name(&target.sheet),
                    Reference::new(target.row, target.column)?,
                    Value::Text(target.after.clone()),
                )?;
            },
            Operation::NoopGeneric => {
                let target = numeric.as_ref().ok_or("fixture has no editable numeric cell")?;
                transaction.set_numeric(
                    Selector::Name(&target.sheet),
                    Reference::new(target.row, target.column)?,
                    target.before,
                )?;
            },
            Operation::Inventory | Operation::Open => unreachable!(),
        }
        let stage_ns = stage_started.elapsed().as_nanos();

        let commit_started = Instant::now();
        let published: u64 = match operation {
            Operation::NumberPlan => {
                let commit = transaction.commit_source_backed_plan()?;
                let commit_ns = commit_started.elapsed().as_nanos();
                if measured && iteration == warmups {
                    diagnostics_json = render_diagnostics(commit.diagnostics());
                }
                let publish_started = Instant::now();
                commit.write_to(&mut sink)?;
                let publish_ns = publish_started.elapsed().as_nanos();
                if measured {
                    collected.push(Sample {
                        open_ns,
                        stage_ns,
                        commit_ns,
                        publish_ns,
                        published_bytes: sink.bytes,
                    });
                }
                sink.bytes
            },
            Operation::NumberSourceBacked => {
                let commit = transaction.commit_source_backed()?;
                let commit_ns = commit_started.elapsed().as_nanos();
                if measured && iteration == warmups {
                    diagnostics_json = render_diagnostics(commit.diagnostics());
                }
                let publish_started = Instant::now();
                commit.write_to(&mut sink)?;
                let publish_ns = publish_started.elapsed().as_nanos();
                if measured {
                    collected.push(Sample {
                        open_ns,
                        stage_ns,
                        commit_ns,
                        publish_ns,
                        published_bytes: sink.bytes,
                    });
                }
                sink.bytes
            },
            Operation::NumberGeneric | Operation::StringGeneric | Operation::NoopGeneric => {
                let commit = transaction.commit()?;
                let commit_ns = commit_started.elapsed().as_nanos();
                let publish_started = Instant::now();
                let (target, _patch, _diagnostics) = commit.into_parts();
                use std::io::Write as _;
                for chunk in target.bytes().chunks(64 * 1024) {
                    sink.write_all(chunk)?;
                }
                sink.flush()?;
                let publish_ns = publish_started.elapsed().as_nanos();
                if measured {
                    collected.push(Sample {
                        open_ns,
                        stage_ns,
                        commit_ns,
                        publish_ns,
                        published_bytes: sink.bytes,
                    });
                }
                sink.bytes
            },
            Operation::Inventory | Operation::Open => unreachable!(),
        };
        if measured && iteration == warmups {
            COUNTING.store(0, Ordering::Relaxed);
            counted_allocs = ALLOCS.load(Ordering::Relaxed);
            counted_alloc_bytes = ALLOC_BYTES.load(Ordering::Relaxed);
            counted_peak = PEAK_BYTES.load(Ordering::Relaxed);
        }
        match digest {
            None => digest = Some(published),
            Some(previous) if previous == published => {},
            Some(_) => return Err("published byte count changed between samples".into()),
        }
    }

    let open: Vec<u128> = collected.iter().map(|sample| sample.open_ns).collect();
    let stage: Vec<u128> = collected.iter().map(|sample| sample.stage_ns).collect();
    let commit: Vec<u128> = collected.iter().map(|sample| sample.commit_ns).collect();
    let publish: Vec<u128> = collected.iter().map(|sample| sample.publish_ns).collect();
    let render = |values: &[u128]| {
        values
            .iter()
            .map(u128::to_string)
            .collect::<Vec<_>>()
            .join(",")
    };
    println!(
        "{{\"input\":\"{}\",\"operation\":\"{}\",\"source_bytes\":{source_bytes},\"workbook_stream_bytes\":{workbook_stream_bytes},\"samples\":{},\"published_bytes\":{},\"allocations\":{counted_allocs},\"allocated_bytes\":{counted_alloc_bytes},\"peak_live_bytes\":{counted_peak},\"open_allocations\":{open_allocs},\"open_allocated_bytes\":{open_alloc_bytes},\"open_peak_live_bytes\":{open_peak},\"diagnostics\":{diagnostics_json},\"open_ns\":[{}],\"stage_ns\":[{}],\"commit_ns\":[{}],\"publish_ns\":[{}]}}",
        escape(&input),
        operation.name(),
        collected.len(),
        digest.unwrap_or(0),
        render(&open),
        render(&stage),
        render(&commit),
        render(&publish)
    );
    Ok(())
}
