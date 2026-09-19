//! Counting allocator companion for the 0686 XLS retry probe.

use std::alloc::{GlobalAlloc, Layout, System};
use std::error::Error;
use std::hint::black_box;
use std::sync::Arc;
use std::sync::atomic::{AtomicUsize, Ordering};

use litchi_core::{FileSource, OwnedSource, ReadAt};
use litchi_xls::{SourceBackedLimits, SourceBackedWorkbook};

static ALLOCATION_CALLS: AtomicUsize = AtomicUsize::new(0);
static ALLOCATED_BYTES: AtomicUsize = AtomicUsize::new(0);
static DEALLOCATION_CALLS: AtomicUsize = AtomicUsize::new(0);
static DEALLOCATED_BYTES: AtomicUsize = AtomicUsize::new(0);
static LIVE_BYTES: AtomicUsize = AtomicUsize::new(0);
static PEAK_LIVE_BYTES: AtomicUsize = AtomicUsize::new(0);

struct Counting;

impl Counting {
    fn allocation(size: usize) {
        ALLOCATION_CALLS.fetch_add(1, Ordering::Relaxed);
        ALLOCATED_BYTES.fetch_add(size, Ordering::Relaxed);
        let live = LIVE_BYTES.fetch_add(size, Ordering::Relaxed) + size;
        PEAK_LIVE_BYTES.fetch_max(live, Ordering::Relaxed);
    }

    fn deallocation(size: usize) {
        DEALLOCATION_CALLS.fetch_add(1, Ordering::Relaxed);
        DEALLOCATED_BYTES.fetch_add(size, Ordering::Relaxed);
        LIVE_BYTES.fetch_sub(size, Ordering::Relaxed);
    }

    fn reset() {
        ALLOCATION_CALLS.store(0, Ordering::Relaxed);
        ALLOCATED_BYTES.store(0, Ordering::Relaxed);
        DEALLOCATION_CALLS.store(0, Ordering::Relaxed);
        DEALLOCATED_BYTES.store(0, Ordering::Relaxed);
        PEAK_LIVE_BYTES.store(LIVE_BYTES.load(Ordering::Relaxed), Ordering::Relaxed);
    }
}

unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        // Count only the request that the system allocator accepted. A null
        // allocation leaves live and peak accounting unchanged.
        let pointer = unsafe { System.alloc(layout) };
        if !pointer.is_null() {
            Counting::allocation(layout.size());
        }
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        Counting::deallocation(layout.size());
        unsafe { System.dealloc(pointer, layout) };
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        let result = unsafe { System.realloc(pointer, layout, new_size) };
        if !result.is_null() {
            // `allocated_bytes` is requested bytes: each successful realloc
            // contributes its new requested size, while logical live bytes
            // changes only by the size delta.
            ALLOCATION_CALLS.fetch_add(1, Ordering::Relaxed);
            ALLOCATED_BYTES.fetch_add(new_size, Ordering::Relaxed);
            if new_size >= layout.size() {
                let delta = new_size - layout.size();
                let live = LIVE_BYTES.fetch_add(delta, Ordering::Relaxed) + delta;
                PEAK_LIVE_BYTES.fetch_max(live, Ordering::Relaxed);
            } else {
                LIVE_BYTES.fetch_sub(layout.size() - new_size, Ordering::Relaxed);
            }
        }
        result
    }
}

#[global_allocator]
static ALLOCATOR: Counting = Counting;

fn signed_delta(after: usize, before: usize) -> i128 {
    i128::try_from(after).unwrap_or(i128::MAX) - i128::try_from(before).unwrap_or(i128::MAX)
}

fn main() -> Result<(), Box<dyn Error>> {
    let args: Vec<String> = std::env::args().skip(1).collect();
    if args.len() != 7 {
        return Err(
            "usage: xls0686-alloc MODE OP FILE SHEET ROW COLUMN BUDGET; MODE=owned|file OP=q1|q2|q3|q8|visit"
                .into(),
        );
    }
    let mode = &args[0];
    let op = &args[1];
    let sheet: usize = args[3].parse()?;
    let row: u32 = args[4].parse()?;
    let column: u32 = args[5].parse()?;
    let source: Arc<dyn ReadAt> = match mode.as_str() {
        "owned" => Arc::new(OwnedSource::new(std::fs::read(&args[2])?)),
        "file" => Arc::new(FileSource::open(&args[2])?),
        _ => return Err("invalid mode".into()),
    };
    let budget: u64 = args[6].parse()?;
    let limits = SourceBackedLimits::default().with_max_query_index_bytes(budget);
    let owner = SourceBackedWorkbook::from_read_at_with_limits(source, limits)?;
    let warm = match op.as_str() {
        "q1" | "visit" => 0,
        "q2" => 1,
        "q3" => 2,
        "q8" => 7,
        _ => return Err("invalid operation".into()),
    };
    for _ in 0..warm {
        let _ = black_box(owner.cell_value_by_index(sheet, row, column));
    }
    Counting::reset();
    let baseline_live = LIVE_BYTES.load(Ordering::Relaxed);
    let mut callbacks = 0u64;
    // All owners and returned values/errors remain alive until after gauge reads.
    let query = if op != "visit" {
        Some(owner.cell_value_by_index(sheet, row, column))
    } else {
        None
    };
    let visit = if op == "visit" {
        Some(owner.worksheet_by_index(sheet).and_then(|sheet| {
            sheet
                .ok_or_else(|| litchi_xls::SourceBackedError::WorksheetNotFound("probe".into()))?
                .visit_cells(|cell| {
                    callbacks += 1;
                    black_box(cell);
                    Ok(())
                })
        }))
    } else {
        None
    };
    let live_after = LIVE_BYTES.load(Ordering::Relaxed);
    let peak = PEAK_LIVE_BYTES.load(Ordering::Relaxed);
    let calls = ALLOCATION_CALLS.load(Ordering::Relaxed);
    let bytes = ALLOCATED_BYTES.load(Ordering::Relaxed);
    let frees = DEALLOCATION_CALLS.load(Ordering::Relaxed);
    let freed_bytes = DEALLOCATED_BYTES.load(Ordering::Relaxed);
    let outcome = match &query {
        Some(Ok(Some(value))) => format!("value:{value:?}"),
        Some(Ok(None)) => "missing".to_owned(),
        Some(Err(error)) => format!("error:{error}"),
        None => match &visit {
            Some(Ok(())) => "ok".to_owned(),
            Some(Err(error)) => format!("error:{error}"),
            None => unreachable!(),
        },
    };
    println!(
        "{}",
        serde_json::json!({
            "mode": mode, "operation": op, "budget": budget, "worksheet": sheet, "row": row, "column": column,
            "outcome": outcome, "callbacks": callbacks, "allocation_calls": calls,
            "allocated_bytes": bytes, "deallocation_calls": frees, "deallocated_bytes": freed_bytes,
            "peak_live_delta": signed_delta(peak, baseline_live),
            "retained_live_delta": signed_delta(live_after, baseline_live),
        })
    );
    black_box((&owner, &query, &visit));
    Ok(())
}
