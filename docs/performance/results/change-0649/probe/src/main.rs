//! Change 0649 probe: where the 133.61 ms opened-transaction shape-text edit
//! of change 0638's `pptx_real_file_ordinary_save_edit` row goes.
//!
//! The edit phase change 0638 times is exactly four documented public calls:
//!
//! ```text
//! package.opened_presentation_transaction()      // capture + working clone
//! edit.set_shape_text(slide, shape, MARKER)      // stage one slide
//! edit.commit()                                  // compact, fingerprint, patch, recapture
//! package.apply_opened_presentation_commit(c)    // candidate, validate, assign
//! ```
//!
//! so the phase boundary is a public API boundary and no crate source has to
//! move to decompose it.
//!
//! Subcommands:
//!   shape   <source>                    archive census of the deck under test
//!   phases  <source> <iterations>       per-call wall clock, one line per iteration
//!   prefix  <source> <stage> <iters>    `iters` x (open + every stage up to
//!                                       <stage>); for callgrind isolation pairs
//!   target  <source>                    the (slide, shape) position 0638 derives
//!
//! `<source>` is either a `.pptx` path or `generated:<slides>x<boxes>`, which
//! reproduces the shape of the harness's `build_semantic_pptx_corpus(Medium)`
//! (12 slides x 8 text boxes) that 0638's generated PPTX row measures.
//!
//! Stages, in prefix order: open, capture, transaction, settext, commit, apply.

use std::env;
use std::error::Error;
use std::time::Instant;

use litchi_pptx::Package;

/// A counting global allocator, so the per-phase allocation counts are
/// measured in the same process that measures the phase wall clock. This is a
/// retained scratch probe, not harness code: `tools/perf-baseline`'s source
/// policy forbids `#[global_allocator]` there, and change 0519's
/// `allocator-probe` is the precedent for measuring it here instead.
mod counting_allocator {
    use std::alloc::{GlobalAlloc, Layout, System};
    use std::sync::atomic::{AtomicU64, Ordering};

    pub static ALLOCATIONS: AtomicU64 = AtomicU64::new(0);
    pub static ALLOCATED_BYTES: AtomicU64 = AtomicU64::new(0);
    pub static REALLOCATIONS: AtomicU64 = AtomicU64::new(0);

    pub struct Counting;

    // SAFETY-free wrapper: every method forwards to `System` unchanged and only
    // adds relaxed counter arithmetic.
    unsafe impl GlobalAlloc for Counting {
        unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
            ALLOCATIONS.fetch_add(1, Ordering::Relaxed);
            ALLOCATED_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
            unsafe { System.alloc(layout) }
        }
        unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
            unsafe { System.dealloc(ptr, layout) }
        }
        unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
            REALLOCATIONS.fetch_add(1, Ordering::Relaxed);
            ALLOCATED_BYTES.fetch_add(new_size.saturating_sub(layout.size()) as u64, Ordering::Relaxed);
            unsafe { System.realloc(ptr, layout, new_size) }
        }
    }

    #[must_use]
    pub fn snapshot() -> [u64; 3] {
        [
            ALLOCATIONS.load(Ordering::Relaxed),
            ALLOCATED_BYTES.load(Ordering::Relaxed),
            REALLOCATIONS.load(Ordering::Relaxed),
        ]
    }
}

#[global_allocator]
static ALLOCATOR: counting_allocator::Counting = counting_allocator::Counting;

/// The marker change 0638's ordinary-save family writes.
const EDIT_MARKER: &str = "litchi-perf-0638-ordinary-save";

type Fallible<T> = Result<T, Box<dyn Error>>;

/// Materialize the deck under test as an in-memory archive, so that every
/// iteration opens identical bytes and no filesystem read is timed twice.
fn source_bytes(source: &str) -> Fallible<Vec<u8>> {
    if let Some(shape) = source.strip_prefix("generated:") {
        let (slides, boxes) = shape
            .split_once('x')
            .ok_or("generated shape must be <slides>x<boxes>")?;
        let slides: usize = slides.parse()?;
        let boxes: usize = boxes.parse()?;
        let mut package = Package::new()?;
        let presentation = package.presentation_mut()?;
        for slide_index in 0..slides {
            let slide = presentation.add_slide()?;
            for shape_index in 0..boxes {
                slide.add_text_box(
                    &format!("litchi-perf-baseline-pptx-semantic-v1-source-{slide_index:05}-{shape_index:05}"),
                    36 + i64::try_from(shape_index % 4)? * 180,
                    36 + i64::try_from(shape_index / 4)? * 90,
                    144,
                    54,
                );
            }
        }
        return Ok(package.to_bytes()?);
    }
    Ok(std::fs::read(source)?)
}

/// The first (slide, shape) position the documented transaction admits — the
/// same derivation change 0638's `pptx_edit_target` runs.
fn edit_target(archive: &[u8]) -> Fallible<(usize, usize)> {
    const MAX_SLIDES: usize = 64;
    const MAX_SHAPES: usize = 32;
    let package = Package::from_bytes(archive)?;
    let slide_count = package.opened_presentation()?.slides().len().min(MAX_SLIDES);
    for slide in 0..slide_count {
        for shape in 0..MAX_SHAPES {
            let mut edit = package.opened_presentation_transaction()?;
            if let Ok(true) = edit.set_shape_text(slide, shape, EDIT_MARKER) {
                return Ok((slide, shape));
            }
        }
    }
    Err("no admitted shape position".into())
}

/// Allocation counts for one open plus every stage up to and including each
/// stage, differenced the way the deterministic counters are.
fn allocations(source: &str) -> Fallible<()> {
    let archive = source_bytes(source)?;
    let (slide, shape) = edit_target(&archive)?;
    println!("source\t{source}");
    println!("stage\tallocations\tallocated_bytes\treallocations");
    for stage in ["open", "capture", "transaction", "settext", "commit", "apply"] {
        // One untimed warm-up so lazily initialized statics are not charged to
        // the first measured stage.
        prefix_once(&archive, stage, slide, shape)?;
        let before = counting_allocator::snapshot();
        prefix_once(&archive, stage, slide, shape)?;
        let after = counting_allocator::snapshot();
        println!(
            "{stage}\t{}\t{}\t{}",
            after[0] - before[0],
            after[1] - before[1],
            after[2] - before[2]
        );
    }
    Ok(())
}

/// One complete edit, timed per documented call. Returns nanoseconds for
/// capture, working-clone, set_shape_text, commit and apply.
fn timed_edit(archive: &[u8], slide: usize, shape: usize) -> Fallible<[u128; 5]> {
    let mut package = Package::from_bytes(archive)?;

    let started = Instant::now();
    let snapshot = package.opened_presentation()?;
    let capture = started.elapsed().as_nanos();

    let started = Instant::now();
    let mut edit = snapshot.edit();
    let clone = started.elapsed().as_nanos();

    let started = Instant::now();
    let changed = edit.set_shape_text(slide, shape, EDIT_MARKER)?;
    let settext = started.elapsed().as_nanos();
    if !changed {
        return Err("the derived shape position reported no change".into());
    }

    let started = Instant::now();
    let commit = edit.commit()?;
    let commit_ns = started.elapsed().as_nanos();
    if !commit.is_changed() {
        return Err("the commit reports no change".into());
    }

    let started = Instant::now();
    let published = package.apply_opened_presentation_commit(commit)?;
    let apply = started.elapsed().as_nanos();
    std::hint::black_box(&published);
    std::hint::black_box(&package);

    Ok([capture, clone, settext, commit_ns, apply])
}

fn percentile(sorted: &[u128], fraction: f64) -> u128 {
    if sorted.is_empty() {
        return 0;
    }
    let rank = (fraction * (sorted.len() - 1) as f64).round() as usize;
    sorted[rank.min(sorted.len() - 1)]
}

fn phases(source: &str, iterations: usize) -> Fallible<()> {
    let archive = source_bytes(source)?;
    let (slide, shape) = edit_target(&archive)?;
    let mut columns: [Vec<u128>; 5] = Default::default();
    for _ in 0..iterations {
        let sample = timed_edit(&archive, slide, shape)?;
        for (column, value) in columns.iter_mut().zip(sample) {
            column.push(value);
        }
    }
    let names = [
        "opened_presentation (capture)",
        "Snapshot::edit (working clone)",
        "set_shape_text",
        "Transaction::commit",
        "apply_opened_presentation_commit",
    ];
    println!("source\t{source}");
    println!("archive_bytes\t{}", archive.len());
    println!("target\tslide:{slide}/shape:{shape}");
    println!("iterations\t{iterations}");
    let mut totals = vec![0u128; iterations];
    for column in &columns {
        for (total, value) in totals.iter_mut().zip(column) {
            *total += value;
        }
    }
    println!("phase\tp50_ns\tmean_ns\tp95_ns\tmin_ns");
    for (name, column) in names.iter().zip(&columns) {
        let mut sorted = column.clone();
        sorted.sort_unstable();
        let mean = column.iter().sum::<u128>() / column.len() as u128;
        println!(
            "{name}\t{}\t{mean}\t{}\t{}",
            percentile(&sorted, 0.50),
            percentile(&sorted, 0.95),
            sorted[0]
        );
    }
    let mut sorted = totals.clone();
    sorted.sort_unstable();
    println!(
        "edit total\t{}\t{}\t{}\t{}",
        percentile(&sorted, 0.50),
        totals.iter().sum::<u128>() / totals.len() as u128,
        percentile(&sorted, 0.95),
        sorted[0]
    );
    Ok(())
}

fn prefix(source: &str, stage: &str, iterations: usize) -> Fallible<()> {
    let archive = source_bytes(source)?;
    let stage_rank = stage_rank(stage)?;
    // Deriving the edit target costs one transaction per probed position, so it
    // is skipped for the stages that never reach `set_shape_text`. Under
    // callgrind that derivation would otherwise dominate the profile.
    let (slide, shape) = if stage_rank >= 3 {
        edit_target(&archive)?
    } else {
        (0, 0)
    };
    for _ in 0..iterations {
        let mut package = Package::from_bytes(&archive)?;
        if stage_rank == 0 {
            std::hint::black_box(&package);
            continue;
        }
        let snapshot = package.opened_presentation()?;
        if stage_rank == 1 {
            std::hint::black_box(&snapshot);
            continue;
        }
        let mut edit = snapshot.edit();
        if stage_rank == 2 {
            std::hint::black_box(&edit);
            continue;
        }
        let changed = edit.set_shape_text(slide, shape, EDIT_MARKER)?;
        if !changed {
            return Err("the derived shape position reported no change".into());
        }
        if stage_rank == 3 {
            std::hint::black_box(&edit);
            continue;
        }
        let commit = edit.commit()?;
        if stage_rank == 4 {
            std::hint::black_box(&commit);
            continue;
        }
        let published = package.apply_opened_presentation_commit(commit)?;
        std::hint::black_box(&published);
    }
    println!("{stage}\t{iterations}\tdone");
    Ok(())
}

fn shape_census(source: &str) -> Fallible<()> {
    let archive = source_bytes(source)?;
    let package = Package::from_bytes(&archive)?;
    let snapshot = package.opened_presentation()?;
    println!("source\t{source}");
    println!("archive_bytes\t{}", archive.len());
    println!("opened_slides\t{}", snapshot.slides().len());
    for slide in snapshot.slides() {
        println!("slide\t{}\t{:?}", slide.id(), slide.name());
    }
    Ok(())
}

/// Measurement-only: per-stage deterministic counters from the instrumented
/// leg (`litchi_ooxml_common::perf0649`). Prints one block per stage, each the
/// count for ONE open + the stages up to and including it.
#[cfg(feature = "counts")]
fn counts(source: &str) -> Fallible<()> {
    use litchi_ooxml_common::perf0649;
    let archive = source_bytes(source)?;
    let (slide, shape) = edit_target(&archive)?;
    println!("source\t{source}");
    println!("target\tslide:{slide}/shape:{shape}");
    let stages = ["open", "capture", "transaction", "settext", "commit", "apply"];
    for stage in stages {
        perf0649::reset();
        let _ = perf0649::take_trace();
        prefix_once(&archive, stage, slide, shape)?;
        for (name, value) in perf0649::snapshot() {
            println!("{stage}\t{name}\t{value}");
        }
        let trace = perf0649::take_trace();
        println!("{stage}\tMCE_TRACE\t{}", trace.iter().map(u64::to_string).collect::<Vec<_>>().join(","));
    }
    Ok(())
}

/// One open plus every stage up to and including `stage`.
fn prefix_once(archive: &[u8], stage: &str, slide: usize, shape: usize) -> Fallible<()> {
    let rank = stage_rank(stage)?;
    let mut package = Package::from_bytes(archive)?;
    if rank == 0 {
        std::hint::black_box(&package);
        return Ok(());
    }
    let snapshot = package.opened_presentation()?;
    if rank == 1 {
        std::hint::black_box(&snapshot);
        return Ok(());
    }
    let mut edit = snapshot.edit();
    if rank == 2 {
        std::hint::black_box(&edit);
        return Ok(());
    }
    if !edit.set_shape_text(slide, shape, EDIT_MARKER)? {
        return Err("the derived shape position reported no change".into());
    }
    if rank == 3 {
        std::hint::black_box(&edit);
        return Ok(());
    }
    let commit = edit.commit()?;
    if rank == 4 {
        std::hint::black_box(&commit);
        return Ok(());
    }
    let published = package.apply_opened_presentation_commit(commit)?;
    std::hint::black_box(&published);
    Ok(())
}

fn stage_rank(stage: &str) -> Fallible<usize> {
    Ok(match stage {
        "open" => 0,
        "capture" => 1,
        "transaction" => 2,
        "settext" => 3,
        "commit" => 4,
        "apply" => 5,
        other => return Err(format!("unknown stage {other}").into()),
    })
}

/// Write the deck under test to a path, so a generated corpus can be inspected
/// with the same census the real fixture gets.
fn dump(source: &str, destination: &str) -> Fallible<()> {
    std::fs::write(destination, source_bytes(source)?)?;
    println!("wrote\t{destination}");
    Ok(())
}

fn main() -> Fallible<()> {
    let arguments: Vec<String> = env::args().skip(1).collect();
    match arguments.as_slice() {
        [mode, source] if mode == "shape" => shape_census(source),
        [mode, source] if mode == "target" => {
            let archive = source_bytes(source)?;
            let (slide, shape) = edit_target(&archive)?;
            println!("target\tslide:{slide}/shape:{shape}");
            Ok(())
        },
        [mode, source, iterations] if mode == "phases" => phases(source, iterations.parse()?),
        [mode, source, destination] if mode == "dump" => dump(source, destination),
        [mode, source] if mode == "allocations" => allocations(source),
        #[cfg(feature = "counts")]
        [mode, source] if mode == "counts" => counts(source),
        [mode, source, stage, iterations] if mode == "prefix" => {
            prefix(source, stage, iterations.parse()?)
        },
        _ => Err("usage: probe0649 shape|target <source> | phases <source> <n> | prefix <source> <stage> <n>".into()),
    }
}
