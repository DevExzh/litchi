//! Counting-allocator companion for the 0606 driver.
//!
//! `ppt0606_alloc <mode> <path>` reports allocation count, allocated bytes,
//! peak live bytes and live bytes still retained while the operation's result
//! is held, for one operation.

use std::alloc::{GlobalAlloc, Layout, System};
use std::hint::black_box;
use std::io::Cursor;
use std::path::Path;
use std::sync::atomic::{AtomicUsize, Ordering};

type BoxError = Box<dyn std::error::Error>;

static ALLOCATIONS: AtomicUsize = AtomicUsize::new(0);
static ALLOCATED_BYTES: AtomicUsize = AtomicUsize::new(0);
static LIVE_BYTES: AtomicUsize = AtomicUsize::new(0);
static PEAK_BYTES: AtomicUsize = AtomicUsize::new(0);

struct Counting;

impl Counting {
    fn record(size: usize) {
        ALLOCATIONS.fetch_add(1, Ordering::Relaxed);
        ALLOCATED_BYTES.fetch_add(size, Ordering::Relaxed);
        let live = LIVE_BYTES.fetch_add(size, Ordering::Relaxed) + size;
        PEAK_BYTES.fetch_max(live, Ordering::Relaxed);
    }
}

unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        Self::record(layout.size());
        unsafe { System.alloc(layout) }
    }

    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        LIVE_BYTES.fetch_sub(layout.size(), Ordering::Relaxed);
        unsafe { System.dealloc(ptr, layout) }
    }

    unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        ALLOCATIONS.fetch_add(1, Ordering::Relaxed);
        if new_size > layout.size() {
            ALLOCATED_BYTES.fetch_add(new_size - layout.size(), Ordering::Relaxed);
            let live = LIVE_BYTES.fetch_add(new_size - layout.size(), Ordering::Relaxed)
                + (new_size - layout.size());
            PEAK_BYTES.fetch_max(live, Ordering::Relaxed);
        } else {
            LIVE_BYTES.fetch_sub(layout.size() - new_size, Ordering::Relaxed);
        }
        unsafe { System.realloc(ptr, layout, new_size) }
    }
}

#[global_allocator]
static ALLOCATOR: Counting = Counting;

fn reset() {
    ALLOCATIONS.store(0, Ordering::Relaxed);
    ALLOCATED_BYTES.store(0, Ordering::Relaxed);
    PEAK_BYTES.store(LIVE_BYTES.load(Ordering::Relaxed), Ordering::Relaxed);
}

fn report(mode: &str, path: &Path, baseline: usize, retained: usize) {
    println!(
        "{}\t{}\tallocations={}\tallocated_bytes={}\tpeak_live_delta={}\tretained_live_delta={}",
        mode,
        path.display(),
        ALLOCATIONS.load(Ordering::Relaxed),
        ALLOCATED_BYTES.load(Ordering::Relaxed),
        PEAK_BYTES.load(Ordering::Relaxed).saturating_sub(baseline),
        retained.saturating_sub(baseline),
    );
}

fn main() -> Result<(), BoxError> {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let mode = args[0].clone();
    let path = std::path::PathBuf::from(&args[1]);
    let bytes = std::fs::read(&path)?;

    // One untimed warm-up so lazily-initialized globals are not attributed.
    {
        let mut package = litchi_ppt::Package::from_reader(Cursor::new(bytes.clone()))?;
        let presentation = package.presentation()?;
        black_box(&presentation);
    }

    match mode.as_str() {
        "eager-open" => {
            let owned = bytes.clone();
            let baseline = LIVE_BYTES.load(Ordering::Relaxed);
            reset();
            let mut package = litchi_ppt::Package::from_reader(Cursor::new(owned))?;
            let presentation = package.presentation()?;
            let retained = LIVE_BYTES.load(Ordering::Relaxed);
            report(&mode, &path, baseline, retained);
            black_box(&presentation);
        },
        "eager-text" => {
            let owned = bytes.clone();
            let baseline = LIVE_BYTES.load(Ordering::Relaxed);
            reset();
            let mut package = litchi_ppt::Package::from_reader(Cursor::new(owned))?;
            let presentation = package.presentation()?;
            let text = presentation.text()?;
            let retained = LIVE_BYTES.load(Ordering::Relaxed);
            report(&mode, &path, baseline, retained);
            black_box(&presentation);
            black_box(text);
        },
        "eager-slides" => {
            let owned = bytes.clone();
            let baseline = LIVE_BYTES.load(Ordering::Relaxed);
            reset();
            let mut package = litchi_ppt::Package::from_reader(Cursor::new(owned))?;
            let presentation = package.presentation()?;
            let slides = presentation.slides()?;
            let retained = LIVE_BYTES.load(Ordering::Relaxed);
            report(&mode, &path, baseline, retained);
            black_box(&slides);
        },
        "eager-notes" => {
            let owned = bytes.clone();
            let baseline = LIVE_BYTES.load(Ordering::Relaxed);
            reset();
            let mut package = litchi_ppt::Package::from_reader(Cursor::new(owned))?;
            let presentation = package.presentation()?;
            let slides = presentation.slides()?;
            let mut notes = Vec::new();
            for slide in &slides {
                notes.push(slide.speaker_notes()?.map(|value| value.text().map(str::to_string)));
            }
            let retained = LIVE_BYTES.load(Ordering::Relaxed);
            report(&mode, &path, baseline, retained);
            black_box(&slides);
            black_box(&notes);
        },
        "source-open" => {
            let baseline = LIVE_BYTES.load(Ordering::Relaxed);
            reset();
            let package = litchi_ppt::SourceBackedPackage::from_path(&path)?;
            let presentation = package.presentation()?;
            let retained = LIVE_BYTES.load(Ordering::Relaxed);
            report(&mode, &path, baseline, retained);
            black_box(&presentation);
        },
        "owned-edit-save" => {
            let owned = bytes.clone();
            let baseline = LIVE_BYTES.load(Ordering::Relaxed);
            reset();
            let snapshot = litchi_ppt::text_edit::Snapshot::from_bytes(owned)?;
            let commit = snapshot
                .edit_text(litchi_ppt::text_edit::Target::new(litchi_ppt::text_edit::Position::new(0), litchi_ppt::text_edit::Position::new(0)))?
                .commit()?;
            let retained = LIVE_BYTES.load(Ordering::Relaxed);
            report(&mode, &path, baseline, retained);
            black_box(&commit);
        },
        other => return Err(format!("unknown mode {other}").into()),
    }
    Ok(())
}
