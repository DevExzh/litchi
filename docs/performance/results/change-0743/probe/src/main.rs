//! Probe for record 0743: loops one timed region of the semantic PPTX cases
//! over the harness's own deterministic corpus, so that a profile of this
//! process is dominated by the timed work rather than corpus construction.
#![allow(clippy::all)]

use std::time::Instant;

#[cfg(feature = "count-alloc")]
mod counting {
    //! Probe-only counting allocator: calls and requested bytes on this
    //! process, read around one region. Not used by any production code.
    use std::alloc::{GlobalAlloc, Layout, System};
    use std::sync::atomic::{AtomicU64, Ordering};

    pub static CALLS: AtomicU64 = AtomicU64::new(0);
    pub static BYTES: AtomicU64 = AtomicU64::new(0);

    pub struct Counting;

    unsafe impl GlobalAlloc for Counting {
        unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
            CALLS.fetch_add(1, Ordering::Relaxed);
            BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
            unsafe { System.alloc(layout) }
        }

        unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
            unsafe { System.dealloc(ptr, layout) }
        }

        unsafe fn realloc(&self, ptr: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
            CALLS.fetch_add(1, Ordering::Relaxed);
            BYTES.fetch_add(new_size as u64, Ordering::Relaxed);
            unsafe { System.realloc(ptr, layout, new_size) }
        }
    }

    #[global_allocator]
    static GLOBAL: Counting = Counting;

    pub fn read() -> (u64, u64) {
        (CALLS.load(Ordering::Relaxed), BYTES.load(Ordering::Relaxed))
    }
}

#[cfg(feature = "count-alloc")]
fn allocations() -> (u64, u64) {
    counting::read()
}

#[cfg(not(feature = "count-alloc"))]
fn allocations() -> (u64, u64) {
    (0, 0)
}

fn semantic_pptx_text(slide: usize, shape: usize, updated: bool) -> String {
    let state = if updated { "updated" } else { "source" };
    format!("litchi-perf-baseline-pptx-semantic-v1-{state}-{slide:03}-{shape:03}")
}

fn shape_dims(name: &str) -> (usize, usize) {
    match name {
        "tiny" => (3, 4),
        "medium" => (12, 8),
        "large" => (100, 100),
        _ => panic!("unknown shape"),
    }
}

fn build(slides: usize, boxes: usize) -> Vec<u8> {
    let mut package = litchi_pptx::Package::new().unwrap();
    let presentation = package.presentation_mut().unwrap();
    for slide_index in 0..slides {
        let slide = presentation.add_slide().unwrap();
        for shape_index in 0..boxes {
            slide.add_text_box(
                &semantic_pptx_text(slide_index, shape_index, false),
                36 + i64::try_from(shape_index % 4).unwrap() * 180,
                36 + i64::try_from(shape_index / 4).unwrap() * 90,
                144,
                54,
            );
        }
    }
    package.to_bytes().unwrap()
}

fn update_indices(count: usize) -> Vec<usize> {
    let updates = (count + 99) / 100;
    (0..updates).map(|index| index * count / updates).collect()
}

fn median(mut values: Vec<u128>) -> u128 {
    values.sort_unstable();
    values[values.len() / 2]
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let mode = args.get(1).map(String::as_str).unwrap_or("stats");
    let shape = args.get(2).map(String::as_str).unwrap_or("large");
    let iters: usize = args.get(3).map(|v| v.parse().unwrap()).unwrap_or(10);
    let edits = args.get(4).map(String::as_str).unwrap_or("one");
    let (slides, boxes) = shape_dims(shape);
    let bytes = build(slides, boxes);
    let total = slides * boxes;
    let edits = if mode.starts_with("dump") { "noop" } else { edits };
    let selected: Vec<usize> = match edits {
        "noop" => Vec::new(),
        "one" => vec![update_indices(total)[0]],
        "pct" => update_indices(total),
        other => {
            let n: usize = other.parse().unwrap();
            update_indices(total).into_iter().take(n).collect()
        },
    };
    match mode {
        "stats" => {
            let package = litchi_pptx::Package::from_bytes(&bytes).unwrap();
            let snapshot = package.opened_presentation().unwrap();
            println!("archive bytes {}", bytes.len());
            let mut sizes = Vec::new();
            for slide in snapshot.slides() {
                let _ = slide;
            }
            let presentation = package.presentation().unwrap();
            for slide in presentation.slides().unwrap() {
                sizes.push(slide.part().part().blob().len());
            }
            let sum: usize = sizes.iter().sum();
            println!(
                "slides {} xml total {} min {} max {} avg {}",
                sizes.len(),
                sum,
                sizes.iter().min().unwrap(),
                sizes.iter().max().unwrap(),
                sum / sizes.len()
            );
            let first = presentation.slides().unwrap()[0].part().part().blob().to_vec();
            println!("{}", String::from_utf8_lossy(&first[..first.len().min(3000)]));
        },
        "dump-slide" => {
            let package = litchi_pptx::Package::from_bytes(&bytes).unwrap();
            let presentation = package.presentation().unwrap();
            let first = presentation.slides().unwrap()[0].part().part().blob().to_vec();
            std::fs::write(args.get(4).unwrap(), &first).unwrap();
        },
        "setup" => {
            // The per-iteration setup of the edit/save cycle, outside its
            // clocks: a fresh owned package from the archive bytes. Counter
            // differences of `cycle` minus `setup` isolate the timed region.
            for _ in 0..iters {
                let package = litchi_pptx::Package::from_vec(bytes.clone()).unwrap();
                std::hint::black_box(package);
            }
            println!("setup done");
        },
        "open" => {
            // The harness's pptx_semantic_open region: Package::from_bytes.
            let mut times = Vec::new();
            for _ in 0..iters {
                let started = Instant::now();
                let package = litchi_pptx::Package::from_bytes(&bytes).unwrap();
                times.push(started.elapsed().as_nanos());
                std::hint::black_box(package);
            }
            println!("open median_ns {}", median(times));
        },
        "fulltext" => {
            let package = litchi_pptx::Package::from_bytes(&bytes).unwrap();
            let presentation = package.presentation().unwrap();
            let mut times = Vec::new();
            for _ in 0..iters {
                let started = Instant::now();
                let text = presentation.text().unwrap();
                times.push(started.elapsed().as_nanos());
                std::hint::black_box(text);
            }
            println!("fulltext median_ns {}", median(times));
        },
        "capture" => {
            let package = litchi_pptx::Package::from_bytes(&bytes).unwrap();
            let mut times = Vec::new();
            for _ in 0..iters {
                let started = Instant::now();
                let snapshot = package.opened_presentation().unwrap();
                times.push(started.elapsed().as_nanos());
                std::hint::black_box(snapshot);
            }
            println!("capture median_ns {}", median(times));
        },
        "settext" => {
            let package = litchi_pptx::Package::from_bytes(&bytes).unwrap();
            let snapshot = package.opened_presentation().unwrap();
            let mut times = Vec::new();
            for _ in 0..iters {
                let mut edit = snapshot.edit();
                let started = Instant::now();
                for linear in &selected {
                    let slide = *linear / boxes;
                    let object = *linear % boxes;
                    assert!(edit
                        .set_shape_text(slide, object, semantic_pptx_text(slide, object, true))
                        .unwrap());
                }
                times.push(started.elapsed().as_nanos());
                std::hint::black_box(edit);
            }
            println!("settext median_ns {}", median(times));
        },
        "commit" => {
            let package = litchi_pptx::Package::from_bytes(&bytes).unwrap();
            let snapshot = package.opened_presentation().unwrap();
            let mut edit = snapshot.edit();
            for linear in &selected {
                let slide = *linear / boxes;
                let object = *linear % boxes;
                edit.set_shape_text(slide, object, semantic_pptx_text(slide, object, true))
                    .unwrap();
            }
            let mut times = Vec::new();
            for _ in 0..iters {
                let staged = edit.clone();
                let started = Instant::now();
                let commit = staged.commit().unwrap();
                times.push(started.elapsed().as_nanos());
                std::hint::black_box(commit);
            }
            println!("commit median_ns {}", median(times));
        },
        "cycle" => {
            let mut phase = [Vec::new(), Vec::new(), Vec::new(), Vec::new(), Vec::new(), Vec::new()];
            let mut totals = Vec::new();
            let mut digest = None;
            for _ in 0..iters {
                let mut package = litchi_pptx::Package::from_vec(bytes.clone()).unwrap();
                let t0 = Instant::now();
                let snapshot = package.opened_presentation().unwrap();
                let t1 = Instant::now();
                let mut edit = snapshot.edit();
                let t2 = Instant::now();
                for linear in &selected {
                    let slide = *linear / boxes;
                    let object = *linear % boxes;
                    assert!(edit
                        .set_shape_text(slide, object, semantic_pptx_text(slide, object, true))
                        .unwrap());
                }
                let t3 = Instant::now();
                let commit = edit.commit().unwrap();
                let t4 = Instant::now();
                package.apply_opened_presentation_commit(commit).unwrap();
                let t5 = Instant::now();
                let out = package.to_bytes().unwrap();
                let t6 = Instant::now();
                phase[0].push((t1 - t0).as_nanos());
                phase[1].push((t2 - t1).as_nanos());
                phase[2].push((t3 - t2).as_nanos());
                phase[3].push((t4 - t3).as_nanos());
                phase[4].push((t5 - t4).as_nanos());
                phase[5].push((t6 - t5).as_nanos());
                totals.push((t6 - t0).as_nanos());
                let len = out.len();
                match digest {
                    None => digest = Some(out),
                    Some(ref previous) => assert_eq!(previous, &out, "nondeterministic output"),
                }
                std::hint::black_box(len);
            }
            let names = ["capture", "edit", "set_text", "commit", "apply", "to_bytes"];
            for (name, values) in names.iter().zip(phase) {
                println!("{name:10} median_ns {}", median(values));
            }
            println!("total      median_ns {}", median(totals));
            let out = digest.unwrap();
            let reopened = litchi_pptx::Package::from_bytes(&out).unwrap();
            let text = reopened.presentation().unwrap().text().unwrap();
            println!("output bytes {} text bytes {}", out.len(), text.len());
        },
        "alloc" => {
            // One region per line: allocation calls and requested bytes of
            // exactly the timed regions, on a warmed process.
            let package = litchi_pptx::Package::from_bytes(&bytes).unwrap();
            let presentation = package.presentation().unwrap();
            let _ = presentation.text().unwrap();
            let before = allocations();
            let text = presentation.text().unwrap();
            let after = allocations();
            std::hint::black_box(text);
            println!("fulltext calls {} bytes {}", after.0 - before.0, after.1 - before.1);
            for (label, chosen) in [("noop", Vec::new()), ("one", vec![update_indices(total)[0]]), ("pct", update_indices(total))] {
                let mut package = litchi_pptx::Package::from_vec(bytes.clone()).unwrap();
                let before = allocations();
                let mut edit = package.opened_presentation_transaction().unwrap();
                for linear in &chosen {
                    let slide = *linear / boxes;
                    let object = *linear % boxes;
                    edit.set_shape_text(slide, object, semantic_pptx_text(slide, object, true)).unwrap();
                }
                let commit = edit.commit().unwrap();
                package.apply_opened_presentation_commit(commit).unwrap();
                let out = package.to_bytes().unwrap();
                let after = allocations();
                std::hint::black_box(out);
                println!("{label} calls {} bytes {}", after.0 - before.0, after.1 - before.1);
            }
        },
        "dump" => {
            // Published bytes and full text, for a byte-for-byte comparison of
            // two builds.
            let directory = std::path::PathBuf::from(args.get(4).expect("dump directory"));
            std::fs::create_dir_all(&directory).unwrap();
            let package = litchi_pptx::Package::from_bytes(&bytes).unwrap();
            let text = package.presentation().unwrap().text().unwrap();
            std::fs::write(directory.join(format!("{shape}-fulltext.txt")), text).unwrap();
            for (label, chosen) in [("noop", Vec::new()), ("one", vec![update_indices(total)[0]]), ("pct", update_indices(total))] {
                let mut package = litchi_pptx::Package::from_vec(bytes.clone()).unwrap();
                let mut edit = package.opened_presentation_transaction().unwrap();
                for linear in &chosen {
                    let slide = *linear / boxes;
                    let object = *linear % boxes;
                    edit.set_shape_text(slide, object, semantic_pptx_text(slide, object, true)).unwrap();
                }
                let commit = edit.commit().unwrap();
                let patch = format!("{:?}", commit.patch().is_empty());
                let snapshot = package.apply_opened_presentation_commit(commit).unwrap();
                let out = package.to_bytes().unwrap();
                std::fs::write(directory.join(format!("{shape}-{label}.pptx")), &out).unwrap();
                std::fs::write(
                    directory.join(format!("{shape}-{label}.revision")),
                    format!("{:?} {patch}\n", snapshot.revision()),
                )
                .unwrap();
            }
        },
        other => panic!("unknown mode {other}"),
    }
}
