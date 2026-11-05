//! Matched native owner qualification probe for the four retained iWork routes.
//!
//! The same source is compiled against the frozen baseline and candidate
//! snapshots. Fixture construction, correctness checks, and output hashing
//! happen before the timed closure. The timed closure therefore measures the
//! owner operation rather than setup, replay, or digest work.
//!
//! Usage:
//!
//! ```text
//! owner-probe ROUTE ITERATIONS FIXTURE_DIR
//! owner-probe validate_outputs FIXTURE_DIR OUTPUT_DIR
//! ```

use std::{hint::black_box, time::Instant};

use sha2::{Digest, Sha256};

#[cfg(feature = "allocations")]
mod allocation {
    use std::alloc::{GlobalAlloc, Layout, System};
    use std::sync::atomic::{AtomicUsize, Ordering};

    pub static CURRENT: AtomicUsize = AtomicUsize::new(0);
    pub static PEAK: AtomicUsize = AtomicUsize::new(0);
    pub static BYTES: AtomicUsize = AtomicUsize::new(0);
    pub static CALLS: AtomicUsize = AtomicUsize::new(0);

    pub struct Counting;

    // SAFETY: every operation forwards to the system allocator with the
    // caller's layout and pointer; the atomics only observe sizes.
    unsafe impl GlobalAlloc for Counting {
        unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
            // SAFETY: the system allocator receives the original layout.
            let pointer = unsafe { System.alloc(layout) };
            if !pointer.is_null() {
                add(layout.size());
            }
            pointer
        }

        unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
            // SAFETY: the pointer and layout came from this allocator.
            unsafe { System.dealloc(pointer, layout) };
            CURRENT.fetch_sub(layout.size(), Ordering::Relaxed);
        }

        unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, size: usize) -> *mut u8 {
            // SAFETY: the pointer and old layout came from this allocator.
            let moved = unsafe { System.realloc(pointer, layout, size) };
            if !moved.is_null() {
                let live = if size >= layout.size() {
                    CURRENT.fetch_add(size - layout.size(), Ordering::Relaxed) + size
                        - layout.size()
                } else {
                    CURRENT.fetch_sub(layout.size() - size, Ordering::Relaxed)
                        - (layout.size() - size)
                };
                PEAK.fetch_max(live, Ordering::Relaxed);
                BYTES.fetch_add(size, Ordering::Relaxed);
                CALLS.fetch_add(1, Ordering::Relaxed);
            }
            moved
        }
    }

    fn add(size: usize) {
        let live = CURRENT.fetch_add(size, Ordering::Relaxed) + size;
        PEAK.fetch_max(live, Ordering::Relaxed);
        BYTES.fetch_add(size, Ordering::Relaxed);
        CALLS.fetch_add(1, Ordering::Relaxed);
    }

    pub fn start() -> usize {
        let live = CURRENT.load(Ordering::Relaxed);
        PEAK.store(live, Ordering::Relaxed);
        BYTES.store(0, Ordering::Relaxed);
        CALLS.store(0, Ordering::Relaxed);
        live
    }

    pub fn result(base: usize) -> (usize, usize, usize, usize) {
        (
            PEAK.load(Ordering::Relaxed).saturating_sub(base),
            CURRENT.load(Ordering::Relaxed).saturating_sub(base),
            BYTES.load(Ordering::Relaxed),
            CALLS.load(Ordering::Relaxed),
        )
    }
}

#[cfg(feature = "allocations")]
#[global_allocator]
static ALLOCATOR: allocation::Counting = allocation::Counting;

const ROUTES: [&str; 4] = [
    "numbers_headers",
    "keynote_media",
    "pages_footnote_insert",
    "pages_section_name",
];
const SECTION_SELECTOR: litchi_pages::SectionSelector<'static> =
    litchi_pages::SectionSelector::index(0);
const SECTION_TARGET: &str = "Native Pages section name with a deliberately longer UTF-8 value";
const SECTION_TARGET_ALT: &str =
    "Native Pages section name with a deliberately longer UTF-8 value 2";

fn sha256(bytes: &[u8]) -> String {
    let mut text = String::with_capacity(64);
    for byte in Sha256::digest(bytes) {
        text.push_str(&format!("{byte:02x}"));
    }
    text
}

fn write_numbers(package: &litchi_numbers::Package) -> Vec<u8> {
    let mut output = Vec::new();
    package
        .write_to(&mut output)
        .expect("Numbers package write");
    output
}

fn write_keynote(package: &litchi_keynote::Package) -> Vec<u8> {
    let mut output = Vec::new();
    package
        .write_to(&mut output)
        .expect("Keynote package write");
    output
}

fn write_pages(package: &litchi_pages::Package) -> Vec<u8> {
    let mut output = Vec::new();
    package.write_to(&mut output).expect("Pages package write");
    output
}

#[derive(Debug, Clone)]
struct Correctness {
    source_sha256: String,
    changed_sha256: String,
    replay_sha256: String,
    inverse_sha256: String,
    noop_sha256: String,
    source_bytes: usize,
    changed_bytes: usize,
    noop_bytes: usize,
    changed_differs: bool,
    replay_matches_changed: bool,
    inverse_matches_source: bool,
    noop_is_noop: bool,
    noop_matches_expected: bool,
    semantic_ok: bool,
}

impl Correctness {
    fn assert_valid(&self) {
        assert!(self.changed_differs, "owner route did not change bytes");
        assert!(
            self.replay_matches_changed,
            "replayed patch did not reproduce the changed package"
        );
        assert!(
            self.inverse_matches_source,
            "inverse patch did not reproduce the source package"
        );
        assert!(self.noop_is_noop, "owner no-op patch was not marked no-op");
        assert!(
            self.noop_matches_expected,
            "owner no-op did not preserve the expected package bytes"
        );
        assert!(
            self.semantic_ok,
            "owner route did not prove its semantic result"
        );
    }

    fn json(&self) -> String {
        format!(
            "{{\"source_sha256\":\"{}\",\"changed_sha256\":\"{}\",\"replay_sha256\":\"{}\",\"inverse_sha256\":\"{}\",\"noop_sha256\":\"{}\",\"source_bytes\":{},\"changed_bytes\":{},\"noop_bytes\":{},\"changed_differs\":{},\"replay_matches_changed\":{},\"inverse_matches_source\":{},\"noop_is_noop\":{},\"noop_matches_expected\":{},\"semantic_ok\":{}}}",
            self.source_sha256,
            self.changed_sha256,
            self.replay_sha256,
            self.inverse_sha256,
            self.noop_sha256,
            self.source_bytes,
            self.changed_bytes,
            self.noop_bytes,
            self.changed_differs,
            self.replay_matches_changed,
            self.inverse_matches_source,
            self.noop_is_noop,
            self.noop_matches_expected,
            self.semantic_ok,
        )
    }
}

fn numbers_correctness(
    source: &[u8],
    package: &litchi_numbers::Package,
    changed: litchi_numbers::table::headers::Settings,
) -> Correctness {
    let before = package.table_header_settings(0usize, 0usize).unwrap();
    let commit = package
        .edit_table_headers(0usize, 0usize)
        .unwrap()
        .set(changed)
        .commit()
        .unwrap();
    let returned = commit
        .package()
        .table_header_settings(0usize, 0usize)
        .unwrap();
    let semantic_ok = returned == changed && before != changed;
    let changed_bytes = write_numbers(commit.package());
    let replay = litchi_numbers::Package::from_bytes(source)
        .unwrap()
        .apply_table_headers(commit.patch())
        .unwrap();
    let replay_bytes = write_numbers(replay.package());
    let inverse_patch = commit.patch().inverse();
    let inverse = commit
        .package()
        .apply_table_headers(&inverse_patch)
        .unwrap();
    let inverse_bytes = write_numbers(inverse.package());
    let noop = package
        .edit_table_headers(0usize, 0usize)
        .unwrap()
        .set(before)
        .commit()
        .unwrap();
    let noop_bytes = write_numbers(noop.package());
    Correctness {
        source_sha256: sha256(source),
        changed_sha256: sha256(&changed_bytes),
        replay_sha256: sha256(&replay_bytes),
        inverse_sha256: sha256(&inverse_bytes),
        noop_sha256: sha256(&noop_bytes),
        source_bytes: source.len(),
        changed_bytes: changed_bytes.len(),
        noop_bytes: noop_bytes.len(),
        changed_differs: changed_bytes != source,
        replay_matches_changed: replay_bytes == changed_bytes,
        inverse_matches_source: inverse_bytes == source,
        noop_is_noop: noop.patch().is_noop(),
        noop_matches_expected: noop_bytes == source,
        semantic_ok,
    }
}

fn keynote_correctness(
    source: &[u8],
    package: &litchi_keynote::Package,
    replacement: &[u8],
) -> Correctness {
    use litchi_keynote::slide::media::data::MediaPart;
    let selector = litchi_keynote::SlideSelector::index(0);
    let movie = litchi_keynote::MovieSelector::index(0);
    let part = MediaPart::Content;
    let before = package
        .slide_media_data(selector, movie, part)
        .unwrap()
        .to_vec();
    let commit = package
        .edit_slide_media_data(selector, movie, part)
        .unwrap()
        .set(replacement)
        .unwrap()
        .commit()
        .unwrap();
    let semantic_ok = commit
        .package()
        .slide_media_data(selector, movie, part)
        .unwrap()
        == replacement
        && before != replacement;
    let changed_bytes = write_keynote(commit.package());
    let replay = litchi_keynote::Package::from_bytes(source)
        .unwrap()
        .apply_slide_media_data(commit.patch())
        .unwrap();
    let replay_bytes = write_keynote(replay.package());
    let inverse_patch = commit.patch().inverse();
    let inverse = commit
        .package()
        .apply_slide_media_data(&inverse_patch)
        .unwrap();
    let inverse_bytes = write_keynote(inverse.package());
    let noop = package
        .edit_slide_media_data(selector, movie, part)
        .unwrap()
        .set(&before)
        .unwrap()
        .commit()
        .unwrap();
    let noop_bytes = write_keynote(noop.package());
    Correctness {
        source_sha256: sha256(source),
        changed_sha256: sha256(&changed_bytes),
        replay_sha256: sha256(&replay_bytes),
        inverse_sha256: sha256(&inverse_bytes),
        noop_sha256: sha256(&noop_bytes),
        source_bytes: source.len(),
        changed_bytes: changed_bytes.len(),
        noop_bytes: noop_bytes.len(),
        changed_differs: changed_bytes != source,
        replay_matches_changed: replay_bytes == changed_bytes,
        inverse_matches_source: inverse_bytes == source,
        noop_is_noop: noop.patch().is_noop(),
        noop_matches_expected: noop_bytes == source,
        semantic_ok,
    }
}

fn pages_section_noop(package: &litchi_pages::Package) -> (bool, Vec<u8>) {
    let name = package.section_name(SECTION_SELECTOR).unwrap();
    let mut edit = package.edit_section_name(SECTION_SELECTOR).unwrap();
    edit.set_name(name).unwrap();
    let commit = edit.commit().unwrap();
    (commit.patch().is_noop(), write_pages(commit.package()))
}

fn pages_footnote_correctness(source: &[u8], package: &litchi_pages::Package) -> Correctness {
    let position = litchi_pages::footnote::body::Position::from_utf16_index(1).unwrap();
    let commit = package
        .insert_body_footnote(position, "Inserted", Some("\u{2021}"))
        .unwrap();
    let semantic_ok = commit
        .package()
        .body_footnotes()
        .unwrap()
        .iter()
        .any(|note| {
            note.position == position
                && note.text.as_ref() == "Inserted"
                && note.custom_mark.as_deref() == Some("\u{2021}")
        });
    let changed_bytes = write_pages(commit.package());
    let replay = litchi_pages::Package::from_bytes(source)
        .unwrap()
        .apply_body_footnote(commit.patch())
        .unwrap();
    let replay_bytes = write_pages(replay.package());
    let inverse_patch = commit.patch().inverse();
    let inverse = commit
        .package()
        .apply_body_footnote(&inverse_patch)
        .unwrap();
    let inverse_bytes = write_pages(inverse.package());
    // Exercise the same body-footnote family for the no-op check. The
    // insertion above creates a note at `position`; reopening that exact
    // anchor and committing without edits must preserve the changed package.
    let noop = commit
        .package()
        .edit_body_footnote(litchi_pages::footnote::body::Selector::at(position))
        .unwrap()
        .commit()
        .unwrap();
    let noop_is_noop = noop.patch().is_noop();
    let noop_bytes = write_pages(noop.package());
    Correctness {
        source_sha256: sha256(source),
        changed_sha256: sha256(&changed_bytes),
        replay_sha256: sha256(&replay_bytes),
        inverse_sha256: sha256(&inverse_bytes),
        noop_sha256: sha256(&noop_bytes),
        source_bytes: source.len(),
        changed_bytes: changed_bytes.len(),
        noop_bytes: noop_bytes.len(),
        changed_differs: changed_bytes != source,
        replay_matches_changed: replay_bytes == changed_bytes,
        inverse_matches_source: inverse_bytes == source,
        noop_is_noop,
        noop_matches_expected: noop_bytes == changed_bytes,
        semantic_ok,
    }
}

fn pages_section_correctness(source: &[u8], package: &litchi_pages::Package) -> Correctness {
    let old_name = package.section_name(SECTION_SELECTOR).unwrap();
    let name = match old_name {
        Some(value) if value == SECTION_TARGET => SECTION_TARGET_ALT,
        _ => SECTION_TARGET,
    };
    let mut edit = package.edit_section_name(SECTION_SELECTOR).unwrap();
    edit.set_name(Some(name)).unwrap();
    let commit = edit.commit().unwrap();
    let semantic_ok = commit.package().section_name(SECTION_SELECTOR).unwrap() == Some(name)
        && old_name != Some(name);
    let changed_bytes = write_pages(commit.package());
    let replay = litchi_pages::Package::from_bytes(source)
        .unwrap()
        .apply_section_name(commit.patch())
        .unwrap();
    let replay_bytes = write_pages(replay.package());
    let inverse_patch = commit.patch().inverse();
    let inverse = commit.package().apply_section_name(&inverse_patch).unwrap();
    let inverse_bytes = write_pages(inverse.package());
    let (noop_is_noop, noop_bytes) = pages_section_noop(package);
    Correctness {
        source_sha256: sha256(source),
        changed_sha256: sha256(&changed_bytes),
        replay_sha256: sha256(&replay_bytes),
        inverse_sha256: sha256(&inverse_bytes),
        noop_sha256: sha256(&noop_bytes),
        source_bytes: source.len(),
        changed_bytes: changed_bytes.len(),
        noop_bytes: noop_bytes.len(),
        changed_differs: changed_bytes != source,
        replay_matches_changed: replay_bytes == changed_bytes,
        inverse_matches_source: inverse_bytes == source,
        noop_is_noop,
        noop_matches_expected: noop_bytes == source,
        semantic_ok,
    }
}

fn owner_correctness(route: &str, directory: &std::path::Path) -> Correctness {
    match route {
        "numbers_headers" => {
            let source = std::fs::read(directory.join("basic.numbers")).unwrap();
            let package = litchi_numbers::Package::from_bytes(&source).unwrap();
            let before = package.table_header_settings(0usize, 0usize).unwrap();
            let mut changed = before;
            changed.header_rows_frozen = Some(!before.header_rows_are_frozen());
            numbers_correctness(&source, &package, changed)
        },
        "keynote_media" => {
            let source = std::fs::read(directory.join("media-replacement-native.key")).unwrap();
            let package = litchi_keynote::Package::from_bytes(&source).unwrap();
            let replacement = std::fs::read(directory.join("tone.wav")).unwrap();
            keynote_correctness(&source, &package, &replacement)
        },
        "pages_footnote_insert" => {
            let source = std::fs::read(directory.join("footnote-fixture.pages")).unwrap();
            let package = litchi_pages::Package::from_bytes(&source).unwrap();
            pages_footnote_correctness(&source, &package)
        },
        "pages_section_name" => {
            let source = std::fs::read(directory.join("basic.pages")).unwrap();
            let package = litchi_pages::Package::from_bytes(&source).unwrap();
            pages_section_correctness(&source, &package)
        },
        other => panic!("unknown owner route {other}"),
    }
}

#[derive(Debug, Clone, Copy)]
struct AllocationMetrics {
    peak: usize,
    retained: usize,
    allocated: usize,
    allocations: usize,
}

#[cfg(feature = "allocations")]
fn allocation_start() -> usize {
    allocation::start()
}

#[cfg(not(feature = "allocations"))]
fn allocation_start() -> usize {
    0
}

#[cfg(feature = "allocations")]
fn allocation_finish(base: usize) -> AllocationMetrics {
    let (peak, retained, allocated, allocations) = allocation::result(base);
    AllocationMetrics {
        peak,
        retained,
        allocated,
        allocations,
    }
}

#[cfg(not(feature = "allocations"))]
fn allocation_finish(_base: usize) -> AllocationMetrics {
    AllocationMetrics {
        peak: 0,
        retained: 0,
        allocated: 0,
        allocations: 0,
    }
}

fn print_report(
    route: &str,
    iterations: usize,
    ns: u128,
    checksum: u64,
    output: &[u8],
    source_sha256: &str,
    correctness: &Correctness,
    metrics: AllocationMetrics,
) {
    assert_eq!(sha256(output), correctness.changed_sha256);
    assert_eq!(output.len(), correctness.changed_bytes);
    println!(
        "{{\"schema\":\"litchi.iwa.owner.probe.v2\",\"route\":\"{route}\",\"iterations\":{iterations},\"ns\":{ns},\"checksum\":{checksum},\"output_bytes\":{},\"source_sha256\":\"{source_sha256}\",\"output_sha256\":\"{}\",\"output_differs\":true,\"peak\":{},\"retained\":{},\"allocated\":{},\"allocations\":{},\"correctness\":{{\"status\":\"ok\",\"replay_matches_changed\":{},\"inverse_matches_source\":{},\"noop_is_noop\":{},\"noop_matches_expected\":{},\"semantic_ok\":{},\"source_bytes\":{},\"changed_bytes\":{},\"noop_bytes\":{},\"changed_sha256\":\"{}\",\"replay_sha256\":\"{}\",\"inverse_sha256\":\"{}\",\"noop_sha256\":\"{}\"}}}}",
        output.len(),
        correctness.changed_sha256,
        metrics.peak,
        metrics.retained,
        metrics.allocated,
        metrics.allocations,
        correctness.replay_matches_changed,
        correctness.inverse_matches_source,
        correctness.noop_is_noop,
        correctness.noop_matches_expected,
        correctness.semantic_ok,
        correctness.source_bytes,
        correctness.changed_bytes,
        correctness.noop_bytes,
        correctness.changed_sha256,
        correctness.replay_sha256,
        correctness.inverse_sha256,
        correctness.noop_sha256,
    );
}

fn run(route: &str, iterations: usize, directory: &std::path::Path) {
    let correctness = owner_correctness(route, directory);
    correctness.assert_valid();
    let source_sha256 = correctness.source_sha256.clone();
    match route {
        "numbers_headers" => {
            let source = std::fs::read(directory.join("basic.numbers")).unwrap();
            let package = litchi_numbers::Package::from_bytes(&source).unwrap();
            let before = package.table_header_settings(0usize, 0usize).unwrap();
            let mut changed = before;
            changed.header_rows_frozen = Some(!before.header_rows_are_frozen());
            let allocation_base = allocation_start();
            let started = Instant::now();
            let mut checksum = 0u64;
            for _ in 0..iterations {
                let commit = black_box(&package)
                    .edit_table_headers(0usize, 0usize)
                    .unwrap()
                    .set(black_box(changed))
                    .commit()
                    .unwrap();
                checksum = checksum.wrapping_add(commit.diagnostics().touched_components() as u64);
                black_box(commit);
            }
            let elapsed = started.elapsed().as_nanos();
            let metrics = allocation_finish(allocation_base);
            let output = package
                .edit_table_headers(0usize, 0usize)
                .unwrap()
                .set(changed)
                .commit()
                .unwrap();
            let bytes = write_numbers(output.package());
            print_report(
                route,
                iterations,
                elapsed,
                checksum,
                &bytes,
                &source_sha256,
                &correctness,
                metrics,
            );
        },
        "keynote_media" => {
            let source = std::fs::read(directory.join("media-replacement-native.key")).unwrap();
            let package = litchi_keynote::Package::from_bytes(&source).unwrap();
            let replacement = std::fs::read(directory.join("tone.wav")).unwrap();
            let run = |package: &litchi_keynote::Package| {
                package
                    .edit_slide_media_data(
                        litchi_keynote::SlideSelector::index(0),
                        litchi_keynote::MovieSelector::index(0),
                        litchi_keynote::slide::media::data::MediaPart::Content,
                    )
                    .and_then(|edit| edit.set(&replacement))
                    .and_then(|edit| edit.commit())
                    .unwrap()
            };
            let allocation_base = allocation_start();
            let started = Instant::now();
            let mut checksum = 0u64;
            for _ in 0..iterations {
                let commit = run(black_box(&package));
                checksum = checksum.wrapping_add(commit.diagnostics().touched_components() as u64);
                black_box(commit);
            }
            let elapsed = started.elapsed().as_nanos();
            let metrics = allocation_finish(allocation_base);
            let commit = run(&package);
            let bytes = write_keynote(commit.package());
            print_report(
                route,
                iterations,
                elapsed,
                checksum,
                &bytes,
                &source_sha256,
                &correctness,
                metrics,
            );
        },
        "pages_footnote_insert" => {
            let source = std::fs::read(directory.join("footnote-fixture.pages")).unwrap();
            let package = litchi_pages::Package::from_bytes(&source).unwrap();
            let position = litchi_pages::footnote::body::Position::from_utf16_index(1).unwrap();
            let allocation_base = allocation_start();
            let started = Instant::now();
            let mut checksum = 0u64;
            for _ in 0..iterations {
                let commit = black_box(&package)
                    .insert_body_footnote(position, black_box("Inserted"), Some("\u{2021}"))
                    .unwrap();
                checksum = checksum.wrapping_add(commit.diagnostics().touched_components() as u64);
                black_box(commit);
            }
            let elapsed = started.elapsed().as_nanos();
            let metrics = allocation_finish(allocation_base);
            let commit = package
                .insert_body_footnote(position, "Inserted", Some("\u{2021}"))
                .unwrap();
            let bytes = write_pages(commit.package());
            print_report(
                route,
                iterations,
                elapsed,
                checksum,
                &bytes,
                &source_sha256,
                &correctness,
                metrics,
            );
        },
        "pages_section_name" => {
            let source = std::fs::read(directory.join("basic.pages")).unwrap();
            let package = litchi_pages::Package::from_bytes(&source).unwrap();
            let old_name = package.section_name(SECTION_SELECTOR).unwrap();
            let name = match old_name {
                Some(value) if value == SECTION_TARGET => SECTION_TARGET_ALT,
                _ => SECTION_TARGET,
            };
            let run = |package: &litchi_pages::Package| {
                let mut edit = package.edit_section_name(SECTION_SELECTOR).unwrap();
                edit.set_name(Some(black_box(name))).unwrap();
                edit.commit().unwrap()
            };
            let allocation_base = allocation_start();
            let started = Instant::now();
            let mut checksum = 0u64;
            for _ in 0..iterations {
                let commit = run(black_box(&package));
                checksum = checksum.wrapping_add(commit.diagnostics().touched_components() as u64);
                black_box(commit);
            }
            let elapsed = started.elapsed().as_nanos();
            let metrics = allocation_finish(allocation_base);
            let commit = run(&package);
            let bytes = write_pages(commit.package());
            print_report(
                route,
                iterations,
                elapsed,
                checksum,
                &bytes,
                &source_sha256,
                &correctness,
                metrics,
            );
        },
        other => panic!("unknown owner route {other}"),
    }
}

fn validate_outputs(directory: &std::path::Path, output: &std::path::Path) {
    std::fs::create_dir_all(output).expect("validation output directory");
    let mut rows = Vec::with_capacity(ROUTES.len());
    for route in ROUTES {
        let correctness = owner_correctness(route, directory);
        correctness.assert_valid();
        rows.push(format!(
            "{{\"route\":\"{route}\",\"correctness\":{}}}",
            correctness.json()
        ));
    }
    let report = format!(
        "{{\"schema\":\"litchi.iwa.owner.validation.v1\",\"status\":\"ok\",\"routes\":[{}],\"route_checks\":[{}]}}\n",
        ROUTES
            .iter()
            .map(|route| format!("\"{route}\""))
            .collect::<Vec<_>>()
            .join(","),
        rows.join(",")
    );
    std::fs::write(output.join("validation.json"), report).expect("validation report");
}

fn main() {
    let mut args = std::env::args();
    let _program = args.next();
    let route = args.next().expect("route argument");
    if route == "validate_outputs" {
        let directory = args.next().expect("fixture directory");
        let output = args.next().expect("validation output directory");
        validate_outputs(
            std::path::Path::new(&directory),
            std::path::Path::new(&output),
        );
        return;
    }
    assert!(
        ROUTES.contains(&route.as_str()),
        "unknown owner route {route}"
    );
    let iterations = args
        .next()
        .expect("iteration argument")
        .parse::<usize>()
        .expect("iterations must be an integer");
    assert!(iterations > 0, "iterations must be positive");
    let directory = args.next().expect("fixture directory");
    run(&route, iterations, std::path::Path::new(&directory));
}
