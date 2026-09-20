//! Standalone native phase probe for the opened-PresentationML transaction.
//!
//! The default binary has no counting allocator.  It times only the five
//! public calls that make up one opened-presentation edit after the source
//! archive has been materialized and the edit target has been derived:
//! capture, working clone, `set_shape_text`, commit, and publication.  Source
//! package parsing is setup and is deliberately outside every sample timer.
//!
//! Sources are either a caller supplied `.pptx` path or a deterministic
//! `generated:<slides>x<text-boxes>` specification.  The `shape`, `target`,
//! and `prefix` modes are small correctness and profiling controls.  The
//! `allocations` feature adds a measurement-only global allocator; see the
//! packet README for its perturbation and accounting scope.

#[cfg(feature = "allocations")]
mod counting_allocator {
    use std::alloc::{GlobalAlloc, Layout, System};
    use std::sync::atomic::{AtomicU64, Ordering};

    static LIVE: AtomicU64 = AtomicU64::new(0);
    static PEAK: AtomicU64 = AtomicU64::new(0);
    static PHASE_PEAK: AtomicU64 = AtomicU64::new(0);
    static ALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
    static REQUESTED: AtomicU64 = AtomicU64::new(0);
    static REALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
    static REALLOC_REQUESTED: AtomicU64 = AtomicU64::new(0);

    pub struct Counting;

    #[derive(Clone, Copy)]
    pub struct Snapshot {
        pub live: u64,
        pub peak: u64,
        pub alloc_calls: u64,
        pub requested: u64,
        pub realloc_calls: u64,
        pub realloc_requested: u64,
    }

    #[derive(Clone, Copy, Debug)]
    pub struct Metrics {
        pub baseline_live: u64,
        pub current_live: u64,
        pub peak_live: u64,
        pub alloc_calls: u64,
        pub requested: u64,
        pub realloc_calls: u64,
        pub realloc_requested: u64,
    }

    #[inline]
    fn successful_alloc(size: usize) {
        let size = size as u64;
        ALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        REQUESTED.fetch_add(size, Ordering::Relaxed);
        let live = LIVE.fetch_add(size, Ordering::Relaxed).saturating_add(size);
        PEAK.fetch_max(live, Ordering::Relaxed);
        PHASE_PEAK.fetch_max(live, Ordering::Relaxed);
    }

    #[inline]
    fn successful_realloc(old_size: usize, new_size: usize) {
        let old_size = old_size as u64;
        let new_size = new_size as u64;
        REALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        REALLOC_REQUESTED.fetch_add(new_size, Ordering::Relaxed);
        REQUESTED.fetch_add(new_size, Ordering::Relaxed);
        if new_size >= old_size {
            let live = LIVE
                .fetch_add(new_size - old_size, Ordering::Relaxed)
                .saturating_add(new_size - old_size);
            PEAK.fetch_max(live, Ordering::Relaxed);
            PHASE_PEAK.fetch_max(live, Ordering::Relaxed);
        } else {
            LIVE.fetch_sub(old_size - new_size, Ordering::Relaxed);
        }
    }

    #[inline]
    fn successful_dealloc(size: usize) {
        LIVE.fetch_sub(size as u64, Ordering::Relaxed);
    }

    // SAFETY: every operation forwards to `System`; the atomics only observe
    // successful storage transitions and never alter the returned pointer.
    unsafe impl GlobalAlloc for Counting {
        unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
            let pointer = unsafe { System.alloc(layout) };
            if !pointer.is_null() {
                successful_alloc(layout.size());
            }
            pointer
        }

        unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
            let pointer = unsafe { System.alloc_zeroed(layout) };
            if !pointer.is_null() {
                successful_alloc(layout.size());
            }
            pointer
        }

        unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
            unsafe { System.dealloc(pointer, layout) };
            successful_dealloc(layout.size());
        }

        unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
            let replacement = unsafe { System.realloc(pointer, layout, new_size) };
            if !replacement.is_null() {
                successful_realloc(layout.size(), new_size);
            }
            replacement
        }
    }

    pub fn snapshot() -> Snapshot {
        Snapshot {
            live: LIVE.load(Ordering::Relaxed),
            peak: PEAK.load(Ordering::Relaxed),
            alloc_calls: ALLOC_CALLS.load(Ordering::Relaxed),
            requested: REQUESTED.load(Ordering::Relaxed),
            realloc_calls: REALLOC_CALLS.load(Ordering::Relaxed),
            realloc_requested: REALLOC_REQUESTED.load(Ordering::Relaxed),
        }
    }

    pub fn total_start() -> Snapshot {
        let snapshot = snapshot();
        PEAK.store(snapshot.live, Ordering::Relaxed);
        snapshot
    }

    pub fn phase_start() -> Snapshot {
        let snapshot = snapshot();
        PHASE_PEAK.store(snapshot.live, Ordering::Relaxed);
        snapshot
    }

    pub fn phase_end(start: Snapshot) -> Metrics {
        let end = snapshot();
        Metrics {
            baseline_live: start.live,
            current_live: end.live,
            peak_live: PHASE_PEAK.load(Ordering::Relaxed),
            alloc_calls: end.alloc_calls.saturating_sub(start.alloc_calls),
            requested: end.requested.saturating_sub(start.requested),
            realloc_calls: end.realloc_calls.saturating_sub(start.realloc_calls),
            realloc_requested: end
                .realloc_requested
                .saturating_sub(start.realloc_requested),
        }
    }

    pub fn total_end(start: Snapshot) -> Metrics {
        let end = snapshot();
        Metrics {
            baseline_live: start.live,
            current_live: end.live,
            peak_live: end.peak,
            alloc_calls: end.alloc_calls.saturating_sub(start.alloc_calls),
            requested: end.requested.saturating_sub(start.requested),
            realloc_calls: end.realloc_calls.saturating_sub(start.realloc_calls),
            realloc_requested: end
                .realloc_requested
                .saturating_sub(start.realloc_requested),
        }
    }
}

#[cfg(feature = "allocations")]
#[global_allocator]
static ALLOCATOR: counting_allocator::Counting = counting_allocator::Counting;

use litchi_pptx::Package;
use sha2::{Digest, Sha256};
use std::env;
use std::error::Error;
use std::hint::black_box;
use std::time::Instant;

const EDIT_MARKER: &str = "litchi-perf-0691-opened-phase";
const PROBE_ID: &str = "0702";
const MAX_SAMPLES: usize = 100_000;
const MAX_WARMUPS: usize = 10_000;
const MAX_PREFIX_ITERS: usize = 1_000_000;
const MAX_GENERATED_SLIDES: usize = 64;
const MAX_GENERATED_SHAPES: usize = 64;

type Fallible<T> = Result<T, Box<dyn Error + Send + Sync>>;

#[derive(Clone, Copy, Debug)]
struct Target {
    slide: usize,
    shape: usize,
}

#[derive(Clone, Copy, Debug)]
struct Sample {
    capture_ns: u128,
    clone_ns: u128,
    settext_ns: u128,
    commit_ns: u128,
    apply_ns: u128,
    total_ns: u128,
    #[cfg(feature = "allocations")]
    allocations: AllocationReport,
}

#[cfg(feature = "allocations")]
#[derive(Clone, Copy, Debug)]
struct AllocationReport {
    capture: counting_allocator::Metrics,
    clone: counting_allocator::Metrics,
    settext: counting_allocator::Metrics,
    commit: counting_allocator::Metrics,
    apply: counting_allocator::Metrics,
    total: counting_allocator::Metrics,
}

#[derive(Debug)]
struct SemanticDigest {
    digest: String,
    payload_bytes: usize,
    slides: usize,
    shapes: usize,
}

#[derive(Debug)]
struct Validation {
    before_revision: String,
    after_revision: String,
    before_semantic: SemanticDigest,
    after_semantic: SemanticDigest,
    candidate_bytes: usize,
    candidate_sha256: String,
    reopened_target_text_sha256: String,
}

#[derive(Debug)]
struct NoopValidation {
    before_revision: String,
    after_revision: String,
    before_archive_bytes: usize,
    after_archive_bytes: usize,
    before_archive_sha256: String,
    after_archive_sha256: String,
    before_semantic: SemanticDigest,
    after_semantic: SemanticDigest,
}

#[derive(Debug)]
struct TwoEditValidation {
    before_revision: String,
    after_revision: String,
    before_semantic: SemanticDigest,
    after_semantic: SemanticDigest,
    candidate_bytes: usize,
    candidate_sha256: String,
    reopened_target_text_sha256: [String; 2],
}

fn source_bytes(source: &str) -> Fallible<Vec<u8>> {
    if let Some(shape) = source.strip_prefix("generated:") {
        let (slides, shapes) = shape
            .split_once('x')
            .ok_or("generated source must be generated:<slides>x<shapes>")?;
        let slides = parse_bounded(slides, "generated slide count", MAX_GENERATED_SLIDES)?;
        let shapes = parse_bounded(shapes, "generated shape count", MAX_GENERATED_SHAPES)?;
        let mut package = Package::new()?;
        {
            let presentation = package.presentation_mut()?;
            for slide_index in 0..slides {
                let slide = presentation.add_slide()?;
                for shape_index in 0..shapes {
                    slide.add_text_box(
                        &format!(
                            "litchi-perf-baseline-pptx-semantic-v1-source-{slide_index:05}-{shape_index:05}"
                        ),
                        36 + i64::try_from(shape_index % 4)? * 180,
                        36 + i64::try_from(shape_index / 4)? * 90,
                        144,
                        54,
                    );
                }
            }
        }
        return Ok(package.to_bytes()?);
    }
    Ok(std::fs::read(source)?)
}

fn parse_bounded(value: &str, label: &str, maximum: usize) -> Fallible<usize> {
    let parsed = value.parse::<usize>()?;
    if parsed == 0 || parsed > maximum {
        return Err(format!("{label} must be between 1 and {maximum}, got {parsed}").into());
    }
    Ok(parsed)
}

fn sha256(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    hex(digest.as_ref())
}

fn hex(bytes: &[u8]) -> String {
    const DIGITS: &[u8; 16] = b"0123456789abcdef";
    let mut result = String::with_capacity(bytes.len() * 2);
    for &byte in bytes {
        result.push(char::from(DIGITS[(byte >> 4) as usize]));
        result.push(char::from(DIGITS[(byte & 0x0f) as usize]));
    }
    result
}

fn update_field(hasher: &mut Sha256, bytes: &mut usize, label: &str, value: &[u8]) {
    hasher.update((label.len() as u64).to_le_bytes());
    hasher.update(label.as_bytes());
    hasher.update((value.len() as u64).to_le_bytes());
    hasher.update(value);
    *bytes = bytes.saturating_add(value.len());
}

/// Hash the logical slide and shape payloads through the public semantic API.
/// The complete package revision is reported separately from this digest; the
/// latter makes the changed text payload and all untouched slide text visible
/// in the correctness record without depending on ZIP member order.
fn semantic_digest(package: &Package) -> Fallible<SemanticDigest> {
    let presentation = package.presentation()?;
    let slides = presentation.slides()?;
    let mut hasher = Sha256::new();
    let mut payload_bytes = 0usize;
    let mut shape_count = 0usize;
    for (slide_index, slide) in slides.iter().enumerate() {
        update_field(
            &mut hasher,
            &mut payload_bytes,
            "slide-index",
            &(slide_index as u64).to_le_bytes(),
        );
        let name = slide.name()?;
        update_field(
            &mut hasher,
            &mut payload_bytes,
            "slide-name",
            name.as_bytes(),
        );
        let text = slide.text()?;
        update_field(
            &mut hasher,
            &mut payload_bytes,
            "slide-text",
            text.as_bytes(),
        );
        let scene = slide.shapes()?;
        shape_count = shape_count.saturating_add(scene.len());
        for (shape_index, shape) in scene.iter().enumerate() {
            update_field(
                &mut hasher,
                &mut payload_bytes,
                "shape-index",
                &(shape_index as u64).to_le_bytes(),
            );
            update_field(
                &mut hasher,
                &mut payload_bytes,
                "shape-name",
                shape.common().name().unwrap_or("").as_bytes(),
            );
            update_field(
                &mut hasher,
                &mut payload_bytes,
                "shape-text",
                shape.common().text().unwrap_or("").as_bytes(),
            );
        }
    }
    Ok(SemanticDigest {
        digest: hex(hasher.finalize().as_ref()),
        payload_bytes,
        slides: slides.len(),
        shapes: shape_count,
    })
}

fn shape_text(package: &Package, target: Target) -> Fallible<String> {
    let presentation = package.presentation()?;
    let slides = presentation.slides()?;
    let slide = slides
        .get(target.slide)
        .ok_or_else(|| format!("slide {} is out of range", target.slide))?;
    let scene = slide.shapes()?;
    let shape = scene
        .at(target.shape)
        .map_err(|error| format!("target shape lookup failed: {error:?}"))?;
    Ok(shape.common().text().unwrap_or("").to_owned())
}

fn derive_target(archive: &[u8]) -> Fallible<Target> {
    let package = Package::from_bytes(archive)?;
    let snapshot = package.opened_presentation()?;
    for slide in 0..snapshot.slides().len().min(MAX_GENERATED_SLIDES) {
        for shape in 0..MAX_GENERATED_SHAPES {
            let mut edit = snapshot.edit();
            match edit.set_shape_text(slide, shape, EDIT_MARKER) {
                Ok(true) => return Ok(Target { slide, shape }),
                Ok(false) | Err(_) => {},
            }
        }
    }
    Err("no editable text shape found in the first bounded target search".into())
}

fn derive_second_target(archive: &[u8], first: Target) -> Fallible<Target> {
    let package = Package::from_bytes(archive)?;
    let snapshot = package.opened_presentation()?;
    for slide in 0..snapshot.slides().len().min(MAX_GENERATED_SLIDES) {
        if slide == first.slide {
            continue;
        }
        for shape in 0..MAX_GENERATED_SHAPES {
            let mut edit = snapshot.edit();
            match edit.set_shape_text(slide, shape, EDIT_MARKER) {
                Ok(true) => return Ok(Target { slide, shape }),
                Ok(false) | Err(_) => {},
            }
        }
    }
    Err("no second editable text shape found on a distinct bounded slide".into())
}

fn revision(snapshot: &litchi_pptx::opened::Snapshot) -> String {
    hex(&snapshot.revision())
}

// Untimed oracle: retain every part's identity, content type and relationships,
// and exact payload bytes for every part except the one explicitly edited.
fn untouched_digest_excluding(package: &Package, edited_parts: &[&str]) -> Fallible<String> {
    let opc = package.opc()?;
    let mut parts = Vec::new();
    for part in opc.try_iter_parts() {
        let part = part?;
        let edited = edited_parts
            .iter()
            .any(|name| *name == part.partname().as_str());
        let mut relationships = part
            .rels()
            .iter()
            .map(|r| {
                (
                    r.r_id().to_owned(),
                    r.reltype().to_owned(),
                    r.target_ref().to_owned(),
                    r.is_external(),
                )
            })
            .collect::<Vec<_>>();
        relationships.sort();
        parts.push((
            part.partname().as_str().to_owned(),
            part.content_type().to_owned(),
            (!edited).then(|| sha256(part.blob())),
            relationships,
        ));
    }
    parts.sort();
    let mut root_relationships = opc
        .rels()
        .iter()
        .map(|r| {
            (
                r.r_id().to_owned(),
                r.reltype().to_owned(),
                r.target_ref().to_owned(),
                r.is_external(),
            )
        })
        .collect::<Vec<_>>();
    root_relationships.sort();
    let mut non_parts = opc
        .non_part_members()
        .iter()
        .map(|m| (m.name().to_owned(), format!("{:?}", m.reason())))
        .collect::<Vec<_>>();
    non_parts.sort();
    Ok(sha256(
        format!("{parts:?}{root_relationships:?}{non_parts:?}").as_bytes(),
    ))
}

fn untouched_digest(package: &Package, edited_part: &str) -> Fallible<String> {
    untouched_digest_excluding(package, &[edited_part])
}

fn validate_edit(archive: &[u8], target: Target) -> Fallible<Validation> {
    let mut package = Package::from_bytes(archive)?;
    let before_snapshot = package.opened_presentation()?;
    let before_revision = revision(&before_snapshot);
    let before_semantic = semantic_digest(&package)?;
    let edited_part = before_snapshot.slides()[target.slide]
        .part_name()
        .as_str()
        .to_owned();
    let preserved = untouched_digest(&package, &edited_part)?;

    let mut edit = before_snapshot.edit();
    if !edit.set_shape_text(target.slide, target.shape, EDIT_MARKER)? {
        return Err("correctness edit did not change the selected shape".into());
    }
    let commit = edit.commit()?;
    if !commit.is_changed() {
        return Err("correctness commit reports no change".into());
    }
    let after_snapshot = package.apply_opened_presentation_commit(commit)?;
    let after_revision = revision(&after_snapshot);
    if before_revision == after_revision || untouched_digest(&package, &edited_part)? != preserved {
        return Err("edit revision or untouched part preservation oracle failed".into());
    }
    let candidate = package.to_bytes()?;
    let candidate_sha256 = sha256(&candidate);

    let reopened = Package::from_bytes(&candidate)?;
    let reopened_text = shape_text(&reopened, target)?;
    if reopened_text != EDIT_MARKER {
        return Err(format!(
            "semantic reopen returned {:?} instead of {:?}",
            reopened_text, EDIT_MARKER
        )
        .into());
    }
    let reopened_target_text_sha256 = sha256(reopened_text.as_bytes());
    let after_semantic = semantic_digest(&reopened)?;
    if before_semantic.slides != after_semantic.slides
        || before_semantic.shapes != after_semantic.shapes
        || untouched_digest(&reopened, &edited_part)? != preserved
    {
        return Err("reopened counts or untouched part preservation oracle failed".into());
    }
    Ok(Validation {
        before_revision,
        after_revision,
        before_semantic,
        after_semantic,
        candidate_bytes: candidate.len(),
        candidate_sha256,
        reopened_target_text_sha256,
    })
}

fn validate_noop(archive: &[u8], target: Target, original_text: &str) -> Fallible<NoopValidation> {
    let mut package = Package::from_bytes(archive)?;
    let before_snapshot = package.opened_presentation()?;
    let revision_before = revision(&before_snapshot);
    let semantic = semantic_digest(&package)?;
    let output_before = package.to_bytes()?;

    let mut edit = before_snapshot.edit();
    if edit.set_shape_text(target.slide, target.shape, original_text)? {
        return Err("no-op correctness edit unexpectedly changed the selected shape".into());
    }
    let commit = edit.commit()?;
    if commit.is_changed() {
        return Err("no-op correctness commit reports a change".into());
    }
    let after_snapshot = package.apply_opened_presentation_commit(commit)?;
    let revision_after = revision(&after_snapshot);
    let output_after = package.to_bytes()?;
    let after_semantic = semantic_digest(&package)?;
    if revision_before != revision_after
        || output_before != output_after
        || semantic.digest != after_semantic.digest
    {
        return Err("no-op revision or serialized output changed".into());
    }
    Ok(NoopValidation {
        before_revision: revision_before,
        after_revision: revision_after,
        before_archive_bytes: output_before.len(),
        after_archive_bytes: output_after.len(),
        before_archive_sha256: sha256(&output_before),
        after_archive_sha256: sha256(&output_after),
        before_semantic: semantic,
        after_semantic,
    })
}

fn validate_two_edit(archive: &[u8], targets: [Target; 2]) -> Fallible<TwoEditValidation> {
    if targets[0].slide == targets[1].slide {
        return Err("two-edit targets must be on distinct slides".into());
    }
    let mut package = Package::from_bytes(archive)?;
    let before_snapshot = package.opened_presentation()?;
    let before_revision = revision(&before_snapshot);
    let before_semantic = semantic_digest(&package)?;
    let edited_parts = [
        before_snapshot.slides()[targets[0].slide]
            .part_name()
            .as_str()
            .to_owned(),
        before_snapshot.slides()[targets[1].slide]
            .part_name()
            .as_str()
            .to_owned(),
    ];
    if edited_parts[0] == edited_parts[1] {
        return Err("two-edit targets resolved to one part".into());
    }
    let preserved = untouched_digest_excluding(
        &package,
        &[edited_parts[0].as_str(), edited_parts[1].as_str()],
    )?;

    let mut edit = before_snapshot.edit();
    if !edit.set_shape_text(targets[0].slide, targets[0].shape, EDIT_MARKER)?
        || !edit.set_shape_text(targets[1].slide, targets[1].shape, EDIT_MARKER)?
    {
        return Err("two-edit correctness edit did not change both selected shapes".into());
    }
    let commit = edit.commit()?;
    if !commit.is_changed() {
        return Err("two-edit correctness commit reports no change".into());
    }
    let after_snapshot = package.apply_opened_presentation_commit(commit)?;
    let after_revision = revision(&after_snapshot);
    if before_revision == after_revision
        || untouched_digest_excluding(
            &package,
            &[edited_parts[0].as_str(), edited_parts[1].as_str()],
        )? != preserved
    {
        return Err("two-edit revision or untouched part preservation oracle failed".into());
    }
    let candidate = package.to_bytes()?;
    let candidate_sha256 = sha256(&candidate);
    let reopened = Package::from_bytes(&candidate)?;
    let reopened_texts = [
        shape_text(&reopened, targets[0])?,
        shape_text(&reopened, targets[1])?,
    ];
    if reopened_texts.iter().any(|text| text != EDIT_MARKER) {
        return Err("semantic reopen did not preserve both edited target markers".into());
    }
    let after_semantic = semantic_digest(&reopened)?;
    if before_semantic.slides != after_semantic.slides
        || before_semantic.shapes != after_semantic.shapes
        || untouched_digest_excluding(
            &reopened,
            &[edited_parts[0].as_str(), edited_parts[1].as_str()],
        )? != preserved
    {
        return Err("reopened counts or two-edit preservation oracle failed".into());
    }
    Ok(TwoEditValidation {
        before_revision,
        after_revision,
        before_semantic,
        after_semantic,
        candidate_bytes: candidate.len(),
        candidate_sha256,
        reopened_target_text_sha256: [
            sha256(reopened_texts[0].as_bytes()),
            sha256(reopened_texts[1].as_bytes()),
        ],
    })
}

fn timed_edit(archive: &[u8], target: Target) -> Fallible<Sample> {
    // Opening the source package is deliberately outside the total and every
    // named phase. Each sample still starts from the same immutable bytes.
    let mut package = Package::from_bytes(archive)?;
    #[cfg(feature = "allocations")]
    let total_alloc_start = counting_allocator::total_start();
    let total_start = Instant::now();

    #[cfg(feature = "allocations")]
    let capture_alloc_start = counting_allocator::phase_start();
    let started = Instant::now();
    let snapshot = package.opened_presentation()?;
    let capture_ns = started.elapsed().as_nanos();
    #[cfg(feature = "allocations")]
    let capture_alloc = counting_allocator::phase_end(capture_alloc_start);

    #[cfg(feature = "allocations")]
    let clone_alloc_start = counting_allocator::phase_start();
    let started = Instant::now();
    let mut edit = snapshot.edit();
    let clone_ns = started.elapsed().as_nanos();
    #[cfg(feature = "allocations")]
    let clone_alloc = counting_allocator::phase_end(clone_alloc_start);

    #[cfg(feature = "allocations")]
    let settext_alloc_start = counting_allocator::phase_start();
    let started = Instant::now();
    let changed = edit.set_shape_text(target.slide, target.shape, EDIT_MARKER)?;
    let settext_ns = started.elapsed().as_nanos();
    #[cfg(feature = "allocations")]
    let settext_alloc = counting_allocator::phase_end(settext_alloc_start);
    if !changed {
        return Err("timed edit target reported no change".into());
    }

    #[cfg(feature = "allocations")]
    let commit_alloc_start = counting_allocator::phase_start();
    let started = Instant::now();
    let commit = edit.commit()?;
    let commit_ns = started.elapsed().as_nanos();
    #[cfg(feature = "allocations")]
    let commit_alloc = counting_allocator::phase_end(commit_alloc_start);
    if !commit.is_changed() {
        return Err("timed commit reports no change".into());
    }

    #[cfg(feature = "allocations")]
    let apply_alloc_start = counting_allocator::phase_start();
    let started = Instant::now();
    let published = package.apply_opened_presentation_commit(commit)?;
    let apply_ns = started.elapsed().as_nanos();
    let total_ns = total_start.elapsed().as_nanos();
    #[cfg(feature = "allocations")]
    let apply_alloc = counting_allocator::phase_end(apply_alloc_start);
    #[cfg(feature = "allocations")]
    let total_alloc = counting_allocator::total_end(total_alloc_start);
    black_box((&published, &package));

    Ok(Sample {
        capture_ns,
        clone_ns,
        settext_ns,
        commit_ns,
        apply_ns,
        total_ns,
        #[cfg(feature = "allocations")]
        allocations: AllocationReport {
            capture: capture_alloc,
            clone: clone_alloc,
            settext: settext_alloc,
            commit: commit_alloc,
            apply: apply_alloc,
            total: total_alloc,
        },
    })
}

fn timed_noop(archive: &[u8], target: Target, original_text: &str) -> Fallible<Sample> {
    let mut package = Package::from_bytes(archive)?;
    #[cfg(feature = "allocations")]
    let total_alloc_start = counting_allocator::total_start();
    let total_start = Instant::now();

    #[cfg(feature = "allocations")]
    let capture_alloc_start = counting_allocator::phase_start();
    let started = Instant::now();
    let snapshot = package.opened_presentation()?;
    let capture_ns = started.elapsed().as_nanos();
    #[cfg(feature = "allocations")]
    let capture_alloc = counting_allocator::phase_end(capture_alloc_start);

    #[cfg(feature = "allocations")]
    let clone_alloc_start = counting_allocator::phase_start();
    let started = Instant::now();
    let mut edit = snapshot.edit();
    let clone_ns = started.elapsed().as_nanos();
    #[cfg(feature = "allocations")]
    let clone_alloc = counting_allocator::phase_end(clone_alloc_start);

    #[cfg(feature = "allocations")]
    let settext_alloc_start = counting_allocator::phase_start();
    let started = Instant::now();
    let changed = edit.set_shape_text(target.slide, target.shape, original_text)?;
    let settext_ns = started.elapsed().as_nanos();
    #[cfg(feature = "allocations")]
    let settext_alloc = counting_allocator::phase_end(settext_alloc_start);
    if changed {
        return Err("no-op target unexpectedly reported a change".into());
    }

    #[cfg(feature = "allocations")]
    let commit_alloc_start = counting_allocator::phase_start();
    let started = Instant::now();
    let commit = edit.commit()?;
    let commit_ns = started.elapsed().as_nanos();
    #[cfg(feature = "allocations")]
    let commit_alloc = counting_allocator::phase_end(commit_alloc_start);
    if commit.is_changed() {
        return Err("no-op commit reports a change".into());
    }

    #[cfg(feature = "allocations")]
    let apply_alloc_start = counting_allocator::phase_start();
    let started = Instant::now();
    let published = package.apply_opened_presentation_commit(commit)?;
    let apply_ns = started.elapsed().as_nanos();
    let total_ns = total_start.elapsed().as_nanos();
    #[cfg(feature = "allocations")]
    let apply_alloc = counting_allocator::phase_end(apply_alloc_start);
    #[cfg(feature = "allocations")]
    let total_alloc = counting_allocator::total_end(total_alloc_start);
    if revision(&snapshot) != revision(&published) {
        return Err("no-op publication changed the revision".into());
    }
    black_box((&published, &package));

    Ok(Sample {
        capture_ns,
        clone_ns,
        settext_ns,
        commit_ns,
        apply_ns,
        total_ns,
        #[cfg(feature = "allocations")]
        allocations: AllocationReport {
            capture: capture_alloc,
            clone: clone_alloc,
            settext: settext_alloc,
            commit: commit_alloc,
            apply: apply_alloc,
            total: total_alloc,
        },
    })
}

fn timed_two_edit(archive: &[u8], targets: [Target; 2]) -> Fallible<Sample> {
    let mut package = Package::from_bytes(archive)?;
    #[cfg(feature = "allocations")]
    let total_alloc_start = counting_allocator::total_start();
    let total_start = Instant::now();

    #[cfg(feature = "allocations")]
    let capture_alloc_start = counting_allocator::phase_start();
    let started = Instant::now();
    let snapshot = package.opened_presentation()?;
    let capture_ns = started.elapsed().as_nanos();
    #[cfg(feature = "allocations")]
    let capture_alloc = counting_allocator::phase_end(capture_alloc_start);

    #[cfg(feature = "allocations")]
    let clone_alloc_start = counting_allocator::phase_start();
    let started = Instant::now();
    let mut edit = snapshot.edit();
    let clone_ns = started.elapsed().as_nanos();
    #[cfg(feature = "allocations")]
    let clone_alloc = counting_allocator::phase_end(clone_alloc_start);

    #[cfg(feature = "allocations")]
    let settext_alloc_start = counting_allocator::phase_start();
    let started = Instant::now();
    let first_changed = edit.set_shape_text(targets[0].slide, targets[0].shape, EDIT_MARKER)?;
    let second_changed = edit.set_shape_text(targets[1].slide, targets[1].shape, EDIT_MARKER)?;
    let settext_ns = started.elapsed().as_nanos();
    #[cfg(feature = "allocations")]
    let settext_alloc = counting_allocator::phase_end(settext_alloc_start);
    if !first_changed || !second_changed {
        return Err("two-edit target did not change both selected shapes".into());
    }

    #[cfg(feature = "allocations")]
    let commit_alloc_start = counting_allocator::phase_start();
    let started = Instant::now();
    let commit = edit.commit()?;
    let commit_ns = started.elapsed().as_nanos();
    #[cfg(feature = "allocations")]
    let commit_alloc = counting_allocator::phase_end(commit_alloc_start);
    if !commit.is_changed() {
        return Err("two-edit commit reports no change".into());
    }

    #[cfg(feature = "allocations")]
    let apply_alloc_start = counting_allocator::phase_start();
    let started = Instant::now();
    let published = package.apply_opened_presentation_commit(commit)?;
    let apply_ns = started.elapsed().as_nanos();
    let total_ns = total_start.elapsed().as_nanos();
    #[cfg(feature = "allocations")]
    let apply_alloc = counting_allocator::phase_end(apply_alloc_start);
    #[cfg(feature = "allocations")]
    let total_alloc = counting_allocator::total_end(total_alloc_start);
    black_box((&published, &package));

    Ok(Sample {
        capture_ns,
        clone_ns,
        settext_ns,
        commit_ns,
        apply_ns,
        total_ns,
        #[cfg(feature = "allocations")]
        allocations: AllocationReport {
            capture: capture_alloc,
            clone: clone_alloc,
            settext: settext_alloc,
            commit: commit_alloc,
            apply: apply_alloc,
            total: total_alloc,
        },
    })
}

fn print_provenance(source: &str, archive: &[u8], target: Target, samples: usize, warmups: usize) {
    println!("probe\t{PROBE_ID}");
    println!("source\t{source}");
    println!("source_archive_bytes\t{}", archive.len());
    println!("source_archive_sha256\t{}", sha256(archive));
    println!("target\tslide:{}\tshape:{}", target.slide, target.shape);
    println!("samples\t{samples}");
    println!("warmups\t{warmups}");
}

fn print_workflow_provenance(
    source: &str,
    archive: &[u8],
    workflow: &str,
    targets: &[Target],
    samples: usize,
    warmups: usize,
) {
    println!("probe\t{PROBE_ID}");
    println!("workflow\t{workflow}");
    println!("source\t{source}");
    println!("source_archive_bytes\t{}", archive.len());
    println!("source_archive_sha256\t{}", sha256(archive));
    for (index, target) in targets.iter().enumerate() {
        println!(
            "target{}\tslide:{}\tshape:{}",
            index + 1,
            target.slide,
            target.shape
        );
    }
    println!("samples\t{samples}");
    println!("warmups\t{warmups}");
}

#[cfg(feature = "allocations")]
fn allocation_header() -> String {
    let mut header = String::new();
    for phase in ["capture", "clone", "settext", "commit", "apply", "total"] {
        for field in [
            "alloc_calls",
            "requested_bytes",
            "baseline_live_bytes",
            "peak_live_bytes",
            "current_live_bytes",
            "realloc_calls",
            "realloc_requested_bytes",
        ] {
            header.push('\t');
            header.push_str(phase);
            header.push('_');
            header.push_str(field);
        }
    }
    header
}

#[cfg(feature = "allocations")]
fn allocation_values(report: &AllocationReport) -> String {
    let mut values = String::new();
    for metrics in [
        report.capture,
        report.clone,
        report.settext,
        report.commit,
        report.apply,
        report.total,
    ] {
        values.push_str(&format!(
            "\t{}\t{}\t{}\t{}\t{}\t{}\t{}",
            metrics.alloc_calls,
            metrics.requested,
            metrics.baseline_live,
            metrics.peak_live,
            metrics.current_live,
            metrics.realloc_calls,
            metrics.realloc_requested,
        ));
    }
    values
}

fn print_sample_table(samples_out: &[(usize, Sample)]) {
    let sample_header =
        "sample\tcapture_ns\tclone_ns\tsettext_ns\tcommit_ns\tapply_ns\ttotal_ns".to_owned();
    #[cfg(feature = "allocations")]
    let sample_header = {
        let mut sample_header = sample_header;
        sample_header.push_str(&allocation_header());
        sample_header
    };
    println!("{sample_header}");
    for (sample_index, sample) in samples_out {
        let row = format!(
            "{sample_index}\t{}\t{}\t{}\t{}\t{}\t{}",
            sample.capture_ns,
            sample.clone_ns,
            sample.settext_ns,
            sample.commit_ns,
            sample.apply_ns,
            sample.total_ns,
        );
        #[cfg(feature = "allocations")]
        let row = {
            let mut row = row;
            row.push_str(&allocation_values(&sample.allocations));
            row
        };
        println!("{row}");
    }
}

fn phases(source: &str, samples: usize, warmups: usize) -> Fallible<()> {
    let archive = source_bytes(source)?;
    let target = derive_target(&archive)?;
    for _ in 0..warmups {
        black_box(timed_edit(&archive, target)?);
    }
    let validation = validate_edit(&archive, target)?;
    print_provenance(source, &archive, target, samples, warmups);
    println!(
        "before_revision_sha256\t{}\nbefore_semantic_sha256\t{}\nbefore_semantic_payload_bytes\t{}\nbefore_slides\t{}\nbefore_shapes\t{}",
        validation.before_revision,
        validation.before_semantic.digest,
        validation.before_semantic.payload_bytes,
        validation.before_semantic.slides,
        validation.before_semantic.shapes,
    );
    let mut samples_out = Vec::with_capacity(samples);
    for sample_index in 0..samples {
        let sample = timed_edit(&archive, target)?;
        samples_out.push((sample_index, sample));
    }
    print_sample_table(&samples_out);
    println!(
        "after_revision_sha256\t{}\nafter_semantic_sha256\t{}\nafter_semantic_payload_bytes\t{}\nafter_slides\t{}\nafter_shapes\t{}\ncandidate_archive_bytes\t{}\ncandidate_archive_sha256\t{}\nreopened_target_text_sha256\t{}\ncorrectness_target_text\t{}",
        validation.after_revision,
        validation.after_semantic.digest,
        validation.after_semantic.payload_bytes,
        validation.after_semantic.slides,
        validation.after_semantic.shapes,
        validation.candidate_bytes,
        validation.candidate_sha256,
        validation.reopened_target_text_sha256,
        EDIT_MARKER,
    );
    Ok(())
}

fn noop(source: &str, samples: usize, warmups: usize) -> Fallible<()> {
    let archive = source_bytes(source)?;
    let target = derive_target(&archive)?;
    let source_package = Package::from_bytes(&archive)?;
    let original_text = shape_text(&source_package, target)?;
    for _ in 0..warmups {
        black_box(timed_noop(&archive, target, &original_text)?);
    }
    let validation = validate_noop(&archive, target, &original_text)?;
    print_workflow_provenance(source, &archive, "noop", &[target], samples, warmups);
    println!(
        "before_revision_sha256\t{}\nbefore_semantic_sha256\t{}\nbefore_semantic_payload_bytes\t{}\nbefore_slides\t{}\nbefore_shapes\t{}\nbefore_archive_bytes\t{}\nbefore_archive_sha256\t{}",
        validation.before_revision,
        validation.before_semantic.digest,
        validation.before_semantic.payload_bytes,
        validation.before_semantic.slides,
        validation.before_semantic.shapes,
        validation.before_archive_bytes,
        validation.before_archive_sha256,
    );
    let mut samples_out = Vec::with_capacity(samples);
    for sample_index in 0..samples {
        let sample = timed_noop(&archive, target, &original_text)?;
        samples_out.push((sample_index, sample));
    }
    print_sample_table(&samples_out);
    println!(
        "after_revision_sha256\t{}\nafter_semantic_sha256\t{}\nafter_semantic_payload_bytes\t{}\nafter_slides\t{}\nafter_shapes\t{}\nafter_archive_bytes\t{}\nafter_archive_sha256\t{}\ncommit_is_changed\tfalse\nrevision_identical\ttrue\noutput_identical\ttrue\ncorrectness_target_text_sha256\t{}",
        validation.after_revision,
        validation.after_semantic.digest,
        validation.after_semantic.payload_bytes,
        validation.after_semantic.slides,
        validation.after_semantic.shapes,
        validation.after_archive_bytes,
        validation.after_archive_sha256,
        sha256(original_text.as_bytes()),
    );
    Ok(())
}

fn two_edit(source: &str, samples: usize, warmups: usize) -> Fallible<()> {
    let archive = source_bytes(source)?;
    let first = derive_target(&archive)?;
    let second = derive_second_target(&archive, first)?;
    let targets = [first, second];
    for _ in 0..warmups {
        black_box(timed_two_edit(&archive, targets)?);
    }
    let validation = validate_two_edit(&archive, targets)?;
    print_workflow_provenance(source, &archive, "two", &targets, samples, warmups);
    println!(
        "before_revision_sha256\t{}\nbefore_semantic_sha256\t{}\nbefore_semantic_payload_bytes\t{}\nbefore_slides\t{}\nbefore_shapes\t{}",
        validation.before_revision,
        validation.before_semantic.digest,
        validation.before_semantic.payload_bytes,
        validation.before_semantic.slides,
        validation.before_semantic.shapes,
    );
    let mut samples_out = Vec::with_capacity(samples);
    for sample_index in 0..samples {
        let sample = timed_two_edit(&archive, targets)?;
        samples_out.push((sample_index, sample));
    }
    print_sample_table(&samples_out);
    println!(
        "after_revision_sha256\t{}\nafter_semantic_sha256\t{}\nafter_semantic_payload_bytes\t{}\nafter_slides\t{}\nafter_shapes\t{}\ncandidate_archive_bytes\t{}\ncandidate_archive_sha256\t{}\nreopened_target1_text_sha256\t{}\nreopened_target2_text_sha256\t{}\ncorrectness_target_text\t{}",
        validation.after_revision,
        validation.after_semantic.digest,
        validation.after_semantic.payload_bytes,
        validation.after_semantic.slides,
        validation.after_semantic.shapes,
        validation.candidate_bytes,
        validation.candidate_sha256,
        validation.reopened_target_text_sha256[0],
        validation.reopened_target_text_sha256[1],
        EDIT_MARKER,
    );
    Ok(())
}

fn stage_rank(stage: &str) -> Fallible<usize> {
    Ok(match stage {
        "open" => 0,
        "capture" => 1,
        "transaction" | "clone" => 2,
        "settext" => 3,
        "commit" => 4,
        "apply" => 5,
        other => return Err(format!("unknown prefix stage {other:?}").into()),
    })
}

fn prefix(source: &str, stage: &str, iterations: usize) -> Fallible<()> {
    let archive = source_bytes(source)?;
    let rank = stage_rank(stage)?;
    let target = if rank >= 3 {
        derive_target(&archive)?
    } else {
        Target { slide: 0, shape: 0 }
    };
    println!("probe\t{PROBE_ID}");
    println!("mode\tprefix");
    println!("source\t{source}");
    println!("source_archive_bytes\t{}", archive.len());
    println!("source_archive_sha256\t{}", sha256(&archive));
    println!("stage\t{stage}");
    println!("iterations\t{iterations}");
    let mut completed = 0usize;
    for iteration in 0..iterations {
        let mut package = Package::from_bytes(&archive)?;
        if rank == 0 {
            black_box(&package);
        } else {
            let snapshot = package.opened_presentation()?;
            if rank == 1 {
                black_box(&snapshot);
            } else {
                let mut edit = snapshot.edit();
                if rank == 2 {
                    black_box(&edit);
                } else {
                    if !edit.set_shape_text(target.slide, target.shape, EDIT_MARKER)? {
                        return Err("prefix target reported no change".into());
                    }
                    if rank == 3 {
                        black_box(&edit);
                    } else {
                        let commit = edit.commit()?;
                        if rank == 4 {
                            black_box(&commit);
                        } else {
                            let published = package.apply_opened_presentation_commit(commit)?;
                            black_box(&published);
                        }
                    }
                }
            }
        }
        black_box(iteration);
        completed = completed.saturating_add(1);
    }
    println!("completed\t{completed}");
    Ok(())
}

fn shape(source: &str) -> Fallible<()> {
    let archive = source_bytes(source)?;
    let package = Package::from_bytes(&archive)?;
    let semantic = semantic_digest(&package)?;
    let presentation = package.presentation()?;
    println!("probe\t{PROBE_ID}");
    println!("mode\tshape");
    println!("source\t{source}");
    println!("archive_bytes\t{}", archive.len());
    println!("archive_sha256\t{}", sha256(&archive));
    println!("slides\t{}", semantic.slides);
    println!("shapes\t{}", semantic.shapes);
    println!("semantic_payload_bytes\t{}", semantic.payload_bytes);
    println!("semantic_sha256\t{}", semantic.digest);
    for (index, slide) in presentation.slides()?.iter().enumerate() {
        println!(
            "slide\t{index}\t{}\t{}",
            slide.name()?,
            slide.shape_count()?
        );
    }
    Ok(())
}

fn target(source: &str) -> Fallible<()> {
    let archive = source_bytes(source)?;
    let target = derive_target(&archive)?;
    println!("probe\t{PROBE_ID}");
    println!("mode\ttarget");
    println!("source\t{source}");
    println!("archive_bytes\t{}", archive.len());
    println!("archive_sha256\t{}", sha256(&archive));
    println!("target\tslide:{}\tshape:{}", target.slide, target.shape);
    Ok(())
}

fn parse_count(value: &str, label: &str, maximum: usize) -> Fallible<usize> {
    parse_bounded(value, label, maximum)
}

fn usage() -> &'static str {
    "usage: probe0702 phases <source|generated:SxB> <samples> [warmups [one|noop|two]]\n       probe0702 noop <source|generated:SxB> <samples> [warmups]\n       probe0702 two <source|generated:SxB> <samples> [warmups]\n       probe0702 prefix <source|generated:SxB> <open|capture|transaction|settext|commit|apply> <iterations>\n       probe0702 shape <source|generated:SxB>\n       probe0702 target <source|generated:SxB>"
}

fn workflow(workflow: &str, source: &str, samples: usize, warmups: usize) -> Fallible<()> {
    match workflow {
        "one" => phases(source, samples, warmups),
        "noop" => noop(source, samples, warmups),
        "two" => two_edit(source, samples, warmups),
        other => {
            Err(format!("unknown phases workflow {other:?}; expected one, noop, or two").into())
        },
    }
}

fn main() -> Fallible<()> {
    let arguments: Vec<String> = env::args().skip(1).collect();
    match arguments.as_slice() {
        [mode, source, samples] if mode == "phases" => {
            let samples = parse_count(samples, "samples", MAX_SAMPLES)?;
            phases(source, samples, 0)
        },
        [mode, source, samples, warmups] if mode == "phases" => {
            let samples = parse_count(samples, "samples", MAX_SAMPLES)?;
            let warmups = parse_count(warmups, "warmups", MAX_WARMUPS)?;
            phases(source, samples, warmups)
        },
        [mode, source, samples, warmups, selected] if mode == "phases" => {
            let samples = parse_count(samples, "samples", MAX_SAMPLES)?;
            let warmups = parse_count(warmups, "warmups", MAX_WARMUPS)?;
            workflow(selected, source, samples, warmups)
        },
        [mode, source, samples] if mode == "noop" || mode == "two" => {
            let samples = parse_count(samples, "samples", MAX_SAMPLES)?;
            workflow(mode, source, samples, 0)
        },
        [mode, source, samples, warmups] if mode == "noop" || mode == "two" => {
            let samples = parse_count(samples, "samples", MAX_SAMPLES)?;
            let warmups = parse_count(warmups, "warmups", MAX_WARMUPS)?;
            workflow(mode, source, samples, warmups)
        },
        [mode, source, stage, iterations] if mode == "prefix" => {
            let iterations = parse_count(iterations, "prefix iterations", MAX_PREFIX_ITERS)?;
            prefix(source, stage, iterations)
        },
        [mode, source] if mode == "shape" => shape(source),
        [mode, source] if mode == "target" => target(source),
        _ => Err(usage().into()),
    }
}
