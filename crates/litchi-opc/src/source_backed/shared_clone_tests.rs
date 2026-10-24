//! Internal checks of what a [`SourceBackedPackage`] clone shares (record
//! 0012).
//!
//! The public integration tests in `tests/source_backed_shared_clones.rs`
//! cover reads, publication and source fences through the public API. These
//! tests need private access: they compare the shared allocations by pointer,
//! follow the budget of a managed package across clones, and run concurrent
//! compression waves on the one write session that every clone shares.

#![allow(
    clippy::expect_used,
    clippy::panic,
    clippy::unwrap_used,
    reason = "focused internal sharing tests deliberately panic on fixture errors"
)]

use super::*;
use crate::SourceCacheLimits;
use litchi_core::{Budget, CancellationSource, ExecutionLimits, Limits};
use soapberry_zip::office::StreamingArchiveWriter;
use std::num::{NonZeroU64, NonZeroUsize};

const RELATIONSHIPS: &[u8] = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
const CONTENT_TYPES: &[u8] = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#;

fn uri(name: &str) -> PackURI {
    PackURI::new(name).expect("fixture URI must be valid")
}

/// Deterministic XML-shaped payload; distinct seeds give distinct bytes.
fn payload(seed: usize, bytes: usize) -> Vec<u8> {
    let mut data = Vec::with_capacity(bytes + 32);
    data.extend_from_slice(b"<p>");
    let mut state = (seed as u64).wrapping_mul(0x9e37_79b9_7f4a_7c15).max(1);
    while data.len() < bytes {
        state ^= state << 13;
        state ^= state >> 7;
        state ^= state << 17;
        data.extend_from_slice(format!("<c r=\"{}\"/>", state % 8192).as_bytes());
    }
    data.extend_from_slice(b"</p>");
    data
}

/// A main document plus `parts` custom XML parts, all deflated, so that a
/// publication replacing them regenerates deflate members.
fn package_bytes(parts: usize, bytes: usize) -> Vec<u8> {
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    writer
        .write_stored("[Content_Types].xml", CONTENT_TYPES)
        .unwrap();
    writer.write_stored("_rels/.rels", RELATIONSHIPS).unwrap();
    writer
        .write_deflated("word/document.xml", &payload(1, bytes))
        .unwrap();
    for index in 0..parts {
        writer
            .write_deflated(
                &format!("custom/part-{index:04}.xml"),
                &payload(index + 2, bytes),
            )
            .unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn replacements(parts: usize, bytes: usize) -> Vec<(PackURI, Arc<Vec<u8>>)> {
    (0..parts)
        .map(|index| {
            (
                uri(&format!("/custom/part-{index:04}.xml")),
                Arc::new(payload(index + 100, bytes)),
            )
        })
        .collect()
}

fn context(workers: usize) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "opc-shared-clone-test",
        Limits::new(
            256 * 1024 * 1024,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(workers).unwrap(),
        NonZeroUsize::new(workers.max(64)).unwrap(),
        NonZeroU64::new(64 * 1024 * 1024).unwrap(),
        0,
    )
    .unwrap();
    let context = ExecutionContext::new(budget.clone(), cancellation, limits);
    (budget, cancellation_source, context)
}

fn managed(source: &Arc<dyn ReadAt>, context: ExecutionContext) -> SourceBackedPackage {
    SourceBackedPackage::from_read_at(
        Arc::clone(source),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap()
}

fn assert_shares_everything(original: &SourceBackedPackage, clone: &SourceBackedPackage) {
    let SourceBackedPackage {
        source,
        fingerprint_memo,
        archive,
        read_ahead,
        limits,
        catalog,
        cache,
        catalog_object_reservation,
    } = clone;
    assert!(Arc::ptr_eq(&original.source.source, &source.source));
    assert!(Arc::ptr_eq(
        &original.source.monitor_reads,
        &source.monitor_reads
    ));
    assert_eq!(original.source.lineage, source.lineage);
    assert_eq!(original.source.version, source.version);
    assert_eq!(original.source.length, source.length);
    match (&original.fingerprint_memo, fingerprint_memo) {
        (None, None) => {},
        (Some(left), Some(right)) => assert!(Arc::ptr_eq(left, right)),
        _ => panic!("a clone must keep the fingerprint memo of its package"),
    }
    assert!(Arc::ptr_eq(&original.archive, archive));
    assert!(Arc::ptr_eq(&original.catalog, catalog));
    assert!(Arc::ptr_eq(&original.cache, cache));
    assert_eq!(original.limits, *limits);
    match (&original.read_ahead, read_ahead) {
        (None, None) => {},
        (Some(left), Some(right)) => assert!(Arc::ptr_eq(left, right)),
        _ => panic!("a clone must keep the read-ahead window of its package"),
    }
    match (
        &original.catalog_object_reservation,
        catalog_object_reservation,
    ) {
        (None, None) => {},
        (Some(left), Some(right)) => assert!(Arc::ptr_eq(left, right)),
        _ => panic!("a clone must keep the catalog reservation of its package"),
    }
    assert!(original.is_same_opened_package(clone));
    assert!(clone.is_same_opened_package(original));
}

#[test]
fn owned_ingress_clone_retains_the_shared_fingerprint_after_original_drop() {
    let original = SourceBackedPackage::from_vec(
        package_bytes(3, 2 * 1024),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    let clone = original.clone();
    assert_shares_everything(&original, &clone);
    let memo = clone.fingerprint_memo.as_ref().unwrap();
    assert!(!memo.is_initialized());
    let fingerprint = original.source_artifact().fingerprint().unwrap();
    assert!(memo.is_initialized());
    drop(original);
    assert_eq!(clone.source_artifact().fingerprint().unwrap(), fingerprint);
}

#[test]
fn a_clone_shares_every_allocation_of_its_opened_package() {
    let bytes = package_bytes(3, 2 * 1024);
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(bytes));

    let unmanaged = SourceBackedPackage::from_read_at(
        Arc::clone(&source),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let counts = |package: &SourceBackedPackage| {
        [
            Arc::strong_count(&package.archive),
            Arc::strong_count(&package.catalog),
            Arc::strong_count(&package.cache),
            Arc::strong_count(&package.source.source),
            Arc::strong_count(&package.source.monitor_reads),
        ]
    };
    let alone = counts(&unmanaged);
    assert_eq!(alone[..3], [1, 1, 1], "the open holds the only references");
    let clone = unmanaged.clone();
    let clone_of_clone = clone.clone();
    assert_eq!(
        counts(&unmanaged),
        alone.map(|count| count + 2),
        "each clone adds one reference to each shared allocation and copies none"
    );
    assert_shares_everything(&unmanaged, &clone);
    assert_shares_everything(&unmanaged, &clone_of_clone);
    assert!(unmanaged.read_ahead.is_none());
    assert!(unmanaged.catalog_object_reservation.is_none());
    drop(clone_of_clone);
    assert_eq!(counts(&clone), alone.map(|count| count + 1));
    drop(clone);
    assert_eq!(counts(&unmanaged), alone);

    let windowed = SourceBackedPackage::from_read_at(
        Arc::clone(&source),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::forward_start(64 * 1024).unwrap(),
        None,
    )
    .unwrap();
    assert!(windowed.read_ahead.is_some());
    assert_shares_everything(&windowed, &windowed.clone());

    let (_budget, _cancellation, context) = context(1);
    let managed = managed(&source, context);
    assert!(managed.catalog_object_reservation.is_some());
    assert_shares_everything(&managed, &managed.clone());

    // A second open of the same source is a different package: a fresh
    // lineage, catalog, index and cache.
    let reopened = SourceBackedPackage::from_read_at(
        Arc::clone(&source),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    assert!(!unmanaged.is_same_opened_package(&reopened));
    assert_ne!(unmanaged.source_lineage(), reopened.source_lineage());
    assert!(!Arc::ptr_eq(&unmanaged.archive, &reopened.archive));
    assert!(!Arc::ptr_eq(&unmanaged.cache, &reopened.cache));
}

#[test]
fn managed_clones_charge_one_decode_and_release_with_the_last_clone() {
    let bytes = package_bytes(2, 4 * 1024);
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(bytes));
    let (budget, _cancellation, context) = context(1);
    let original = managed(&source, context);
    let catalog_objects = budget.used(Resource::Objects);
    assert!(catalog_objects > 0);
    let memory_after_open = budget.used(Resource::Memory);
    let clone = original.clone();
    assert_eq!(budget.used(Resource::Objects), catalog_objects);
    assert_eq!(budget.used(Resource::Memory), memory_after_open);

    let document = uri("/word/document.xml");
    let first = original.part(&document).unwrap().data().unwrap();
    let charged = budget.used(Resource::Memory);
    let charged_objects = budget.used(Resource::Objects);
    assert!(charged > memory_after_open);
    assert!(charged_objects > catalog_objects);
    let second = clone.part(&document).unwrap().data().unwrap();
    assert_eq!(first.as_bytes(), second.as_bytes());
    assert_eq!(
        (
            budget.used(Resource::Memory),
            budget.used(Resource::Objects)
        ),
        (charged, charged_objects),
        "a payload decoded through one clone is a cache hit for the other"
    );
    let diagnostics = clone.cache_diagnostics();
    assert_eq!(diagnostics.cold_loads, 1);
    assert_eq!(diagnostics.hits, 1);
    assert_eq!(diagnostics, original.cache_diagnostics());
    drop((first, second));

    drop(original);
    assert_eq!(
        (
            budget.used(Resource::Memory),
            budget.used(Resource::Objects)
        ),
        (charged, charged_objects),
        "the catalog and the cached entry stay charged while a clone holds them"
    );
    let again = clone.part(&document).unwrap().data().unwrap();
    assert_eq!(clone.cache_diagnostics().hits, 2);
    drop(again);

    drop(clone);
    assert_eq!(budget.used(Resource::Objects), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

const WAVE_PARTS: usize = 15;
const WAVE_BYTES: usize = 6 * 1024;
const WAVE_WORKERS: usize = 4;
const PUBLISHERS: usize = 4;

/// The fixture, its replacements, and the unmanaged sequential output.
fn wave_fixture() -> (Arc<dyn ReadAt>, Vec<(PackURI, Arc<Vec<u8>>)>, Vec<u8>) {
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(package_bytes(WAVE_PARTS, WAVE_BYTES)));
    let replacements = replacements(WAVE_PARTS, WAVE_BYTES);
    let mut expected = Vec::new();
    SourceBackedPackage::from_read_at(
        Arc::clone(&source),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap()
    .write_part_overlays_shared_to_stream(&mut expected, replacements.clone())
    .unwrap();
    (source, replacements, expected)
}

/// Publish the same replacements from `PUBLISHERS` clones on as many threads.
fn publish_from_clones_concurrently(
    package: &SourceBackedPackage,
    replacements: &[(PackURI, Arc<Vec<u8>>)],
) -> Vec<Vec<u8>> {
    std::thread::scope(|scope| {
        let publishers: Vec<_> = (0..PUBLISHERS)
            .map(|_| {
                let clone = package.clone();
                let replacements = replacements.to_vec();
                scope.spawn(move || {
                    let mut output = Vec::new();
                    clone
                        .write_part_overlays_shared_to_stream(&mut output, replacements)
                        .unwrap();
                    output
                })
            })
            .collect();
        publishers
            .into_iter()
            .map(|publisher| publisher.join().unwrap())
            .collect()
    })
}

#[test]
fn clones_run_compression_waves_concurrently_on_one_write_session() {
    let (source, replacements, expected) = wave_fixture();
    let (budget, _cancellation, context) = context(WAVE_WORKERS);
    let package = managed(&source, context);
    for output in publish_from_clones_concurrently(&package, &replacements) {
        assert_eq!(output, expected, "a wave's width never changes the bytes");
    }
    // One CPU task per regenerated member in each publication.
    assert_eq!(
        budget.used(Resource::CpuTasks),
        (PUBLISHERS * WAVE_PARTS) as u64
    );
    // The pool and the permits that admitted it belong to the package, not to
    // one publication: they stay while any clone is alive. Concurrent waves
    // narrow to the held width instead of reserving more.
    let held_workers = budget.used(Resource::Workers);
    assert_eq!(held_workers, WAVE_WORKERS as u64);
    let mut again = Vec::new();
    package
        .clone()
        .write_part_overlays_shared_to_stream(&mut again, replacements)
        .unwrap();
    assert_eq!(again, expected);
    assert_eq!(budget.used(Resource::Workers), held_workers);

    drop(package);
    assert_eq!(budget.used(Resource::Workers), 0);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[derive(Debug, Default)]
struct CountingCallerWorkers {
    waves: AtomicU64,
    tasks: AtomicU64,
}

impl litchi_core::ScopedWorkers for CountingCallerWorkers {
    fn run_all<'task>(&self, tasks: &mut [&mut (dyn FnMut() + Send + 'task)]) {
        self.waves.fetch_add(1, Ordering::Relaxed);
        self.tasks.fetch_add(tasks.len() as u64, Ordering::Relaxed);
        for task in tasks.iter_mut() {
            task();
        }
    }
}

#[test]
fn clones_run_caller_facility_waves_concurrently_and_release_each_wave() {
    let (source, replacements, expected) = wave_fixture();
    let (budget, _cancellation, context) = context(WAVE_WORKERS);
    let workers = Arc::new(CountingCallerWorkers::default());
    let facility: Arc<dyn litchi_core::ScopedWorkers> = workers.clone();
    let package = managed(&source, context.with_scoped_workers(facility));
    for output in publish_from_clones_concurrently(&package, &replacements) {
        assert_eq!(output, expected);
    }
    // Every publication ran its own wave on the caller's facility, and a
    // caller-facility wave holds its permits only while it runs.
    assert_eq!(workers.waves.load(Ordering::Relaxed), PUBLISHERS as u64);
    assert_eq!(
        workers.tasks.load(Ordering::Relaxed),
        (PUBLISHERS * WAVE_PARTS) as u64
    );
    assert_eq!(budget.used(Resource::Workers), 0);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn a_publication_from_any_clone_fences_the_shared_read_ahead_window() {
    let bytes = package_bytes(2, 2 * 1024);
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(bytes.clone()));
    let original = SourceBackedPackage::from_read_at(
        source,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::forward_start(64 * 1024).unwrap(),
        None,
    )
    .unwrap();
    let document = uri("/word/document.xml");
    let before = original.part(&document).unwrap().data().unwrap();
    let clone = original.clone();
    assert!(
        original
            .source_read_diagnostics()
            .unwrap()
            .expect("forward policy")
            .enabled
    );

    let mut output = Vec::new();
    clone
        .clone()
        .write_part_overlays_to_stream(&mut output, Vec::new())
        .unwrap();
    assert_eq!(
        output, bytes,
        "an empty overlay set is an exact source copy"
    );

    for package in [&original, &clone] {
        let diagnostics = package
            .source_read_diagnostics()
            .unwrap()
            .expect("the window keeps its counters");
        assert!(!diagnostics.enabled);
        assert_eq!(diagnostics.retained_window_bytes, 0);
    }
    assert_eq!(
        original.source_read_diagnostics().unwrap(),
        clone.source_read_diagnostics().unwrap()
    );
    // Exact reads return the same bytes the window returned.
    let after = clone.part(&document).unwrap().data().unwrap();
    assert_eq!(before.as_bytes(), after.as_bytes());
    let uncached = clone
        .part(&uri("/custom/part-0001.xml"))
        .unwrap()
        .data()
        .unwrap();
    assert_eq!(uncached.as_bytes(), payload(3, 2 * 1024).as_slice());
}
