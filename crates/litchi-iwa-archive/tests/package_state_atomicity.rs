use std::sync::Arc;
use std::sync::atomic::{AtomicUsize, Ordering};

use litchi_iwa_archive::package::{GetOrInsertError, PackageState};
use litchi_iwa_cache::{CacheError, WeightError};
use litchi_iwa_core::{Archive, ArchiveLimits};
use litchi_iwa_package::{Entry, Error as EntryStoreError};

fn state() -> PackageState {
    PackageState::from_entries(
        vec![
            Entry::new("Index/Document.iwa".to_owned(), vec![1]),
            Entry::new("Index/Metadata.iwa".to_owned(), vec![2]),
        ],
        ArchiveLimits::default(),
    )
    .unwrap_or_else(|error| panic!("test package state should be valid: {error}"))
}

fn parse(state: &PackageState, name: &str) -> Arc<Archive> {
    state
        .get_or_parse_archive(name, |_| Ok((Archive::default(), 1)))
        .unwrap_or_else(|error| panic!("test archive should parse: {error}"))
}

#[test]
fn structural_ingress_is_atomic_and_keeps_unaffected_cache_entries() {
    let source = state();
    let document = parse(&source, "Index/Document.iwa");
    let metadata = parse(&source, "Index/Metadata.iwa");

    let mut edited = source.clone();

    assert_eq!(
        edited.try_insert_entry_at(3, Entry::new("Index/New.iwa".to_owned(), vec![3]),),
        Err(EntryStoreError::InvalidPosition {
            position: 3,
            len: 2
        })
    );
    assert_eq!(edited.len(), 2);
    assert_eq!(edited.position("Index/New.iwa"), None);

    assert_eq!(
        edited.try_insert_entry_at(1, Entry::new("Index/Document.iwa".to_owned(), vec![9]),),
        Err(EntryStoreError::DuplicateEntry(
            "Index/Document.iwa".to_owned(),
        ))
    );
    assert_eq!(edited.len(), 2);
    assert_eq!(edited.get_at(0).map(Entry::data), Some([1].as_slice()));
    assert_eq!(edited.get_at(1).map(Entry::data), Some([2].as_slice()));

    let document_parse_count = AtomicUsize::new(0);
    let retained_document = edited
        .get_or_parse_archive("Index/Document.iwa", |_| {
            document_parse_count.fetch_add(1, Ordering::SeqCst);
            Ok((Archive::default(), 1))
        })
        .unwrap_or_else(|error| panic!("retained document should parse: {error}"));
    assert!(Arc::ptr_eq(&document, &retained_document));
    assert_eq!(document_parse_count.load(Ordering::SeqCst), 0);

    edited
        .try_insert_entry_at(1, Entry::new("Index/New.iwa".to_owned(), vec![3]))
        .unwrap_or_else(|error| panic!("valid insertion should succeed: {error}"));
    assert_eq!(edited.len(), 3);
    assert_eq!(
        edited.iter().map(Entry::name).collect::<Vec<_>>(),
        ["Index/Document.iwa", "Index/New.iwa", "Index/Metadata.iwa"]
    );

    let metadata_parse_count = AtomicUsize::new(0);
    let retained_metadata = edited
        .get_or_parse_archive("Index/Metadata.iwa", |_| {
            metadata_parse_count.fetch_add(1, Ordering::SeqCst);
            Ok((Archive::default(), 1))
        })
        .unwrap_or_else(|error| panic!("retained metadata should parse: {error}"));
    assert!(Arc::ptr_eq(&metadata, &retained_metadata));
    assert_eq!(metadata_parse_count.load(Ordering::SeqCst), 0);

    let new_parse_count = AtomicUsize::new(0);
    let new_archive = edited
        .get_or_parse_archive("Index/New.iwa", |_| {
            new_parse_count.fetch_add(1, Ordering::SeqCst);
            Ok((Archive::default(), 1))
        })
        .unwrap_or_else(|error| panic!("new archive should parse: {error}"));
    assert_eq!(new_parse_count.load(Ordering::SeqCst), 1);

    let cached_new = edited
        .get_or_parse_archive("Index/New.iwa", |_| {
            panic!("a completed new archive should be served from the cache")
        })
        .unwrap_or_else(|error| panic!("cached new archive should succeed: {error}"));
    assert!(Arc::ptr_eq(&new_archive, &cached_new));
}

#[test]
fn successful_publication_and_failed_publication_preserve_the_source() {
    let source = state();
    let source_document = parse(&source, "Index/Document.iwa");

    let mut edited = source.clone();
    assert_eq!(
        edited.replace_entry_data("Index/Document.iwa", vec![9]),
        Some(vec![1])
    );
    let patch = source.patch_to(&edited);

    let mut stale = state();
    stale
        .entry_data_mut("Index/Document.iwa")
        .unwrap_or_else(|| panic!("stale document should exist"))[0] = 7;
    assert!(matches!(
        stale.apply_patch(&patch),
        Err(EntryStoreError::PatchSourceMismatch)
    ));

    // A rejected publication must not evict or replace the source cache.
    let still_source = source
        .get_or_parse_archive("Index/Document.iwa", |_| {
            panic!("failed publication must not disturb the source cache")
        })
        .unwrap_or_else(|error| panic!("source cache should remain usable: {error}"));
    assert!(Arc::ptr_eq(&source_document, &still_source));
    assert_eq!(
        source.get("Index/Document.iwa").map(Entry::data),
        Some([1].as_slice())
    );

    let published = source
        .apply_patch(&patch)
        .unwrap_or_else(|error| panic!("matching publication should succeed: {error}"));
    assert_eq!(
        published.get("Index/Document.iwa").map(Entry::data),
        Some([9].as_slice())
    );

    // The target owns a detached entry generation; changing it cannot mutate
    // the source snapshot or its cached parsed archive.
    let mut changed_target = published;
    assert_eq!(
        changed_target.replace_entry_data("Index/Document.iwa", vec![4]),
        Some(vec![9])
    );
    assert_eq!(
        source.get("Index/Document.iwa").map(Entry::data),
        Some([1].as_slice())
    );
    let source_again = source
        .get_or_parse_archive("Index/Document.iwa", |_| {
            panic!("target mutation must not evict the source cache")
        })
        .unwrap_or_else(|error| panic!("source cache should remain usable: {error}"));
    assert!(Arc::ptr_eq(&source_document, &source_again));
}

#[test]
fn patch_publication_forks_cache_and_invalidates_only_changed_names() {
    let source = state();
    let document = parse(&source, "Index/Document.iwa");
    let metadata = parse(&source, "Index/Metadata.iwa");

    let mut edited = source.clone();
    assert_eq!(
        edited.replace_entry_data("Index/Metadata.iwa", vec![8]),
        Some(vec![2])
    );
    let patch = source.patch_to(&edited);
    let published = source
        .apply_patch(&patch)
        .unwrap_or_else(|error| panic!("matching patch should publish: {error}"));

    let document_parse_count = AtomicUsize::new(0);
    let retained_document = published
        .get_or_parse_archive("Index/Document.iwa", |_| {
            document_parse_count.fetch_add(1, Ordering::SeqCst);
            Ok((Archive::default(), 1))
        })
        .unwrap_or_else(|error| panic!("unchanged document should parse: {error}"));
    assert!(Arc::ptr_eq(&document, &retained_document));
    assert_eq!(document_parse_count.load(Ordering::SeqCst), 0);

    let metadata_parse_count = AtomicUsize::new(0);
    let replacement_metadata = published
        .get_or_parse_archive("Index/Metadata.iwa", |_| {
            metadata_parse_count.fetch_add(1, Ordering::SeqCst);
            Ok((Archive::default(), 1))
        })
        .unwrap_or_else(|error| panic!("changed metadata should parse: {error}"));
    assert!(!Arc::ptr_eq(&metadata, &replacement_metadata));
    assert_eq!(metadata_parse_count.load(Ordering::SeqCst), 1);

    // Publishing the target must not evict the source generation's cache.
    let source_document = source
        .get_or_parse_archive("Index/Document.iwa", |_| {
            panic!("source document cache should remain populated")
        })
        .unwrap_or_else(|error| panic!("source document cache should remain usable: {error}"));
    let source_metadata = source
        .get_or_parse_archive("Index/Metadata.iwa", |_| {
            panic!("source metadata cache should remain populated")
        })
        .unwrap_or_else(|error| panic!("source metadata cache should remain usable: {error}"));
    assert!(Arc::ptr_eq(&document, &source_document));
    assert!(Arc::ptr_eq(&metadata, &source_metadata));
}

#[test]
fn patch_publication_does_not_import_unaffected_source_flights() {
    let source = Arc::new(state());
    let source_started = Arc::new(std::sync::Barrier::new(2));
    let release_source = Arc::new(std::sync::Barrier::new(2));
    let source_thread = {
        let source = Arc::clone(&source);
        let source_started = Arc::clone(&source_started);
        let release_source = Arc::clone(&release_source);
        std::thread::spawn(move || {
            source.get_or_parse_archive("Index/Document.iwa", |_| {
                source_started.wait();
                release_source.wait();
                Ok((Archive::default(), 1))
            })
        })
    };

    source_started.wait();

    let mut edited = (*source).clone();
    assert_eq!(
        edited.replace_entry_data("Index/Metadata.iwa", vec![8]),
        Some(vec![2])
    );
    let patch = source.patch_to(&edited);
    let published = source
        .apply_patch(&patch)
        .unwrap_or_else(|error| panic!("matching patch should publish: {error}"));

    let target_parse_count = AtomicUsize::new(0);
    let target_document = published
        .get_or_parse_archive("Index/Document.iwa", |_| {
            target_parse_count.fetch_add(1, Ordering::SeqCst);
            Ok((Archive::default(), 1))
        })
        .unwrap_or_else(|error| panic!("target should start its own unaffected parse: {error}"));
    assert_eq!(target_parse_count.load(Ordering::SeqCst), 1);

    release_source.wait();
    let source_document = source_thread
        .join()
        .unwrap_or_else(|_| panic!("source parser thread should not panic"))
        .unwrap_or_else(|error| panic!("source parser should succeed: {error}"));
    assert!(!Arc::ptr_eq(&target_document, &source_document));
}

#[test]
fn exact_noop_reuses_completed_cache_and_inverse_restores_ordered_bytes() {
    let source = state();
    let source_document = parse(&source, "Index/Document.iwa");

    let mut same_bytes = source.clone();
    assert_eq!(
        same_bytes.replace_entry_data("Index/Document.iwa", vec![1]),
        Some(vec![1])
    );
    let retained_after_replace = same_bytes
        .get_or_parse_archive("Index/Document.iwa", |_| {
            panic!("an exact replacement should retain the parsed archive")
        })
        .unwrap_or_else(|error| panic!("exact replacement cache lookup should succeed: {error}"));
    assert!(Arc::ptr_eq(&source_document, &retained_after_replace));

    let no_op = source.patch_to(&source);
    assert!(no_op.is_empty());
    assert!(no_op.inverse().is_empty());
    let no_op_result = source
        .apply_patch(&no_op)
        .unwrap_or_else(|error| panic!("exact no-op should apply: {error}"));
    let retained = no_op_result
        .get_or_parse_archive("Index/Document.iwa", |_| {
            panic!("an exact no-op should retain completed cache values")
        })
        .unwrap_or_else(|error| panic!("no-op cache lookup should succeed: {error}"));
    assert!(Arc::ptr_eq(&source_document, &retained));

    let mut target = source.clone();
    assert_eq!(
        target.replace_entry_data("Index/Metadata.iwa", vec![8, 9]),
        Some(vec![2])
    );
    let forward = source.patch_to(&target);
    let inverse = forward.inverse();
    assert_eq!(inverse.len(), forward.len());
    assert_eq!(inverse.inverse().changes(), forward.changes());

    let published = source
        .apply_patch(&forward)
        .unwrap_or_else(|error| panic!("forward patch should apply: {error}"));
    let restored = published
        .apply_patch(&inverse)
        .unwrap_or_else(|error| panic!("inverse patch should apply to its target: {error}"));
    assert_eq!(
        restored.iter().map(Entry::name).collect::<Vec<_>>(),
        source.iter().map(Entry::name).collect::<Vec<_>>()
    );
    assert_eq!(
        restored.iter().map(Entry::data).collect::<Vec<_>>(),
        source.iter().map(Entry::data).collect::<Vec<_>>()
    );
}

#[test]
fn reordered_patch_retains_name_keyed_archive_cache() {
    let source = state();
    let source_document = parse(&source, "Index/Document.iwa");
    let source_metadata = parse(&source, "Index/Metadata.iwa");

    let mut reordered = source.clone();
    let document = reordered
        .remove_entry("Index/Document.iwa")
        .unwrap_or_else(|| panic!("document entry should exist"));
    reordered
        .try_insert_entry_at(1, document)
        .unwrap_or_else(|error| panic!("reinserted document should be accepted: {error}"));
    assert_eq!(
        reordered.iter().map(Entry::name).collect::<Vec<_>>(),
        ["Index/Metadata.iwa", "Index/Document.iwa"]
    );

    let patch = source.patch_to(&reordered);
    assert_eq!(patch.len(), 2);
    assert!(
        patch
            .changes()
            .iter()
            .all(|change| change.kind() == litchi_iwa_package::EntryChangeKind::Reordered)
    );
    let published = source
        .apply_patch(&patch)
        .unwrap_or_else(|error| panic!("reorder patch should publish: {error}"));

    let retained_document = published
        .get_or_parse_archive("Index/Document.iwa", |_| {
            panic!("reordering must retain the document archive cache")
        })
        .unwrap_or_else(|error| panic!("retained document cache should be usable: {error}"));
    let retained_metadata = published
        .get_or_parse_archive("Index/Metadata.iwa", |_| {
            panic!("reordering must retain the metadata archive cache")
        })
        .unwrap_or_else(|error| panic!("retained metadata cache should be usable: {error}"));
    assert!(Arc::ptr_eq(&source_document, &retained_document));
    assert!(Arc::ptr_eq(&source_metadata, &retained_metadata));
}

#[test]
fn mixed_reorder_patch_retains_only_unchanged_name_caches() {
    let source = state();
    let source_document = parse(&source, "Index/Document.iwa");
    let source_metadata = parse(&source, "Index/Metadata.iwa");

    let mut target = source.clone();
    assert_eq!(
        target.replace_entry_data("Index/Metadata.iwa", vec![8]),
        Some(vec![2])
    );
    let document = target
        .remove_entry("Index/Document.iwa")
        .unwrap_or_else(|| panic!("document entry should exist"));
    target
        .try_insert_entry_at(1, document)
        .unwrap_or_else(|error| panic!("reinserted document should be accepted: {error}"));

    let patch = source.patch_to(&target);
    assert_eq!(patch.len(), 2);
    assert_eq!(
        patch
            .changes()
            .iter()
            .map(|change| (change.name(), change.kind()))
            .collect::<Vec<_>>(),
        vec![
            (
                "Index/Document.iwa",
                litchi_iwa_package::EntryChangeKind::Reordered
            ),
            (
                "Index/Metadata.iwa",
                litchi_iwa_package::EntryChangeKind::Replaced
            ),
        ]
    );
    let published = source
        .apply_patch(&patch)
        .unwrap_or_else(|error| panic!("mixed patch should publish: {error}"));
    assert_eq!(
        published.iter().map(Entry::name).collect::<Vec<_>>(),
        ["Index/Metadata.iwa", "Index/Document.iwa"]
    );

    let retained_document = published
        .get_or_parse_archive("Index/Document.iwa", |_| {
            panic!("reordered unchanged entry must retain its archive cache")
        })
        .unwrap_or_else(|error| panic!("retained document cache should be usable: {error}"));
    assert!(Arc::ptr_eq(&source_document, &retained_document));

    let metadata_parse_count = AtomicUsize::new(0);
    let replacement_metadata = published
        .get_or_parse_archive("Index/Metadata.iwa", |_| {
            metadata_parse_count.fetch_add(1, Ordering::SeqCst);
            Ok((Archive::default(), 1))
        })
        .unwrap_or_else(|error| panic!("replaced metadata should parse: {error}"));
    assert!(!Arc::ptr_eq(&source_metadata, &replacement_metadata));
    assert_eq!(metadata_parse_count.load(Ordering::SeqCst), 1);
}

#[test]
fn inverse_mixed_patch_retains_reordered_cache_and_reparses_replacement() {
    let source = state();
    let source_document = parse(&source, "Index/Document.iwa");
    let source_metadata = parse(&source, "Index/Metadata.iwa");

    let mut target = source.clone();
    assert_eq!(
        target.replace_entry_data("Index/Metadata.iwa", vec![8]),
        Some(vec![2])
    );
    let document = target
        .remove_entry("Index/Document.iwa")
        .unwrap_or_else(|| panic!("document entry should exist"));
    target
        .try_insert_entry_at(1, document)
        .unwrap_or_else(|error| panic!("reinserted document should be accepted: {error}"));

    let patch = source.patch_to(&target);
    let published = source
        .apply_patch(&patch)
        .unwrap_or_else(|error| panic!("mixed patch should publish: {error}"));
    let published_document = published
        .get_or_parse_archive("Index/Document.iwa", |_| {
            panic!("forward reorder must retain the document archive cache")
        })
        .unwrap_or_else(|error| panic!("published document cache should be usable: {error}"));
    assert!(Arc::ptr_eq(&source_document, &published_document));

    let published_metadata = published
        .get_or_parse_archive("Index/Metadata.iwa", |_| Ok((Archive::default(), 1)))
        .unwrap_or_else(|error| panic!("published metadata should parse: {error}"));
    assert!(!Arc::ptr_eq(&source_metadata, &published_metadata));

    let restored = published
        .apply_patch(&patch.inverse())
        .unwrap_or_else(|error| panic!("inverse mixed patch should publish: {error}"));
    assert_eq!(
        restored.iter().map(Entry::name).collect::<Vec<_>>(),
        ["Index/Document.iwa", "Index/Metadata.iwa"]
    );

    let restored_document = restored
        .get_or_parse_archive("Index/Document.iwa", |_| {
            panic!("inverse reorder must retain the document archive cache")
        })
        .unwrap_or_else(|error| panic!("restored document cache should be usable: {error}"));
    assert!(Arc::ptr_eq(&source_document, &restored_document));

    let restored_metadata_parse_count = AtomicUsize::new(0);
    let restored_metadata = restored
        .get_or_parse_archive("Index/Metadata.iwa", |_| {
            restored_metadata_parse_count.fetch_add(1, Ordering::SeqCst);
            Ok((Archive::default(), 1))
        })
        .unwrap_or_else(|error| panic!("restored metadata should parse: {error}"));
    assert!(!Arc::ptr_eq(&published_metadata, &restored_metadata));
    assert!(!Arc::ptr_eq(&source_metadata, &restored_metadata));
    assert_eq!(
        restored_metadata_parse_count.load(Ordering::SeqCst),
        1,
        "inverse replacement must detach the forward-generation cache"
    );
}

#[test]
fn speculative_and_oversized_archives_are_rejected_with_typed_limits() {
    let limits = ArchiveLimits::default()
        .with_archive_bytes(4)
        .unwrap_or_else(|error| panic!("small archive limits should be valid: {error}"));
    let mut state = PackageState::from_entries(Vec::new(), limits)
        .unwrap_or_else(|error| panic!("empty package state should be valid: {error}"));

    // The low-level state permits parsing before structural insertion, but a
    // successful insertion must detach that speculative value.
    let stale = state
        .get_or_parse_archive("Index/Later.iwa", |_| Ok((Archive::default(), 1)))
        .unwrap_or_else(|error| panic!("speculative parse should succeed: {error}"));
    state
        .try_insert_entry_at(0, Entry::new("Index/Later.iwa".to_owned(), vec![1]))
        .unwrap_or_else(|error| panic!("entry insertion should succeed: {error}"));
    let parse_count = AtomicUsize::new(0);
    let fresh = state
        .get_or_parse_archive("Index/Later.iwa", |_| {
            parse_count.fetch_add(1, Ordering::SeqCst);
            Ok((Archive::default(), 1))
        })
        .unwrap_or_else(|error| panic!("inserted entry should parse: {error}"));
    assert_eq!(parse_count.load(Ordering::SeqCst), 1);
    assert!(!Arc::ptr_eq(&stale, &fresh));

    let zero_weight = state.get_or_parse_archive("Index/Later.iwa", |_| {
        // The completed value is already retained, so this closure should not
        // run; use a different key for the typed rejection below.
        panic!("cached entry should be returned before parsing")
    });
    assert!(zero_weight.is_ok());

    let zero_weight = state.get_or_parse_archive("Index/Zero.iwa", |_| Ok((Archive::default(), 0)));
    assert!(matches!(
        zero_weight,
        Err(GetOrInsertError::Cache(CacheError::Weight(
            WeightError::ZeroWeight
        )))
    ));

    let too_large = limits.max_archive_bytes() + 1;
    let oversized = state.get_or_parse_archive("Index/Oversized.iwa", |_| {
        Ok((Archive::default(), too_large))
    });
    assert!(matches!(
        oversized,
        Err(GetOrInsertError::Cache(CacheError::Weight(
            WeightError::ExceedsCapacity {
                weight,
                capacity,
            }
        ))) if weight == too_large && capacity == limits.max_archive_bytes()
    ));
}
