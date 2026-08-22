//! Shared indexed storage for immutable iWork package snapshots.
//!
//! The physical package owner retains the ordered entry table together with a
//! bounded cache of parsed IWA components. Higher-level package facades retain
//! their own validation and error mapping, while this module keeps cache
//! lifetime and copy-on-write generation behavior alongside the physical
//! package data it protects.

use std::sync::Arc;

use litchi_iwa_cache::WeightedCache;
pub use litchi_iwa_cache::{GetOrInsertError, ParseError};
use litchi_iwa_core::{Archive, ArchiveLimits};
use litchi_iwa_package::{Entry, EntryStore, Error as EntryStoreError, Patch};

/// Package entries plus an index and bounded parsed-component cache.
///
/// The state is intended for copy-on-write package owners. Structural edits
/// clone the entry store once, then rebuild its small name index; read-only
/// clones keep sharing entries, payloads, and completed parsed components.
/// Active parser flights remain with their source generation so detached state
/// can never observe a parse of stale component bytes.
#[derive(Debug)]
pub struct PackageState {
    entries: EntryStore,
    archive_limits: ArchiveLimits,
    parsed_archive: WeightedCache<String, Archive>,
}

impl Default for PackageState {
    fn default() -> Self {
        Self::from_store(EntryStore::default(), ArchiveLimits::default())
    }
}

impl Clone for PackageState {
    fn clone(&self) -> Self {
        Self {
            entries: self.entries.clone(),
            archive_limits: self.archive_limits,
            // Completed values share immutable Arcs; active flights are tied
            // to the source generation's exact entry bytes.
            parsed_archive: self.parsed_archive.fork(),
        }
    }
}

impl PackageState {
    /// Build state from ordered entries after checking the name index.
    ///
    /// # Errors
    ///
    /// Returns an entry-store error when names are duplicated or the index
    /// cannot reserve storage.
    pub fn from_entries(
        entries: Vec<Entry>,
        archive_limits: ArchiveLimits,
    ) -> Result<Self, EntryStoreError> {
        let entry_store = EntryStore::try_from_entries(entries)?;
        Ok(Self::from_store(entry_store, archive_limits))
    }

    /// Build state from a previously validated entry store.
    #[must_use]
    fn from_store(entries: EntryStore, archive_limits: ArchiveLimits) -> Self {
        Self {
            entries,
            archive_limits,
            parsed_archive: new_archive_cache(archive_limits),
        }
    }

    /// Build the reversible entry patch between two package-state snapshots.
    ///
    /// The entry table and its parsed-component cache are one ownership unit.
    /// Keeping patch construction here prevents a facade from comparing one
    /// generation's entries while publishing another generation's cache.
    #[must_use]
    pub fn patch_to(&self, target: &Self) -> Patch {
        Patch::between(&self.entries, &target.entries)
    }

    /// Apply a source-checked patch and publish a fresh cache generation.
    ///
    /// Patch validation happens before the target state is constructed. A
    /// stale source therefore has no observable effect, and a successful
    /// target never inherits parsed archives for a different entry generation.
    ///
    /// # Errors
    ///
    /// Returns [`EntryStoreError::PatchSourceMismatch`] when the patch was not
    /// created from this state's exact ordered entry table and payloads.
    pub fn apply_patch(&self, patch: &Patch) -> Result<Self, EntryStoreError> {
        let entries = patch.apply(&self.entries)?;
        // A patch publication is a copy-on-write generation just like a
        // facade edit through `Arc::make_mut`: fork completed values so
        // unchanged names keep their immutable `Arc`s, then detach only
        // names whose entry bytes/topology changed. `fork` intentionally
        // leaves active flights with the source generation, so a target can
        // never publish a parse that observed the source's old bytes.
        let mut target = self.clone();
        target.entries = entries;
        for change in patch.changes() {
            // Reordering does not change the bytes associated with a name;
            // its parsed archive remains valid and is therefore retained.
            if !matches!(
                change.kind(),
                litchi_iwa_package::EntryChangeKind::Reordered
            ) {
                target.invalidate_archive(change.name());
            }
        }
        Ok(target)
    }

    /// Return the number of ordered package members.
    #[must_use]
    pub fn len(&self) -> usize {
        self.entries.len()
    }

    /// Return whether this package state has no members.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    /// Iterate over ordered package members without exposing the entry store.
    pub fn iter(&self) -> impl Iterator<Item = &Entry> + Clone {
        self.entries.iter()
    }

    /// Enumerate package members that contain object-oriented IWA archives.
    ///
    /// Legacy `OperationStorage.iwa` members use a separate `bvxn` operation
    /// log format. They remain in the raw package state but are excluded from
    /// object-archive scans at this physical package boundary.
    pub fn iwa_entry_names(&self) -> impl Iterator<Item = &str> {
        self.entries
            .iter()
            .filter(|entry| {
                entry.name().ends_with(".iwa")
                    && !is_legacy_operation_storage(entry.name(), entry.data())
            })
            .map(Entry::name)
    }

    /// Borrow one exact-name package member without exposing the entry store.
    #[must_use]
    pub fn get(&self, name: &str) -> Option<&Entry> {
        self.entries.get(name)
    }

    /// Borrow one package member by its preserved order.
    #[must_use]
    pub fn get_at(&self, position: usize) -> Option<&Entry> {
        self.entries.get_at(position)
    }

    /// Look up the position of one exact package entry name.
    #[must_use]
    pub fn position(&self, name: &str) -> Option<usize> {
        self.entries.position(name)
    }

    /// Return a cached parsed archive or parse it once for concurrent callers.
    ///
    /// Parsing runs without the cache mutex. Concurrent requests for the same
    /// entry share a single parser invocation, while failures wake all waiters
    /// and remain retryable by a later request.
    ///
    /// # Errors
    ///
    /// Returns cache or parser errors without imposing a facade-specific
    /// error contract.
    pub fn get_or_parse_archive<F>(
        &self,
        name: &str,
        parse: F,
    ) -> Result<Arc<Archive>, GetOrInsertError>
    where
        F: FnOnce(ArchiveLimits) -> Result<(Archive, usize), ParseError>,
    {
        self.parsed_archive.get_or_try_insert_with_weight_borrowed(
            name,
            || name.to_owned(),
            || parse(self.archive_limits),
        )
    }

    /// Borrow one entry's payload for an owner-controlled mutation.
    ///
    /// The matching parsed component is evicted before the payload is exposed,
    /// so a caller cannot retain a stale parse after changing the bytes.
    pub fn entry_data_mut(&mut self, name: &str) -> Option<&mut Vec<u8>> {
        let position = self.position(name)?;
        self.invalidate_archive(name);
        self.entries.get_at_mut(position).map(Entry::data_mut)
    }

    /// Replace an entry payload and evict the matching parsed component.
    /// Exact byte-identical replacements are treated as no-ops and retain
    /// their existing parsed component.
    pub fn replace_entry_data(&mut self, name: &str, data: Vec<u8>) -> Option<Vec<u8>> {
        let position = self.position(name)?;
        if self
            .entries
            .get_at(position)
            .is_some_and(|entry| entry.data() == data.as_slice())
        {
            // An exact byte-for-byte replacement is a true no-op. Keep both
            // the entry allocation and its parsed archive: evicting here
            // would make a failed/empty publication observable as a cache
            // miss even though the source bytes did not change.
            return Some(data);
        }
        let previous = self.entries.replace_data(position, data);
        self.invalidate_archive(name);
        previous
    }

    /// Remove one entry and evict the matching parsed component.
    pub fn remove_entry(&mut self, name: &str) -> Option<Entry> {
        let position = self.position(name)?;
        let removed = self.entries.remove_at(position);
        self.invalidate_archive(name);
        removed
    }

    /// Insert a new entry at an ordered position.
    ///
    /// # Errors
    ///
    /// Returns an entry-store error when the position is invalid, the name is
    /// duplicated, or the entry store cannot reserve storage.
    pub fn try_insert_entry_at(
        &mut self,
        position: usize,
        entry: Entry,
    ) -> Result<(), EntryStoreError> {
        let name = entry.name().to_owned();
        self.entries.try_insert_at(position, entry)?;
        // The low-level API permits callers to parse a key before its entry
        // is inserted. Detach that speculative flight/value only after the
        // fallible structural insertion succeeds, so failed publication is
        // still a no-op while a successful insertion cannot expose stale
        // bytes for the newly occupied name.
        self.invalidate_archive(&name);
        Ok(())
    }

    fn invalidate_archive(&mut self, name: &str) {
        let _removed = self.parsed_archive.invalidate_borrowed(name);
    }
}

/// Return whether one package member is the legacy operation-log encoding.
///
/// The name check is intentionally basename-only to match historical package
/// behavior; callers pass the bounded, preserved member bytes from this
/// package state.
#[must_use]
pub fn is_legacy_operation_storage(name: &str, data: &[u8]) -> bool {
    name.rsplit('/').next() == Some("OperationStorage.iwa") && data.starts_with(b"bvxn")
}

fn new_archive_cache(archive_limits: ArchiveLimits) -> WeightedCache<String, Archive> {
    WeightedCache::new(archive_limits.max_archive_bytes()).unwrap_or_else(|error| {
        unreachable!("validated IWA archive limits cannot create a zero-weight cache: {error}")
    })
}

#[cfg(test)]
mod tests {
    use std::sync::Arc;
    use std::sync::Barrier;
    use std::sync::atomic::{AtomicUsize, Ordering};
    use std::thread;

    use litchi_iwa_core::{Archive, ArchiveLimits};
    use litchi_iwa_package::Entry;

    use super::{EntryStoreError, PackageState};

    fn archive() -> Archive {
        Archive::default()
    }

    #[test]
    fn exposes_ordered_entry_views_without_leaking_the_store() {
        let state = PackageState::from_entries(
            vec![
                Entry::new("first".to_owned(), vec![1]),
                Entry::new("second".to_owned(), vec![2, 3]),
            ],
            ArchiveLimits::default(),
        )
        .unwrap_or_else(|error| panic!("test entries should be valid: {error}"));

        assert_eq!(state.len(), 2);
        assert!(!state.is_empty());
        assert_eq!(state.position("second"), Some(1));
        assert_eq!(state.get("first").map(Entry::data), Some(&[1][..]));
        assert_eq!(state.get_at(1).map(Entry::data), Some(&[2, 3][..]));
        assert_eq!(
            state.iter().map(Entry::name).collect::<Vec<_>>(),
            ["first", "second"]
        );
    }

    #[test]
    fn excludes_legacy_operation_logs_from_iwa_scans() {
        let state = PackageState::from_entries(
            vec![
                Entry::new("Index/Document.iwa".to_owned(), vec![1]),
                Entry::new(
                    "Index/OperationStorage.iwa".to_owned(),
                    b"bvxn log".to_vec(),
                ),
                Entry::new("Index/Other.iwa".to_owned(), vec![2]),
                Entry::new("preview.jpg".to_owned(), vec![3]),
            ],
            ArchiveLimits::default(),
        )
        .unwrap_or_else(|error| panic!("valid entries should be accepted: {error}"));

        assert_eq!(
            state.iwa_entry_names().collect::<Vec<_>>(),
            ["Index/Document.iwa", "Index/Other.iwa"]
        );
    }

    #[test]
    fn patch_publication_restarts_the_component_cache() {
        let source = PackageState::from_entries(
            vec![Entry::new("Index/Document.iwa".to_owned(), vec![1])],
            ArchiveLimits::default(),
        )
        .unwrap_or_else(|error| panic!("test entries should be valid: {error}"));
        let parsed = source
            .get_or_parse_archive("Index/Document.iwa", |_| Ok((archive(), 1)))
            .unwrap_or_else(|error| panic!("source parse should succeed: {error}"));

        let mut edited = source.clone();
        assert_eq!(
            edited.replace_entry_data("Index/Document.iwa", vec![2]),
            Some(vec![1])
        );
        let patch = source.patch_to(&edited);
        let published = source
            .apply_patch(&patch)
            .unwrap_or_else(|error| panic!("matching patch should apply: {error}"));
        assert_eq!(
            published.get("Index/Document.iwa").map(Entry::data),
            Some(&[2][..])
        );

        let parse_count = AtomicUsize::new(0);
        let replacement = published
            .get_or_parse_archive("Index/Document.iwa", |_| {
                parse_count.fetch_add(1, Ordering::SeqCst);
                Ok((archive(), 1))
            })
            .unwrap_or_else(|error| panic!("published parse should succeed: {error}"));
        assert!(!Arc::ptr_eq(&parsed, &replacement));
        assert_eq!(parse_count.load(Ordering::SeqCst), 1);

        let stale = PackageState::from_entries(
            vec![Entry::new("Index/Document.iwa".to_owned(), vec![9])],
            ArchiveLimits::default(),
        )
        .unwrap_or_else(|error| panic!("stale entries should be valid: {error}"));
        assert!(matches!(
            stale.apply_patch(&patch),
            Err(EntryStoreError::PatchSourceMismatch)
        ));
    }

    #[test]
    fn forks_share_completed_archives_and_detach_entry_storage() {
        let source = PackageState::from_entries(
            vec![Entry::new("Index/Document.iwa".to_owned(), vec![1])],
            ArchiveLimits::default(),
        )
        .unwrap_or_else(|error| panic!("test entries should be valid: {error}"));
        let parsed = source
            .get_or_parse_archive("Index/Document.iwa", |_| Ok((archive(), 1)))
            .unwrap_or_else(|error| panic!("initial parse should succeed: {error}"));

        let mut forked = source.clone();
        let parse_count = AtomicUsize::new(0);
        let shared = forked
            .get_or_parse_archive("Index/Document.iwa", |_| {
                parse_count.fetch_add(1, Ordering::SeqCst);
                Ok((archive(), 1))
            })
            .unwrap_or_else(|error| panic!("completed parse should remain shared: {error}"));
        assert!(Arc::ptr_eq(&parsed, &shared));
        assert_eq!(parse_count.load(Ordering::SeqCst), 0);

        let previous = forked.replace_entry_data("Index/Document.iwa", vec![2]);
        assert_eq!(previous, Some(vec![1]));
        assert_eq!(source.get_at(0).map(Entry::data), Some(&[1][..]));
        assert_eq!(forked.get_at(0).map(Entry::data), Some(&[2][..]));
        let replacement_parse_count = AtomicUsize::new(0);
        let replacement = forked
            .get_or_parse_archive("Index/Document.iwa", |_| {
                replacement_parse_count.fetch_add(1, Ordering::SeqCst);
                Ok((archive(), 1))
            })
            .unwrap_or_else(|error| panic!("replacement parse should succeed: {error}"));
        assert!(!Arc::ptr_eq(&parsed, &replacement));
        assert_eq!(replacement_parse_count.load(Ordering::SeqCst), 1);
    }

    #[test]
    fn patch_publication_is_exact_source_checked_and_starts_a_fresh_cache() {
        let source = PackageState::from_entries(
            vec![Entry::new("Index/Document.iwa".to_owned(), vec![1])],
            ArchiveLimits::default(),
        )
        .unwrap_or_else(|error| panic!("source entries should be valid: {error}"));
        let parsed = source
            .get_or_parse_archive("Index/Document.iwa", |_| Ok((archive(), 1)))
            .unwrap_or_else(|error| panic!("source parse should succeed: {error}"));

        let mut edited = source.clone();
        assert_eq!(
            edited.replace_entry_data("Index/Document.iwa", vec![2]),
            Some(vec![1])
        );
        let patch = source.patch_to(&edited);

        let published = source
            .apply_patch(&patch)
            .unwrap_or_else(|error| panic!("exact source patch should apply: {error}"));
        assert_eq!(
            published.get("Index/Document.iwa").map(Entry::data),
            Some([2].as_slice())
        );
        let parse_count = AtomicUsize::new(0);
        let published_archive = published
            .get_or_parse_archive("Index/Document.iwa", |_| {
                parse_count.fetch_add(1, Ordering::SeqCst);
                Ok((archive(), 1))
            })
            .unwrap_or_else(|error| panic!("published parse should succeed: {error}"));
        assert!(!Arc::ptr_eq(&parsed, &published_archive));
        assert_eq!(parse_count.load(Ordering::SeqCst), 1);

        let stale = PackageState::from_entries(
            vec![Entry::new("Index/Document.iwa".to_owned(), vec![9])],
            ArchiveLimits::default(),
        )
        .unwrap_or_else(|error| panic!("stale entries should be valid: {error}"));
        assert!(matches!(
            stale.apply_patch(&patch),
            Err(litchi_iwa_package::Error::PatchSourceMismatch)
        ));
        assert_eq!(
            source.get("Index/Document.iwa").map(Entry::data),
            Some([1].as_slice())
        );
    }

    #[test]
    fn forks_do_not_inherit_active_parser_flights() {
        let source = Arc::new(PackageState::default());
        let source_parse_started = Arc::new(Barrier::new(2));
        let release_source_parse = Arc::new(Barrier::new(2));
        let source_for_thread = Arc::clone(&source);
        let started_for_thread = Arc::clone(&source_parse_started);
        let release_for_thread = Arc::clone(&release_source_parse);
        let source_parse = thread::spawn(move || {
            source_for_thread.get_or_parse_archive("Index/Document.iwa", |_| {
                started_for_thread.wait();
                release_for_thread.wait();
                Ok((archive(), 1))
            })
        });

        source_parse_started.wait();
        let forked = (*source).clone();
        let fork_parse_count = AtomicUsize::new(0);
        let parsed = forked
            .get_or_parse_archive("Index/Document.iwa", |_| {
                fork_parse_count.fetch_add(1, Ordering::SeqCst);
                Ok((archive(), 1))
            })
            .unwrap_or_else(|error| panic!("forked parse should not wait: {error}"));
        assert_eq!(parsed.as_ref(), &archive());
        assert_eq!(fork_parse_count.load(Ordering::SeqCst), 1);

        release_source_parse.wait();
        let source_parsed = source_parse
            .join()
            .unwrap_or_else(|_error| panic!("source parser thread should not panic"))
            .unwrap_or_else(|error| panic!("source parse should succeed: {error}"));
        assert_eq!(source_parsed.as_ref(), &archive());
    }

    #[test]
    fn mutating_a_component_reparses_its_replacement() {
        let mut state = PackageState::from_entries(
            vec![Entry::new("Index/Document.iwa".to_owned(), vec![1])],
            ArchiveLimits::default(),
        )
        .unwrap_or_else(|error| panic!("test entries should be valid: {error}"));
        let first = state
            .get_or_parse_archive("Index/Document.iwa", |_| Ok((archive(), 1)))
            .unwrap_or_else(|error| panic!("initial parse should succeed: {error}"));

        state
            .entry_data_mut("Index/Document.iwa")
            .unwrap_or_else(|| panic!("test component should exist"))[0] = 2;
        let parse_count = AtomicUsize::new(0);
        let replacement = state
            .get_or_parse_archive("Index/Document.iwa", |_| {
                parse_count.fetch_add(1, Ordering::SeqCst);
                Ok((archive(), 1))
            })
            .unwrap_or_else(|error| panic!("replacement parse should succeed: {error}"));
        assert!(!Arc::ptr_eq(&first, &replacement));
        assert_eq!(parse_count.load(Ordering::SeqCst), 1);
    }

    #[test]
    fn single_flight_shares_one_parsed_archive_across_threads() {
        const CALLERS: usize = 8;
        let state = Arc::new(PackageState::default());
        let ready = Arc::new(Barrier::new(CALLERS + 1));
        let first_parse_started = Arc::new(Barrier::new(2));
        let parse_count = Arc::new(AtomicUsize::new(0));
        let handles = (0..CALLERS)
            .map(|_| {
                let state_for_thread = Arc::clone(&state);
                let ready_for_thread = Arc::clone(&ready);
                let first_parse_started_for_thread = Arc::clone(&first_parse_started);
                let parse_count_for_thread = Arc::clone(&parse_count);
                thread::spawn(move || {
                    ready_for_thread.wait();
                    state_for_thread.get_or_parse_archive("Index/Document.iwa", |_| {
                        if parse_count_for_thread.fetch_add(1, Ordering::SeqCst) == 0 {
                            first_parse_started_for_thread.wait();
                        }
                        Ok((archive(), 1))
                    })
                })
            })
            .collect::<Vec<_>>();

        ready.wait();
        first_parse_started.wait();
        let mut results = handles
            .into_iter()
            .map(|handle| {
                handle
                    .join()
                    .unwrap_or_else(|_error| panic!("parser thread should not panic"))
                    .unwrap_or_else(|error| panic!("parse should succeed: {error}"))
            })
            .collect::<Vec<_>>();
        let first = results
            .pop()
            .unwrap_or_else(|| panic!("at least one parser result is required"));
        assert!(results.iter().all(|result| Arc::ptr_eq(&first, result)));
        assert_eq!(parse_count.load(Ordering::SeqCst), 1);
    }

    #[test]
    fn failed_flight_wakes_waiters_and_allows_a_retry() {
        const CALLERS: usize = 6;
        let state = Arc::new(PackageState::default());
        let ready = Arc::new(Barrier::new(CALLERS + 1));
        let first_parse_started = Arc::new(Barrier::new(2));
        let parse_count = Arc::new(AtomicUsize::new(0));
        let handles = (0..CALLERS)
            .map(|_| {
                let state_for_thread = Arc::clone(&state);
                let ready_for_thread = Arc::clone(&ready);
                let first_parse_started_for_thread = Arc::clone(&first_parse_started);
                let parse_count_for_thread = Arc::clone(&parse_count);
                thread::spawn(move || {
                    ready_for_thread.wait();
                    state_for_thread.get_or_parse_archive("Index/Document.iwa", |_| {
                        if parse_count_for_thread.fetch_add(1, Ordering::SeqCst) == 0 {
                            first_parse_started_for_thread.wait();
                        }
                        Err(Box::new(std::io::Error::other("synthetic parse failure")))
                    })
                })
            })
            .collect::<Vec<_>>();

        ready.wait();
        first_parse_started.wait();
        for handle in handles {
            let result = handle
                .join()
                .unwrap_or_else(|_error| panic!("parser thread should not panic"));
            let Err(error) = result else {
                panic!("failed parser flight should return an error");
            };
            assert!(error.to_string().contains("synthetic parse failure"));
        }

        let retry_count = AtomicUsize::new(0);
        let retry = state
            .get_or_parse_archive("Index/Document.iwa", |_| {
                retry_count.fetch_add(1, Ordering::SeqCst);
                Ok((archive(), 1))
            })
            .unwrap_or_else(|error| panic!("retry should succeed: {error}"));
        assert_eq!(retry.as_ref(), &archive());
        assert_eq!(retry_count.load(Ordering::SeqCst), 1);
        assert!(parse_count.load(Ordering::SeqCst) >= 1);
    }
}
