//! Bounded clean-scene retention for source-backed presentations.
//!
//! Each presentation has an independent 16 MiB / 4096-entry weighted LRU.
//! Caller-held scenes pin their snapshots but may be evicted from the memo.
//! Managed execution contexts disable retention, so the memo cannot keep a
//! payload budget reservation alive after the caller drops its handles.
//! Logical weights are conservative cache accounting, not allocator measurements.
//! Failures are never cached. Reads fence source and execution state before
//! lookup and after lookup, parsing and publication.

use std::collections::HashMap;
use std::fmt;
use std::sync::{Mutex, MutexGuard};

use litchi_opc::{PackURI, SourceBackedPackage};

use crate::shape::owned::Scene;

/// Logical ceiling on the scenes one presentation retains: 16 MiB.
const SCENE_CACHE_MAX_BYTES: usize = 16 * 1024 * 1024;
/// Ceiling on the number of scenes one presentation retains.
const SCENE_CACHE_MAX_ENTRIES: usize = 4096;
/// Modeled bookkeeping per retained scene (the slot, its index entry and both
/// key headers), added to twice the key's text and the scene's own weight.
const SCENE_CACHE_ENTRY_OVERHEAD: usize = 192;

/// The logical weight one retained scene charges against the ceiling.
fn entry_weight(part: &PackURI, scene: &Scene) -> usize {
    scene
        .retained_bytes()
        .saturating_add(part.as_str().len().saturating_mul(2))
        .saturating_add(SCENE_CACHE_ENTRY_OVERHEAD)
}

/// Content-free counters and gauges of a source-backed presentation's shape
/// scene memo; see [`super::SourceBackedPresentation::scene_cache_diagnostics`].
///
/// Counters accumulate over the presentation's life. Weights are the logical
/// `Scene` retained weight of each scene plus a modeled per-entry
/// overhead: cache accounting, not allocator measurements.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub struct CacheDiagnostics {
    /// Scene lookups that found a retained scene, before the final source fence.
    pub hits: u64,
    /// Scene requests that found no retained scene and read the part.
    pub misses: u64,
    /// Scenes built after a miss and offered to the cache.
    pub builds: u64,
    /// Built scenes discarded because a concurrent read had already
    /// published one for the same part; the caller received that one.
    pub concurrent_builds: u64,
    /// Retained scenes evicted while no caller held them.
    pub clean_evictions: u64,
    /// Retained scenes evicted while a caller still held them; the scene
    /// lives on with that caller.
    pub pinned_evictions: u64,
    /// Retained scenes released by an explicit trim.
    pub trimmed: u64,
    /// Built scenes heavier than the whole ceiling, returned unretained.
    pub oversized_bypasses: u64,
    /// Built scenes returned unretained because the presentation is
    /// budget-managed and retains no scene.
    pub unretained_builds: u64,
    /// Built scenes returned unretained because the cache could not reserve
    /// a slot for them.
    pub allocation_bypasses: u64,
    /// Published scenes removed after a final source/execution fence failed.
    pub fence_discards: u64,
    /// Scenes currently retained.
    pub retained_entries: usize,
    /// Logical weight of the retained scenes.
    pub retained_bytes: usize,
    /// Retained scenes a caller also holds.
    pub pinned_entries: usize,
    /// Logical weight of the retained scenes a caller also holds.
    pub pinned_bytes: usize,
    /// The retention ceiling in logical bytes; `0` when the presentation
    /// retains no scene.
    pub max_bytes: usize,
    /// The retention ceiling in scenes; `0` when the presentation retains no
    /// scene.
    pub max_entries: usize,
}

/// One retained scene, linked into the recency list by slot position.
struct Slot {
    key: PackURI,
    scene: Scene,
    weight: usize,
    /// The next less recently used slot.
    older: Option<usize>,
    /// The next more recently used slot.
    newer: Option<usize>,
}

/// The memo's entries in a slab with an intrusive recency list, so a hit, an
/// insertion and an eviction are each O(1) and a hit allocates nothing.
#[derive(Default)]
struct State {
    index: HashMap<PackURI, usize>,
    slots: Vec<Slot>,
    newest: Option<usize>,
    oldest: Option<usize>,
    retained_bytes: usize,
    hits: u64,
    misses: u64,
    builds: u64,
    concurrent_builds: u64,
    clean_evictions: u64,
    pinned_evictions: u64,
    trimmed: u64,
    oversized_bypasses: u64,
    unretained_builds: u64,
    allocation_bypasses: u64,
    fence_discards: u64,
}

impl State {
    fn slot_mut(&mut self, slot: usize) -> Option<&mut Slot> {
        self.slots.get_mut(slot)
    }

    /// Detach `slot` from the recency list.
    fn unlink(&mut self, slot: usize) {
        let Some((older, newer)) = self.slots.get(slot).map(|entry| (entry.older, entry.newer))
        else {
            return;
        };
        match older.and_then(|older| self.slot_mut(older)) {
            Some(entry) => entry.newer = newer,
            None => self.oldest = newer,
        }
        match newer.and_then(|newer| self.slot_mut(newer)) {
            Some(entry) => entry.older = older,
            None => self.newest = older,
        }
    }

    /// Attach a detached `slot` as the most recently used.
    fn link_newest(&mut self, slot: usize) {
        let previous = self.newest;
        if let Some(entry) = self.slot_mut(slot) {
            entry.older = previous;
            entry.newer = None;
        }
        match previous.and_then(|previous| self.slot_mut(previous)) {
            Some(entry) => entry.newer = Some(slot),
            None => self.oldest = Some(slot),
        }
        self.newest = Some(slot);
    }

    /// Mark `slot` the most recently used.
    fn touch(&mut self, slot: usize) {
        if self.newest != Some(slot) {
            self.unlink(slot);
            self.link_newest(slot);
        }
    }

    /// Remove `slot`, keeping the slab dense: the last slot moves into its
    /// place and its links and index entry follow it.
    fn remove(&mut self, slot: usize) -> Option<Slot> {
        if slot >= self.slots.len() {
            return None;
        }
        self.unlink(slot);
        let removed = self.slots.swap_remove(slot);
        self.index.remove(&removed.key);
        if let Some((older, newer)) = self.slots.get(slot).map(|moved| (moved.older, moved.newer)) {
            match older.and_then(|older| self.slot_mut(older)) {
                Some(entry) => entry.newer = Some(slot),
                None => self.oldest = Some(slot),
            }
            match newer.and_then(|newer| self.slot_mut(newer)) {
                Some(entry) => entry.older = Some(slot),
                None => self.newest = Some(slot),
            }
            let moved_key = self.slots.get(slot).map(|moved| &moved.key);
            if let Some(position) = moved_key.and_then(|key| self.index.get_mut(key)) {
                *position = slot;
            }
        }
        self.retained_bytes = self.retained_bytes.saturating_sub(removed.weight);
        Some(removed)
    }

    /// Evict the least recently used entry; `false` when there is none.
    fn evict_oldest(&mut self) -> bool {
        let Some(removed) = self.oldest.and_then(|oldest| self.remove(oldest)) else {
            return false;
        };
        if removed.scene.is_unshared() {
            self.clean_evictions = self.clean_evictions.saturating_add(1);
        } else {
            self.pinned_evictions = self.pinned_evictions.saturating_add(1);
        }
        true
    }

    /// The retained scene for `part`, marked most recently used.
    fn get(&mut self, part: &PackURI) -> Option<Scene> {
        let slot = *self.index.get(part)?;
        self.touch(slot);
        self.slots.get(slot).map(|entry| entry.scene.clone())
    }
}

/// The shape-scene memo of one source-backed presentation; see the module
/// documentation.
pub(in crate::presentation) struct SceneCache {
    /// Logical byte ceiling; `0` retains nothing.
    max_bytes: usize,
    /// Entry ceiling; `0` retains nothing.
    max_entries: usize,
    state: Mutex<State>,
}

impl SceneCache {
    /// The cache policy for `package`: bounded retention, or none for a
    /// budget-managed package.
    pub(super) fn for_package(package: &SourceBackedPackage) -> Self {
        if package.execution_context().is_some() {
            Self::with_limits(0, 0)
        } else {
            Self::with_limits(SCENE_CACHE_MAX_BYTES, SCENE_CACHE_MAX_ENTRIES)
        }
    }

    fn with_limits(max_bytes: usize, max_entries: usize) -> Self {
        Self {
            max_bytes,
            max_entries,
            state: Mutex::new(State::default()),
        }
    }

    const fn retains(&self) -> bool {
        self.max_bytes != 0 && self.max_entries != 0
    }

    fn lock(&self) -> MutexGuard<'_, State> {
        // Every critical section leaves the state consistent before any call
        // that could panic, so a poisoned lock still holds a usable state.
        self.state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
    }

    /// The retained scene for `part`, if any.
    pub(super) fn lookup(&self, part: &PackURI) -> Option<Scene> {
        let mut state = self.lock();
        let found = state.get(part);
        if found.is_some() {
            state.hits = state.hits.saturating_add(1);
        } else {
            state.misses = state.misses.saturating_add(1);
        }
        found
    }

    /// Offer a scene just built for `part` and return the scene the caller
    /// should use: the one a concurrent read already published, or `scene`.
    pub(super) fn publish(&self, part: &PackURI, scene: Scene) -> Scene {
        let mut state = self.lock();
        state.builds = state.builds.saturating_add(1);
        if let Some(published) = state.get(part) {
            state.concurrent_builds = state.concurrent_builds.saturating_add(1);
            return published;
        }
        if !self.retains() {
            state.unretained_builds = state.unretained_builds.saturating_add(1);
            return scene;
        }
        let weight = entry_weight(part, &scene);
        if weight > self.max_bytes {
            state.oversized_bypasses = state.oversized_bypasses.saturating_add(1);
            return scene;
        }
        while state.retained_bytes.saturating_add(weight) > self.max_bytes
            || state.slots.len() >= self.max_entries
        {
            if !state.evict_oldest() {
                break;
            }
        }
        if state.slots.try_reserve(1).is_err() || state.index.try_reserve(1).is_err() {
            state.allocation_bypasses = state.allocation_bypasses.saturating_add(1);
            return scene;
        }
        let slot = state.slots.len();
        state.slots.push(Slot {
            key: part.clone(),
            scene: scene.clone(),
            weight,
            older: None,
            newer: None,
        });
        state.index.insert(part.clone(), slot);
        state.link_newest(slot);
        state.retained_bytes = state.retained_bytes.saturating_add(weight);
        scene
    }

    /// Roll back only the exact publication whose final fence failed.
    pub(super) fn discard(&self, part: &PackURI, scene: &Scene) {
        let mut state = self.lock();
        if let Some(slot) = state.index.get(part).copied()
            && state
                .slots
                .get(slot)
                .is_some_and(|entry| Scene::ptr_eq(&entry.scene, scene))
        {
            state.remove(slot);
            state.fence_discards = state.fence_discards.saturating_add(1);
        }
    }

    /// Release every retained scene no caller holds; scenes a caller holds
    /// stay retained. Returns how many were released.
    pub(super) fn trim_clean(&self) -> usize {
        let mut state = self.lock();
        let mut released = 0usize;
        // Walk from the least recently used end. Removing a slot moves the
        // last slot into its place and re-points every link to it, so the
        // only position read before a removal that can go stale is the next
        // one, when the next slot is the one that moved.
        let mut cursor = state.oldest;
        while let Some(slot) = cursor {
            let Some(entry) = state.slots.get(slot) else {
                break;
            };
            let newer = entry.newer;
            if entry.scene.is_unshared() {
                let last = state.slots.len().saturating_sub(1);
                state.remove(slot);
                released = released.saturating_add(1);
                // The slot that moved into `slot` may be the next one to visit.
                cursor = if newer == Some(last) {
                    Some(slot)
                } else {
                    newer
                };
            } else {
                cursor = newer;
            }
        }
        state.trimmed = state
            .trimmed
            .saturating_add(u64::try_from(released).unwrap_or(u64::MAX));
        released
    }

    pub(super) fn diagnostics(&self) -> CacheDiagnostics {
        let state = self.lock();
        let (pinned_entries, pinned_bytes) = state
            .slots
            .iter()
            .filter(|entry| !entry.scene.is_unshared())
            .fold((0usize, 0usize), |(count, bytes), entry| {
                (count.saturating_add(1), bytes.saturating_add(entry.weight))
            });
        CacheDiagnostics {
            hits: state.hits,
            misses: state.misses,
            builds: state.builds,
            concurrent_builds: state.concurrent_builds,
            clean_evictions: state.clean_evictions,
            pinned_evictions: state.pinned_evictions,
            trimmed: state.trimmed,
            oversized_bypasses: state.oversized_bypasses,
            unretained_builds: state.unretained_builds,
            allocation_bypasses: state.allocation_bypasses,
            fence_discards: state.fence_discards,
            retained_entries: state.slots.len(),
            retained_bytes: state.retained_bytes,
            pinned_entries,
            pinned_bytes,
            max_bytes: self.max_bytes,
            max_entries: self.max_entries,
        }
    }
}

// Clones of a presentation share one memo across threads.
const _: () = {
    const fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<SceneCache>();
};

impl fmt::Debug for SceneCache {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SceneCache")
            .field("diagnostics", &self.diagnostics())
            .finish()
    }
}

impl super::SourceBackedPresentation {
    /// Content-free gauges and counters for the bounded scene memo.
    /// Managed execution contexts retain no scenes and report zero ceilings.
    #[must_use]
    pub fn scene_cache_diagnostics(&self) -> CacheDiagnostics {
        self.inner.scenes.diagnostics()
    }

    /// Release retained scenes that no caller currently holds.
    /// Pinned scene snapshots remain valid, including after later eviction.
    pub fn trim_scene_cache(&self) -> usize {
        self.inner.scenes.trim_clean()
    }
}

impl super::SourceSlide {
    /// Read this slide's independently retained shape scene.
    ///
    /// Successful scenes use a bounded presentation-local LRU. Every call
    /// checks source revision and execution state before and after lookup or
    /// parsing; failed reads are not cached. Returned scenes are detached
    /// immutable snapshots and hold no presentation owner. Managed execution
    /// contexts retain no scenes in the memo; a returned source payload keeps
    /// its budget reservation until the last owning handle is dropped.
    ///
    /// # Errors
    /// Returns source, execution, physical XML limit, or typed leaf parse errors.
    pub fn scene(&self) -> crate::Result<Scene> {
        let package = &self.owner.package;
        let fence = || -> crate::Result<()> {
            package.check_execution()?;
            package.source_version()?;
            Ok(())
        };
        fence()?;
        if let Some(scene) = self.owner.scenes.lookup(&self.data.part_uri) {
            fence()?;
            return Ok(scene);
        }
        let result: crate::Result<Scene> = (|| {
            let view = package.part(&self.data.part_uri)?;
            let declared = view.declared_uncompressed_size()?;
            if declared > self.owner.limits.max_part_bytes() {
                return Err(crate::Error::Opc(litchi_opc::OpcError::Limit {
                    kind: litchi_opc::SourceLimitKind::PartBytes,
                    observed: declared,
                    maximum: self.owner.limits.max_part_bytes(),
                }));
            }
            let shape_limits =
                crate::open_options::selected(self.owner.semantic_limits.as_ref()).shape();
            let semantic_maximum = shape_limits.input_bytes();
            if declared > semantic_maximum as u64 {
                return Err(crate::Error::Limit {
                    kind: crate::opened::LimitKind::Codec(
                        crate::opened::CodecLimitKind::InputBytes,
                    ),
                    observed: usize::try_from(declared).unwrap_or(usize::MAX),
                    maximum: semantic_maximum,
                });
            }
            let part = self.load_part(&view)?;
            crate::source_limits::validate_xml_limits(
                part.source_data().as_bytes(),
                self.owner.limits,
            )?;
            let slide = crate::parts::SlidePart::from_part(&part)?;
            let scene = Scene::from_source(
                slide.shapes_with_limits(shape_limits)?,
                part.source_data().clone(),
            )?;
            fence()?;
            Ok(scene)
        })();
        fence()?;
        let scene = result?;
        let scene = self.owner.scenes.publish(&self.data.part_uri, scene);
        if let Err(error) = fence() {
            self.owner.scenes.discard(&self.data.part_uri, &scene);
            return Err(error);
        }
        Ok(scene)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn uri(index: usize) -> PackURI {
        PackURI::new(format!("/ppt/slides/slide{index}.xml")).unwrap()
    }

    fn scene(shapes: usize) -> Scene {
        let body: String = (0..shapes)
            .map(|index| {
                format!(
                    r#"<p:sp><p:nvSpPr><p:cNvPr id="{index}" name="s{index}"/></p:nvSpPr></p:sp>"#
                )
            })
            .collect();
        let xml = format!(
            r#"<p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:nvGrpSpPr/><p:grpSpPr/>{body}</p:spTree>"#
        );
        Scene::read(xml.into_bytes(), crate::shape::Limits::DEFAULT).unwrap()
    }

    fn weight(scene: &Scene, part: &PackURI) -> usize {
        entry_weight(part, scene)
    }

    #[test]
    fn rollback_does_not_remove_a_different_scene_for_the_same_part() {
        let cache = SceneCache::with_limits(SCENE_CACHE_MAX_BYTES, SCENE_CACHE_MAX_ENTRIES);
        let retained = cache.publish(&uri(1), scene(1));
        cache.discard(&uri(1), &scene(1));
        assert_eq!(cache.diagnostics().retained_entries, 1);
        assert_eq!(cache.diagnostics().fence_discards, 0);
        cache.discard(&uri(1), &retained);
        assert_eq!(cache.diagnostics().retained_entries, 0);
        assert_eq!(cache.diagnostics().fence_discards, 1);
        assert_eq!(retained.len(), 1);
    }

    #[test]
    fn a_published_scene_is_returned_by_later_lookups() {
        let cache = SceneCache::with_limits(SCENE_CACHE_MAX_BYTES, SCENE_CACHE_MAX_ENTRIES);
        assert!(cache.lookup(&uri(1)).is_none());
        let built = scene(2);
        let published = cache.publish(&uri(1), built.clone());
        assert!(Scene::ptr_eq(&built, &published));
        let hit = cache.lookup(&uri(1)).unwrap();
        assert!(Scene::ptr_eq(&built, &hit));
        let diagnostics = cache.diagnostics();
        assert_eq!((diagnostics.hits, diagnostics.misses), (1, 1));
        assert_eq!(diagnostics.builds, 1);
        assert_eq!(diagnostics.retained_entries, 1);
        assert_eq!(diagnostics.retained_bytes, weight(&built, &uri(1)));
        // `built`, `published` and `hit` are callers' shares.
        assert_eq!(diagnostics.pinned_entries, 1);
        drop((built, published, hit));
        assert_eq!(cache.diagnostics().pinned_entries, 0);
    }

    #[test]
    fn a_concurrent_second_build_returns_the_first_published_scene() {
        let cache = SceneCache::with_limits(SCENE_CACHE_MAX_BYTES, SCENE_CACHE_MAX_ENTRIES);
        let first = cache.publish(&uri(1), scene(1));
        let second = cache.publish(&uri(1), scene(1));
        assert!(Scene::ptr_eq(&first, &second));
        let diagnostics = cache.diagnostics();
        assert_eq!(diagnostics.builds, 2);
        assert_eq!(diagnostics.concurrent_builds, 1);
        assert_eq!(diagnostics.retained_entries, 1);
    }

    #[test]
    fn eviction_is_least_recently_used_and_never_frees_a_held_scene() {
        let probe = scene(4);
        let one = weight(&probe, &uri(1));
        // Room for exactly two entries of this weight.
        let cache = SceneCache::with_limits(one * 2 + one / 2, SCENE_CACHE_MAX_ENTRIES);
        let held = cache.publish(&uri(1), scene(4));
        drop(cache.publish(&uri(2), scene(4)));
        // Touch slide 1 so slide 2 is the least recently used.
        assert!(cache.lookup(&uri(1)).is_some());
        drop(cache.publish(&uri(3), scene(4)));
        assert!(cache.lookup(&uri(2)).is_none());
        assert!(cache.lookup(&uri(1)).is_some());
        assert!(cache.lookup(&uri(3)).is_some());
        let diagnostics = cache.diagnostics();
        assert_eq!(diagnostics.clean_evictions, 1);
        assert_eq!(diagnostics.retained_entries, 2);
        assert!(diagnostics.retained_bytes <= diagnostics.max_bytes);

        // Slide 1 is now the least recently used and still held: evicting it
        // drops only the cache's share.
        drop(cache.publish(&uri(4), scene(4)));
        drop(cache.publish(&uri(5), scene(4)));
        let diagnostics = cache.diagnostics();
        assert_eq!(diagnostics.pinned_evictions, 1);
        assert!(cache.lookup(&uri(1)).is_none());
        assert_eq!(held.len(), 4);
        assert_eq!(held.shape("s3").unwrap().unwrap().view().id(), Some(3));
    }

    #[test]
    fn the_entry_ceiling_bounds_the_number_of_scenes() {
        let cache = SceneCache::with_limits(SCENE_CACHE_MAX_BYTES, 3);
        for index in 0..10 {
            drop(cache.publish(&uri(index), scene(1)));
        }
        let diagnostics = cache.diagnostics();
        assert_eq!(diagnostics.retained_entries, 3);
        assert_eq!(diagnostics.clean_evictions, 7);
        for index in 7..10 {
            assert!(cache.lookup(&uri(index)).is_some(), "{index}");
        }
    }

    #[test]
    fn an_oversized_scene_is_returned_but_not_retained() {
        let big = scene(64);
        let cache = SceneCache::with_limits(weight(&big, &uri(1)) - 1, SCENE_CACHE_MAX_ENTRIES);
        drop(cache.publish(&uri(2), scene(1)));
        let returned = cache.publish(&uri(1), big.clone());
        assert!(Scene::ptr_eq(&returned, &big));
        assert!(cache.lookup(&uri(1)).is_none());
        let diagnostics = cache.diagnostics();
        assert_eq!(diagnostics.oversized_bypasses, 1);
        // The bypass evicted nothing.
        assert_eq!(diagnostics.retained_entries, 1);
        assert_eq!(diagnostics.clean_evictions, 0);
    }

    #[test]
    fn a_non_retaining_cache_hands_every_build_back() {
        let cache = SceneCache::with_limits(0, 0);
        let built = scene(1);
        let returned = cache.publish(&uri(1), built.clone());
        assert!(Scene::ptr_eq(&returned, &built));
        assert!(cache.lookup(&uri(1)).is_none());
        let diagnostics = cache.diagnostics();
        assert_eq!(diagnostics.unretained_builds, 1);
        assert_eq!(diagnostics.retained_entries, 0);
        assert_eq!((diagnostics.max_bytes, diagnostics.max_entries), (0, 0));
    }

    #[test]
    fn trimming_releases_clean_scenes_and_keeps_held_ones() {
        let cache = SceneCache::with_limits(SCENE_CACHE_MAX_BYTES, SCENE_CACHE_MAX_ENTRIES);
        let held = cache.publish(&uri(1), scene(2));
        drop(cache.publish(&uri(2), scene(2)));
        drop(cache.publish(&uri(3), scene(2)));
        assert_eq!(cache.trim_clean(), 2);
        let diagnostics = cache.diagnostics();
        assert_eq!(diagnostics.trimmed, 2);
        assert_eq!(diagnostics.retained_entries, 1);
        assert_eq!(diagnostics.retained_bytes, weight(&held, &uri(1)));
        assert_eq!(diagnostics.pinned_entries, 1);
        assert!(Scene::ptr_eq(&cache.lookup(&uri(1)).unwrap(), &held));
        drop(held);
        assert_eq!(cache.trim_clean(), 1);
        assert_eq!(cache.diagnostics().retained_bytes, 0);
    }

    /// The recency list from least to most recently used, checking every
    /// link, the index and the weight total on the way.
    fn checked_order(cache: &SceneCache) -> Vec<PackURI> {
        let state = cache.lock();
        let mut order = Vec::new();
        let mut previous = None;
        let mut cursor = state.oldest;
        while let Some(slot) = cursor {
            let entry = &state.slots[slot];
            assert_eq!(entry.older, previous, "back link of slot {slot}");
            assert_eq!(
                state.index.get(&entry.key),
                Some(&slot),
                "index of slot {slot}"
            );
            order.push(entry.key.clone());
            previous = Some(slot);
            cursor = entry.newer;
            assert!(order.len() <= state.slots.len(), "the list has a cycle");
        }
        assert_eq!(state.newest, previous);
        assert_eq!(order.len(), state.slots.len());
        assert_eq!(state.index.len(), state.slots.len());
        assert_eq!(
            state.retained_bytes,
            state.slots.iter().map(|entry| entry.weight).sum::<usize>()
        );
        order
    }

    /// A pseudo-random mix of lookups, publications, holds, releases and
    /// trims keeps the slab list identical to a plain least-recently-used
    /// model.
    #[test]
    fn the_recency_list_matches_a_reference_model() {
        let scenes: Vec<Scene> = (1..=4).map(scene).collect();
        let largest = scenes
            .iter()
            .map(|scene| entry_weight(&uri(10), scene))
            .max()
            .unwrap();
        let max_entries = 6;
        let cache = SceneCache::with_limits(largest * 4, max_entries);
        let mut model: Vec<(PackURI, usize)> = Vec::new();
        let mut held: Vec<Scene> = Vec::new();
        let mut seed = 0x2545_f491_4f6c_dd1d_u64;
        for step in 0..4_000 {
            seed ^= seed << 13;
            seed ^= seed >> 7;
            seed ^= seed << 17;
            let key = uri(10 + usize::try_from(seed % 9).unwrap());
            match (seed >> 8) % 6 {
                0 | 1 => {
                    let hit = cache.lookup(&key);
                    let position = model.iter().position(|(entry, _)| *entry == key);
                    assert_eq!(hit.is_some(), position.is_some(), "step {step}");
                    if let Some(position) = position {
                        let entry = model.remove(position);
                        model.push(entry);
                    }
                    if let Some(hit) = hit
                        && seed % 5 == 0
                    {
                        held.push(hit);
                    }
                },
                2 | 3 => {
                    // Every build is a fresh scene, as a part read is.
                    let candidate = scene(1 + usize::try_from((seed >> 16) % 4).unwrap());
                    let weight = entry_weight(&key, &candidate);
                    let published = cache.publish(&key, candidate);
                    if let Some(position) = model.iter().position(|(entry, _)| *entry == key) {
                        let entry = model.remove(position);
                        model.push(entry);
                    } else if weight <= largest * 4 {
                        while model.iter().map(|(_, weight)| weight).sum::<usize>() + weight
                            > largest * 4
                            || model.len() >= max_entries
                        {
                            model.remove(0);
                        }
                        model.push((key.clone(), weight));
                    }
                    if seed % 3 == 0 {
                        held.push(published);
                    }
                },
                4 => {
                    if !held.is_empty() {
                        let index = usize::try_from(seed % held.len() as u64).unwrap();
                        held.swap_remove(index);
                    }
                },
                _ => {
                    cache.trim_clean();
                    let live: Vec<PackURI> = {
                        let state = cache.lock();
                        state.slots.iter().map(|entry| entry.key.clone()).collect()
                    };
                    model.retain(|(entry, _)| live.contains(entry));
                },
            }
            let order = checked_order(&cache);
            let expected: Vec<PackURI> = model.iter().map(|(entry, _)| entry.clone()).collect();
            assert_eq!(order, expected, "step {step}");
        }
        // Once no caller holds a scene, a trim releases every entry.
        drop(held);
        cache.trim_clean();
        assert!(checked_order(&cache).is_empty());
        assert_eq!(cache.diagnostics().retained_bytes, 0);
    }

    #[test]
    fn concurrent_publishers_of_one_part_converge_on_one_scene() {
        let cache = SceneCache::with_limits(SCENE_CACHE_MAX_BYTES, SCENE_CACHE_MAX_ENTRIES);
        let scenes: Vec<Scene> = std::thread::scope(|scope| {
            let workers: Vec<_> = (0..8)
                .map(|_| {
                    scope.spawn(|| match cache.lookup(&uri(1)) {
                        Some(hit) => hit,
                        None => cache.publish(&uri(1), scene(3)),
                    })
                })
                .collect();
            workers
                .into_iter()
                .map(|worker| worker.join().unwrap())
                .collect()
        });
        for scene in &scenes[1..] {
            assert!(Scene::ptr_eq(&scenes[0], scene));
        }
        assert_eq!(cache.diagnostics().retained_entries, 1);
    }
}
