//! The eager owner's slide-scene cache (record 0262).
//!
//! The source-backed owner reads slide scenes through the leaf's weighted
//! scene memo; the eager owner has none, so before this cache every path
//! that reached into a slide (`/slide[N]/shape[K]/...`) parsed the slide's
//! XML into a scene again. A unified write resolves its target that way and
//! then hands the native edit positions, so on a large slide the scene read
//! was the whole unified overhead (record 0252: +17-22% at 100-10,000
//! shapes).
//!
//! The cache is keyed by the identity of the slide part's payload
//! (`Arc<Vec<u8>>`): an eager package shares the payload of every part a
//! commit did not touch with its predecessor, so a handle derived by a write
//! keeps the scenes of its unchanged slides, and a changed slide, whose
//! payload is new, is read again. A cached scene holds its payload, so the
//! key cannot be reused by another allocation while the entry lives. The
//! cache holds at most [`CAPACITY`] scenes, the most recently used first,
//! and only scenes that passed the adapter's strict check.

use std::sync::{Arc, Mutex};

use litchi_pptx::shape::owned::Scene;

/// The most scenes one handle lineage keeps.
const CAPACITY: usize = 4;

/// Scenes of eager slides, keyed by their payload's identity.
#[derive(Default)]
pub(super) struct Scenes {
    entries: Mutex<Vec<(Arc<Vec<u8>>, Scene)>>,
}

impl Scenes {
    /// The cached scene of the slide whose payload is `payload`, if any; a
    /// hit becomes the most recently used entry.
    pub(super) fn get(&self, payload: &Arc<Vec<u8>>) -> Option<Scene> {
        let mut entries = self.entries.lock().ok()?;
        let index = entries
            .iter()
            .position(|(key, _)| Arc::ptr_eq(key, payload))?;
        let entry = entries.remove(index);
        let scene = entry.1.clone();
        entries.insert(0, entry);
        Some(scene)
    }

    /// Remembers `scene` as the scene of the slide whose payload is
    /// `payload`, evicting the least recently used entry when full. A
    /// poisoned lock or a failed reservation caches nothing.
    pub(super) fn insert(&self, payload: Arc<Vec<u8>>, scene: Scene) {
        let Ok(mut entries) = self.entries.lock() else {
            return;
        };
        if entries.iter().any(|(key, _)| Arc::ptr_eq(key, &payload)) {
            return;
        }
        if entries.len() >= CAPACITY {
            entries.truncate(CAPACITY - 1);
        } else if entries.try_reserve(1).is_err() {
            return;
        }
        entries.insert(0, (payload, scene));
    }
}

#[cfg(test)]
mod tests {
    use std::sync::Arc;

    use litchi_pptx::shape::owned::Scene;

    use super::{CAPACITY, Scenes};

    fn scene() -> Scene {
        Scene::read(
            Arc::new(
                br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sld>"#
                    .to_vec(),
            ),
            litchi_pptx::shape::Limits::DEFAULT,
        )
        .expect("an empty slide reads")
    }

    #[test]
    fn hits_by_payload_identity_and_evicts_the_least_recent() {
        let cache = Scenes::default();
        let payloads: Vec<Arc<Vec<u8>>> = (0..=CAPACITY).map(|_| Arc::new(vec![0])).collect();
        assert!(cache.get(&payloads[0]).is_none());
        for payload in &payloads[..CAPACITY] {
            cache.insert(Arc::clone(payload), scene());
        }
        // An equal payload at another address is another slide.
        assert!(cache.get(&Arc::new(vec![0])).is_none());
        // Touch the oldest so the second oldest is evicted next.
        assert!(cache.get(&payloads[0]).is_some());
        cache.insert(Arc::clone(&payloads[CAPACITY]), scene());
        assert!(cache.get(&payloads[0]).is_some());
        assert!(cache.get(&payloads[1]).is_none());
        assert!(cache.get(&payloads[CAPACITY]).is_some());
    }
}
