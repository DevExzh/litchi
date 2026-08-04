//! Shared indexed storage for immutable iWork package snapshots.

use std::collections::HashMap;
use std::panic::{self, AssertUnwindSafe};
use std::sync::{Arc, Condvar, Mutex};

use crate::archive::{Archive, ArchiveLimits};
use crate::{Error, Result};

type SharedArchiveResult = std::result::Result<Arc<Archive>, Arc<Error>>;

#[derive(Debug)]
struct ArchiveFlight {
    state: Mutex<ArchiveFlightState>,
    completed: Condvar,
}

#[derive(Debug)]
enum ArchiveFlightState {
    Parsing,
    Complete(SharedArchiveResult),
}

impl ArchiveFlight {
    fn new() -> Self {
        Self {
            state: Mutex::new(ArchiveFlightState::Parsing),
            completed: Condvar::new(),
        }
    }

    fn complete(&self, result: SharedArchiveResult) {
        let mut state = self
            .state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        *state = ArchiveFlightState::Complete(result);
        self.completed.notify_all();
    }

    fn wait(&self) -> SharedArchiveResult {
        let mut state = self
            .state
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        loop {
            match &*state {
                ArchiveFlightState::Parsing => {
                    state = self
                        .completed
                        .wait(state)
                        .unwrap_or_else(|poisoned| poisoned.into_inner());
                },
                ArchiveFlightState::Complete(result) => return result.clone(),
            }
        }
    }
}

#[derive(Debug, Default)]
struct ArchiveCache {
    /// The bounded completed-archive cache retained for this package state.
    ready: Option<(String, Arc<Archive>)>,
    /// Active parses are indexed by name so unrelated archives do not wait
    /// for one another. Entries are removed after their result is published.
    flights: HashMap<String, Arc<ArchiveFlight>>,
}

#[derive(Debug)]
enum ArchiveLookup {
    Cached(Arc<Archive>),
    Wait(Arc<ArchiveFlight>),
    Parse(Arc<ArchiveFlight>),
}

/// Package entries plus an index for validated name lookups.
///
/// The state is kept behind an `Arc` by both the mutable package and immutable
/// snapshot types. Structural edits clone this state once, then rebuild the
/// small name index; read-only clones never duplicate either the entry bytes,
/// the index, or the single bounded parsed-archive cache.
#[derive(Debug, Default)]
pub(crate) struct PackageState {
    pub(crate) entries: Vec<(String, Vec<u8>)>,
    positions: HashMap<String, usize>,
    archive_limits: ArchiveLimits,
    parsed_archive: Mutex<ArchiveCache>,
}

impl Clone for PackageState {
    fn clone(&self) -> Self {
        let ready = self
            .parsed_archive
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .ready
            .clone();
        Self {
            entries: self.entries.clone(),
            positions: self.positions.clone(),
            archive_limits: self.archive_limits,
            // Never carry active flights into a copy-on-write generation. A
            // flight is tied to the exact immutable entry bytes in its source
            // state; the detached state must be able to parse independently.
            parsed_archive: Mutex::new(ArchiveCache {
                ready,
                flights: HashMap::new(),
            }),
        }
    }
}

impl PackageState {
    pub(crate) fn from_entries(
        entries: Vec<(String, Vec<u8>)>,
        archive_limits: ArchiveLimits,
    ) -> Self {
        let mut state = Self {
            entries,
            positions: HashMap::new(),
            archive_limits,
            parsed_archive: Mutex::new(ArchiveCache::default()),
        };
        state.rebuild_positions();
        state
    }

    pub(crate) fn position(&self, name: &str) -> Option<usize> {
        self.positions.get(name).copied()
    }

    pub(crate) fn rebuild_positions(&mut self) {
        self.positions.clear();
        self.positions.reserve(self.entries.len());
        for (position, (name, _)) in self.entries.iter().enumerate() {
            let previous = self.positions.insert(name.clone(), position);
            debug_assert!(previous.is_none(), "package entry names must be unique");
        }
        self.clear_parsed_archive();
    }

    /// Parse or wait for one archive without holding the cache lock during
    /// decompression or protobuf parsing.
    pub(crate) fn get_or_parse_archive<F>(&self, name: &str, parse: F) -> Result<Arc<Archive>>
    where
        F: FnOnce(ArchiveLimits) -> Result<Archive>,
    {
        let lookup = self.lookup_archive(name);
        let flight = match lookup {
            ArchiveLookup::Cached(archive) => return Ok(archive),
            ArchiveLookup::Wait(flight) => {
                return self.wait_for_archive(&flight);
            },
            ArchiveLookup::Parse(flight) => flight,
        };

        let parsed = match panic::catch_unwind(AssertUnwindSafe(|| parse(self.archive_limits))) {
            Ok(result) => result.map(Arc::new),
            Err(payload) => {
                self.publish_archive(
                    name,
                    &flight,
                    Err(Arc::new(Error::Archive(
                        "IWA archive parser panicked".to_owned(),
                    ))),
                );
                panic::resume_unwind(payload);
            },
        };
        let shared = parsed
            .as_ref()
            .map(Arc::clone)
            .map_err(|error| Arc::new(clone_error(error)));
        self.publish_archive(name, &flight, shared);
        parsed
    }

    pub(crate) fn invalidate_archive(&mut self, name: &str) {
        let cache = self
            .parsed_archive
            .get_mut()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if cache
            .ready
            .as_ref()
            .is_some_and(|(cached_name, _)| cached_name == name)
        {
            cache.ready = None;
        }
    }

    fn lookup_archive(&self, name: &str) -> ArchiveLookup {
        let mut cache = self
            .parsed_archive
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if let Some((cached_name, archive)) = cache.ready.as_ref()
            && cached_name == name
        {
            return ArchiveLookup::Cached(Arc::clone(archive));
        }
        if let Some(flight) = cache.flights.get(name) {
            return ArchiveLookup::Wait(Arc::clone(flight));
        }
        let flight = Arc::new(ArchiveFlight::new());
        cache.flights.insert(name.to_owned(), Arc::clone(&flight));
        ArchiveLookup::Parse(flight)
    }

    fn wait_for_archive(&self, flight: &ArchiveFlight) -> Result<Arc<Archive>> {
        flight.wait().map_err(|error| clone_error(error.as_ref()))
    }

    fn publish_archive(
        &self,
        name: &str,
        flight: &Arc<ArchiveFlight>,
        result: SharedArchiveResult,
    ) {
        let current = {
            let mut cache = self
                .parsed_archive
                .lock()
                .unwrap_or_else(|poisoned| poisoned.into_inner());
            let current = cache
                .flights
                .get(name)
                .is_some_and(|current| Arc::ptr_eq(current, flight));
            if current && let Ok(archive) = result.as_ref() {
                cache.ready = Some((name.to_owned(), Arc::clone(archive)));
            }
            current
        };

        // Complete before removing the slot. A caller that arrives in this
        // small window still receives this generation's result; the first
        // caller after removal is the one permitted to retry a failure.
        flight.complete(result);

        if current {
            let mut cache = self
                .parsed_archive
                .lock()
                .unwrap_or_else(|poisoned| poisoned.into_inner());
            if cache
                .flights
                .get(name)
                .is_some_and(|registered| Arc::ptr_eq(registered, flight))
            {
                cache.flights.remove(name);
            }
        }
    }

    fn clear_parsed_archive(&mut self) {
        let cache = self
            .parsed_archive
            .get_mut()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        // `&mut self`/Arc::make_mut guarantees no active reader can be using
        // this state while structural entries are rebuilt. Keeping the map
        // untouched here also avoids stranding a waiter if that invariant is
        // ever violated by an internal caller.
        cache.ready = None;
    }
}

fn clone_error(error: &Error) -> Error {
    match error {
        Error::Io(error) => Error::Io(std::io::Error::new(error.kind(), error.to_string())),
        Error::IwaCore(error) => Error::IwaCore(error.clone()),
        Error::InvalidFormat(message) => Error::InvalidFormat(message.clone()),
        Error::Snappy(message) => Error::Snappy(message.clone()),
        Error::ProtobufDecode(error) => Error::ProtobufDecode(error.clone()),
        Error::UnsupportedMessageType(type_) => Error::UnsupportedMessageType(*type_),
        Error::Archive(message) => Error::Archive(message.clone()),
        Error::Bundle(message) => Error::Bundle(message.clone()),
        Error::ParseError(message) => Error::ParseError(message.clone()),
    }
}
