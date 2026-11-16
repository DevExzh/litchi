//! The `Equation Native` stream of a listed equation, read on demand.
//!
//! Listing equations reads no storage. The first request for an object's
//! stream decompresses its storage once; the stream is memoized per persist
//! identifier on the presentation, so every frame of one object shares it,
//! and the bytes all storages decompress to are bounded by the
//! presentation's aggregate payload limit.

use std::collections::HashMap;
use std::sync::{Arc, Mutex};

use super::read;
use crate::Presentation;
use crate::embedded::storage::Storage;

/// Why an equation's `Equation Native` stream cannot be read. The equation
/// is still listed; only its content is missing.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Unreadable {
    reason: String,
}

impl Unreadable {
    pub(crate) fn new(reason: impl Into<String>) -> Self {
        Self {
            reason: reason.into(),
        }
    }

    /// What is wrong with the storage.
    #[must_use]
    pub fn reason(&self) -> &str {
        &self.reason
    }
}

impl std::fmt::Display for Unreadable {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(formatter, "the equation cannot be read: {}", self.reason)
    }
}

impl std::error::Error for Unreadable {}

/// A listed equation's stream, or why it cannot be read.
pub type Native = Result<Arc<[u8]>, Unreadable>;

/// Where a listed equation's stream comes from.
#[derive(Clone)]
pub(crate) enum Content {
    /// Read from the presentation's storage on first use.
    Deferred(Presentation),
    /// Already read (an edit's staged equations).
    Loaded(Result<Arc<[u8]>, Unreadable>),
}

impl std::fmt::Debug for Content {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Deferred(_) => formatter.write_str("Deferred"),
            Self::Loaded(Ok(native)) => write!(formatter, "Loaded({} bytes)", native.len()),
            Self::Loaded(Err(error)) => write!(formatter, "Loaded({error})"),
        }
    }
}

impl Content {
    pub(crate) fn native(&self, persist_id: u32) -> Result<Arc<[u8]>, Unreadable> {
        match self {
            Self::Loaded(native) => native.clone(),
            Self::Deferred(presentation) => {
                let budget = presentation.source_limits().max_aggregate_payload_bytes();
                presentation
                    .inner
                    .equation_memo
                    .native(persist_id, budget, |maximum| {
                        load(presentation.ole_storage(persist_id), maximum)
                    })
            },
        }
    }
}

/// The stream of a stored object, decompressing at most `maximum` bytes;
/// with the bytes decompressed.
///
/// The bytes the storage declares (its decompressed size, or its stored
/// size when uncompressed) are charged before anything is decompressed or
/// parsed, whether or not the stream is then found, so failing storages
/// spend the budget as well.
pub(crate) fn load(
    storage: crate::package::Result<Option<Storage>>,
    maximum: usize,
) -> (Result<Arc<[u8]>, Unreadable>, usize) {
    let storage = match storage {
        Ok(Some(storage)) => storage,
        Ok(None) => return (Err(Unreadable::new("the object has no storage")), 0),
        Err(error) => return (Err(Unreadable::new(error.to_string())), 0),
    };
    let declared = storage
        .declared_uncompressed_len()
        .map_or(storage.stored_payload_len(), |len| len as usize);
    if declared > maximum {
        return (Err(budget_reached()), 0);
    }
    (parse(&storage).map(Arc::from), declared)
}

/// Why an equation was not read: the storages read before it reached the
/// aggregate payload limit.
pub(crate) fn budget_reached() -> Unreadable {
    Unreadable::new(
        "reading this equation would exceed the aggregate payload limit of the equation storages",
    )
}

fn parse(storage: &Storage) -> Result<Vec<u8>, Unreadable> {
    let bytes = storage
        .decompressed_bytes(read::MAX_STORAGE_BYTES)
        .map_err(|error| Unreadable::new(error.to_string()))?;
    let mut file = litchi_cfb::OleFile::open(
        std::io::Cursor::new(bytes),
        litchi_cfb::OleFileLimits::default(),
    )
    .map_err(|error| Unreadable::new(error.to_string()))?;
    if !file.exists(&[read::EQUATION_NATIVE]) {
        return Err(Unreadable::new(
            "the object's storage holds no Equation Native stream",
        ));
    }
    file.open_stream(&[read::EQUATION_NATIVE])
        .map_err(|error| Unreadable::new(error.to_string()))
}

/// The streams read from one presentation, by persist identifier.
#[derive(Default)]
pub(crate) struct Memo {
    state: Mutex<MemoState>,
}

#[derive(Default)]
struct MemoState {
    /// Bytes decompressed so far.
    spent: usize,
    /// Storages decompressed so far.
    loads: usize,
    streams: HashMap<u32, Result<Arc<[u8]>, Unreadable>>,
}

impl Memo {
    /// The memoized stream of `persist_id`, loading it with at most the
    /// budget's remaining bytes on first use.
    fn native(
        &self,
        persist_id: u32,
        budget: usize,
        load: impl FnOnce(usize) -> (Result<Arc<[u8]>, Unreadable>, usize),
    ) -> Result<Arc<[u8]>, Unreadable> {
        let mut state = match self.state.lock() {
            Ok(state) => state,
            Err(poisoned) => poisoned.into_inner(),
        };
        if let Some(native) = state.streams.get(&persist_id) {
            return native.clone();
        }
        let remaining = budget.saturating_sub(state.spent);
        let result = if remaining == 0 {
            Err(budget_reached())
        } else {
            let (result, spent) = load(remaining);
            if spent > 0 {
                state.loads = state.loads.saturating_add(1);
            }
            state.spent = state.spent.saturating_add(spent);
            result
        };
        state.streams.insert(persist_id, result.clone());
        result
    }

    /// The number of storages decompressed so far.
    #[cfg(all(test, feature = "equation"))]
    pub(crate) fn loads(&self) -> usize {
        match self.state.lock() {
            Ok(state) => state.loads,
            Err(poisoned) => poisoned.into_inner().loads,
        }
    }
}
