//! Bounded conversion from native TSWP storage messages to semantic text.
//!
//! This crate is the narrow wire adapter between generated IWA protobufs and
//! [`litchi_iwa_text`]. It owns no package traversal, archive state, object
//! lookup, or application semantics. Callers retain context-specific error
//! wording and transaction ownership at their format boundary.

#![forbid(unsafe_code)]

use litchi_iwa_protos::tswp::StorageArchive;
use litchi_iwa_text::storage::{Error as StorageError, MAX_RUNS, Run, Storage};

/// Maximum native text fragments accepted by one conversion.
pub const MAX_FRAGMENTS: usize = MAX_RUNS;

/// Why a native text-storage payload could not become a semantic value.
#[derive(Debug, Clone, PartialEq, Eq, thiserror::Error)]
#[non_exhaustive]
pub enum Error {
    /// The native payload contains more fragments than the semantic range
    /// representation can retain.
    #[error("text storage contains {actual} fragments; maximum is {limit}")]
    TooManyFragments { actual: usize, limit: usize },
    /// The aggregate UTF-8 length cannot be represented by the host address
    /// space.
    #[error("text storage text length overflows the host address space")]
    TextLengthOverflow,
    /// A fallible text or run allocation failed.
    #[error(transparent)]
    Common(#[from] litchi_iwa_common::Error),
    /// The resulting text/range relation failed semantic validation.
    #[error("semantic text storage is invalid: {0}")]
    Storage(#[source] StorageError),
}

/// Result type for native text-storage conversion.
pub type Result<T> = std::result::Result<T, Error>;

/// Convert one decoded native TSWP storage payload without retaining wire
/// fragments or allocating a second concatenated text buffer.
///
/// # Errors
///
/// Returns a typed error when the fragment budget, aggregate text length, or
/// allocation budget is exceeded, or when the semantic storage ranges cannot
/// be validated.
pub fn from_archive(archive: StorageArchive) -> Result<Storage> {
    if archive.text.len() > MAX_FRAGMENTS {
        return Err(Error::TooManyFragments {
            actual: archive.text.len(),
            limit: MAX_FRAGMENTS,
        });
    }

    let text_len = archive.text.iter().try_fold(0usize, |length, fragment| {
        length
            .checked_add(fragment.len())
            .ok_or(Error::TextLengthOverflow)
    })?;

    let mut text = String::new();
    text.try_reserve_exact(text_len).map_err(|_allocation| {
        litchi_iwa_common::Error::Allocation {
            resource: "native text storage",
            amount: text_len,
        }
    })?;

    let mut runs = Vec::new();
    runs.try_reserve_exact(archive.text.len())
        .map_err(|_allocation| litchi_iwa_common::Error::Allocation {
            resource: "native text storage runs",
            amount: archive.text.len(),
        })?;

    for fragment in archive.text {
        let start = text.len();
        let length = fragment.len();
        text.push_str(&fragment);
        runs.push(Run::new(start, length));
    }

    Storage::try_from_parts(text, runs).map_err(Error::Storage)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn conversion_keeps_fragment_ranges_in_one_owned_text_buffer() {
        let storage = from_archive(StorageArchive {
            text: vec!["Hello".to_owned(), " ".to_owned(), "world".to_owned()],
            ..StorageArchive::default()
        })
        .unwrap_or_else(|error| panic!("valid storage should convert: {error}"));

        assert_eq!(storage.text(), "Hello world");
        assert_eq!(
            storage.runs(),
            [Run::new(0, 5), Run::new(5, 1), Run::new(6, 5)]
        );
        assert_eq!(
            storage
                .fragments()
                .map(|fragment| fragment.text())
                .collect::<Vec<_>>(),
            ["Hello", " ", "world"]
        );
    }

    #[test]
    fn empty_native_storage_has_no_run_allocation() {
        let storage = from_archive(StorageArchive::default())
            .unwrap_or_else(|error| panic!("empty storage should convert: {error}"));

        assert!(storage.is_empty());
        assert!(storage.runs().is_empty());
    }

    #[test]
    fn fragment_limit_is_checked_before_materialization() {
        let archive = StorageArchive {
            text: vec![String::new(); MAX_FRAGMENTS + 1],
            ..StorageArchive::default()
        };

        assert_eq!(
            from_archive(archive),
            Err(Error::TooManyFragments {
                actual: MAX_FRAGMENTS + 1,
                limit: MAX_FRAGMENTS,
            })
        );
    }
}
