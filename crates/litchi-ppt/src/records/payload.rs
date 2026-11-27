//! Storage for one PPT record's payload bytes.
//!
//! A payload is either owned by its record or a span of the shared stream
//! buffer the record tree was parsed from. Both forms expose exactly the same
//! bytes: [`RecordPayload`] dereferences to `[u8]`, compares by value, and
//! formats like a byte slice, so a reader cannot observe which form it holds.

use std::borrow::Borrow;
use std::fmt;
use std::ops::{Deref, DerefMut};
use std::sync::Arc;

/// The payload bytes of one PPT record.
///
/// Records parsed from a shared stream buffer borrow their bytes from it; every
/// other record owns them. [`Self::to_mut`] materializes an owned buffer on
/// demand, so a borrowed payload can still be edited in place.
#[derive(Clone, Default)]
pub struct RecordPayload {
    /// Where this payload's bytes live.
    storage: Storage,
}

/// The byte source behind a [`RecordPayload`].
#[derive(Clone)]
enum Storage {
    /// Bytes this record allocated for itself.
    Owned(Vec<u8>),
    /// A `start..end` span of a shared stream buffer.
    Shared {
        /// The stream the span indexes.
        buffer: Arc<Vec<u8>>,
        /// Inclusive span start within `buffer`.
        start: usize,
        /// Exclusive span end within `buffer`.
        end: usize,
    },
}

impl Default for Storage {
    fn default() -> Self {
        Self::Owned(Vec::new())
    }
}

impl RecordPayload {
    /// Creates an empty owned payload.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Borrows the payload bytes.
    #[must_use]
    pub fn as_slice(&self) -> &[u8] {
        match &self.storage {
            Storage::Owned(bytes) => bytes.as_slice(),
            Storage::Shared { buffer, start, end } => buffer.get(*start..*end).unwrap_or(&[][..]),
        }
    }

    /// Returns a mutable owned buffer, copying a borrowed span first.
    ///
    /// This follows `std::borrow::Cow::to_mut`: the borrowed form is replaced
    /// by an owned copy and then re-matched.
    pub fn to_mut(&mut self) -> &mut Vec<u8> {
        match self.storage {
            Storage::Owned(ref mut bytes) => bytes,
            ref mut storage => {
                *storage = Storage::Owned(match storage {
                    Storage::Shared { buffer, start, end } => {
                        buffer.get(*start..*end).unwrap_or(&[][..]).to_vec()
                    },
                    // `Storage::Owned` returned from the first arm.
                    Storage::Owned(bytes) => std::mem::take(bytes),
                });
                match storage {
                    Storage::Owned(bytes) => bytes,
                    // The assignment above stored `Storage::Owned`, exactly as
                    // `Cow::to_mut` re-matches what it just wrote.
                    _ => unreachable!("payload storage was just made owned"),
                }
            },
        }
    }

    /// Returns the payload as an owned vector, reusing an owned buffer.
    #[must_use]
    pub fn into_vec(self) -> Vec<u8> {
        match self.storage {
            Storage::Owned(bytes) => bytes,
            Storage::Shared { buffer, start, end } => {
                buffer.get(start..end).unwrap_or(&[][..]).to_vec()
            },
        }
    }

    /// Creates a payload that borrows `buffer[start..end]`.
    ///
    /// The span is validated by the record parser before this is called; a span
    /// outside `buffer` reads as empty rather than panicking.
    pub(crate) fn shared(buffer: &Arc<Vec<u8>>, start: usize, end: usize) -> Self {
        Self {
            storage: Storage::Shared {
                buffer: Arc::clone(buffer),
                start,
                end,
            },
        }
    }
}

impl Deref for RecordPayload {
    type Target = [u8];

    fn deref(&self) -> &Self::Target {
        self.as_slice()
    }
}

impl DerefMut for RecordPayload {
    /// Editing a borrowed payload in place copies its bytes first.
    fn deref_mut(&mut self) -> &mut Self::Target {
        self.to_mut().as_mut_slice()
    }
}

impl AsRef<[u8]> for RecordPayload {
    fn as_ref(&self) -> &[u8] {
        self.as_slice()
    }
}

impl Borrow<[u8]> for RecordPayload {
    fn borrow(&self) -> &[u8] {
        self.as_slice()
    }
}

impl fmt::Debug for RecordPayload {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        fmt::Debug::fmt(self.as_slice(), formatter)
    }
}

impl PartialEq for RecordPayload {
    fn eq(&self, other: &Self) -> bool {
        self.as_slice() == other.as_slice()
    }
}

impl Eq for RecordPayload {}

impl PartialEq<[u8]> for RecordPayload {
    fn eq(&self, other: &[u8]) -> bool {
        self.as_slice() == other
    }
}

impl PartialEq<&[u8]> for RecordPayload {
    fn eq(&self, other: &&[u8]) -> bool {
        self.as_slice() == *other
    }
}

impl PartialEq<Vec<u8>> for RecordPayload {
    fn eq(&self, other: &Vec<u8>) -> bool {
        self.as_slice() == other.as_slice()
    }
}

impl PartialEq<RecordPayload> for Vec<u8> {
    fn eq(&self, other: &RecordPayload) -> bool {
        self.as_slice() == other.as_slice()
    }
}

impl<const N: usize> PartialEq<[u8; N]> for RecordPayload {
    fn eq(&self, other: &[u8; N]) -> bool {
        self.as_slice() == other.as_slice()
    }
}

impl From<Vec<u8>> for RecordPayload {
    fn from(value: Vec<u8>) -> Self {
        Self {
            storage: Storage::Owned(value),
        }
    }
}

impl From<&[u8]> for RecordPayload {
    fn from(value: &[u8]) -> Self {
        Self::from(value.to_vec())
    }
}

impl<const N: usize> From<[u8; N]> for RecordPayload {
    fn from(value: [u8; N]) -> Self {
        Self::from(value.to_vec())
    }
}

impl<const N: usize> From<&[u8; N]> for RecordPayload {
    fn from(value: &[u8; N]) -> Self {
        Self::from(value.to_vec())
    }
}

impl From<&Vec<u8>> for RecordPayload {
    fn from(value: &Vec<u8>) -> Self {
        Self::from(value.clone())
    }
}

impl From<RecordPayload> for Vec<u8> {
    fn from(value: RecordPayload) -> Self {
        value.into_vec()
    }
}

impl FromIterator<u8> for RecordPayload {
    fn from_iter<I: IntoIterator<Item = u8>>(iter: I) -> Self {
        Self::from(Vec::from_iter(iter))
    }
}

impl<'a> IntoIterator for &'a RecordPayload {
    type IntoIter = std::slice::Iter<'a, u8>;
    type Item = &'a u8;

    fn into_iter(self) -> Self::IntoIter {
        self.as_slice().iter()
    }
}

#[cfg(test)]
mod tests {
    use super::{Arc, RecordPayload};

    #[test]
    fn shared_and_owned_payloads_are_value_identical() {
        let buffer = Arc::new(vec![1u8, 2, 3, 4, 5]);
        let shared = RecordPayload::shared(&buffer, 1, 4);
        let owned = RecordPayload::from(vec![2u8, 3, 4]);
        assert_eq!(shared, owned);
        assert_eq!(shared.as_slice(), owned.as_slice());
        assert_eq!(format!("{shared:?}"), format!("{owned:?}"));
        assert_eq!(shared.len(), 3);
        assert_eq!(shared.to_vec(), vec![2u8, 3, 4]);
    }

    #[test]
    fn to_mut_copies_a_shared_span_before_editing() {
        let buffer = Arc::new(vec![1u8, 2, 3]);
        let mut payload = RecordPayload::shared(&buffer, 0, 3);
        payload.to_mut().push(4);
        assert_eq!(payload.as_slice(), &[1, 2, 3, 4]);
        assert_eq!(buffer.as_slice(), &[1, 2, 3]);
    }

    #[test]
    fn an_out_of_range_span_reads_as_empty() {
        let buffer = Arc::new(vec![1u8, 2, 3]);
        let payload = RecordPayload::shared(&buffer, 2, 9);
        assert!(payload.is_empty());
    }

    #[test]
    fn a_payload_adds_one_word_to_a_record() {
        // The shared span must not make every record noticeably larger; the
        // enum is the size of its largest variant plus a tag.
        assert_eq!(size_of::<RecordPayload>(), 32);
    }

    #[test]
    fn into_vec_reuses_an_owned_buffer() {
        let payload = RecordPayload::from(vec![7u8; 4]);
        assert_eq!(payload.into_vec(), vec![7u8; 4]);
        let buffer = Arc::new(vec![9u8; 4]);
        assert_eq!(RecordPayload::shared(&buffer, 1, 3).into_vec(), vec![9, 9]);
    }
}
