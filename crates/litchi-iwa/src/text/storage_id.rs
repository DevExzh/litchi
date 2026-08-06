//! Validated identity for one native iWork text storage.
//!
//! The identity belongs to the IWA adapter because it is only meaningful
//! while resolving a native storage object. The archive-free
//! `litchi-iwa-text` value layer deliberately does not depend on it.

use std::fmt;
use std::num::NonZeroU64;
use std::str::FromStr;

/// A non-null native identifier for one writable iWork text storage.
///
/// This is a distinct identity domain from drawable, graph, sheet, and
/// package object identifiers. Keeping it separate makes it impossible to
/// accidentally pass one of those other native identities to a text API.
#[repr(transparent)]
#[derive(Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct TextStorageId(NonZeroU64);

impl TextStorageId {
    /// Construct a text-storage identity, returning `None` for iWork's null
    /// sentinel.
    #[must_use]
    pub const fn new(raw: u64) -> Option<Self> {
        match NonZeroU64::new(raw) {
            Some(value) => Some(Self(value)),
            None => None,
        }
    }

    /// Return the native value at the IWA adapter boundary.
    #[must_use]
    pub(crate) const fn get(self) -> u64 {
        self.0.get()
    }
}

impl TryFrom<u64> for TextStorageId {
    type Error = TextStorageIdError;

    fn try_from(raw: u64) -> Result<Self, Self::Error> {
        Self::new(raw).ok_or(TextStorageIdError)
    }
}

impl fmt::Debug for TextStorageId {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_tuple("TextStorageId").field(&self.get()).finish()
    }
}

impl FromStr for TextStorageId {
    type Err = TextStorageIdError;

    fn from_str(value: &str) -> Result<Self, Self::Err> {
        value
            .parse::<u64>()
            .map_err(|_| TextStorageIdError)
            .and_then(Self::try_from)
    }
}

impl fmt::Display for TextStorageId {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.get().fmt(formatter)
    }
}

/// Error returned when a null native storage reference is used as an ID.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextStorageIdError;

impl fmt::Display for TextStorageIdError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("text storage identifier must be non-zero")
    }
}

impl std::error::Error for TextStorageIdError {}

#[cfg(test)]
mod tests {
    use super::TextStorageId;
    use std::mem::size_of;

    #[test]
    fn rejects_the_native_null_sentinel() {
        assert_eq!(TextStorageId::new(0), None);
        assert!(TextStorageId::try_from(0).is_err());
    }

    #[test]
    fn retains_the_full_native_range() {
        let id = TextStorageId::try_from(u64::MAX).expect("non-zero ID");
        assert_eq!(id.to_string(), u64::MAX.to_string());
        assert_eq!(format!("{id:?}"), format!("TextStorageId({})", u64::MAX));
    }

    #[test]
    fn remains_compact_and_option_optimized() {
        assert_eq!(size_of::<TextStorageId>(), size_of::<u64>());
        assert_eq!(size_of::<Option<TextStorageId>>(), size_of::<u64>());
    }
}
