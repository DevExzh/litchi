//! Archive-free comment values shared by the iWork format adapters.
//!
//! This module deliberately contains no package, archive, protobuf, or
//! application knowledge.  The concrete IWA adapter owns native object
//! traversal and wire updates; this leaf owns the validated values exchanged
//! by Pages, Numbers, and Keynote.

#![allow(
    clippy::module_name_repetitions,
    reason = "DrawableComment and TableCellComment are canonical semantic record names"
)]

use std::fmt;
use std::num::{NonZeroU32, NonZeroU64};

macro_rules! nonzero_id {
    ($name:ident, $raw:ty, $inner:ty, $kind:literal) => {
        /// A validated native identifier used by the comment semantic model.
        #[repr(transparent)]
        #[derive(Clone, Copy, Debug, Eq, Hash, Ord, PartialEq, PartialOrd)]
        pub struct $name($inner);

        impl $name {
            /// Construct an identifier, returning `None` for the native zero
            /// absence sentinel.
            #[must_use]
            pub const fn new(raw: $raw) -> Option<Self> {
                match <$inner>::new(raw) {
                    Some(value) => Some(Self(value)),
                    None => None,
                }
            }

            /// Construct an identifier and report a typed validation error.
            ///
            /// # Errors
            ///
            /// Returns [`Error::ZeroIdentifier`] when `raw` is zero.
            pub fn from_raw(raw: $raw) -> Result<Self> {
                Self::new(raw).ok_or(Error::ZeroIdentifier { kind: $kind })
            }

            /// Return the native identifier used at the adapter boundary.
            #[must_use]
            pub const fn get(self) -> $raw {
                self.0.get()
            }
        }

        impl TryFrom<$raw> for $name {
            type Error = Error;

            fn try_from(raw: $raw) -> Result<Self> {
                Self::from_raw(raw)
            }
        }

        impl From<$name> for $raw {
            fn from(value: $name) -> Self {
                value.get()
            }
        }

        impl fmt::Display for $name {
            fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
                self.get().fmt(formatter)
            }
        }
    };
}

/// Validation failures for a compact comment identifier.
#[derive(Clone, Copy, Debug, PartialEq, Eq, thiserror::Error)]
pub enum Error {
    /// A native zero sentinel was supplied where an object must exist.
    #[error("{kind} identifier must be non-zero")]
    ZeroIdentifier {
        /// Semantic kind of the rejected identifier.
        kind: &'static str,
    },
    /// A native all-zero UUID was supplied where a storage UUID must exist.
    #[error("comment UUID must be non-zero")]
    ZeroUuid,
}

/// Result type for comment-value validation.
pub type Result<T> = std::result::Result<T, Error>;

nonzero_id!(DrawableId, u64, NonZeroU64, "drawable");
nonzero_id!(StorageId, u64, NonZeroU64, "comment-storage");
nonzero_id!(AuthorId, u64, NonZeroU64, "comment-author");
nonzero_id!(ListId, u32, NonZeroU32, "comment-list");

/// A stable UUID stored on a comment-storage archive.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub struct Uuid {
    lower: u64,
    upper: u64,
}

impl Uuid {
    /// Construct a UUID, returning `None` for the native all-zero value.
    #[must_use]
    pub const fn new(lower: u64, upper: u64) -> Option<Self> {
        if lower == 0 && upper == 0 {
            None
        } else {
            Some(Self { lower, upper })
        }
    }

    /// Construct a UUID and report a typed validation error for zero.
    ///
    /// # Errors
    ///
    /// Returns [`Error::ZeroUuid`] when both native UUID halves are zero.
    pub const fn from_parts(lower: u64, upper: u64) -> Result<Self> {
        match Self::new(lower, upper) {
            Some(uuid) => Ok(uuid),
            None => Err(Error::ZeroUuid),
        }
    }

    /// Return the lower native UUID half.
    #[must_use]
    pub const fn lower(self) -> u64 {
        self.lower
    }

    /// Return the upper native UUID half.
    #[must_use]
    pub const fn upper(self) -> u64 {
        self.upper
    }
}

/// Archive-free value of one iWork comment-storage record.
#[derive(Clone, Debug, PartialEq)]
pub struct Comment {
    /// Comment text.
    pub text: String,
    /// Seconds since Apple's 2001-01-01 reference date.
    pub creation_date_seconds: Option<f64>,
    /// Native annotation-author identity, when present.
    pub author_id: Option<AuthorId>,
    /// Direct reply storage identities in native order.
    pub reply_ids: Box<[StorageId]>,
    /// Stable storage UUID, when present.
    pub storage_uuid: Option<Uuid>,
}

/// One drawable and its direct comment-storage attachment, if present.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct DrawableInfo {
    /// Drawable object identity.
    pub id: DrawableId,
    /// Native drawable message type retained for adapter dispatch.
    pub message_type: u32,
    /// Direct comment-storage identity, if attached.
    pub comment_id: Option<StorageId>,
}

/// A resolved direct comment attached to a drawable.
#[derive(Clone, Debug, PartialEq)]
pub struct DrawableComment {
    /// Drawable object identity.
    pub drawable_id: DrawableId,
    /// Root comment-storage identity.
    pub storage_id: StorageId,
    /// Decoded comment value.
    pub comment: Comment,
}

/// A resolved direct reply in a drawable comment thread.
#[derive(Clone, Debug, PartialEq)]
pub struct DrawableReply {
    /// Drawable object identity.
    pub drawable_id: DrawableId,
    /// Root comment-storage identity.
    pub root_storage_id: StorageId,
    /// Reply comment-storage identity.
    pub storage_id: StorageId,
    /// Decoded reply value.
    pub comment: Comment,
}

/// A resolved comment attached to one table cell.
#[derive(Clone, Debug, PartialEq)]
pub struct TableCellComment {
    /// Table model object identity.
    pub table_id: DrawableId,
    /// Zero-based row address.
    pub row: usize,
    /// Zero-based column address.
    pub column: usize,
    /// Native table comment-list identity.
    pub list_id: ListId,
    /// Root comment-storage identity.
    pub storage_id: StorageId,
    /// Decoded comment value.
    pub comment: Comment,
}

/// A resolved direct reply in a table-cell comment thread.
#[derive(Clone, Debug, PartialEq)]
pub struct TableCellReply {
    /// Table model object identity.
    pub table_id: DrawableId,
    /// Zero-based row address.
    pub row: usize,
    /// Zero-based column address.
    pub column: usize,
    /// Root comment-storage identity.
    pub root_storage_id: StorageId,
    /// Reply comment-storage identity.
    pub storage_id: StorageId,
    /// Decoded reply value.
    pub comment: Comment,
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::{AuthorId, Comment, DrawableId, DrawableInfo, ListId, StorageId, Uuid};

    #[test]
    fn identifiers_reject_zero_and_round_trip_native_values() {
        assert_eq!(DrawableId::new(0), None);
        assert_eq!(StorageId::new(0), None);
        assert_eq!(AuthorId::new(0), None);
        assert_eq!(ListId::new(0), None);
        assert_eq!(DrawableId::from_raw(7).unwrap().get(), 7);
        assert_eq!(StorageId::try_from(9).unwrap().get(), 9);
        assert_eq!(ListId::try_from(11).unwrap().get(), 11);
    }

    #[test]
    fn uuid_rejects_all_zero_and_is_two_words() {
        assert_eq!(Uuid::new(0, 0), None);
        let uuid = Uuid::from_parts(1, 2).unwrap();
        assert_eq!((uuid.lower(), uuid.upper()), (1, 2));
        assert_eq!(size_of::<Uuid>(), size_of::<(u64, u64)>());
        assert_eq!(size_of::<DrawableId>(), size_of::<u64>());
        assert_eq!(size_of::<StorageId>(), size_of::<u64>());
        assert_eq!(size_of::<ListId>(), size_of::<u32>());
    }

    #[test]
    fn records_use_canonical_typed_values() {
        let comment = Comment {
            text: "hello".to_owned(),
            creation_date_seconds: Some(1.0),
            author_id: Some(AuthorId::new(3).unwrap()),
            reply_ids: vec![StorageId::new(4).unwrap()].into_boxed_slice(),
            storage_uuid: Some(Uuid::new(5, 6).unwrap()),
        };
        let drawable = DrawableInfo {
            id: DrawableId::new(7).unwrap(),
            message_type: 8,
            comment_id: Some(StorageId::new(9).unwrap()),
        };
        assert_eq!(drawable.id.get(), 7);
        assert_eq!(comment.reply_ids[0].get(), 4);
    }
}
