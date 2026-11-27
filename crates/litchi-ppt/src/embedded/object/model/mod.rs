//! Semantic metadata models for `PowerPoint` OLE object records.
//!
//! The model layer deliberately contains no record traversal or mutation
//! policy.  Wire-facing implementations live under `crate::codec`, while
//! collection invariants and snapshots live under `crate::validation`.

mod collection;
mod containers;
mod metadata;
mod unknown;

pub use collection::Collection;
pub use containers::{ContainerKind, Control, Definition, ExternalObject};
pub use metadata::{
    ColorFollow, DimensionPolicy, DrawAspect, EmbedPreferences, LinkInfo, Metadata, ObjectSubtype,
    ObjectType, UpdateMode,
};
pub(crate) use metadata::{MAX_METAFILE_BYTES, MAX_OLE_NAME_UNITS, MAX_OLE_OBJECTS};
pub use unknown::UnknownRecord;
