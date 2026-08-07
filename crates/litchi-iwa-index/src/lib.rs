#![forbid(unsafe_code)]
#![allow(
    clippy::module_name_repetitions,
    reason = "Focused modules retain explicit public names when re-exported at the crate root."
)]

//! Archive-free object-location indexing for iWork format adapters.
//!
//! This crate deliberately models only neutral identity, byte locations, and
//! references. It does not open packages, decode IWA frames, know protobuf
//! message types, or expose the archive's native identifiers. A concrete
//! adapter is responsible for validating a native object and translating it
//! into [`ObjectRecord`] plus [`Reference`] values before building a snapshot.

pub mod error;
pub mod fragment;
pub mod index;
pub mod object;
pub mod reference;
pub mod span;

pub use error::{AllocationKind, IndexError};
pub use fragment::{FragmentId, FragmentIdError};
pub use index::{IndexBuilder, ObjectIndex};
pub use litchi_iwa_graph::{ObjectId, ObjectIdError, ObjectIdIter, ReferenceGraphSnapshot};
pub use object::ObjectRecord;
pub use reference::{Reference, ReferenceError};
pub use span::{ByteSpan, ByteSpanError};
