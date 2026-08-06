//! Typed Word caption metadata.
//!
//! The owner covers the two extended string tables addressed by the Word 97+
//! FIB: `SttbfCaption` and `SttbfAutoCaption`. It is deliberately a metadata
//! boundary. Caption labels and OLE ProgID references are inspected and
//! serialized, but no caption is inserted, no field is refreshed, and no host
//! or macro is opened or executed.

mod codec;
mod model;
mod semantic;
mod validation;

#[cfg(test)]
mod tests;

/// Shared MS-OSHARED numbering vocabulary used by caption numbers.
pub use crate::parts::numbering::NumberFormat as Format;

pub use model::{
    AutoEntry, AutoTable, Definition, Heading, Info, LabelTable, Location, Numbering, Separator,
};
pub use semantic::Tables;

// The document state predates the contextual facade and uses this name
// internally. It is crate-visible only; the public path is `captions::Tables`.
pub(crate) use semantic::Tables as CaptionTables;
