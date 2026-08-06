//! Typed OfficeArt solver-rule atoms.
//!
//! The owner is deliberately limited to the fixed-layout rule records defined
//! by `[MS-ODRAW]` sections 2.2.34--2.2.36.  Known records are decoded into
//! semantic fields, while unrecognized OfficeArt record types remain borrowed
//! and byte-for-byte writable through [`Opaque`].
//!
//! The implementation is split by responsibility: [`model`] owns the typed
//! rule values, [`codec`] performs record decoding and writing, and
//! [`validation`] checks the record header and fixed body lengths.

mod codec;
mod model;
mod validation;

pub use codec::{from_record, parse};
pub use model::{Arc, Callout, Connector, Kind, Opaque, Rule};

#[cfg(test)]
mod tests;
