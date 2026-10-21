//! Typed picture-name metadata from `[MS-ODRAW]` section 2.3.23.
//!
//! [`Package`] owns an immutable admitted record; [`View`] is its read-only
//! projection. Edits publish the shared commit triad and preserve every
//! untouched descriptor and complex payload. Patches replay on matching bytes.
//!
//! ```
//! use litchi_odraw::prop::picture::Package;
//!
//! // An empty Opt record: version 3, no properties, no body bytes.
//! let input = [3, 0, 0x0b, 0xf0, 0, 0, 0, 0];
//! let source = Package::open(&input)?;
//! let mut edit = source.edit();
//! edit.set_name("diagram")?;
//! let committed = edit.commit()?;
//! assert!(committed.diagnostics().changed());
//! let mut output = Vec::new();
//! committed.package().write_to(&mut output)?;
//! let reopened = Package::open(&output)?;
//! let restored = reopened.apply(&committed.patch().inverse())?;
//! assert_eq!(restored.package().to_bytes(), input);
//! # Ok::<(), litchi_odraw::prop::picture::Error>(())
//! ```

mod error;
mod model;
mod options;
mod package;
mod validation;

pub use error::{Error, LimitKind, Refusal, RefusalKind, Result};
pub use model::{Change, Flags, Kind, Metadata, Name};
pub use options::{EditOptions, Limits, OpenOptions};
pub use package::{Commit, Diagnostics, Edit, Package, Patch, View};

pub(crate) fn parse_properties<'data>(
    properties: &super::Props<'data>,
    limits: Limits,
) -> Result<Option<Metadata<'data>>> {
    validation::decode(properties, limits)
}

#[cfg(test)]
mod tests;
