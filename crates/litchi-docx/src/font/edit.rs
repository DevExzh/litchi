//! The font-table family (record 0355): the main document's font table and
//! its embedded-font resources.
//!
//! [`Package::fonts`](crate::Package::fonts) reads the table;
//! [`Package::edit_fonts`](crate::Package::edit_fonts) stages an [`Edit`] whose
//! commit publishes a new package, and
//! [`Package::apply_fonts`](crate::Package::apply_fonts) replays its [`Patch`].
//!
//! ```
//! use litchi_docx::{Package, font};
//!
//! let package = Package::new()?;
//! let mut table = package.fonts()?.unwrap_or_default();
//! table.add(font::Font::new("Example Sans")?)?;
//! let mut edit = package.edit_fonts();
//! edit.put(table, font::Conformance::Transitional)?;
//! let commit = edit.commit()?;
//! assert!(commit.package().fonts()?.unwrap().get("Example Sans")?.is_some());
//! # Ok::<(), litchi_docx::Error>(())
//! ```

use super::{Conformance, Table};
use crate::package::family::{graph_edit, graph_family};

graph_family!("font-table", own_limits);

pub use super::limit::LimitKind;
graph_edit!("font-table", edit_fonts, apply_fonts);

impl Edit {
    /// Move a complete font table into the package; returns whether the
    /// package changed.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid table and the signed-package refusal
    /// for a change to a signed package (ADR 0006).
    pub fn put(&mut self, table: Table, conformance: Conformance) -> Result<bool> {
        self.stage(|package| package.put_fonts(table, conformance))
    }

    /// Remove the font table and the font resources that become
    /// unreferenced; returns whether the package changed.
    ///
    /// # Errors
    ///
    /// Returns the signed-package refusal for a change to a signed package.
    pub fn remove(&mut self) -> Result<bool> {
        self.stage(crate::Package::remove_fonts)
    }
}
