//! The document-variable family (record 0355): the settings part's
//! `w:docVars`.
//!
//! [`Package::document_variables`](crate::Package::document_variables) reads
//! them; [`Package::edit_document_variables`](crate::Package::edit_document_variables)
//! stages an [`Edit`] whose commit publishes a new package, and
//! [`Package::apply_document_variables`](crate::Package::apply_document_variables)
//! replays its [`Patch`]. Only the variables element is rewritten; every other
//! settings byte is kept.
//!
//! ```
//! use litchi_docx::Package;
//!
//! let package = Package::new()?;
//! let mut edit = package.edit_document_variables();
//! assert_eq!(edit.set("Client", "Acme")?, None);
//! let commit = edit.commit()?;
//! let variables = commit.package().document_variables()?.unwrap();
//! assert_eq!(variables.get("Client"), Some("Acme"));
//! # Ok::<(), litchi_docx::Error>(())
//! ```

use crate::package::family::{graph_edit, graph_family};

graph_family!("document-variable", own_limits);

/// The budgets of this family: the settings part it rewrites.
pub use crate::settings::LimitKind;
graph_edit!(
    "document-variable",
    edit_document_variables,
    apply_document_variables
);

impl Edit {
    /// Set a document variable; returns its previous value.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid name, value or settings part.
    pub fn set(
        &mut self,
        name: impl Into<String>,
        value: impl Into<String>,
    ) -> Result<Option<String>> {
        let (name, value) = (name.into(), value.into());
        self.stage(|package| package.set_document_variable(name, value))
    }

    /// Remove a document variable; returns its value, if it was set.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid settings part.
    pub fn remove(&mut self, name: &str) -> Result<Option<String>> {
        self.stage(|package| package.remove_document_variable(name))
    }

    /// Remove every document variable; returns how many were set.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid settings part.
    pub fn clear(&mut self) -> Result<usize> {
        self.stage(crate::Package::clear_document_variables)
    }
}
