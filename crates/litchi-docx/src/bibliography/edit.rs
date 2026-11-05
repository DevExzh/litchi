//! The bibliography-source family (record 0355): the sources of the
//! document's bibliography Custom XML store.
//!
//! [`Package::bibliography_sources`](crate::Package::bibliography_sources)
//! reads them; [`Package::edit_bibliography`](crate::Package::edit_bibliography)
//! stages an [`Edit`] whose commit publishes a new package, and
//! [`Package::apply_bibliography`](crate::Package::apply_bibliography) replays
//! its [`Patch`].

use super::BibliographySourceBuilder;
use crate::package::family::{graph_edit, graph_family};

graph_family!("bibliography-source", own_limits);

pub use super::limit::LimitKind;
graph_edit!("bibliography-source", edit_bibliography, apply_bibliography);

impl Edit {
    /// Add a source, creating the bibliography store when the package has
    /// none; returns the store's item identifier.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid source or store.
    pub fn add_source(&mut self, source: BibliographySourceBuilder) -> Result<String> {
        self.stage(|package| package.add_bibliography_source(source))
    }

    /// Remove the source tagged `tag`; returns whether it existed.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid store.
    pub fn remove_source(&mut self, tag: &str) -> Result<bool> {
        self.stage(|package| package.remove_bibliography_source(tag))
    }

    /// Replace the source tagged `tag`.
    ///
    /// # Errors
    ///
    /// Returns an error when there is no store or no such source.
    pub fn replace_source(
        &mut self,
        tag: &str,
        source: BibliographySourceBuilder,
    ) -> Result<&mut Self> {
        self.stage(|package| package.replace_bibliography_source(tag, source))?;
        Ok(self)
    }
}
