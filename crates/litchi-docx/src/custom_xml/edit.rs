//! The Custom XML data-store family (record 0355).
//!
//! [`Package::custom_xml`](crate::Package::custom_xml) reads the stores;
//! [`Package::edit_custom_xml`](crate::Package::edit_custom_xml) stages an
//! [`Edit`] whose commit publishes a new package, and
//! [`Package::apply_custom_xml`](crate::Package::apply_custom_xml) replays its
//! [`Patch`].
//!
//! ```
//! use litchi_docx::{Package, custom_xml};
//!
//! let package = Package::new()?;
//! let mut edit = package.edit_custom_xml();
//! edit.add(custom_xml::NewStore {
//!     xml: b"<data/>".to_vec(),
//!     content_type: "application/xml".to_owned(),
//!     id: "{6C3C8BC8-F283-45A2-878A-457727A1A6AF}".to_owned(),
//!     schemas: Vec::new(),
//!     conformance: custom_xml::Conformance::Transitional,
//! })?;
//! let commit = edit.commit()?;
//! assert_eq!(commit.package().custom_xml()?.len(), 1);
//! assert!(package.custom_xml()?.is_empty());
//! # Ok::<(), litchi_docx::Error>(())
//! ```

use super::NewStore;
use crate::package::family::{graph_edit, graph_family};
use litchi_ooxml_common::custom_xml::Item;

graph_family!("Custom XML", own_limits);

/// The typed budgets of the Custom XML family: its limit errors are
/// [`litchi_ooxml_common::Error::Limit`] carried by
/// [`crate::Error::Common`] (record 0360).
pub use litchi_ooxml_common::custom_xml::LimitKind;
graph_edit!("Custom XML", edit_custom_xml, apply_custom_xml);

impl Edit {
    /// Add a data store; returns the stored item.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid payload or identifier and the
    /// signed-package refusal for a change to a signed package.
    pub fn add(&mut self, store: NewStore) -> Result<Item> {
        self.stage(|package| package.add_custom_xml(store))
    }

    /// Replace the payload of the store `id`.
    ///
    /// # Errors
    ///
    /// Returns an error for an unknown store or an invalid payload.
    pub fn set(&mut self, id: &str, xml: Vec<u8>) -> Result<&mut Self> {
        self.stage(|package| package.set_custom_xml(id, xml))?;
        Ok(self)
    }

    /// Replace the store `id` with `replacement`.
    ///
    /// # Errors
    ///
    /// Returns an error for an unknown store or an invalid replacement.
    pub fn replace(&mut self, id: &str, replacement: NewStore) -> Result<&mut Self> {
        self.stage(|package| package.replace_custom_xml(id, replacement))?;
        Ok(self)
    }

    /// Remove the store `id`; returns whether it existed.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid store graph.
    pub fn remove(&mut self, id: &str) -> Result<bool> {
        self.stage(|package| package.remove_custom_xml(id))
    }

    /// Reorder the stores to `ordered_ids`.
    ///
    /// # Errors
    ///
    /// Returns an error unless `ordered_ids` names every store exactly once.
    pub fn order(&mut self, ordered_ids: &[String]) -> Result<&mut Self> {
        self.stage(|package| package.order_custom_xml(ordered_ids))?;
        Ok(self)
    }
}
