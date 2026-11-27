//! Transactional ODP package ownership for the singleton handout master.

use super::{Master, Resolved, codec};
use crate::model::page_layout::Collection;
use crate::package::Presentation;
use litchi_core::{Error, Result};

impl Presentation {
    /// Read the package's optional, singleton handout master.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn handout_master(&self) -> Result<Option<Master>> {
        codec::read(styles(self)?)
    }

    /// Resolve the optional presentation-layout layer once and return an owned
    /// view.  ODF has no recursive handout-master inheritance chain.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn resolved_handout_master(&self) -> Result<Option<Resolved>> {
        let Some(master) = self.handout_master()? else {
            return Ok(None);
        };
        let layouts: Collection = self.layouts()?;
        master.resolve(&layouts).map(Some)
    }

    /// Insert or replace the singleton handout master (edit pipeline only).
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub(crate) fn set_handout_master(&mut self, master: &Master) -> Result<()> {
        let fragment = master.to_xml_fragment()?;
        let current = styles(self)?;
        let updated = codec::replace_in_styles(current, &fragment)?;
        if updated == current {
            return Ok(());
        }
        let content = self.content_xml().to_string();
        self.commit_design(&updated, &content)
    }

    /// Remove the optional handout master (edit pipeline only).
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub(crate) fn clear_handout_master(&mut self) -> Result<()> {
        let current = styles(self)?;
        let updated = codec::remove_from_styles(current)?;
        if updated == current {
            return Ok(());
        }
        let content = self.content_xml().to_string();
        self.commit_design(&updated, &content)
    }
}

fn error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn styles(presentation: &Presentation) -> Result<&str> {
    presentation
        .styles_xml()
        .ok_or_else(|| error("ODP package has no styles.xml"))
}
