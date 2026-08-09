//! Transactional ODP package ownership for the singleton handout master.

use super::{Master, Resolved, codec};
use crate::model::page_layout::Collection;
use crate::package::Presentation;
use litchi_core::{Error, Result};

fn error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

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

    /// Insert or replace the singleton handout master atomically.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn set_handout_master(&mut self, master: &Master) -> Result<()> {
        let fragment = master.to_xml_fragment()?;
        let current = styles(self)?;
        let updated = codec::replace_in_styles(current, &fragment)?;
        if updated == current {
            return Ok(());
        }
        let content = self.content_xml().to_string();
        self.commit_design(&updated, &content)
    }

    /// Insert a handout master, failing if the package already has one.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn add_handout_master(&mut self, master: &Master) -> Result<()> {
        if self.handout_master()?.is_some() {
            return Err(error("ODP package already has a handout master"));
        }
        self.set_handout_master(master)
    }

    /// Replace a handout master, failing if the package has none.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn replace_handout_master(&mut self, master: &Master) -> Result<()> {
        if self.handout_master()?.is_none() {
            return Err(error("ODP package has no handout master"));
        }
        self.set_handout_master(master)
    }

    /// Remove the optional handout master atomically.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn clear_handout_master(&mut self) -> Result<()> {
        let current = styles(self)?;
        let updated = codec::remove_from_styles(current)?;
        if updated == current {
            return Ok(());
        }
        let content = self.content_xml().to_string();
        self.commit_design(&updated, &content)
    }

    /// Alias with the same verb used by other singleton package owners.
    ///
    /// # Errors
    /// Returns an error when a package part is missing, malformed, or exceeds a configured limit.
    pub fn remove_handout_master(&mut self) -> Result<()> {
        self.clear_handout_master()
    }
}

fn styles(presentation: &Presentation) -> Result<&str> {
    presentation
        .styles_xml()
        .ok_or_else(|| error("ODP package has no styles.xml"))
}
