//! Flat-document and package adapters for bibliography metadata.

use super::codec::{parse_bibliography_configuration, parse_bibliography_configuration_parts};
use super::model::Configuration;
use crate::variable_declaration::Part;
use crate::{FlatDocument, Package};
use litchi_core::Result;

impl Package {
    /// Return the stored document-wide bibliography formatting policy.
    ///
    /// The policy is metadata in `styles.xml`. This method does not generate
    /// bibliography entries, resolve citations, or access external sources.
    pub fn bibliography_configuration(&self) -> Result<Option<Configuration>> {
        self.styles_xml()?
            .map_or_else(|| Ok(None), |xml| parse_bibliography_configuration(&xml))
    }
}

impl FlatDocument {
    /// Return the stored document-wide bibliography formatting policy.
    ///
    /// The policy is metadata in the flat document's `office:styles` element.
    /// This method does not generate bibliography entries or resolve citations.
    pub fn bibliography_configuration(&self) -> Result<Option<Configuration>> {
        parse_bibliography_configuration_parts(&[(self.xml(), Part::Flat)])
    }
}
