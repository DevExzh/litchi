//! Flat-document and package adapters for bibliography metadata.

use super::codec::{
    parse_bibliography_configuration_parts_with_budget_and_xml_limits,
    parse_bibliography_configuration_with_xml_limits,
};
use super::model::Configuration;
use crate::variable_declaration::Part;
use crate::{FlatDocument, GenericPackage as Package};
use litchi_core::Result;

impl Package {
    /// Return the stored document-wide bibliography formatting policy.
    ///
    /// The policy is metadata in `styles.xml`. This method does not generate
    /// bibliography entries, resolve citations, or access external sources.
    pub fn bibliography_configuration(&self) -> Result<Option<Configuration>> {
        self.styles_xml()?.map_or_else(
            || Ok(None),
            |xml| parse_bibliography_configuration_with_xml_limits(&xml, self.xml_limits_value()),
        )
    }
}

impl FlatDocument {
    /// Return the stored document-wide bibliography formatting policy.
    ///
    /// The policy is metadata in the flat document's `office:styles` element.
    /// This method does not generate bibliography entries or resolve citations.
    pub fn bibliography_configuration(&self) -> Result<Option<Configuration>> {
        parse_bibliography_configuration_parts_with_budget_and_xml_limits(
            &[(self.xml(), Part::Flat)],
            None,
            self.xml_limits_value(),
        )
    }
}
