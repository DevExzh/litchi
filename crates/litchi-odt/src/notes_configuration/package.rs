//! Flat and packaged `OpenDocument` access for note configurations.

use super::Configurations;
use super::codec::parse;
use crate::{FlatDocument, GenericPackage as Package};
use litchi_core::Result;

impl Package {
    pub fn notes_configurations(&self) -> Result<Configurations> {
        self.styles_xml()?.map_or_else(
            || Ok(Configurations::default()),
            |xml| parse(&xml, &self.xml_limits_value()),
        )
    }
}

impl FlatDocument {
    pub fn notes_configurations(&self) -> Result<Configurations> {
        parse(self.xml(), &self.xml_limits_value())
    }
}
