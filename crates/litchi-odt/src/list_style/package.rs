//! Package and flat-document facades for list-style declarations.

use super::{Styles, parse};
use crate::{FlatDocument, GenericPackage as Package};
use litchi_core::Result;

impl Package {
    /// Parse the `text:list-style` declarations of the package `styles.xml`.
    pub fn styles(&self) -> Result<Styles> {
        self.styles_xml()?.map_or_else(
            || Ok(Styles::default()),
            |xml| parse(&xml, &self.xml_limits_value()),
        )
    }
}

impl FlatDocument {
    /// Parse the `text:list-style` declarations of a flat XML document.
    pub fn styles(&self) -> Result<Styles> {
        parse(self.xml(), &self.xml_limits_value())
    }
}
