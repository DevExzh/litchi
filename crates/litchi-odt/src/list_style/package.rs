//! Package and flat-document facades for list-style declarations.

use super::{Styles, parse};
use crate::{FlatDocument, Package};
use litchi_core::Result;

impl Package {
    /// Parse the `text:list-style` declarations of the package `styles.xml`.
    pub fn styles(&self) -> Result<Styles> {
        self.styles_xml()?
            .map_or_else(|| Ok(Styles::default()), |xml| parse(&xml))
    }
}

impl FlatDocument {
    /// Parse the `text:list-style` declarations of a flat XML document.
    pub fn styles(&self) -> Result<Styles> {
        parse(self.xml())
    }
}
