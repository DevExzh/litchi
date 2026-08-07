//! Flat and packaged `OpenDocument` access for note configurations.

use super::Configurations;
use super::codec::parse;
use crate::{FlatDocument, Package};
use litchi_core::Result;

impl Package {
    pub fn notes_configurations(&self) -> Result<Configurations> {
        self.styles_xml()?
            .map_or_else(|| Ok(Configurations::default()), |xml| parse(&xml))
    }
}

impl FlatDocument {
    pub fn notes_configurations(&self) -> Result<Configurations> {
        parse(self.xml())
    }
}
