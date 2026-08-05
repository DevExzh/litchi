//! Package and flat-document adapters for list-level alignments.

use super::codec::parse;
use super::model::Styles;
use crate::{FlatDocument, Package};
use litchi_core::Result;

impl Package {
    pub fn alignments(&self) -> Result<Styles> {
        self.styles_xml()?
            .map_or_else(|| Ok(Default::default()), |xml| parse(&xml))
    }
}

impl FlatDocument {
    pub fn alignments(&self) -> Result<Styles> {
        parse(self.xml())
    }
}
