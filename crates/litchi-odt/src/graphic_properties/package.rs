//! Flat and packaged `OpenDocument` access for graphic styles.

use super::codec::parse_graphic_style_properties;
use super::model::Styles;
use crate::{FlatDocument, Package};
use litchi_core::Result;

impl Package {
    pub fn graphic_style_properties(&self) -> Result<Styles> {
        self.styles_xml()?.map_or_else(
            || Ok(Default::default()),
            |xml| parse_graphic_style_properties(&xml),
        )
    }
}
impl FlatDocument {
    pub fn graphic_style_properties(&self) -> Result<Styles> {
        parse_graphic_style_properties(self.xml())
    }
}
