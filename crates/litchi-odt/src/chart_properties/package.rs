//! Flat and packaged `OpenDocument` access for chart styles.

use super::codec::parse_chart_style_properties;
use super::model::StylePropertiesSet;
use crate::{FlatDocument, Package};
use litchi_core::Result;

impl Package {
    pub fn chart_style_properties(&self) -> Result<StylePropertiesSet> {
        self.styles_xml()?.map_or_else(
            || Ok(Default::default()),
            |xml| parse_chart_style_properties(&xml),
        )
    }
}
impl FlatDocument {
    pub fn chart_style_properties(&self) -> Result<StylePropertiesSet> {
        parse_chart_style_properties(self.xml())
    }
}
