//! Flat and packaged ODT access for page-layout header/footer properties.

use super::codec::parse_page_layout_header_footer_properties;
use super::model::Properties;
use crate::{FlatDocument, Package};
use litchi_core::Result;

impl Package {
    pub fn page_layout_header_footer_properties(&self) -> Result<Vec<Properties>> {
        self.styles_xml()?.map_or_else(
            || Ok(Vec::new()),
            |xml| parse_page_layout_header_footer_properties(&xml),
        )
    }
}

impl FlatDocument {
    pub fn page_layout_header_footer_properties(&self) -> Result<Vec<Properties>> {
        parse_page_layout_header_footer_properties(self.xml())
    }
}
