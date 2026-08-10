//! Flat and packaged ODT access for page-layout header/footer properties.

use super::codec::parse_page_layout_header_footer_properties;
use super::model::Properties;
use crate::{FlatDocument, Package};
use litchi_core::Result;

impl Package {
    /// Capture the optional `styles.xml` advanced-layout owner for isolated edits.
    ///
    /// The snapshot retains exact XML for source-checked patches while exposing
    /// typed master pages, page layouts, and section layouts. Package
    /// publication remains an explicit caller step so a layout patch cannot
    /// silently invalidate a signed or encrypted container.
    pub fn layout_snapshot(&self) -> Result<Option<super::Snapshot>> {
        self.styles_xml()?.map(super::Snapshot::parse).transpose()
    }

    pub fn page_layout_header_footer_properties(&self) -> Result<Vec<Properties>> {
        self.styles_xml()?.map_or_else(
            || Ok(Vec::new()),
            |xml| parse_page_layout_header_footer_properties(&xml),
        )
    }
}

impl FlatDocument {
    /// Capture this flat document's advanced-layout owner for isolated edits.
    pub fn layout_snapshot(&self) -> Result<super::Snapshot> {
        super::Snapshot::parse(self.xml().to_string())
    }

    pub fn page_layout_header_footer_properties(&self) -> Result<Vec<Properties>> {
        parse_page_layout_header_footer_properties(self.xml())
    }
}
