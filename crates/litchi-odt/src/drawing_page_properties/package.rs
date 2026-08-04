//! Flat and packaged OpenDocument access for drawing-page styles.

use super::{Styles, parse_drawing_page_style_properties};
use crate::{FlatOpenDocument, OpenDocumentPackage};
use litchi_core::Result;

impl OpenDocumentPackage {
    pub fn drawing_page_style_properties(&self) -> Result<Styles> {
        self.styles_xml()?.map_or_else(
            || Ok(Default::default()),
            |xml| parse_drawing_page_style_properties(&xml),
        )
    }
}

impl FlatOpenDocument {
    pub fn drawing_page_style_properties(&self) -> Result<Styles> {
        parse_drawing_page_style_properties(self.xml())
    }
}
