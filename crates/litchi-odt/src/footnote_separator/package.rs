//! Flat and packaged `OpenDocument` access for footnote separators.

use litchi_core::Result;

use super::{MAX_SEPARATORS, Separator, codec::parse, invalid};

impl crate::Package {
    pub fn style_footnote_separators(&self) -> Result<Vec<Separator>> {
        let mut values = parse(self.styles_xml()?.as_deref().unwrap_or_default())?;
        values.extend(parse(&self.content_xml()?)?);
        if values.len() > MAX_SEPARATORS {
            return invalid("package exceeds 65536 style:footnote-sep values");
        }
        Ok(values)
    }
}

impl crate::FlatDocument {
    pub fn style_footnote_separators(&self) -> Result<Vec<Separator>> {
        parse(self.xml())
    }
}
