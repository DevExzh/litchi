//! Package and flat-document accessors for line-numbering metadata.

use litchi_core::Result;

use super::{Configuration, parse};
use crate::{FlatDocument, Package};

impl Package {
    /// Return stored document line-numbering configuration from styles XML.
    ///
    /// The declaration is presentation metadata only. It is never used to
    /// paginate a document or generate line numbers.
    pub fn line_numbering_configuration(&self) -> Result<Option<Configuration>> {
        self.styles_xml()?
            .map_or_else(|| Ok(None), |xml| parse(&xml))
    }
}

impl FlatDocument {
    /// Return stored document line-numbering configuration from flat ODF XML.
    ///
    /// The declaration is presentation metadata only. It is never used to
    /// paginate a document or generate line numbers.
    pub fn line_numbering_configuration(&self) -> Result<Option<Configuration>> {
        parse(self.xml())
    }
}
