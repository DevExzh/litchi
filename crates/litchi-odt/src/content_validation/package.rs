//! Flat and packaged ODF access for content-validation metadata.

use super::codec::parse_part;
use super::{ContentValidationPart, ContentValidations};
use crate::{FlatDocument, Package};
use litchi_core::Result;

impl Package {
    pub fn content_validations(&self) -> Result<ContentValidations> {
        parse_part(&self.content_xml()?, ContentValidationPart::Content)
    }
}

impl FlatDocument {
    pub fn content_validations(&self) -> Result<ContentValidations> {
        super::parse_content_validations(self.xml())
    }
}
