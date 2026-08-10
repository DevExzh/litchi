//! Drawing package authoring.

use litchi_core::Result;
use litchi_odf_common::{
    compact_xml,
    core::{PackageWriter, Profile},
    signature::DocumentSigner,
};

/// Detached builder; publication validates through the package facade.
#[derive(Clone, Debug)]
pub struct Builder {
    content_xml: String,
}

impl Builder {
    /// Creates a builder pre-filled with an empty drawing content document.
    #[must_use]
    pub fn new() -> Self {
        Self {
            content_xml: empty_content().to_owned(),
        }
    }

    /// Replaces the `content.xml` payload.
    #[must_use]
    pub fn content_xml(mut self, xml: impl Into<String>) -> Self {
        self.content_xml = xml.into();
        self
    }

    /// Validates and packages the document bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if content validation or package writing fails.
    pub fn build(self) -> Result<Vec<u8>> {
        self.build_with_mimetype(crate::package::MIMETYPE)
    }

    /// Validates and packages an OTG drawing template.
    ///
    /// # Errors
    ///
    /// Returns an error if content validation or package writing fails.
    pub fn build_template(self) -> Result<Vec<u8>> {
        self.build_with_mimetype(crate::package::TEMPLATE_MIMETYPE)
    }

    /// Validates and packages a freshly authored password-encrypted drawing.
    ///
    /// # Errors
    ///
    /// Returns an error if validation, encryption configuration, or package writing fails.
    pub fn build_encrypted(self, password: impl Into<String>, profile: Profile) -> Result<Vec<u8>> {
        self.build_protected(
            crate::package::MIMETYPE,
            Some((password.into(), profile)),
            None,
        )
    }

    /// Validates and packages a freshly authored signed drawing.
    ///
    /// # Errors
    ///
    /// Returns an error if validation, signer configuration, or package writing fails.
    pub fn build_signed(self, signer: DocumentSigner) -> Result<Vec<u8>> {
        self.build_protected(crate::package::MIMETYPE, None, Some(signer))
    }

    /// Validates and packages a freshly authored encrypted and signed drawing.
    ///
    /// # Errors
    ///
    /// Returns an error if validation, protection configuration, or package writing fails.
    pub fn build_encrypted_and_signed(
        self,
        password: impl Into<String>,
        profile: Profile,
        signer: DocumentSigner,
    ) -> Result<Vec<u8>> {
        self.build_protected(
            crate::package::MIMETYPE,
            Some((password.into(), profile)),
            Some(signer),
        )
    }

    fn build_with_mimetype(self, mimetype: &str) -> Result<Vec<u8>> {
        self.build_protected(mimetype, None, None)
    }

    fn build_protected(
        self,
        mimetype: &str,
        encryption: Option<(String, Profile)>,
        signer: Option<DocumentSigner>,
    ) -> Result<Vec<u8>> {
        compact_xml::validate(self.content_xml.as_bytes())?;
        crate::codec::validate(&self.content_xml)?;
        let mut writer = PackageWriter::new();
        writer.set_mimetype(mimetype)?;
        if let Some(configured_signer) = signer {
            writer.set_document_signer(configured_signer)?;
        }
        if let Some((configured_password, profile)) = encryption {
            writer.set_encryption(configured_password, profile)?;
        }
        writer.add_file("content.xml", self.content_xml.as_bytes())?;
        writer.finish_to_bytes()
    }
}

impl Default for Builder {
    fn default() -> Self {
        Self::new()
    }
}

fn empty_content() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.3"><office:body><office:drawing/></office:body></office:document-content>"#
}
