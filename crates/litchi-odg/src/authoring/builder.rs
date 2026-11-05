//! Drawing package authoring.

use crate::package::Package;
use litchi_core::Result;
use litchi_odf_common::{
    Error as OdfError, compact_xml,
    core::{PackageWriter, Profile},
    signature::DocumentSigner,
};

/// Detached builder; [`Builder::build`] validates and publishes a [`Package`].
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

    /// Validates the content and publishes it as an immutable ODG [`Package`].
    ///
    /// Identical input produces identical bytes; no clock or randomness is
    /// consulted.
    ///
    /// # Errors
    ///
    /// Returns an error if content validation, package writing or readback
    /// fails.
    pub fn build(self) -> std::result::Result<Package, OdfError> {
        let bytes = self.build_protected(crate::package::MIMETYPE, None, None)?;
        Ok(Package::from_vec(bytes)?)
    }

    /// Validates the content and publishes it as an OTG drawing template.
    ///
    /// # Errors
    ///
    /// Returns an error if content validation, package writing or readback
    /// fails.
    pub fn build_template(self) -> std::result::Result<Package, OdfError> {
        let bytes = self.build_protected(crate::package::TEMPLATE_MIMETYPE, None, None)?;
        Ok(Package::from_template_vec(bytes)?)
    }

    /// Validates the content and publishes a freshly authored
    /// password-encrypted drawing. The returned package was reopened with
    /// `password` and is read-only like any encrypted package.
    ///
    /// # Errors
    ///
    /// Returns an error if validation, encryption configuration, package
    /// writing or readback fails.
    pub fn build_encrypted(
        self,
        password: impl Into<String>,
        profile: Profile,
    ) -> std::result::Result<Package, OdfError> {
        let password = password.into();
        let bytes = self.build_protected(
            crate::package::MIMETYPE,
            Some((password.clone(), profile)),
            None,
        )?;
        Ok(Package::from_encrypted_vec(bytes, password)?)
    }

    /// Validates the content and publishes a freshly authored signed drawing.
    ///
    /// # Errors
    ///
    /// Returns an error if validation, signer configuration, package writing
    /// or readback fails.
    pub fn build_signed(self, signer: DocumentSigner) -> std::result::Result<Package, OdfError> {
        let bytes = self.build_protected(crate::package::MIMETYPE, None, Some(signer))?;
        Ok(Package::from_vec(bytes)?)
    }

    /// Validates the content and publishes a freshly authored encrypted and
    /// signed drawing, reopened with `password`.
    ///
    /// # Errors
    ///
    /// Returns an error if validation, protection configuration, package
    /// writing or readback fails.
    pub fn build_encrypted_and_signed(
        self,
        password: impl Into<String>,
        profile: Profile,
        signer: DocumentSigner,
    ) -> std::result::Result<Package, OdfError> {
        let password = password.into();
        let bytes = self.build_protected(
            crate::package::MIMETYPE,
            Some((password.clone(), profile)),
            Some(signer),
        )?;
        Ok(Package::from_encrypted_vec(bytes, password)?)
    }

    fn build_protected(
        self,
        mimetype: &str,
        encryption: Option<(String, Profile)>,
        signer: Option<DocumentSigner>,
    ) -> Result<Vec<u8>> {
        compact_xml::validate(self.content_xml.as_bytes(), compact_xml::Limits::default())?;
        crate::codec::validate(&self.content_xml)?;
        let mut writer =
            PackageWriter::new(litchi_odf_common::core::StreamingArchiveLimits::default());
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
