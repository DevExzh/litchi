//! Format-neutral access to any packaged OpenDocument document family.

use crate::constants;
use crate::core::{Meta, OwnedPackage};
use litchi_core::{Error, Metadata, Result};
use std::io::Read;
use std::path::Path;

/// Standard packaged OpenDocument document family.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum OpenDocumentFamily {
    /// Text document or template.
    Text,
    /// Spreadsheet document or template.
    Spreadsheet,
    /// Presentation document or template.
    Presentation,
    /// Drawing document or template.
    Drawing,
    /// Standalone chart document or template.
    Chart,
    /// Mathematical formula document or template.
    Formula,
    /// Image document or template.
    Image,
    /// Text master document.
    Master,
    /// Web-oriented text document.
    Web,
}

/// Validated, format-neutral OpenDocument package.
///
/// This provides lossless package access for every standard OpenDocument
/// family, including document types that do not yet have a specialized object
/// model. Saving an unmodified package returns the original bytes exactly.
pub struct OpenDocumentPackage {
    package: OwnedPackage,
    family: OpenDocumentFamily,
    template: bool,
    mimetype: String,
}

impl OpenDocumentPackage {
    /// Open and validate a packaged OpenDocument file.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }

    /// Read and validate a packaged OpenDocument file.
    pub fn from_reader(reader: impl Read) -> Result<Self> {
        let package = OwnedPackage::from_reader(reader)?;
        Self::from_owned(package)
    }

    /// Validate a packaged OpenDocument file from owned bytes.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = OwnedPackage::from_bytes(bytes)?;
        Self::from_owned(package)
    }

    fn from_owned(package: OwnedPackage) -> Result<Self> {
        // Constructing the borrowed package validates the mimetype file and
        // META-INF/manifest.xml together.
        let mimetype = package.mimetype()?;
        let (family, template) = classify_mimetype(&mimetype).ok_or_else(|| {
            Error::InvalidFormat(format!("unsupported OpenDocument mimetype '{mimetype}'"))
        })?;
        if !package.has_file(constants::ODF_CONTENT)? {
            return Err(Error::InvalidFormat(
                "OpenDocument package has no content.xml".to_string(),
            ));
        }
        // ODF XML parts are UTF-8 XML. Validate the mandatory part eagerly so
        // callers never receive a nominally valid package with binary content.
        let content = package.get_file(constants::ODF_CONTENT)?;
        std::str::from_utf8(&content)
            .map_err(|_| Error::InvalidFormat("invalid UTF-8 in content.xml".to_string()))?;

        Ok(Self {
            package,
            family,
            template,
            mimetype,
        })
    }

    /// Return the standard document family.
    pub fn family(&self) -> OpenDocumentFamily {
        self.family
    }

    /// Whether this package uses a template MIME type.
    pub fn is_template(&self) -> bool {
        self.template
    }

    /// Return the exact package MIME type.
    pub fn mimetype(&self) -> &str {
        &self.mimetype
    }

    /// Return the mandatory `content.xml` part.
    pub fn content_xml(&self) -> Result<String> {
        decode_xml_part(
            self.package.get_file(constants::ODF_CONTENT)?,
            constants::ODF_CONTENT,
        )
    }

    /// Return the optional `styles.xml` part.
    pub fn styles_xml(&self) -> Result<Option<String>> {
        self.optional_xml_part(constants::ODF_STYLES)
    }

    /// Return the optional `settings.xml` part.
    pub fn settings_xml(&self) -> Result<Option<String>> {
        self.optional_xml_part(constants::ODF_SETTINGS)
    }

    /// Extract common document metadata, or an empty value when `meta.xml` is absent.
    pub fn metadata(&self) -> Result<Metadata> {
        let Some(xml) = self.optional_xml_part(constants::ODF_META)? else {
            return Ok(Metadata::default());
        };
        Ok(Meta::from_bytes(xml.as_bytes())?.extract_metadata())
    }

    fn optional_xml_part(&self, path: &str) -> Result<Option<String>> {
        if !self.package.has_file(path)? {
            return Ok(None);
        }
        self.package
            .get_file(path)
            .and_then(|bytes| decode_xml_part(bytes, path))
            .map(Some)
    }

    /// List every path stored in the package.
    pub fn files(&self) -> Result<Vec<String>> {
        self.package.files()
    }

    /// Return whether a package path exists.
    pub fn has_file(&self, path: &str) -> Result<bool> {
        self.package.has_file(path)
    }

    /// Extract an arbitrary package part without interpreting it.
    pub fn get_file(&self, path: &str) -> Result<Vec<u8>> {
        self.package.get_file(path)
    }

    /// List package paths that appear to contain embedded media.
    pub fn media_files(&self) -> Result<Vec<String>> {
        self.package.media_files()
    }

    /// Return the original package bytes.
    pub fn as_bytes(&self) -> &[u8] {
        self.package.as_bytes()
    }

    /// Clone the original package bytes for lossless saving or transfer.
    pub fn to_bytes(&self) -> Vec<u8> {
        self.package.as_bytes().to_vec()
    }

    /// Consume this wrapper and return the original package bytes.
    pub fn into_bytes(self) -> Vec<u8> {
        self.package.into_inner()
    }

    /// Save the package without reconstructing or discarding any parts.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        std::fs::write(path, self.as_bytes())?;
        Ok(())
    }
}

fn decode_xml_part(bytes: Vec<u8>, path: &str) -> Result<String> {
    String::from_utf8(bytes).map_err(|_| Error::InvalidFormat(format!("invalid UTF-8 in {path}")))
}

fn classify_mimetype(mimetype: &str) -> Option<(OpenDocumentFamily, bool)> {
    Some(match mimetype {
        constants::ODF_TEXT => (OpenDocumentFamily::Text, false),
        constants::ODF_TEXT_TEMPLATE => (OpenDocumentFamily::Text, true),
        constants::ODF_SPREADSHEET => (OpenDocumentFamily::Spreadsheet, false),
        constants::ODF_SPREADSHEET_TEMPLATE => (OpenDocumentFamily::Spreadsheet, true),
        constants::ODF_PRESENTATION => (OpenDocumentFamily::Presentation, false),
        constants::ODF_PRESENTATION_TEMPLATE => (OpenDocumentFamily::Presentation, true),
        constants::ODF_DRAWING => (OpenDocumentFamily::Drawing, false),
        constants::ODF_DRAWING_TEMPLATE => (OpenDocumentFamily::Drawing, true),
        constants::ODF_CHART => (OpenDocumentFamily::Chart, false),
        constants::ODF_CHART_TEMPLATE => (OpenDocumentFamily::Chart, true),
        constants::ODF_FORMULA => (OpenDocumentFamily::Formula, false),
        constants::ODF_FORMULA_TEMPLATE => (OpenDocumentFamily::Formula, true),
        constants::ODF_IMAGE => (OpenDocumentFamily::Image, false),
        constants::ODF_IMAGE_TEMPLATE => (OpenDocumentFamily::Image, true),
        constants::ODF_MASTER => (OpenDocumentFamily::Master, false),
        constants::ODF_WEB => (OpenDocumentFamily::Web, false),
        _ => return None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::core::PackageWriter;

    fn package(mimetype: &str) -> Vec<u8> {
        let mut writer = PackageWriter::new();
        writer.set_mimetype(mimetype).unwrap();
        writer
            .add_file(
                constants::ODF_CONTENT,
                br#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body/></office:document-content>"#,
            )
            .unwrap();
        writer
            .add_file_with_media_type("Pictures/pixel.png", b"PNG", "image/png")
            .unwrap();
        writer.finish_to_bytes().unwrap()
    }

    #[test]
    fn opens_every_document_family_and_template_losslessly() {
        for (mimetype, family, template) in [
            (constants::ODF_TEXT, OpenDocumentFamily::Text, false),
            (constants::ODF_TEXT_TEMPLATE, OpenDocumentFamily::Text, true),
            (
                constants::ODF_SPREADSHEET,
                OpenDocumentFamily::Spreadsheet,
                false,
            ),
            (
                constants::ODF_SPREADSHEET_TEMPLATE,
                OpenDocumentFamily::Spreadsheet,
                true,
            ),
            (
                constants::ODF_PRESENTATION,
                OpenDocumentFamily::Presentation,
                false,
            ),
            (
                constants::ODF_PRESENTATION_TEMPLATE,
                OpenDocumentFamily::Presentation,
                true,
            ),
            (constants::ODF_DRAWING, OpenDocumentFamily::Drawing, false),
            (
                constants::ODF_DRAWING_TEMPLATE,
                OpenDocumentFamily::Drawing,
                true,
            ),
            (constants::ODF_CHART, OpenDocumentFamily::Chart, false),
            (
                constants::ODF_CHART_TEMPLATE,
                OpenDocumentFamily::Chart,
                true,
            ),
            (constants::ODF_FORMULA, OpenDocumentFamily::Formula, false),
            (
                constants::ODF_FORMULA_TEMPLATE,
                OpenDocumentFamily::Formula,
                true,
            ),
            (constants::ODF_IMAGE, OpenDocumentFamily::Image, false),
            (
                constants::ODF_IMAGE_TEMPLATE,
                OpenDocumentFamily::Image,
                true,
            ),
            (constants::ODF_MASTER, OpenDocumentFamily::Master, false),
            (constants::ODF_WEB, OpenDocumentFamily::Web, false),
        ] {
            let bytes = package(mimetype);
            let document = OpenDocumentPackage::from_bytes(bytes.clone()).unwrap();
            assert_eq!(document.family(), family);
            assert_eq!(document.is_template(), template);
            assert_eq!(document.mimetype(), mimetype);
            assert!(document.content_xml().unwrap().contains("office:body"));
            assert_eq!(document.media_files().unwrap(), ["Pictures/pixel.png"]);
            assert_eq!(document.to_bytes(), bytes);
            assert_eq!(document.into_bytes(), bytes);
        }
    }

    #[test]
    fn rejects_non_odf_missing_content_and_invalid_xml_bytes() {
        let mut writer = PackageWriter::new();
        writer.set_mimetype("application/zip").unwrap();
        writer.add_file(constants::ODF_CONTENT, b"<x/>").unwrap();
        assert!(OpenDocumentPackage::from_bytes(writer.finish_to_bytes().unwrap()).is_err());

        let mut writer = PackageWriter::new();
        writer.set_mimetype(constants::ODF_DRAWING).unwrap();
        assert!(OpenDocumentPackage::from_bytes(writer.finish_to_bytes().unwrap()).is_err());

        let mut writer = PackageWriter::new();
        writer.set_mimetype(constants::ODF_CHART).unwrap();
        writer.add_file(constants::ODF_CONTENT, &[0xff]).unwrap();
        assert!(OpenDocumentPackage::from_bytes(writer.finish_to_bytes().unwrap()).is_err());
    }
}
