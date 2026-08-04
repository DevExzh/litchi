//! Read-only access to legacy OpenDocument web templates.

use crate::constants;
use crate::core::OwnedPackage;
use crate::odt::Document;
use litchi_core::{Error, Metadata, Result};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::io::Read;
use std::path::Path;

const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const MAX_DEPTH: usize = 128;

/// A validated `.oth` Writer/Web template.
///
/// The `application/vnd.oasis.opendocument.text-web` MIME type is a legacy
/// LibreOffice/odfpy/odfdo convention rather than an ODF 1.3 or 1.4 conforming
/// text MIME type. Its package content uses the standard `office:text` model,
/// exposed through [`Self::document`].
pub struct WebDocument {
    document: Document,
}

impl WebDocument {
    /// Open a web template from a path.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        let file = std::fs::File::open(path)?;
        Self::from_reader(file)
    }

    /// Read a web template from a stream.
    pub fn from_reader(mut reader: impl Read) -> Result<Self> {
        let mut bytes = Vec::new();
        reader.read_to_end(&mut bytes)?;
        Self::from_bytes(bytes)
    }

    /// Validate a web template from owned package bytes.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        let package = OwnedPackage::from_bytes(bytes)?;
        let mimetype = package.mimetype()?;
        if mimetype != constants::ODF_WEB {
            return Err(Error::InvalidFormat(format!(
                "not an OpenDocument web template: MIME type is '{mimetype}'"
            )));
        }
        let content = package.get_file(constants::ODF_CONTENT)?;
        let content = std::str::from_utf8(&content)
            .map_err(|_| Error::InvalidFormat("invalid UTF-8 in content.xml".to_string()))?;
        validate_web_content(content)?;
        Ok(Self {
            document: Document::from_owned_package(package)?,
        })
    }

    /// Return the producer MIME type.
    pub fn mimetype(&self) -> &'static str {
        constants::ODF_WEB
    }

    /// Return `true`; `.oth` is a Writer/Web template format.
    pub fn is_template(&self) -> bool {
        true
    }

    /// Return the complete text-document semantic reader.
    pub fn document(&self) -> &Document {
        &self.document
    }

    /// Extract visible template text.
    pub fn text(&self) -> Result<String> {
        self.document.text()
    }

    /// Extract common package metadata.
    pub fn metadata(&self) -> Result<Metadata> {
        self.document.metadata()
    }

    /// Extract complete OpenDocument metadata.
    pub fn odf_metadata(&self) -> Result<Option<crate::Metadata>> {
        self.document.odf_metadata()
    }

    /// Return the exact original package bytes.
    pub fn as_bytes(&self) -> &[u8] {
        self.document.original_bytes()
    }

    /// Convert this template into an atomic mutable web template.
    pub fn into_mutable(self) -> Result<super::MutableWebDocument> {
        super::MutableWebDocument::from_document(self.document)
    }

    /// Clone this template into an atomic mutable web template.
    pub fn to_mutable(&self) -> Result<super::MutableWebDocument> {
        let document = Document::from_bytes(self.document.original_bytes().to_vec())?;
        super::MutableWebDocument::from_document(document)
    }

    /// Clone the exact original package bytes.
    pub fn to_bytes(&self) -> Vec<u8> {
        self.as_bytes().to_vec()
    }

    /// Save without reconstructing the package.
    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        std::fs::write(path, self.as_bytes())?;
        Ok(())
    }
}

fn validate_web_content(xml: &str) -> Result<()> {
    let mut reader = NsReader::from_str(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut root_seen = false;
    let mut root_closed = false;
    let mut body_seen = false;
    let mut body_depth = None;
    let mut text_seen = false;
    let mut text_depth = None;
    loop {
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| Error::InvalidFormat(format!("invalid web-template XML: {error}")))?;
        match event {
            Event::Start(ref element) => {
                if depth == 0 {
                    if root_seen
                        || root_closed
                        || !bound_to(&namespace, OFFICE_NAMESPACE)
                        || element.local_name().as_ref() != b"document-content"
                    {
                        return Err(Error::InvalidFormat(
                            "web template must have one office:document-content root".to_string(),
                        ));
                    }
                    root_seen = true;
                } else if depth == 1
                    && bound_to(&namespace, OFFICE_NAMESPACE)
                    && element.local_name().as_ref() == b"body"
                {
                    if body_seen || body_depth.is_some() {
                        return Err(Error::InvalidFormat("duplicate office:body".to_string()));
                    }
                    body_seen = true;
                    body_depth = Some(2);
                } else if depth == 2 && body_depth == Some(2) {
                    if !bound_to(&namespace, OFFICE_NAMESPACE)
                        || element.local_name().as_ref() != b"text"
                        || text_seen
                    {
                        return Err(Error::InvalidFormat(
                            "web-template body must contain exactly one office:text".to_string(),
                        ));
                    }
                    text_seen = true;
                    text_depth = Some(3);
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    Error::InvalidFormat("web-template XML nesting overflow".to_string())
                })?;
                if depth > MAX_DEPTH {
                    return Err(Error::InvalidFormat(format!(
                        "web-template XML nesting exceeds {MAX_DEPTH} levels"
                    )));
                }
            },
            Event::Empty(ref element) => {
                if depth == 0 {
                    return Err(Error::InvalidFormat(
                        "web-template content root cannot be empty".to_string(),
                    ));
                }
                if depth == 1
                    && bound_to(&namespace, OFFICE_NAMESPACE)
                    && element.local_name().as_ref() == b"body"
                {
                    if body_seen {
                        return Err(Error::InvalidFormat("duplicate office:body".to_string()));
                    }
                    body_seen = true;
                } else if depth == 2 && body_depth == Some(2) {
                    if !bound_to(&namespace, OFFICE_NAMESPACE)
                        || element.local_name().as_ref() != b"text"
                        || text_seen
                    {
                        return Err(Error::InvalidFormat(
                            "web-template body must contain exactly one office:text".to_string(),
                        ));
                    }
                    text_seen = true;
                }
            },
            Event::End(ref element) => {
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("unexpected web-template XML closing tag".to_string())
                })?;
                if bound_to(&namespace, OFFICE_NAMESPACE)
                    && element.local_name().as_ref() == b"text"
                    && depth == 2
                {
                    text_depth = None;
                } else if bound_to(&namespace, OFFICE_NAMESPACE)
                    && element.local_name().as_ref() == b"body"
                    && depth == 1
                {
                    body_depth = None;
                }
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::Text(ref text) if depth == 0 && !text.iter().all(u8::is_ascii_whitespace) => {
                return Err(Error::InvalidFormat(
                    "text is not allowed outside the web-template root".to_string(),
                ));
            },
            Event::CData(_) | Event::GeneralRef(_) if depth == 0 => {
                return Err(Error::InvalidFormat(
                    "content is not allowed outside the web-template root".to_string(),
                ));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !root_seen
        || !root_closed
        || depth != 0
        || !body_seen
        || body_depth.is_some()
        || !text_seen
        || text_depth.is_some()
    {
        return Err(Error::InvalidFormat(
            "incomplete OpenDocument web-template structure".to_string(),
        ));
    }
    Ok(())
}

fn bound_to(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(uri)) if *uri == expected)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::core::PackageWriter;
    use std::io::Cursor;

    fn package(mimetype: &str, content: &str) -> Vec<u8> {
        let mut writer = PackageWriter::new();
        writer.set_mimetype(mimetype).unwrap();
        writer
            .add_file(constants::ODF_CONTENT, content.as_bytes())
            .unwrap();
        writer.finish_to_bytes().unwrap()
    }

    fn web_xml() -> &'static str {
        r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:automatic-styles/><office:body><office:text><text:h text:outline-level="1">Web template</text:h><text:p>Body &amp; links</text:p></office:text></office:body></office:document-content>"#
    }

    #[test]
    fn opens_writer_web_templates_with_complete_text_model_losslessly() {
        let bytes = package(constants::ODF_WEB, web_xml());
        let document = WebDocument::from_bytes(bytes.clone()).unwrap();
        assert!(document.is_template());
        assert_eq!(document.mimetype(), constants::ODF_WEB);
        assert_eq!(document.text().unwrap(), "Web template\nBody & links");
        assert_eq!(document.document().paragraph_count().unwrap(), 1);
        assert_eq!(document.as_bytes(), bytes);
        assert_eq!(document.to_bytes(), bytes);
    }

    #[test]
    fn accepts_readers_empty_text_and_arbitrary_container_prefixes() {
        let xml = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><o:body><o:text><t:p>A &amp; B<t:s t:c="2"/>C</t:p></o:text></o:body></o:document-content>"#;
        let bytes = package(constants::ODF_WEB, xml);
        let document = WebDocument::from_reader(Cursor::new(bytes.clone())).unwrap();
        assert_eq!(document.text().unwrap(), "A & B  C");
        assert_eq!(document.as_bytes(), bytes);
    }

    #[test]
    fn rejects_other_mime_types_and_invalid_body_hierarchy() {
        assert!(WebDocument::from_bytes(package(constants::ODF_TEXT, web_xml())).is_err());
        for xml in [
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:spreadsheet/></o:body></o:document-content>"#,
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body/></o:document-content>"#,
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:text/><o:text/></o:body></o:document-content>"#,
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:text/></o:body>"#,
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:text/></o:body></o:document-content><o:x/>"#,
        ] {
            assert!(
                WebDocument::from_bytes(package(constants::ODF_WEB, xml)).is_err(),
                "accepted {xml}"
            );
        }
    }

    #[test]
    fn rejects_excessive_nesting() {
        let nested = "<v:x>".repeat(MAX_DEPTH) + &"</v:x>".repeat(MAX_DEPTH);
        let xml = format!(
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:v="urn:vendor"><o:body><o:text>{nested}</o:text></o:body></o:document-content>"#
        );
        assert!(WebDocument::from_bytes(package(constants::ODF_WEB, &xml)).is_err());
    }
}
