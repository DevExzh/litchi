//! Image package authoring.

use crate::{frame::Frame, source::Source};
use litchi_core::{Error, Result};
use litchi_odf_common::{compact_xml, core::PackageWriter};

const MAX_RESOURCES: usize = 100_000;

#[derive(Clone, Debug)]
struct Resource {
    path: String,
    media_type: String,
    bytes: Vec<u8>,
}

/// Detached builder; publication validates through the package facade.
#[derive(Clone, Debug)]
pub struct Builder {
    content_xml: String,
    resources: Vec<Resource>,
}

impl Builder {
    /// Creates a builder pre-filled with an empty image content document.
    #[must_use]
    pub fn new() -> Self {
        Self {
            content_xml: empty_content().to_owned(),
            resources: Vec::new(),
        }
    }

    /// Replaces the `content.xml` payload.
    #[must_use]
    pub fn content_xml(mut self, xml: impl Into<String>) -> Self {
        self.content_xml = xml.into();
        self
    }

    /// Replaces the content with one semantically authored image frame.
    ///
    /// The generated XML is deterministic and contains no formatting
    /// whitespace. Linked sources remain inert; embedded sources are encoded
    /// directly as `office:binary-data`.
    #[must_use]
    pub fn frame(mut self, frame: &Frame) -> Self {
        self.content_xml = frame_content(frame);
        self
    }

    /// Adds a package-local resource and its explicit manifest media type.
    ///
    /// Use [`Self::frame`] with a linked source whose `xlink:href` resolves to
    /// this path. Path safety, duplicate names, and package budgets are checked
    /// when [`Self::build`] publishes the package.
    #[must_use]
    pub fn resource(
        mut self,
        path: impl Into<String>,
        media_type: impl Into<String>,
        bytes: impl Into<Vec<u8>>,
    ) -> Self {
        self.resources.push(Resource {
            path: path.into(),
            media_type: media_type.into(),
            bytes: bytes.into(),
        });
        self
    }

    /// Validates and packages the document bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if content validation or package writing fails.
    pub fn build(self) -> Result<Vec<u8>> {
        if self.resources.len() > MAX_RESOURCES {
            return Err(Error::InvalidFormat(format!(
                "ODI resource count exceeds {MAX_RESOURCES}"
            )));
        }
        compact_xml::validate(self.content_xml.as_bytes())?;
        crate::codec::validate(&self.content_xml)?;
        let mut writer = PackageWriter::new_bounded(256 * 1024 * 1024);
        writer.set_mimetype(crate::package::MIMETYPE)?;
        writer.add_file("content.xml", self.content_xml.as_bytes())?;
        for resource in self.resources {
            writer.add_file_with_media_type(
                &resource.path,
                &resource.bytes,
                &resource.media_type,
            )?;
        }
        let bytes = writer.finish_to_bounded_bytes()?;
        crate::package::Snapshot::from_bytes(bytes).map(crate::package::Snapshot::into_bytes)
    }
}

impl Default for Builder {
    fn default() -> Self {
        Self::new()
    }
}

fn empty_content() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" office:version="1.4"><office:body><office:image><draw:frame><draw:image draw:mime-type="image/png"><office:binary-data>iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAF/gL+X9p9WQAAAABJRU5ErkJggg==</office:binary-data></draw:image></draw:frame></office:image></office:body></office:document-content>"#
}

fn frame_content(frame: &Frame) -> String {
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.4"><office:body><office:image><draw:frame"#,
    );
    if let Some(name) = frame.name() {
        push_attribute(&mut xml, "draw:name", name);
    }
    xml.push('>');
    if let Some(title) = frame.title() {
        xml.push_str("<svg:title>");
        xml.push_str(&quick_xml::escape::escape(title));
        xml.push_str("</svg:title>");
    }
    if let Some(description) = frame.description() {
        xml.push_str("<svg:desc>");
        xml.push_str(&quick_xml::escape::escape(description));
        xml.push_str("</svg:desc>");
    }
    match frame.source() {
        Source::Linked(href) => {
            xml.push_str("<draw:image");
            push_attribute(&mut xml, "xlink:href", href);
            xml.push_str("/>");
        },
        Source::Embedded(bytes) => {
            xml.push_str("<draw:image><office:binary-data>");
            xml.push_str(&base64(bytes));
            xml.push_str("</office:binary-data></draw:image>");
        },
    }
    xml.push_str("</draw:frame></office:image></office:body></office:document-content>");
    xml
}

fn push_attribute(xml: &mut String, name: &str, value: &str) {
    xml.push(' ');
    xml.push_str(name);
    xml.push_str("=\"");
    xml.push_str(&quick_xml::escape::escape(value));
    xml.push('"');
}

fn base64(bytes: &[u8]) -> String {
    const TABLE: &[u8; 64] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    let mut output = String::with_capacity(bytes.len().div_ceil(3).saturating_mul(4));
    for chunk in bytes.chunks(3) {
        let first = chunk[0];
        let second = chunk.get(1).copied().unwrap_or(0);
        let third = chunk.get(2).copied().unwrap_or(0);
        output.push(TABLE[(first >> 2) as usize] as char);
        output.push(TABLE[((first & 3) << 4 | second >> 4) as usize] as char);
        output.push(if chunk.len() > 1 {
            TABLE[((second & 15) << 2 | third >> 6) as usize] as char
        } else {
            '='
        });
        output.push(if chunk.len() > 2 {
            TABLE[(third & 63) as usize] as char
        } else {
            '='
        });
    }
    output
}
