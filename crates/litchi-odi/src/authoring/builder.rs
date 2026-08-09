//! Image package authoring.

use crate::{
    frame::Frame,
    map::{Area, AreaKind, ImageMap},
    source::Source,
};
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
    styles_xml: Option<String>,
    meta_xml: Option<String>,
    resources: Vec<Resource>,
}

impl Builder {
    /// Creates a builder pre-filled with an empty image content document.
    #[must_use]
    pub fn new() -> Self {
        Self {
            content_xml: empty_content().to_owned(),
            styles_xml: None,
            meta_xml: None,
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

    /// Adds an exact compact `styles.xml` part.
    #[must_use]
    pub fn styles_xml(mut self, xml: impl Into<String>) -> Self {
        self.styles_xml = Some(xml.into());
        self
    }

    /// Adds an exact compact `meta.xml` part.
    #[must_use]
    pub fn meta_xml(mut self, xml: impl Into<String>) -> Self {
        self.meta_xml = Some(xml.into());
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
        for xml in [&self.styles_xml, &self.meta_xml].into_iter().flatten() {
            compact_xml::validate(xml.as_bytes())?;
        }
        let mut writer = PackageWriter::new_bounded(256 * 1024 * 1024);
        writer.set_mimetype(crate::package::MIMETYPE)?;
        writer.add_file("content.xml", self.content_xml.as_bytes())?;
        if let Some(styles_xml) = self.styles_xml {
            writer.add_file("styles.xml", styles_xml.as_bytes())?;
        }
        if let Some(meta_xml) = self.meta_xml {
            writer.add_file("meta.xml", meta_xml.as_bytes())?;
        }
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
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.4"><office:body><office:image><draw:frame"#,
    );
    if let Some(name) = frame.name() {
        push_attribute(&mut xml, "draw:name", name);
    }
    if let Some(xml_id) = frame.xml_id() {
        push_attribute(&mut xml, "xml:id", xml_id);
    }
    for (name, value) in [
        ("draw:style-name", frame.style_name()),
        ("draw:text-style-name", frame.text_style_name()),
        ("draw:layer", frame.layer()),
        ("draw:transform", frame.transform()),
        ("text:anchor-type", frame.anchor_type()),
        ("svg:x", frame.x()),
        ("svg:y", frame.y()),
        ("svg:width", frame.width()),
        ("svg:height", frame.height()),
        ("style:rel-width", frame.relative_width()),
        ("style:rel-height", frame.relative_height()),
        ("draw:copy-of", frame.copy_of()),
    ] {
        if let Some(value) = value {
            push_attribute(&mut xml, name, value);
        }
    }
    if let Some(z_index) = frame.z_index() {
        push_attribute(&mut xml, "draw:z-index", &z_index.to_string());
    }
    xml.push('>');
    match frame.source() {
        Source::Linked(href) => {
            xml.push_str("<draw:image");
            push_attribute(&mut xml, "xlink:href", href);
            push_image_attributes(&mut xml, frame);
            xml.push_str("/>");
        },
        Source::Embedded(bytes) => {
            xml.push_str("<draw:image");
            push_image_attributes(&mut xml, frame);
            xml.push_str("><office:binary-data>");
            xml.push_str(&base64(bytes));
            xml.push_str("</office:binary-data></draw:image>");
        },
    }
    if let Some(image_map) = frame.image_map() {
        push_image_map(&mut xml, image_map);
    }
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
    xml.push_str("</draw:frame></office:image></office:body></office:document-content>");
    xml
}

fn push_image_attributes(xml: &mut String, frame: &Frame) {
    for (name, value) in [
        ("draw:mime-type", frame.media_type()),
        ("xml:id", frame.image_xml_id()),
        ("draw:filter-name", frame.filter_name()),
        ("xlink:type", frame.link_type()),
        ("xlink:show", frame.show()),
        ("xlink:actuate", frame.actuate()),
    ] {
        if let Some(value) = value {
            push_attribute(xml, name, value);
        }
    }
}

pub(crate) fn image_map_xml(image_map: &ImageMap) -> String {
    let mut xml = String::new();
    push_image_map(&mut xml, image_map);
    xml
}

fn push_image_map(xml: &mut String, image_map: &ImageMap) {
    xml.push_str("<draw:image-map>");
    for area in image_map.areas() {
        push_area(xml, area);
    }
    xml.push_str("</draw:image-map>");
}

fn push_area(xml: &mut String, area: &Area) {
    let element = match area.kind() {
        AreaKind::Rectangle {
            x,
            y,
            width,
            height,
        } => {
            xml.push_str("<draw:area-rectangle");
            for (name, value) in [
                ("svg:x", x.as_str()),
                ("svg:y", y.as_str()),
                ("svg:width", width.as_str()),
                ("svg:height", height.as_str()),
            ] {
                push_attribute(xml, name, value);
            }
            "draw:area-rectangle"
        },
        AreaKind::Circle {
            center_x,
            center_y,
            radius,
        } => {
            xml.push_str("<draw:area-circle");
            push_attribute(xml, "svg:cx", center_x);
            push_attribute(xml, "svg:cy", center_y);
            push_attribute(xml, "svg:r", radius);
            "draw:area-circle"
        },
        AreaKind::Polygon {
            x,
            y,
            width,
            height,
            view_box,
            points,
        } => {
            xml.push_str("<draw:area-polygon");
            for (name, value) in [
                ("svg:x", x.as_str()),
                ("svg:y", y.as_str()),
                ("svg:width", width.as_str()),
                ("svg:height", height.as_str()),
                ("svg:viewBox", view_box.as_str()),
                ("draw:points", points.as_str()),
            ] {
                push_attribute(xml, name, value);
            }
            "draw:area-polygon"
        },
    };
    if let Some(href) = area.href() {
        if let Some(link_type) = area.link_type() {
            push_attribute(xml, "xlink:type", link_type);
        }
        push_attribute(xml, "xlink:href", href);
    }
    if let Some(show) = area.show() {
        push_attribute(xml, "xlink:show", show);
    }
    if let Some(actuate) = area.actuate() {
        push_attribute(xml, "xlink:actuate", actuate);
    }
    if let Some(target) = area.target_frame_name() {
        push_attribute(xml, "office:target-frame-name", target);
    }
    if let Some(name) = area.name() {
        push_attribute(xml, "office:name", name);
    }
    if area.has_no_href() {
        push_attribute(xml, "draw:nohref", "nohref");
    }
    if area.title().is_none() && area.description().is_none() {
        xml.push_str("/>");
        return;
    }
    xml.push('>');
    if let Some(title) = area.title() {
        xml.push_str("<svg:title>");
        xml.push_str(&quick_xml::escape::escape(title));
        xml.push_str("</svg:title>");
    }
    if let Some(description) = area.description() {
        xml.push_str("<svg:desc>");
        xml.push_str(&quick_xml::escape::escape(description));
        xml.push_str("</svg:desc>");
    }
    xml.push_str("</");
    xml.push_str(element);
    xml.push('>');
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
