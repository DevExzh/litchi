//! Flat (single-XML) `OpenDocument` documents as package parts.
//!
//! A flat document (`.fodt`, `.fods`, `.fodp`, `.fodg`, ...) is one
//! `office:document` element whose direct children are the parts a package
//! stores separately. [`Parts::split`] partitions the root's children into
//! `content.xml` (scripts, font faces, automatic styles, body), `styles.xml`
//! (font faces, styles, automatic styles, master styles), `meta.xml` and
//! `settings.xml`, each wrapped in its package root element with the flat
//! root's namespace declarations and version. [`Parts::to_bytes`] publishes
//! them as an in-memory package, so a family's packaged reader can project a
//! flat document without a second reader. The flat bytes stay the document:
//! nothing here is written back.
//!
//! Every child's bytes are copied verbatim; nested namespace declarations,
//! text and inline binary data are kept. Only the root start tag is
//! rewritten, dropping its `office:mimetype` attribute.

use quick_xml::Reader;
use quick_xml::events::Event;

use crate::core::PackageWriter;
use litchi_core::{Error, Result};

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";

/// The parts of one flat document.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Parts {
    mimetype: String,
    content: String,
    styles: String,
    meta: Option<String>,
    settings: Option<String>,
}

/// One direct child of the flat root.
struct Child<'a> {
    local: &'a [u8],
    xml: &'a str,
}

impl Parts {
    /// Split a flat document.
    ///
    /// # Errors
    ///
    /// Returns an error when `xml` is not UTF-8, not well-formed, has no
    /// `office:document` root in the `OpenDocument` office namespace, or
    /// declares no `office:mimetype`.
    pub fn split(xml: &[u8]) -> Result<Self> {
        let xml = std::str::from_utf8(xml)
            .map_err(|error| invalid(format!("flat ODF is not UTF-8: {error}")))?;
        let xml = xml.strip_prefix('\u{feff}').unwrap_or(xml);
        let mut reader = Reader::from_str(xml);
        reader.config_mut().check_end_names = true;
        let mut depth = 0usize;
        let mut root: Option<(&str, String, String)> = None;
        let mut children: Vec<Child<'_>> = Vec::new();
        let mut open_child: Option<(usize, &[u8])> = None;
        loop {
            let start = offset(reader.buffer_position());
            let event = reader
                .read_event()
                .map_err(|error| invalid(format!("flat ODF is not well-formed XML: {error}")))?;
            let end = offset(reader.buffer_position());
            match event {
                Event::Start(element) => {
                    if depth == 0 {
                        root = Some(root_tag(&xml[start..end], &element)?);
                    } else if depth == 1 {
                        open_child =
                            Some((start, local(xml, start, element.name().as_ref().len())));
                    }
                    depth += 1;
                },
                Event::Empty(element) => {
                    if depth == 0 {
                        return Err(invalid("the flat ODF root element is empty"));
                    }
                    if depth == 1 {
                        children.push(Child {
                            local: local(xml, start, element.name().as_ref().len()),
                            xml: &xml[start..end],
                        });
                    }
                },
                Event::End(_) => {
                    depth = depth.saturating_sub(1);
                    if depth == 1
                        && let Some((child_start, name)) = open_child.take()
                    {
                        children.push(Child {
                            local: name,
                            xml: &xml[child_start..end],
                        });
                    }
                },
                Event::Eof => break,
                Event::DocType(_) => {
                    return Err(invalid("flat ODF with a document type declaration"));
                },
                _ => {},
            }
        }
        let (prefix, attributes, mimetype) =
            root.ok_or_else(|| invalid("flat ODF has no root element"))?;
        let attributes = attributes.as_str();
        let part = |root_local: &str, locals: &[&[u8]]| {
            let mut out = String::from(r#"<?xml version="1.0" encoding="UTF-8"?>"#);
            let name = qualified(prefix, root_local);
            out.push('<');
            out.push_str(&name);
            out.push_str(attributes);
            out.push('>');
            for wanted in locals {
                for child in children.iter().filter(|child| child.local == *wanted) {
                    out.push_str(child.xml);
                }
            }
            out.push_str("</");
            out.push_str(&name);
            out.push('>');
            out
        };
        let has = |wanted: &[u8]| children.iter().any(|child| child.local == wanted);
        Ok(Self {
            content: part(
                "document-content",
                &[b"scripts", b"font-face-decls", b"automatic-styles", b"body"],
            ),
            styles: part(
                "document-styles",
                &[
                    b"font-face-decls",
                    b"styles",
                    b"automatic-styles",
                    b"master-styles",
                ],
            ),
            meta: has(b"meta").then(|| part("document-meta", &[b"meta"])),
            settings: has(b"settings").then(|| part("document-settings", &[b"settings"])),
            mimetype,
        })
    }

    /// The document's declared media type (`office:mimetype`).
    #[must_use]
    pub fn mimetype(&self) -> &str {
        &self.mimetype
    }

    /// The `content.xml` part.
    #[must_use]
    pub fn content_xml(&self) -> &str {
        &self.content
    }

    /// The `styles.xml` part.
    #[must_use]
    pub fn styles_xml(&self) -> &str {
        &self.styles
    }

    /// The `meta.xml` part, when the document has `office:meta`.
    #[must_use]
    pub fn meta_xml(&self) -> Option<&str> {
        self.meta.as_deref()
    }

    /// The `settings.xml` part, when the document has `office:settings`.
    #[must_use]
    pub fn settings_xml(&self) -> Option<&str> {
        self.settings.as_deref()
    }

    /// Publish the parts as an in-memory package.
    ///
    /// # Errors
    ///
    /// Returns an error when the package writer refuses a part.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        let mut writer = PackageWriter::new();
        writer.set_mimetype(&self.mimetype)?;
        // The parts carry the producer's bytes (indentation included), so
        // they publish as exact source rather than freshly authored XML.
        for (path, part) in [
            ("content.xml", Some(&self.content)),
            ("styles.xml", Some(&self.styles)),
            ("meta.xml", self.meta.as_ref()),
            ("settings.xml", self.settings.as_ref()),
        ] {
            if let Some(part) = part {
                writer.write_exact_source_file(path, part.as_bytes(), "text/xml")?;
            }
        }
        writer.finish()
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

fn offset(position: u64) -> usize {
    usize::try_from(position).unwrap_or(usize::MAX)
}

/// The local name of the element whose tag starts at `start` with a
/// qualified name of `length` bytes.
fn local(xml: &str, start: usize, length: usize) -> &[u8] {
    let name = &xml.as_bytes()[start + 1..start + 1 + length];
    name.iter()
        .position(|byte| *byte == b':')
        .map_or(name, |colon| &name[colon + 1..])
}

fn qualified(prefix: &str, local: &str) -> String {
    if prefix.is_empty() {
        local.to_owned()
    } else {
        format!("{prefix}:{local}")
    }
}

/// The root's prefix, its attribute text without `office:mimetype`, and
/// the media type it declares.
fn root_tag<'a>(
    tag: &'a str,
    element: &quick_xml::events::BytesStart<'_>,
) -> Result<(&'a str, String, String)> {
    let name_length = element.name().as_ref().len();
    let name = &tag[1..=name_length];
    let (prefix, local) = name.split_once(':').unwrap_or(("", name));
    if local != "document" {
        return Err(invalid("the flat ODF root is not office:document"));
    }
    let mimetype_key = qualified(prefix, "mimetype");
    let mut declared = false;
    let mut mimetype = None;
    for attribute in element.attributes() {
        let attribute =
            attribute.map_err(|error| invalid(format!("flat ODF root attribute: {error}")))?;
        let key = attribute.key.as_ref();
        let binds_prefix = if prefix.is_empty() {
            key == b"xmlns"
        } else {
            key.strip_prefix(b"xmlns:") == Some(prefix.as_bytes())
        };
        if binds_prefix && attribute.value.as_ref() == OFFICE {
            declared = true;
        }
        if key == mimetype_key.as_bytes() {
            // Media types are ASCII; a reference in one is not one.
            mimetype = Some(String::from_utf8_lossy(attribute.value.as_ref()).into_owned());
        }
    }
    if !declared {
        return Err(invalid(
            "the flat ODF root is not in the OpenDocument office namespace",
        ));
    }
    let mimetype = mimetype.ok_or_else(|| invalid("the flat ODF root declares no mimetype"))?;
    let mut attributes = tag[1 + name_length..tag.len() - 1].to_owned();
    // Drop `prefix:mimetype="…"` (either quote) from the attribute text.
    let needle = format!("{mimetype_key}=");
    if let Some(at) = attributes.find(&needle) {
        let value_start = at + needle.len();
        if let Some(quote) = attributes[value_start..].chars().next()
            && let Some(close) = attributes[value_start + 1..].find(quote)
        {
            attributes.replace_range(at..value_start + 1 + close + 1, "");
        }
    }
    Ok((prefix, attributes, mimetype))
}

#[cfg(test)]
mod tests {
    use super::Parts;

    const FLAT: &str = concat!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n",
        "<office:document xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" ",
        "xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" office:version=\"1.3\" ",
        "office:mimetype=\"application/vnd.oasis.opendocument.text\">\n",
        " <office:meta><dc:title xmlns:dc=\"http://purl.org/dc/elements/1.1/\">T</dc:title></office:meta>\n",
        " <office:styles><style:style xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" style:name=\"S\"/></office:styles>\n",
        " <office:automatic-styles/>\n",
        " <office:body>\n  <office:text><text:p>Hello</text:p></office:text>\n </office:body>\n",
        "</office:document>\n",
    );

    #[test]
    fn splits_children_into_package_parts_keeping_their_bytes() {
        let parts = Parts::split(FLAT.as_bytes()).unwrap();
        assert_eq!(parts.mimetype(), "application/vnd.oasis.opendocument.text");
        let root = " xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" office:version=\"1.3\" ";
        assert_eq!(
            parts.content_xml(),
            format!(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content{root}><office:automatic-styles/><office:body>\n  <office:text><text:p>Hello</text:p></office:text>\n </office:body></office:document-content>"
            )
        );
        assert!(parts.styles_xml().contains("<style:style xmlns:style="));
        assert!(parts.styles_xml().contains("<office:document-styles"));
        assert!(parts.meta_xml().unwrap().contains("<dc:title"));
        assert_eq!(parts.settings_xml(), None);
        let package = parts.to_bytes().unwrap();
        assert_eq!(&package[..2], b"PK");
    }

    #[test]
    fn refuses_what_is_not_a_flat_document() {
        assert!(Parts::split(b"<root/>").is_err());
        assert!(Parts::split(b"<office:document xmlns:office=\"x\"/>").is_err());
        assert!(
            Parts::split(
                b"<office:document xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"><office:body/></office:document>"
            )
            .is_err(),
            "no mimetype"
        );
        assert!(Parts::split(b"<office:document><unclosed>").is_err());
        assert!(Parts::split(&[0xff, 0xfe]).is_err());
    }
}
