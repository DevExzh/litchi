//! Flat (single-XML) `OpenDocument` documents as package parts.
//!
//! A flat document (`.fodt`, `.fods`, `.fodp`, `.fodg`, ...) is one
//! `office:document` element whose direct children are the parts a package
//! stores separately. [`Parts::split`] partitions the root's children into
//! `content.xml` (scripts, font faces, automatic styles, body), `styles.xml`
//! (font faces, styles, automatic styles, master styles), `meta.xml` and
//! `settings.xml`, each wrapped in its package root element with the flat
//! root's namespace declarations and version. [`Parts::to_bytes`] publishes
//! them as an in-memory package, so a family's packaged reader (and editor)
//! can work on a flat document without a second engine.
//!
//! Every child's bytes are copied verbatim; nested namespace declarations,
//! text and inline binary data are kept. Only the root start tag is
//! rewritten, dropping its `office:mimetype` attribute.
//!
//! [`Parts::merge`] maps an edit of those parts back: every top-level child
//! a part's edit changed replaces that child's bytes in the flat document,
//! and every other byte of the flat document is kept. An edit that changes
//! a part's root, adds or removes top-level children, or changes one shared
//! child (font faces, automatic styles) differently in two parts is refused.

use std::ops::Range;

use quick_xml::Reader;
use quick_xml::events::Event;

use crate::core::PackageWriter;
use litchi_core::{Error, Result};

const OFFICE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";

/// The top-level children each part holds, in order.
const CONTENT: &[&[u8]] = &[b"scripts", b"font-face-decls", b"automatic-styles", b"body"];
const STYLES: &[&[u8]] = &[
    b"font-face-decls",
    b"styles",
    b"automatic-styles",
    b"master-styles",
];
const META: &[&[u8]] = &[b"meta"];
const SETTINGS: &[&[u8]] = &[b"settings"];

/// One package part of a flat document.
#[derive(Debug, Clone, PartialEq, Eq)]
struct Part {
    xml: String,
    /// The flat ranges of the children the part holds, in part order.
    sources: Vec<Range<usize>>,
}

/// The parts of one flat document.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Parts {
    mimetype: String,
    content: Part,
    styles: Part,
    meta: Option<Part>,
    settings: Option<Part>,
    /// The flat document's length, which [`Parts::merge`] checks.
    length: usize,
    /// Where the flat root's content starts (after its start tag), for a
    /// part an edit creates.
    root_inner: usize,
}

/// The root start tag and the top-level children of one XML document.
struct TopLevel<'a> {
    root: &'a str,
    root_end: usize,
    children: Vec<(&'a [u8], Range<usize>)>,
}

/// Scan `xml`'s root start tag and its direct children.
fn top_level(xml: &str) -> Result<TopLevel<'_>> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut depth = 0usize;
    let mut root: Option<(&str, usize)> = None;
    let mut children = Vec::new();
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
                    root = Some((&xml[start..end], end));
                } else if depth == 1 {
                    open_child = Some((start, local(xml, start, element.name().as_ref().len())));
                }
                depth += 1;
            },
            Event::Empty(element) => {
                if depth == 0 {
                    return Err(invalid("the flat ODF root element is empty"));
                }
                if depth == 1 {
                    children.push((local(xml, start, element.name().as_ref().len()), start..end));
                }
            },
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 1
                    && let Some((child_start, name)) = open_child.take()
                {
                    children.push((name, child_start..end));
                }
            },
            Event::Eof => break,
            Event::DocType(_) => {
                return Err(invalid("flat ODF with a document type declaration"));
            },
            _ => {},
        }
    }
    let (root, root_end) = root.ok_or_else(|| invalid("flat ODF has no root element"))?;
    Ok(TopLevel {
        root,
        root_end,
        children,
    })
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
        let text = std::str::from_utf8(xml)
            .map_err(|error| invalid(format!("flat ODF is not UTF-8: {error}")))?;
        let bom = if text.starts_with('\u{feff}') { 3 } else { 0 };
        let body = &text[bom..];
        let scanned = top_level(body)?;
        let mut reader = Reader::from_str(scanned.root);
        let Ok(Event::Start(element)) = reader.read_event() else {
            return Err(invalid("the flat ODF root start tag is malformed"));
        };
        let (prefix, attributes, mimetype) = root_tag(scanned.root, &element)?;
        let part = |root_local: &str, locals: &[&[u8]]| {
            let mut xml = String::from(r#"<?xml version="1.0" encoding="UTF-8"?>"#);
            let name = qualified(prefix, root_local);
            xml.push('<');
            xml.push_str(&name);
            xml.push_str(&attributes);
            xml.push('>');
            let mut sources = Vec::new();
            for wanted in locals {
                for (_, range) in scanned.children.iter().filter(|child| child.0 == *wanted) {
                    xml.push_str(&body[range.clone()]);
                    sources.push(range.start + bom..range.end + bom);
                }
            }
            xml.push_str("</");
            xml.push_str(&name);
            xml.push('>');
            Part { xml, sources }
        };
        let has = |wanted: &[u8]| scanned.children.iter().any(|child| child.0 == wanted);
        Ok(Self {
            content: part("document-content", CONTENT),
            styles: part("document-styles", STYLES),
            meta: has(b"meta").then(|| part("document-meta", META)),
            settings: has(b"settings").then(|| part("document-settings", SETTINGS)),
            mimetype,
            length: xml.len(),
            root_inner: scanned.root_end + bom,
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
        &self.content.xml
    }

    /// The `styles.xml` part.
    #[must_use]
    pub fn styles_xml(&self) -> &str {
        &self.styles.xml
    }

    /// The `meta.xml` part, when the document has `office:meta`.
    #[must_use]
    pub fn meta_xml(&self) -> Option<&str> {
        self.meta.as_ref().map(|part| part.xml.as_str())
    }

    /// The `settings.xml` part, when the document has `office:settings`.
    #[must_use]
    pub fn settings_xml(&self) -> Option<&str> {
        self.settings.as_ref().map(|part| part.xml.as_str())
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
            ("content.xml", Some(self.content_xml())),
            ("styles.xml", Some(self.styles_xml())),
            ("meta.xml", self.meta_xml()),
            ("settings.xml", self.settings_xml()),
        ] {
            if let Some(part) = part {
                writer.write_exact_source_file(path, part.as_bytes(), "text/xml")?;
            }
        }
        writer.finish()
    }

    /// Republish `flat` (the document these parts were split from) with an
    /// edit of the parts: each top-level child an edited part changed
    /// replaces that child's bytes; every other flat byte is kept. A part
    /// passed as `None` is unchanged. An edit that creates `meta.xml` for a
    /// flat document without `office:meta` inserts its children first in
    /// the flat root.
    ///
    /// # Errors
    ///
    /// Returns an error when `flat` is not the split source, and
    /// [`Error::Unsupported`] when an edited part's root start tag changed,
    /// it gained or lost top-level children, a part disappeared, or two
    /// parts changed one shared child differently.
    pub fn merge(
        &self,
        flat: &[u8],
        content: Option<&str>,
        styles: Option<&str>,
        meta: Option<&str>,
        settings: Option<&str>,
    ) -> Result<Vec<u8>> {
        if flat.len() != self.length {
            return Err(invalid(
                "the flat document is not the one these parts were split from",
            ));
        }
        let mut replacements: Vec<(Range<usize>, String)> = Vec::new();
        for (before, after) in [
            (Some(&self.content), content),
            (Some(&self.styles), styles),
            (self.meta.as_ref(), meta),
            (self.settings.as_ref(), settings),
        ] {
            let Some(after) = after else {
                continue;
            };
            match before {
                Some(before) => changed_children(before, after, &mut replacements)?,
                None => {
                    // A part the edit created (`meta.xml` of a flat document
                    // without `office:meta`): its children go first.
                    let scanned = top_level(after)?;
                    let mut inserted = String::new();
                    for (_, range) in &scanned.children {
                        inserted.push_str(&after[range.clone()]);
                    }
                    if !inserted.is_empty() {
                        push_replacement(
                            &mut replacements,
                            self.root_inner..self.root_inner,
                            inserted,
                        )?;
                    }
                },
            }
        }
        // Last first; at one offset, a replaced child before an insertion.
        replacements.sort_by_key(|(range, _)| std::cmp::Reverse((range.start, range.end)));
        let mut out = flat.to_vec();
        for (range, text) in replacements {
            out.splice(range, text.into_bytes());
        }
        Ok(out)
    }
}

impl Parts {
    /// [`Parts::merge`] with the parts of an edited package: the package a
    /// family's editor published from [`Parts::to_bytes`]. Its members must
    /// be the parts and the manifest; a member a flat document cannot hold
    /// (a picture, an embedded object) is refused.
    ///
    /// # Errors
    ///
    /// As [`Parts::merge`]; [`Error::Unsupported`] for an added member;
    /// the package reader's error for an unreadable package.
    pub fn merge_package(&self, flat: &[u8], package: &[u8]) -> Result<Vec<u8>> {
        let owned = crate::OwnedPackage::from_bytes(package.to_vec())?;
        for member in owned.files()? {
            if !matches!(
                member.as_str(),
                "mimetype"
                    | "content.xml"
                    | "styles.xml"
                    | "meta.xml"
                    | "settings.xml"
                    | "META-INF/"
                    | "META-INF/manifest.xml"
            ) {
                return Err(Error::Refused(litchi_core::edit::Refusal::new(
                    litchi_core::edit::RefusalKind::UnsupportedOperation,
                    "odf-common.flat",
                    "the edit added the package member a member which a flat ODF document cannot hold",
                )));
            }
        }
        let read = |path: &str| -> Result<Option<String>> {
            if !owned.has_file(path)? {
                return Ok(None);
            }
            String::from_utf8(owned.get_file(path)?)
                .map(Some)
                .map_err(|error| invalid(format!("edited {path} is not UTF-8: {error}")))
        };
        let content =
            read("content.xml")?.ok_or_else(|| invalid("the edited package has no content.xml"))?;
        let styles = read("styles.xml")?;
        let meta = read("meta.xml")?;
        let settings = read("settings.xml")?;
        self.merge(
            flat,
            Some(&content),
            styles.as_deref(),
            meta.as_deref(),
            settings.as_deref(),
        )
    }
}

/// Record the children `after` changed relative to `before`.
fn changed_children(
    before: &Part,
    after: &str,
    replacements: &mut Vec<(Range<usize>, String)>,
) -> Result<()> {
    let old = top_level(&before.xml)?;
    let new = top_level(after)?;
    if old.root != new.root {
        return Err(Error::Refused(litchi_core::edit::Refusal::new(
            litchi_core::edit::RefusalKind::UnsupportedOperation,
            "odf-common.flat",
            "the edit changed a flat ODF part's root element",
        )));
    }
    if old.children.len() != new.children.len()
        || old.children.len() != before.sources.len()
        || old
            .children
            .iter()
            .zip(&new.children)
            .any(|(left, right)| left.0 != right.0)
    {
        return Err(Error::Refused(litchi_core::edit::Refusal::new(
            litchi_core::edit::RefusalKind::UnsupportedOperation,
            "odf-common.flat",
            "the edit added or removed a top-level element of a flat ODF part",
        )));
    }
    for ((old_child, new_child), source) in
        old.children.iter().zip(&new.children).zip(&before.sources)
    {
        let old_text = &before.xml[old_child.1.clone()];
        let new_text = &after[new_child.1.clone()];
        if old_text != new_text {
            push_replacement(replacements, source.clone(), new_text.to_owned())?;
        }
    }
    Ok(())
}

/// Add one replacement; one shared child changed two ways is refused.
fn push_replacement(
    replacements: &mut Vec<(Range<usize>, String)>,
    range: Range<usize>,
    text: String,
) -> Result<()> {
    if let Some(existing) = replacements.iter().find(|(existing, _)| *existing == range) {
        if existing.1 == text {
            return Ok(());
        }
        return Err(Error::Refused(litchi_core::edit::Refusal::new(
            litchi_core::edit::RefusalKind::UnsupportedOperation,
            "odf-common.flat",
            "the edit changed one shared flat ODF element differently in two parts",
        )));
    }
    replacements.push((range, text));
    Ok(())
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

    #[test]
    fn merges_changed_children_back_and_keeps_every_other_byte() {
        let parts = Parts::split(FLAT.as_bytes()).unwrap();
        let content = parts.content_xml().replace("Hello", "Bye");
        let merged = parts
            .merge(FLAT.as_bytes(), Some(&content), None, None, None)
            .unwrap();
        assert_eq!(
            String::from_utf8(merged).unwrap(),
            FLAT.replace("Hello", "Bye")
        );
        // An unchanged edit is the flat document itself.
        let same = parts
            .merge(
                FLAT.as_bytes(),
                Some(parts.content_xml()),
                Some(parts.styles_xml()),
                parts.meta_xml(),
                None,
            )
            .unwrap();
        assert_eq!(same, FLAT.as_bytes());
    }

    #[test]
    fn refuses_root_and_structure_changes_and_shared_conflicts() {
        let parts = Parts::split(FLAT.as_bytes()).unwrap();
        let rooted = parts
            .content_xml()
            .replace("office:version=\"1.3\"", "office:version=\"1.4\"");
        assert!(matches!(
            parts.merge(FLAT.as_bytes(), Some(&rooted), None, None, None),
            Err(litchi_core::Error::Unsupported(_) | litchi_core::Error::Refused(_))
        ));
        let added = parts.content_xml().replace(
            "<office:automatic-styles/>",
            "<office:automatic-styles/><office:scripts/>",
        );
        assert!(
            parts
                .merge(FLAT.as_bytes(), Some(&added), None, None, None)
                .is_err()
        );
        let one = parts.content_xml().replace(
            "<office:automatic-styles/>",
            "<office:automatic-styles><a/></office:automatic-styles>",
        );
        let two = parts.styles_xml().replace(
            "<office:automatic-styles/>",
            "<office:automatic-styles><b/></office:automatic-styles>",
        );
        assert!(
            parts
                .merge(FLAT.as_bytes(), Some(&one), Some(&two), None, None)
                .is_err()
        );
        assert!(parts.merge(b"short", Some(&one), None, None, None).is_err());
    }

    #[test]
    fn merges_an_edited_package() {
        let parts = Parts::split(FLAT.as_bytes()).unwrap();
        let package = parts.to_bytes().unwrap();
        assert_eq!(
            parts.merge_package(FLAT.as_bytes(), &package).unwrap(),
            FLAT.as_bytes()
        );
    }

    #[test]
    fn a_created_meta_part_goes_first() {
        let flat = FLAT.replace(
            " <office:meta><dc:title xmlns:dc=\"http://purl.org/dc/elements/1.1/\">T</dc:title></office:meta>\n",
            "",
        );
        let parts = Parts::split(flat.as_bytes()).unwrap();
        assert_eq!(parts.meta_xml(), None);
        let meta = "<?xml version=\"1.0\"?><office:document-meta xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"><office:meta><x/></office:meta></office:document-meta>";
        let merged = String::from_utf8(
            parts
                .merge(flat.as_bytes(), None, None, Some(meta), None)
                .unwrap(),
        )
        .unwrap();
        assert!(merged.contains("application/vnd.oasis.opendocument.text\"><office:meta><x/></office:meta>\n <office:styles>"));
    }
}
