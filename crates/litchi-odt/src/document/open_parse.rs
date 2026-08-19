//! Fused single-tokenization open parse for ODT `content.xml`.
//!
//! [`crate::document::SourceBackedDocument`] opening historically tokenized
//! `content.xml` twice: once for the package-content structure validation
//! ([`litchi_odf_common::core::validate_content_document_part`]) and once for
//! the automatic-styles scan ([`StyleRegistry::from_xml`]).  This module runs
//! both over one shared `NsReader` event stream.
//!
//! Observable behavior is identical to the sequential passes:
//!
//! - The validation handler replicates every check of
//!   `validate_content_document_part` for the ODT root and family with
//!   identical error messages, applied to the same events in the same order;
//!   its mid-stream errors abort the scan and return immediately, exactly
//!   where the historical first pass early-returned before `styles.xml` was
//!   even fetched.
//! - The style handler replicates `StyleRegistry::from_xml` byte-exactly,
//!   including its literal raw-qualified-name `style:style` match (no
//!   namespace resolution), its raw undecoded attribute values, and its
//!   error messages; its first error is recorded and later events no longer
//!   mutate its state, matching the historical early return.
//! - A tokenization failure surfaces with the validation mapping
//!   (`invalid ODT content.xml: …`): both historical scans tokenize the same
//!   bytes with the same quick-xml core from the start, and validation ran
//!   first and read to EOF, so it reported every read failure before the
//!   style scan ever started.  The fused reader keeps the validation
//!   configuration (`check_end_names` and `check_comments`), which is a
//!   superset of the historical style-scan checks, so no tokenization error
//!   can be introduced or masked relative to history.
//!
//! Error precedence after a clean scan: the validation end-of-stream check
//! runs first (still ahead of any `styles.xml` work), then the caller parses
//! `styles.xml`, then [`OpenParse::finish`] surfaces the recorded
//! content-styles error, and finally `try_extend` runs — the historical
//! call order.
//!
//! The standalone passes stay byte-identical: `validate_content_document_part`
//! serves other callers, and the owned [`crate::document::Document`] facade
//! keeps the sequential sequence as the equivalence oracle.

use litchi_core::{Error, Result};
use quick_xml::events::{BytesRef, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::source::{CONTENT_ROOT, FAMILY_NAME};
use crate::elements::element::Element;
use crate::elements::style::{Style, StyleRegistry};

// Parity with the `validate_content_document_part` pre-scan limit in
// litchi-odf-common (`MAX_CONTENT_BYTES` there is crate-private).
const MAX_CONTENT_BYTES: usize = 256 * 1024 * 1024;
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";

/// The two-handler fused open parse over `content.xml`.
pub(crate) struct OpenParse {
    registry: StyleRegistry,
    style_error: Option<Error>,
}

impl OpenParse {
    /// Tokenize `content_xml` once, driving the content-validation and
    /// automatic-styles handlers over the shared event stream.
    ///
    /// A validation error (mid-stream, read failure, or end-of-stream) is
    /// returned immediately, before the caller fetches `styles.xml`, exactly
    /// where the historical first pass early-returned.  A content-styles
    /// error is only recorded; [`OpenParse::finish`] reports it after the
    /// `styles.xml` parse, matching the historical pass order.
    pub(crate) fn run(content_xml: &str) -> Result<Self> {
        if content_xml.len() > MAX_CONTENT_BYTES {
            return Err(Error::InvalidFormat(format!(
                "{FAMILY_NAME} content.xml exceeds the family limit"
            )));
        }
        let mut validate = ValidateHandler::new()?;
        let mut styles = StyleHandler::default();
        let mut style_error = None;

        let mut reader = NsReader::from_str(content_xml);
        reader.config_mut().check_end_names = true;
        reader.config_mut().check_comments = true;
        let mut buffer = Vec::new();
        loop {
            let (namespace, event) = match reader.read_resolved_event_into(&mut buffer) {
                Ok(resolved) => resolved,
                Err(error) => {
                    // Validation historically tokenized the same bytes first
                    // and read to EOF, so every read failure surfaced with
                    // its mapping before the style scan ever started.
                    return Err(Error::InvalidFormat(format!(
                        "invalid {FAMILY_NAME} content.xml: {error}"
                    )));
                },
            };
            // The resolved namespace borrows the reader mutably, so classify
            // it before the handlers touch the reader again, exactly where
            // the historical validation loop classified it.
            let office = matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == OFFICE_NAMESPACE);
            let is_eof = matches!(event, Event::Eof);
            // Validation errors early-returned historically, discarding any
            // style state; abort the scan at the same event.
            validate.on_event(office, &event)?;
            if style_error.is_none()
                && let Err(error) = styles.on_event(&event)
            {
                style_error = Some(error);
            }
            buffer.clear();
            if is_eof {
                break;
            }
        }
        // The historical validation pass completed (including its
        // end-of-stream checks) before the style scan started, so a
        // structural failure here still beats any recorded style error.
        validate.finish()?;

        Ok(Self {
            registry: styles.registry,
            style_error,
        })
    }

    /// Surface the recorded content-styles error or return the registry the
    /// historical `StyleRegistry::from_xml` pass would have built.
    pub(crate) fn finish(self) -> Result<StyleRegistry> {
        if let Some(error) = self.style_error {
            return Err(error);
        }
        Ok(self.registry)
    }
}

/// Streaming event handler replicating `validate_content_document_part` for
/// the ODT root (`office:document-content`), body (`office:body`), and family
/// element (`office:text`, derived from [`CONTENT_ROOT`]).
///
/// Every check, error message, and error position matches the standalone
/// pass, which keeps its own inline loop for the other callers.
#[derive(Debug)]
struct ValidateHandler {
    expected_local: String,
    depth: usize,
    root_closed: bool,
    body_seen: bool,
    expected_seen: bool,
    in_body: bool,
    declaration_seen: bool,
    first_event: bool,
}

impl ValidateHandler {
    fn new() -> Result<Self> {
        // Same marker derivation as the standalone pass; unreachable for the
        // fixed ODT root but kept at the same position.
        let expected_local = CONTENT_ROOT
            .strip_prefix("<office:")
            .and_then(|marker| {
                marker
                    .split(|character: char| !character.is_ascii_alphanumeric() && character != '-')
                    .next()
            })
            .filter(|local| !local.is_empty())
            .ok_or_else(|| {
                Error::InvalidFormat(format!("{FAMILY_NAME} content.xml has no expected body"))
            })?;
        Ok(Self {
            expected_local: expected_local.to_string(),
            depth: 0,
            root_closed: false,
            body_seen: false,
            expected_seen: false,
            in_body: false,
            declaration_seen: false,
            first_event: true,
        })
    }

    fn on_event(&mut self, office: bool, event: &Event<'_>) -> Result<()> {
        match event {
            Event::Start(element) => {
                if self.root_closed {
                    return Err(Error::InvalidFormat(format!(
                        "{FAMILY_NAME} content.xml has content after its root"
                    )));
                }
                let local = element.local_name();
                match self.depth {
                    0 if office && local.as_ref() == b"document-content" => self.depth = 1,
                    0 => {
                        return Err(Error::InvalidFormat(format!(
                            "{FAMILY_NAME} content.xml has the wrong root"
                        )));
                    },
                    1 if office && local.as_ref() == b"body" && !self.body_seen => {
                        self.body_seen = true;
                        self.in_body = true;
                        self.depth = 2;
                    },
                    1 if office && local.as_ref() == b"body" => {
                        return Err(Error::InvalidFormat(format!(
                            "{FAMILY_NAME} content.xml has duplicate office:body"
                        )));
                    },
                    1 => {
                        self.depth = self.depth.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormat(format!(
                                "{FAMILY_NAME} content.xml nesting overflows"
                            ))
                        })?;
                    },
                    2 if self.in_body
                        && office
                        && local.as_ref() == b"forms"
                        && !self.expected_seen =>
                    {
                        self.depth = 3;
                    },
                    2 if self.in_body
                        && office
                        && local.as_ref() == self.expected_local.as_bytes()
                        && !self.expected_seen =>
                    {
                        self.expected_seen = true;
                        self.depth = 3;
                    },
                    2 if self.in_body => {
                        return Err(Error::InvalidFormat(format!(
                            "{FAMILY_NAME} content.xml has the wrong office body"
                        )));
                    },
                    _ => {
                        self.depth = self.depth.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormat(format!(
                                "{FAMILY_NAME} content.xml nesting overflows"
                            ))
                        })?;
                    },
                }
            },
            Event::Empty(element) => {
                if self.root_closed || self.depth == 0 {
                    return Err(Error::InvalidFormat(format!(
                        "{FAMILY_NAME} content.xml has an invalid empty root"
                    )));
                }
                let local = element.local_name();
                if self.depth == 1 {
                    if office && local.as_ref() == b"body" && !self.body_seen {
                        self.body_seen = true;
                    } else if office && local.as_ref() == b"body" {
                        return Err(Error::InvalidFormat(format!(
                            "{FAMILY_NAME} content.xml has duplicate office:body"
                        )));
                    }
                } else if self.in_body
                    && self.depth == 2
                    && office
                    && local.as_ref() == b"forms"
                    && !self.expected_seen
                {
                    // `office:forms` may precede the family body.
                } else if self.in_body
                    && self.depth == 2
                    && office
                    && local.as_ref() == self.expected_local.as_bytes()
                    && !self.expected_seen
                {
                    self.expected_seen = true;
                } else if self.in_body && self.depth == 2 {
                    return Err(Error::InvalidFormat(format!(
                        "{FAMILY_NAME} content.xml has the wrong office body"
                    )));
                }
            },
            Event::End(_) => {
                self.depth = self.depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat(format!("{FAMILY_NAME} content.xml has an unexpected end"))
                })?;
                if self.depth == 0 {
                    self.root_closed = true;
                } else if self.in_body && self.depth == 1 {
                    self.in_body = false;
                }
            },
            Event::Text(text) => {
                if (self.depth == 0 || self.root_closed)
                    && !text.iter().all(u8::is_ascii_whitespace)
                {
                    return Err(Error::InvalidFormat(format!(
                        "{FAMILY_NAME} content.xml has unexpected text outside its root"
                    )));
                }
            },
            Event::CData(_) | Event::GeneralRef(_) if self.depth == 0 || self.root_closed => {
                return Err(Error::InvalidFormat(format!(
                    "{FAMILY_NAME} content.xml has content outside its root"
                )));
            },
            Event::DocType(_) => {
                return Err(Error::InvalidFormat(format!(
                    "{FAMILY_NAME} content.xml must not contain a doctype"
                )));
            },
            Event::GeneralRef(reference) if !valid_xml_reference(reference) => {
                return Err(Error::InvalidFormat(format!(
                    "{FAMILY_NAME} content.xml has an invalid character or entity reference"
                )));
            },
            Event::Decl(_)
                if self.declaration_seen
                    || !self.first_event
                    || self.depth != 0
                    || self.root_closed =>
            {
                return Err(Error::InvalidFormat(format!(
                    "{FAMILY_NAME} content.xml has an XML declaration outside its prologue"
                )));
            },
            Event::Decl(_) => self.declaration_seen = true,
            Event::Eof => {},
            Event::Comment(_) | Event::PI(_) | Event::CData(_) | Event::GeneralRef(_) => {},
        }
        self.first_event = false;
        Ok(())
    }

    fn finish(self) -> Result<()> {
        if !self.root_closed || self.depth != 0 || !self.body_seen || !self.expected_seen {
            return Err(Error::InvalidFormat(format!(
                "{FAMILY_NAME} content.xml has no complete expected body"
            )));
        }
        Ok(())
    }
}

// Byte-identical replication of the crate-private reference predicate in
// litchi-odf-common (`validation::valid_xml_reference`), which the standalone
// validation pass applies to general entity and character references.
fn valid_xml_reference(reference: &BytesRef<'_>) -> bool {
    let bytes: &[u8] = reference;
    matches!(bytes, b"amp" | b"lt" | b"gt" | b"apos" | b"quot")
        || (reference.is_char_ref()
            && reference
                .resolve_char_ref()
                .ok()
                .flatten()
                .is_some_and(xml10_character))
}

fn xml10_character(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}')
        || matches!(value as u32, 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
}

/// Streaming event handler replicating [`StyleRegistry::from_xml`]: the raw
/// qualified-name `style:style` match is intentionally literal and
/// prefix-sensitive, attribute values stay undecoded, and the error messages
/// are identical.
#[derive(Debug, Default)]
struct StyleHandler {
    registry: StyleRegistry,
}

impl StyleHandler {
    fn on_event(&mut self, event: &Event<'_>) -> Result<()> {
        if let Event::Start(element) = event
            && element.name().as_ref() == b"style:style"
        {
            let mut style_element = Element::try_new("style:style")?;
            for attribute in element.attributes() {
                let attribute = attribute.map_err(|error| {
                    Error::XmlError(format!("invalid ODT style attribute: {error}"))
                })?;
                let key = std::str::from_utf8(attribute.key.as_ref()).map_err(|error| {
                    Error::XmlError(format!(
                        "invalid UTF-8 in ODT style attribute name: {error}"
                    ))
                })?;
                let value = std::str::from_utf8(attribute.value.as_ref()).map_err(|error| {
                    Error::XmlError(format!(
                        "invalid UTF-8 in ODT style attribute value: {error}"
                    ))
                })?;
                style_element.try_set_attribute(key, value, "ODT style registry attribute")?;
            }
            self.registry
                .try_add_style(Style::from_element(style_element)?)?;
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::{CONTENT_ROOT, FAMILY_NAME, OpenParse};
    use crate::elements::style::StyleElements;
    use litchi_core::Result;
    use std::path::{Path, PathBuf};

    const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
    const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

    fn document(body: &str) -> String {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="{OFFICE}" xmlns:style="{STYLE}" xmlns:text="{TEXT}" office:version="1.4">{body}</office:document-content>"#
        )
    }

    /// Deterministic behavioral projection: `StyleRegistry` is not
    /// `PartialEq` and its `Debug` order follows the hash map, so compare
    /// sorted style names with their extracted attributes and properties.
    fn project(registry: &crate::elements::style::StyleRegistry) -> String {
        let mut names: Vec<&str> = registry.styles.keys().map(String::as_str).collect();
        names.sort_unstable();
        let mut projected = String::new();
        for name in names {
            let style = &registry.styles[name];
            projected.push_str(&format!(
                "{name}|{:?}|{:?}|{}\n",
                style.family(),
                style.parent_style_name(),
                {
                    let properties = style.properties();
                    format!(
                        "{:?}|{:?}|{:?}|{:?}",
                        properties.text, properties.paragraph, properties.table, properties.graphic
                    )
                }
            ));
        }
        projected
    }

    /// The historical sequential passes with their exact error precedence:
    /// content validation first, then the automatic-styles scan.
    fn sequential(xml: &str) -> Result<String> {
        litchi_odf_common::core::validate_content_document_part(xml, CONTENT_ROOT, FAMILY_NAME)?;
        let registry = StyleElements::parse_styles(xml)?;
        Ok(project(&registry))
    }

    /// The fused open parse with the same comparison projection.
    fn fused(xml: &str) -> Result<String> {
        let registry = OpenParse::run(xml)?.finish()?;
        Ok(project(&registry))
    }

    fn assert_equivalent(label: &str, xml: &str) {
        let expected = sequential(xml).map_err(|error| error.to_string());
        let actual = fused(xml).map_err(|error| error.to_string());
        assert_eq!(expected, actual, "{label}: fused and sequential disagree");
    }

    fn corpus_files() -> Vec<PathBuf> {
        let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
        let mut files = Vec::new();
        collect_odt(&root.join("test-data"), &mut files);
        files.sort();
        files
    }

    fn collect_odt(directory: &Path, files: &mut Vec<PathBuf>) {
        let Ok(entries) = std::fs::read_dir(directory) else {
            return;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                collect_odt(&path, files);
            } else if path
                .extension()
                .is_some_and(|extension| extension == "odt" || extension == "fodt")
            {
                files.push(path);
            }
        }
    }

    fn content_xml(path: &Path) -> Option<String> {
        let bytes = std::fs::read(path).ok()?;
        if path
            .extension()
            .is_some_and(|extension| extension == "fodt")
        {
            return String::from_utf8(bytes).ok();
        }
        let reader = soapberry_zip::office::ArchiveReader::new(&bytes).ok()?;
        let entry = reader.read("content.xml").ok()?;
        String::from_utf8(entry).ok()
    }

    #[test]
    fn fused_parse_matches_sequential_passes_on_odt_corpus() {
        let files = corpus_files();
        assert!(!files.is_empty(), "no .odt corpus fixtures discovered");
        let mut compared = 0usize;
        for path in &files {
            let Some(xml) = content_xml(path) else {
                continue;
            };
            assert_equivalent(&path.display().to_string(), &xml);
            compared += 1;
        }
        assert!(compared > 0, "no .odt corpus fixtures yielded content.xml");
    }

    #[test]
    fn synthetic_documents_match_the_sequential_passes() {
        let body_with_style = concat!(
            r#"<office:automatic-styles><style:style style:name="Body" style:family="paragraph"></style:style></office:automatic-styles>"#,
            r#"<office:body><office:text><text:p>hi</text:p></office:text></office:body>"#,
        );
        let cases: Vec<(&str, String)> = vec![
            ("minimal-with-style", document(body_with_style)),
            // Attribute values stay raw and undecoded in both passes.
            (
                "entity-in-style-name",
                document(concat!(
                    r#"<office:automatic-styles><style:style style:name="a&amp;b" style:family="paragraph"></style:style></office:automatic-styles>"#,
                    r#"<office:body><office:text/></office:body>"#,
                )),
            ),
            (
                "forms-before-family-body",
                document(r#"<office:body><office:forms/><office:text/></office:body>"#),
            ),
            (
                "wrong-root",
                document("").replace("document-content", "document"),
            ),
            (
                "incomplete-body",
                format!(
                    r#"<office:document-content xmlns:office="{OFFICE}"></office:document-content>"#
                ),
            ),
            (
                "empty-root",
                format!(r#"<office:document-content xmlns:office="{OFFICE}"/>"#),
            ),
            (
                "wrong-family-body",
                document(r#"<office:body><office:spreadsheet/></office:body>"#),
            ),
            (
                "duplicate-body",
                document(r#"<office:body><office:text/></office:body><office:body/>"#),
            ),
            (
                "doctype",
                format!(
                    r#"<!DOCTYPE office:document-content>{}"#,
                    document(r#"<office:body><office:text/></office:body>"#)
                ),
            ),
            (
                "trailing-text",
                format!(
                    "{}tail",
                    document(r#"<office:body><office:text/></office:body>"#)
                ),
            ),
            (
                "content-after-root",
                format!(
                    "{}<office:document-content/>",
                    document(r#"<office:body><office:text/></office:body>"#)
                ),
            ),
            (
                "late-declaration",
                document(r#"<?xml version="1.1"?><office:body><office:text/></office:body>"#),
            ),
            (
                "invalid-character-reference",
                document(
                    r#"<office:body><office:text><text:p>&#xD800;</text:p></office:text></office:body>"#,
                ),
            ),
            (
                "mismatched-end-tag",
                format!(
                    r#"<office:document-content xmlns:office="{OFFICE}"><office:body><office:text></office:body></office:document-content>"#
                ),
            ),
            (
                "duplicate-style-attribute",
                document(concat!(
                    r#"<office:automatic-styles><style:style style:name="a" style:name="b"></style:style></office:automatic-styles>"#,
                    r#"<office:body><office:text/></office:body>"#,
                )),
            ),
            // A style error recorded before a later tokenization failure
            // still loses to the validation-side read error, matching the
            // historical order in which validation tokenized first.
            (
                "style-error-before-malformed-xml",
                format!(
                    r#"<office:document-content xmlns:office="{OFFICE}" xmlns:style="{STYLE}"><office:automatic-styles><style:style style:name="a" style:name="b"></style:style></office:automatic-styles><office:body><office:text></office:body></office:document-content>"#
                ),
            ),
        ];
        for (label, xml) in &cases {
            assert_equivalent(label, xml);
        }
    }

    #[test]
    fn style_matching_is_literal_and_prefix_sensitive() {
        // A custom prefix bound to the style namespace is NOT collected: the
        // historical scan matches the raw bytes `style:style`.
        let custom_prefix = format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:s="{STYLE}"><office:automatic-styles><s:style s:name="Ghost" s:family="paragraph"></s:style></office:automatic-styles><office:body><office:text/></office:body></office:document-content>"#
        );
        assert_equivalent("custom-style-prefix-not-collected", &custom_prefix);
        let registry = OpenParse::run(&custom_prefix)
            .and_then(OpenParse::finish)
            .expect("custom-prefix document parses");
        assert!(
            registry.get_style("Ghost").is_none(),
            "literal byte matching must not collect a custom prefix"
        );

        // The canonical prefix bound to a FOREIGN namespace IS collected:
        // the match is on bytes, not on the resolved namespace.
        let foreign_binding = format!(
            r#"<office:document-content xmlns:office="{OFFICE}" xmlns:style="urn:foreign"><office:automatic-styles><style:style style:name="Kept" style:family="paragraph"></style:style></office:automatic-styles><office:body><office:text/></office:body></office:document-content>"#
        );
        assert_equivalent("canonical-prefix-foreign-uri-collected", &foreign_binding);
        let registry = OpenParse::run(&foreign_binding)
            .and_then(OpenParse::finish)
            .expect("foreign-binding document parses");
        assert!(
            registry.get_style("Kept").is_some(),
            "literal byte matching must collect the canonical prefix"
        );
    }
}
