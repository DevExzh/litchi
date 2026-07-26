//! Verbatim retention of a presentation's original `content.xml` skeleton.
//!
//! [`MutablePresentation`](crate::odp::MutablePresentation) models only the
//! subset of ODF drawing content it understands. Regenerating `content.xml`
//! purely from that model discards every construct outside the model — nested
//! `table:table` shapes, `draw:image` text alternatives, custom shape trees,
//! automatic styles, and font declarations among them.
//!
//! This module splits the source `content.xml` into reusable byte ranges so the
//! writer can re-emit untouched slides exactly as they were authored and only
//! synthesise markup for slides the caller actually changed.

use litchi_core::{Error, Result};
use std::collections::{BTreeMap, BTreeSet};
mod scanner;

use scanner::Scanner;

/// OASIS `office` namespace.
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
/// OASIS `drawing` namespace.
const DRAW_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
/// OASIS `presentation` namespace.
const PRESENTATION_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
/// OASIS `style` namespace.
const STYLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";

/// Root element of a package `content.xml` stream.
const DOCUMENT_CONTENT_ELEMENT: &str = "document-content";
/// Body wrapper element.
const BODY_ELEMENT: &str = "body";
/// Presentation body element.
const PRESENTATION_ELEMENT: &str = "presentation";
/// Slide element.
const PAGE_ELEMENT: &str = "page";
/// Automatic style container preceding `office:body`.
const AUTOMATIC_STYLES_ELEMENT: &str = "automatic-styles";
/// Qualified name of the automatic style container.
const AUTOMATIC_STYLES_QNAME: &str = "office:automatic-styles";
/// Named style definition inside `office:automatic-styles`.
const STYLE_ELEMENT: &str = "style";
/// Attribute naming a style definition.
const STYLE_NAME_ATTRIBUTE: &[u8] = b"name";
/// Slide-show configuration element regenerated from the settings model.
const SETTINGS_ELEMENT: &str = "settings";
/// Declaration elements regenerated from the declarations model.
const DECLARATION_ELEMENTS: [&str; 3] = ["header-decl", "footer-decl", "date-time-decl"];
/// Default namespace declaration attribute.
const XMLNS_DEFAULT: &[u8] = b"xmlns";
/// Prefixed namespace declaration attribute prefix.
const XMLNS_PREFIX: &[u8] = b"xmlns:";

/// Largest `content.xml` accepted for verbatim retention.
const MAX_XML_BYTES: usize = 64 * 1024 * 1024;
/// Largest element nesting depth accepted while scanning.
const MAX_DEPTH: usize = 512;
/// Largest number of slides accepted while scanning.
const MAX_PAGES: usize = 65_536;
/// Largest number of automatic style names recorded for collision checks.
const MAX_STYLE_NAMES: usize = 262_144;
/// Bytes of `</>` punctuation reserved per emitted end tag.
const CLOSE_TAG_OVERHEAD: usize = 9;
/// UTF-8 byte-order mark, stripped before scanning and restored on write.
const BYTE_ORDER_MARK: &str = "\u{feff}";

/// Depth of `office:document-content` children.
const PROLOGUE_DEPTH: usize = 1;
/// Depth of `office:body` children.
const BODY_CHILD_DEPTH: usize = 2;
/// Depth of `office:presentation` children.
const PRESENTATION_CHILD_DEPTH: usize = 3;

/// Namespace bindings required by markup this crate synthesises.
///
/// A retained root element keeps its own declarations; any binding below that
/// it does not already provide is appended so generated slides stay well
/// formed.
const REQUIRED_NAMESPACES: [(&str, &str); 10] = [
    ("office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"),
    ("style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0"),
    ("text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0"),
    ("draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"),
    (
        "presentation",
        "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0",
    ),
    (
        "anim",
        "urn:oasis:names:tc:opendocument:xmlns:animation:1.0",
    ),
    (
        "smil",
        "urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0",
    ),
    (
        "svg",
        "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
    ),
    ("xlink", "http://www.w3.org/1999/xlink"),
    ("script", "urn:oasis:names:tc:opendocument:xmlns:script:1.0"),
];

/// Half-open byte range into the retained source text.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct XmlSpan {
    start: usize,
    end: usize,
}

/// Where synthesised automatic styles are spliced into the retained prologue.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum AutomaticStylesSite {
    /// The container has children; new styles append before its end tag.
    Content,
    /// The container is self-closing and must be expanded to hold children.
    Empty { span: XmlSpan, name_end: usize },
    /// No container exists; a fresh one is emitted before `office:body`.
    Missing,
}

/// The original `content.xml` split into reusable, verbatim fragments.
///
/// All accessors borrow from the retained source text, so re-emitting an
/// untouched slide costs no allocation beyond the output buffer.
#[derive(Debug, Clone)]
pub(super) struct PresentationContentSource {
    /// Complete original `content.xml` text.
    xml: String,
    /// Everything before the `office:document-content` start tag.
    prolog: XmlSpan,
    /// The `office:document-content` start tag, closing `>` excluded.
    root_open: XmlSpan,
    /// Children of `office:document-content` preceding the automatic-style splice point.
    prologue_head: XmlSpan,
    /// Children of `office:document-content` following the splice point, up to `office:body`.
    prologue_tail: XmlSpan,
    /// How synthesised automatic styles join the retained prologue.
    styles_site: AutomaticStylesSite,
    /// The `office:body` start tag.
    body_open: XmlSpan,
    /// The `office:presentation` start tag.
    presentation_open: XmlSpan,
    /// Presentation children before the first slide that this crate does not model.
    leading_extras: Vec<XmlSpan>,
    /// Each `draw:page` element, in slide order.
    pages: Vec<XmlSpan>,
    /// Presentation children after the last slide that this crate does not model.
    trailing_extras: Vec<XmlSpan>,
    /// Style names already defined in `office:automatic-styles`.
    style_names: BTreeSet<String>,
    /// Namespace prefixes bound on the root element, mapped to their URIs.
    root_namespaces: BTreeMap<String, String>,
    /// Whether the source began with a UTF-8 byte-order mark.
    has_byte_order_mark: bool,
}

impl PresentationContentSource {
    /// Split a presentation `content.xml` into verbatim fragments.
    ///
    /// Returns `Ok(None)` when the stream does not have the expected
    /// presentation shape, letting the caller fall back to full regeneration
    /// rather than failing a save.
    pub(super) fn parse(xml: &str) -> Result<Option<Self>> {
        if xml.len() > MAX_XML_BYTES {
            return Err(invalid(format!(
                "presentation content.xml exceeds {MAX_XML_BYTES} bytes"
            )));
        }
        // `quick-xml` reports byte positions relative to the text after any
        // byte-order mark, so the mark is split off here and re-emitted
        // verbatim when the document is written back.
        let body = xml.strip_prefix(BYTE_ORDER_MARK).unwrap_or(xml);
        let has_byte_order_mark = body.len() != xml.len();
        let mut scanner = Scanner::new(body);
        scanner.run()?;
        Ok(scanner.finish(body, has_byte_order_mark))
    }

    /// Append the byte-order mark, XML declaration, and any other markup that
    /// precedes the root element.
    pub(super) fn write_prolog(&self, output: &mut String) {
        if self.has_byte_order_mark {
            output.push_str(BYTE_ORDER_MARK);
        }
        output.push_str(self.slice(self.prolog));
    }

    /// Number of slides retained from the source document.
    pub(super) fn page_count(&self) -> usize {
        self.pages.len()
    }

    /// The verbatim `draw:page` element at `index`, when it exists.
    pub(super) fn page(&self, index: usize) -> Option<&str> {
        self.pages.get(index).map(|span| self.slice(*span))
    }

    /// The `office:document-content` start tag with `extra_attributes` and,
    /// when `synthesised` markup is present, any namespace binding that markup
    /// relies on but the source does not already declare.
    ///
    /// Fails when synthesised markup is present and the source binds one of the
    /// prefixes this crate emits to a different namespace, because reusing that
    /// prefix would silently change the meaning of the generated elements.
    pub(super) fn root_start_tag(
        &self,
        extra_attributes: &str,
        synthesised: bool,
    ) -> Result<String> {
        let base = self.slice(self.root_open);
        let mut output = String::with_capacity(base.len() + extra_attributes.len() + 64);
        output.push_str(base);
        if !synthesised {
            output.push_str(extra_attributes);
            output.push('>');
            return Ok(output);
        }
        for (prefix, uri) in REQUIRED_NAMESPACES {
            match self.root_namespaces.get(prefix) {
                Some(bound) if bound == uri => continue,
                Some(bound) => {
                    return Err(invalid(format!(
                        "presentation content.xml binds prefix '{prefix}' to '{bound}' instead of '{uri}'"
                    )));
                },
                None => {},
            }
            output.push_str(" xmlns:");
            output.push_str(prefix);
            output.push_str("=\"");
            output.push_str(uri);
            output.push('"');
        }
        output.push_str(extra_attributes);
        output.push('>');
        Ok(output)
    }

    /// Whether `name` is already declared inside `office:automatic-styles`.
    pub(super) fn defines_style(&self, name: &str) -> bool {
        self.style_names.contains(name)
    }

    /// Emit the retained prologue with `generated_styles` spliced into
    /// `office:automatic-styles`.
    pub(super) fn write_prologue(&self, output: &mut String, generated_styles: &str) {
        output.push_str(self.slice(self.prologue_head));
        match self.styles_site {
            AutomaticStylesSite::Content => output.push_str(generated_styles),
            AutomaticStylesSite::Empty { span, name_end } => {
                if generated_styles.is_empty() {
                    output.push_str(self.slice(span));
                } else {
                    output.push_str(self.xml.get(span.start..name_end).unwrap_or_default());
                    output.push('>');
                    output.push_str(generated_styles);
                    output.push_str("</");
                    output.push_str(AUTOMATIC_STYLES_QNAME);
                    output.push('>');
                }
            },
            AutomaticStylesSite::Missing => {
                if !generated_styles.is_empty() {
                    output.push('<');
                    output.push_str(AUTOMATIC_STYLES_QNAME);
                    output.push('>');
                    output.push_str(generated_styles);
                    output.push_str("</");
                    output.push_str(AUTOMATIC_STYLES_QNAME);
                    output.push('>');
                }
            },
        }
        output.push_str(self.slice(self.prologue_tail));
    }

    /// The `office:body` start tag, verbatim.
    pub(super) fn body_start_tag(&self) -> &str {
        self.slice(self.body_open)
    }

    /// The `office:presentation` start tag, verbatim.
    pub(super) fn presentation_start_tag(&self) -> &str {
        self.slice(self.presentation_open)
    }

    /// End tags closing `office:presentation`, `office:body`, and the root.
    ///
    /// The qualified names are taken from the retained start tags so a document
    /// that binds the `office` namespace to a non-conventional prefix still
    /// round-trips.
    pub(super) fn close_tags(&self) -> String {
        let root = qname_of(self.slice(self.root_open));
        let body = qname_of(self.slice(self.body_open));
        let presentation = qname_of(self.slice(self.presentation_open));
        let mut output = String::with_capacity(
            root.len() + body.len() + presentation.len() + CLOSE_TAG_OVERHEAD,
        );
        for name in [presentation, body, root] {
            output.push_str("</");
            output.push_str(name);
            output.push('>');
        }
        output
    }

    /// Append unmodelled presentation children that precede the first slide.
    pub(super) fn write_leading_extras(&self, output: &mut String) {
        for span in &self.leading_extras {
            output.push_str(self.slice(*span));
        }
    }

    /// Append unmodelled presentation children that follow the last slide.
    pub(super) fn write_trailing_extras(&self, output: &mut String) {
        for span in &self.trailing_extras {
            output.push_str(self.slice(*span));
        }
    }

    fn slice(&self, span: XmlSpan) -> &str {
        self.xml.get(span.start..span.end).unwrap_or_default()
    }
}

/// Extract the qualified element name from a raw start tag.
fn qname_of(tag: &str) -> &str {
    let rest = tag.strip_prefix('<').unwrap_or(tag);
    let end = rest
        .find(|character: char| character.is_whitespace() || character == '>' || character == '/')
        .unwrap_or(rest.len());
    &rest[..end]
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

    const CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" office:version="1.3"><office:scripts><office:script office:name="s"/></office:scripts><office:font-face-decls/><office:automatic-styles><style:style style:name="dp1" style:family="drawing-page"/></office:automatic-styles><office:body><office:presentation><draw:page draw:name="page1"><draw:frame><table:table><table:table-row><table:table-cell><text:p>0,0</text:p></table:table-cell></table:table-row></table:table></draw:frame></draw:page><presentation:settings presentation:mouse-visible="false"/></office:presentation></office:body></office:document-content>"#;

    #[test]
    fn retains_pages_prologue_and_style_names() {
        let source = PresentationContentSource::parse(CONTENT)
            .unwrap()
            .expect("presentation skeleton");
        assert_eq!(source.page_count(), 1);
        let page = source.page(0).expect("page");
        assert!(page.starts_with(r#"<draw:page draw:name="page1">"#));
        assert!(page.ends_with("</draw:page>"));
        assert!(page.contains("<text:p>0,0</text:p>"));
        assert!(source.defines_style("dp1"));
        assert!(!source.defines_style("dp2"));
        let mut prolog = String::new();
        source.write_prolog(&mut prolog);
        assert_eq!(prolog, r#"<?xml version="1.0" encoding="UTF-8"?>"#);
        assert_eq!(source.body_start_tag(), "<office:body>");
        assert_eq!(source.presentation_start_tag(), "<office:presentation>");
    }

    #[test]
    fn splices_generated_styles_into_existing_container() {
        let source = PresentationContentSource::parse(CONTENT)
            .unwrap()
            .expect("presentation skeleton");
        let mut output = String::new();
        source.write_prologue(&mut output, r#"<style:style style:name="dpX"/>"#);
        assert!(output.contains(r#"<office:script office:name="s"/>"#));
        assert!(output.contains(
            r#"<style:style style:name="dp1" style:family="drawing-page"/><style:style style:name="dpX"/></office:automatic-styles>"#
        ));
    }

    #[test]
    fn creates_container_when_absent() {
        let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:body><office:presentation><draw:page/></office:presentation></office:body></office:document-content>"#;
        let source = PresentationContentSource::parse(xml)
            .unwrap()
            .expect("presentation skeleton");
        let mut output = String::new();
        source.write_prologue(&mut output, "<style:style/>");
        assert_eq!(
            output,
            "<office:automatic-styles><style:style/></office:automatic-styles>"
        );
        assert_eq!(source.page_count(), 1);
        assert_eq!(source.page(0), Some("<draw:page/>"));
    }

    #[test]
    fn expands_self_closing_container() {
        let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:automatic-styles/><office:body><office:presentation><draw:page/></office:presentation></office:body></office:document-content>"#;
        let source = PresentationContentSource::parse(xml)
            .unwrap()
            .expect("presentation skeleton");
        let mut output = String::new();
        source.write_prologue(&mut output, "<style:style/>");
        assert_eq!(
            output,
            "<office:automatic-styles><style:style/></office:automatic-styles>"
        );
        let mut untouched = String::new();
        source.write_prologue(&mut untouched, "");
        assert_eq!(untouched, "<office:automatic-styles/>");
    }

    #[test]
    fn appends_missing_namespace_bindings() {
        let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><office:body><office:presentation><draw:page/></office:presentation></office:body></office:document-content>"#;
        let source = PresentationContentSource::parse(xml)
            .unwrap()
            .expect("presentation skeleton");
        let tag = source.root_start_tag("", true).unwrap();
        assert!(
            tag.contains(r#"xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0""#)
        );
        assert_eq!(
            tag.matches(r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0""#)
                .count(),
            1
        );
        assert!(tag.ends_with('>'));
    }

    #[test]
    fn rejects_conflicting_prefix_binding() {
        let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:example:not-drawing"><office:body><office:presentation><draw:page/></office:presentation></office:body></office:document-content>"#;
        let source = PresentationContentSource::parse(xml)
            .unwrap()
            .expect("presentation skeleton");
        let error = source.root_start_tag("", true).unwrap_err().to_string();
        assert!(error.contains("binds prefix 'draw'"), "{error}");
    }

    #[test]
    fn rejects_non_presentation_root() {
        let xml = r#"<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>"#;
        assert!(PresentationContentSource::parse(xml).is_err());
    }

    #[test]
    fn reports_unmodelled_presentation_children() {
        let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"><office:body><office:presentation><draw:custom draw:name="x"/><draw:page/><presentation:settings/><draw:other/></office:presentation></office:body></office:document-content>"#;
        let source = PresentationContentSource::parse(xml)
            .unwrap()
            .expect("presentation skeleton");
        let mut leading = String::new();
        source.write_leading_extras(&mut leading);
        let mut trailing = String::new();
        source.write_trailing_extras(&mut trailing);
        assert_eq!(leading, r#"<draw:custom draw:name="x"/>"#);
        assert_eq!(trailing, "<draw:other/>");
    }
}
