//! Single-pass scan that records the byte spans a presentation save reuses.
//!
//! The scanner walks `content.xml` once and remembers where the root element,
//! the prologue, the automatic-style container, and each `draw:page` begin and
//! end, so [`ContentSource`](super::ContentSource) can
//! slice them back out verbatim.

use super::{
    AUTOMATIC_STYLES_ELEMENT, AutomaticStylesSite, BODY_CHILD_DEPTH, BODY_ELEMENT, ContentSource,
    DECLARATION_ELEMENTS, DOCUMENT_CONTENT_ELEMENT, DRAW_NAMESPACE, MAX_DEPTH, MAX_PAGES,
    MAX_STYLE_NAMES, OFFICE_NAMESPACE, PAGE_ELEMENT, PRESENTATION_CHILD_DEPTH,
    PRESENTATION_ELEMENT, PRESENTATION_NAMESPACE, PROLOGUE_DEPTH, SETTINGS_ELEMENT, STYLE_ELEMENT,
    STYLE_NAME_ATTRIBUTE, STYLE_NAMESPACE, XMLNS_DEFAULT, XMLNS_PREFIX, XmlSpan, invalid,
};
use litchi_core::Result;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeMap, BTreeSet};

/// One open element on the scanner stack.
struct Frame {
    kind: ElementKind,
    start: usize,
}

/// Incremental scanner that records the spans a save needs to reuse.
pub(super) struct Scanner<'input> {
    xml: &'input str,
    reader: NsReader<&'input [u8]>,
    stack: Vec<Frame>,
    root_open: Option<XmlSpan>,
    prologue_start: Option<usize>,
    body_open: Option<XmlSpan>,
    presentation_open: Option<XmlSpan>,
    /// Splice offset paired with the container shape, once one is seen.
    styles_site: Option<(usize, AutomaticStylesSite)>,
    pages: Vec<XmlSpan>,
    leading_extras: Vec<XmlSpan>,
    trailing_extras: Vec<XmlSpan>,
    style_names: BTreeSet<String>,
    root_namespaces: BTreeMap<String, String>,
    /// Byte offset of the open `draw:page`, once one is entered.
    open_page: Option<usize>,
    /// Byte offset of an open unmodelled presentation child.
    open_extra: Option<usize>,
    /// Depth at which the automatic-style container was entered.
    styles_depth: Option<usize>,
}

impl<'input> Scanner<'input> {
    pub(super) fn new(xml: &'input str) -> Self {
        let mut reader = NsReader::from_str(xml);
        reader.config_mut().check_end_names = true;
        Self {
            xml,
            reader,
            stack: Vec::new(),
            root_open: None,
            prologue_start: None,
            body_open: None,
            presentation_open: None,
            styles_site: None,
            pages: Vec::new(),
            leading_extras: Vec::new(),
            trailing_extras: Vec::new(),
            style_names: BTreeSet::new(),
            root_namespaces: BTreeMap::new(),
            open_page: None,
            open_extra: None,
            styles_depth: None,
        }
    }

    pub(super) fn run(&mut self) -> Result<()> {
        loop {
            let (kind, event) = {
                let (resolved, event) = self.reader.read_resolved_event().map_err(|error| {
                    invalid(format!("invalid presentation content.xml: {error}"))
                })?;
                let kind = match &event {
                    Event::Start(element) | Event::Empty(element) => {
                        classify(&resolved, element.local_name().as_ref())
                    },
                    _ => ElementKind::Other,
                };
                (kind, event)
            };
            let end = usize::try_from(self.reader.buffer_position()).map_err(|_error| {
                invalid("presentation content.xml position exceeds platform limits")
            })?;
            match event {
                Event::Start(element) => {
                    let start = event_start(self.xml, end)?;
                    if self.stack.is_empty() {
                        self.collect_root_namespaces(&element)?;
                    }
                    self.on_open(kind, start, end, false)?;
                    if self.styles_depth.is_some() && kind == ElementKind::Style {
                        self.record_style_name(&element)?;
                    }
                    if self.stack.len() >= MAX_DEPTH {
                        return Err(invalid(format!(
                            "presentation content.xml nests deeper than {MAX_DEPTH} elements"
                        )));
                    }
                    self.stack.push(Frame { kind, start });
                },
                Event::Empty(element) => {
                    let start = event_start(self.xml, end)?;
                    if self.styles_depth.is_some() && kind == ElementKind::Style {
                        self.record_style_name(&element)?;
                    }
                    self.on_open(kind, start, end, true)?;
                    self.on_close(kind, start, start, end);
                },
                Event::End(_) => {
                    let close_start = event_start(self.xml, end)?;
                    let frame = self
                        .stack
                        .pop()
                        .ok_or_else(|| invalid("presentation content.xml depth underflow"))?;
                    self.on_close(frame.kind, frame.start, close_start, end);
                },
                Event::Eof => break,
                Event::Text(_)
                | Event::CData(_)
                | Event::Comment(_)
                | Event::Decl(_)
                | Event::PI(_)
                | Event::DocType(_)
                | Event::GeneralRef(_) => {},
            }
        }
        if !self.stack.is_empty() {
            return Err(invalid("unterminated presentation content.xml element"));
        }
        Ok(())
    }

    /// Handle an element start, or the start half of an empty element.
    fn on_open(&mut self, kind: ElementKind, start: usize, end: usize, empty: bool) -> Result<()> {
        let depth = self.stack.len();
        if depth == 0 {
            if empty || kind != ElementKind::DocumentContent {
                return Err(invalid(
                    "presentation content.xml root is not office:document-content",
                ));
            }
            let open_end = self
                .xml
                .get(start..end)
                .and_then(|tag| tag.rfind('>'))
                .map(|offset| start + offset)
                .ok_or_else(|| invalid("malformed office:document-content start tag"))?;
            self.root_open = Some(XmlSpan {
                start,
                end: open_end,
            });
            self.prologue_start = Some(end);
            return Ok(());
        }
        if self.open_page.is_some() || self.open_extra.is_some() {
            return Ok(());
        }
        if depth == PROLOGUE_DEPTH && kind == ElementKind::AutomaticStyles {
            if empty {
                let name_end = self
                    .xml
                    .get(start..end)
                    .and_then(|tag| tag.rfind("/>"))
                    .map(|offset| start + offset)
                    .ok_or_else(|| invalid("malformed office:automatic-styles element"))?;
                self.styles_site = Some((
                    start,
                    AutomaticStylesSite::Empty {
                        span: XmlSpan { start, end },
                        name_end,
                    },
                ));
            } else {
                self.styles_depth = Some(depth);
            }
            return Ok(());
        }
        if depth == PROLOGUE_DEPTH && !empty && kind == ElementKind::Body {
            self.body_open = Some(XmlSpan { start, end });
            return Ok(());
        }
        if depth == BODY_CHILD_DEPTH && !empty && kind == ElementKind::Presentation {
            self.presentation_open = Some(XmlSpan { start, end });
            return Ok(());
        }
        if depth == PRESENTATION_CHILD_DEPTH && self.presentation_open.is_some() {
            if kind == ElementKind::Page {
                if self.pages.len() >= MAX_PAGES {
                    return Err(invalid(format!(
                        "presentation content.xml holds more than {MAX_PAGES} slides"
                    )));
                }
                self.open_page = Some(start);
            } else if kind != ElementKind::Regenerated {
                self.open_extra = Some(start);
            }
        }
        Ok(())
    }

    /// Handle an element end, or the end half of an empty element.
    ///
    /// `open_start` is the offset of the matching start tag and `close_start`
    /// the offset of the end tag currently being processed; they coincide for
    /// empty elements.
    fn on_close(&mut self, kind: ElementKind, open_start: usize, close_start: usize, end: usize) {
        if self.open_page == Some(open_start) {
            self.open_page = None;
            self.pages.push(XmlSpan {
                start: open_start,
                end,
            });
            return;
        }
        if self.open_extra == Some(open_start) {
            self.open_extra = None;
            let span = XmlSpan {
                start: open_start,
                end,
            };
            if self.pages.is_empty() {
                self.leading_extras.push(span);
            } else {
                self.trailing_extras.push(span);
            }
            return;
        }
        if self.styles_depth == Some(self.stack.len()) && kind == ElementKind::AutomaticStyles {
            self.styles_depth = None;
            self.styles_site = Some((close_start, AutomaticStylesSite::Content));
        }
    }

    /// Record every namespace binding declared on the root element.
    fn collect_root_namespaces(&mut self, element: &BytesStart<'_>) -> Result<()> {
        for raw_attribute in element.attributes() {
            let attribute = raw_attribute.map_err(|error| {
                invalid(format!(
                    "invalid office:document-content attribute: {error}"
                ))
            })?;
            let key = attribute.key.as_ref();
            let prefix = if key == XMLNS_DEFAULT {
                String::new()
            } else if let Some(rest) = key.strip_prefix(XMLNS_PREFIX) {
                String::from_utf8(rest.to_vec()).map_err(|error| {
                    invalid(format!("non-UTF-8 namespace prefix declaration: {error}"))
                })?
            } else {
                continue;
            };
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, self.reader.decoder())
                .map_err(|error| invalid(format!("invalid namespace declaration: {error}")))?;
            self.root_namespaces.insert(prefix, value.into_owned());
        }
        Ok(())
    }

    fn record_style_name(&mut self, element: &BytesStart<'_>) -> Result<()> {
        if self.style_names.len() >= MAX_STYLE_NAMES {
            return Err(invalid(format!(
                "presentation content.xml declares more than {MAX_STYLE_NAMES} automatic styles"
            )));
        }
        for raw_attribute in element.attributes() {
            let attribute = raw_attribute
                .map_err(|error| invalid(format!("invalid style attribute: {error}")))?;
            let (namespace, local) = self.reader.resolver().resolve_attribute(attribute.key);
            if local.as_ref() != STYLE_NAME_ATTRIBUTE
                || !matches!(namespace, ResolveResult::Bound(Namespace(uri)) if uri == STYLE_NAMESPACE)
            {
                continue;
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, self.reader.decoder())
                .map_err(|error| invalid(format!("invalid style:name value: {error}")))?;
            self.style_names.insert(value.into_owned());
        }
        Ok(())
    }

    /// Assemble the retained skeleton, or `None` when the shape is unexpected.
    pub(super) fn finish(self, xml: &str, has_byte_order_mark: bool) -> Option<ContentSource> {
        let root_open = self.root_open?;
        let body_open = self.body_open?;
        let presentation_open = self.presentation_open?;
        let prologue_start = self.prologue_start?;
        if body_open.start < prologue_start {
            return None;
        }
        let (splice, styles_site) = self
            .styles_site
            .unwrap_or((body_open.start, AutomaticStylesSite::Missing));
        if splice < prologue_start || splice > body_open.start {
            return None;
        }
        let tail_start = match styles_site {
            AutomaticStylesSite::Empty { span, .. } => span.end,
            AutomaticStylesSite::Content | AutomaticStylesSite::Missing => splice,
        };
        Some(ContentSource {
            xml: xml.to_owned(),
            prolog: XmlSpan {
                start: 0,
                end: root_open.start,
            },
            root_open,
            prologue_head: XmlSpan {
                start: prologue_start,
                end: splice,
            },
            prologue_tail: XmlSpan {
                start: tail_start,
                end: body_open.start,
            },
            styles_site,
            body_open,
            presentation_open,
            leading_extras: self.leading_extras,
            pages: self.pages,
            trailing_extras: self.trailing_extras,
            style_names: self.style_names,
            root_namespaces: self.root_namespaces,
            has_byte_order_mark,
        })
    }
}

/// Locate the `<` that opened the event ending at `end`.
fn event_start(xml: &str, end: usize) -> Result<usize> {
    xml.get(..end)
        .and_then(|prefix| prefix.rfind('<'))
        .ok_or_else(|| invalid("invalid presentation content.xml event boundary"))
}

/// Exact distinctions needed after an event's namespace scope has ended.
/// Unknown names and namespaces stay opaque; no original markup is rewritten.
#[derive(Clone, Copy, PartialEq, Eq)]
enum ElementKind {
    DocumentContent,
    AutomaticStyles,
    Body,
    Presentation,
    Page,
    Style,
    Regenerated,
    Other,
}

fn classify(resolved: &ResolveResult<'_>, local: &[u8]) -> ElementKind {
    let ResolveResult::Bound(Namespace(uri)) = resolved else {
        return ElementKind::Other;
    };
    if *uri == OFFICE_NAMESPACE {
        for (name, kind) in [
            (DOCUMENT_CONTENT_ELEMENT, ElementKind::DocumentContent),
            (AUTOMATIC_STYLES_ELEMENT, ElementKind::AutomaticStyles),
            (BODY_ELEMENT, ElementKind::Body),
            (PRESENTATION_ELEMENT, ElementKind::Presentation),
        ] {
            if local == name.as_bytes() {
                return kind;
            }
        }
    } else if *uri == DRAW_NAMESPACE && local == PAGE_ELEMENT.as_bytes() {
        return ElementKind::Page;
    } else if *uri == STYLE_NAMESPACE && local == STYLE_ELEMENT.as_bytes() {
        return ElementKind::Style;
    } else if *uri == PRESENTATION_NAMESPACE
        && (local == SETTINGS_ELEMENT.as_bytes()
            || DECLARATION_ELEMENTS
                .iter()
                .any(|name| local == name.as_bytes()))
    {
        return ElementKind::Regenerated;
    }
    ElementKind::Other
}

#[cfg(test)]
mod tests {
    use super::super::{BYTE_ORDER_MARK, MAX_XML_BYTES, scanner_reference};
    use super::*;

    fn reference(xml: &str) -> Result<Option<ContentSource>> {
        if xml.len() > MAX_XML_BYTES {
            return Err(invalid(format!(
                "presentation content.xml exceeds {MAX_XML_BYTES} bytes"
            )));
        }
        let body = xml.strip_prefix(BYTE_ORDER_MARK).unwrap_or(xml);
        let has_byte_order_mark = body.len() != xml.len();
        let mut scanner = scanner_reference::Scanner::new(body);
        scanner.run()?;
        Ok(scanner.finish(body, has_byte_order_mark))
    }

    fn compare(xml: &str) {
        // Debug includes every private span, root binding, style name and source
        // byte. The reference retains the original complete scanner unchanged.
        let normalize = |value: Result<Option<ContentSource>>| {
            value
                .map(|source| format!("{source:?}"))
                .map_err(|error| format!("{error:?}"))
        };
        assert_eq!(
            normalize(ContentSource::parse(xml)),
            normalize(reference(xml)),
            "{xml:.200}"
        );
    }

    fn wrap(prologue: &str, body: &str) -> String {
        format!(
            "<o:document-content xmlns:o=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:d=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:p=\"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0\" xmlns:s=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:x=\"urn:foreign\">{prologue}<o:body><o:presentation>{body}</o:presentation></o:body></o:document-content>"
        )
    }

    #[test]
    fn compact_frames_match_original_spans_and_namespace_scopes() {
        let fragments = [
            "",
            "<d:page/>",
            "<d:page d:name=\"one\"><x:unknown>A &amp; B</x:unknown></d:page>",
            "<x:page/><x:foreign/>",
            "<p:settings/>",
            "<p:header-decl/>",
            "<p:footer-decl/>",
            "<p:date-time-decl/>",
            "<d:page xmlns:d=\"urn:foreign\"><d:page/></d:page><d:page/>",
            "<page xmlns=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\"/>",
            "<x:foreign><d:page/><p:settings/></x:foreign>",
            "<!-- before --><d:page><![CDATA[<>]]></d:page><?inert tail?>",
        ];
        for left in fragments {
            for right in fragments {
                let xml = wrap(
                    "<o:automatic-styles><s:style s:name=\"style-one\"/></o:automatic-styles>",
                    &format!("{left}{right}"),
                );
                compare(&xml);
                compare(&format!("\u{feff}<?xml version=\"1.0\"?>{xml}"));
                compare(
                    &xml.replace("o:", "office:")
                        .replace("xmlns:o=", "xmlns:office="),
                );
            }
        }
        for prologue in [
            "",
            "<o:automatic-styles/>",
            "<o:automatic-styles></o:automatic-styles>",
            "<o:automatic-styles><s:style s:name=\"outer\"><s:style xmlns:s=\"urn:foreign\" s:name=\"inner\"/></s:style><s:style s:name=\"sibling\"/></o:automatic-styles>",
            "<o:automatic-styles xmlns:o=\"urn:foreign\"/><o:automatic-styles/>",
            "<automatic-styles xmlns=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"><style xmlns=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" s:name=\"default\"/></automatic-styles>",
        ] {
            compare(&wrap(prologue, "<x:before/><d:page/><x:after/>"));
        }
    }

    #[test]
    fn compact_frames_preserve_original_failures_and_fallbacks() {
        for xml in [
            "",
            "<wrong/>",
            "<x:unknown/>",
            "<o:document-content/>",
            "<",
            "<x></wrong>",
            "</x>",
        ] {
            compare(xml);
        }
        for body in [
            "<d:page>",
            "<d:page></x:page>",
            "<d:page broken=\"unterminated>",
            "<x:foreign><d:page></x:foreign>",
            "<d:page xmlns:d=\"urn:foreign\"/>",
            "<d:page a=\"1\" a=\"2\"/>",
            "<d:page>&unknown;</d:page>",
        ] {
            compare(&wrap("", body));
        }
        for prologue in [
            "<o:automatic-styles><s:style s:name=\"&unknown;\"/></o:automatic-styles>",
            "<o:automatic-styles><s:style s:name=\"one\" s:name=\"two\"/></o:automatic-styles>",
            "<o:automatic-styles><s:style broken=\"unterminated/></o:automatic-styles>",
        ] {
            compare(&wrap(prologue, "<d:page/>"));
        }
        compare(&wrap("", "<d:page/>").replace("xmlns:x=\"urn:foreign\"", "xmlns:x=\"&unknown;\""));
    }

    #[test]
    fn compact_frames_preserve_original_depth_page_and_byte_limits() {
        for depth in [MAX_DEPTH - 3, MAX_DEPTH - 2] {
            let xml = wrap(
                "",
                &format!("{}{}", "<x:n>".repeat(depth), "</x:n>".repeat(depth)),
            );
            compare(&xml);
            assert_eq!(ContentSource::parse(&xml).is_err(), depth == MAX_DEPTH - 2);
        }
        for pages in [MAX_PAGES, MAX_PAGES + 1] {
            let xml = wrap("", &"<d:page/>".repeat(pages));
            compare(&xml);
            assert_eq!(ContentSource::parse(&xml).is_err(), pages > MAX_PAGES);
        }
        compare(&" ".repeat(MAX_XML_BYTES));
        let oversized = " ".repeat(MAX_XML_BYTES + 1);
        compare(&oversized);
        assert!(ContentSource::parse(&oversized).is_err());
    }
}
