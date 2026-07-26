//! Single-pass scan that records the byte spans a presentation save reuses.
//!
//! The scanner walks `content.xml` once and remembers where the root element,
//! the prologue, the automatic-style container, and each `draw:page` begin and
//! end, so [`PresentationContentSource`](super::PresentationContentSource) can
//! slice them back out verbatim.

use super::{
    AUTOMATIC_STYLES_ELEMENT, AutomaticStylesSite, BODY_CHILD_DEPTH, BODY_ELEMENT,
    DECLARATION_ELEMENTS, DOCUMENT_CONTENT_ELEMENT, DRAW_NAMESPACE, MAX_DEPTH, MAX_PAGES,
    MAX_STYLE_NAMES, OFFICE_NAMESPACE, PAGE_ELEMENT, PRESENTATION_CHILD_DEPTH,
    PRESENTATION_ELEMENT, PRESENTATION_NAMESPACE, PROLOGUE_DEPTH, PresentationContentSource,
    SETTINGS_ELEMENT, STYLE_ELEMENT, STYLE_NAME_ATTRIBUTE, STYLE_NAMESPACE, XMLNS_DEFAULT,
    XMLNS_PREFIX, XmlSpan, invalid,
};
use litchi_core::Result;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeMap, BTreeSet};

/// One open element on the scanner stack.
struct Frame {
    namespace: Option<Vec<u8>>,
    local: Vec<u8>,
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
            let (namespace, event) = {
                let (resolved, event) = self.reader.read_resolved_event().map_err(|error| {
                    invalid(format!("invalid presentation content.xml: {error}"))
                })?;
                (namespace_of(&resolved), event)
            };
            let end = self.reader.buffer_position() as usize;
            match event {
                Event::Start(element) => {
                    let start = event_start(self.xml, end)?;
                    let local = element.local_name().as_ref().to_vec();
                    if self.stack.is_empty() {
                        self.collect_root_namespaces(&element)?;
                    }
                    self.on_open(&namespace, &local, start, end, false)?;
                    if self.styles_depth.is_some()
                        && is(&namespace, &local, STYLE_NAMESPACE, STYLE_ELEMENT)
                    {
                        self.record_style_name(&element)?;
                    }
                    if self.stack.len() >= MAX_DEPTH {
                        return Err(invalid(format!(
                            "presentation content.xml nests deeper than {MAX_DEPTH} elements"
                        )));
                    }
                    self.stack.push(Frame {
                        namespace,
                        local,
                        start,
                    });
                },
                Event::Empty(element) => {
                    let start = event_start(self.xml, end)?;
                    let local = element.local_name().as_ref().to_vec();
                    if self.styles_depth.is_some()
                        && is(&namespace, &local, STYLE_NAMESPACE, STYLE_ELEMENT)
                    {
                        self.record_style_name(&element)?;
                    }
                    self.on_open(&namespace, &local, start, end, true)?;
                    self.on_close(&namespace, &local, start, start, end);
                },
                Event::End(_) => {
                    let close_start = event_start(self.xml, end)?;
                    let frame = self
                        .stack
                        .pop()
                        .ok_or_else(|| invalid("presentation content.xml depth underflow"))?;
                    self.on_close(
                        &frame.namespace,
                        &frame.local,
                        frame.start,
                        close_start,
                        end,
                    );
                },
                Event::Eof => break,
                _ => {},
            }
        }
        if !self.stack.is_empty() {
            return Err(invalid("unterminated presentation content.xml element"));
        }
        Ok(())
    }

    /// Handle an element start, or the start half of an empty element.
    fn on_open(
        &mut self,
        namespace: &Option<Vec<u8>>,
        local: &[u8],
        start: usize,
        end: usize,
        empty: bool,
    ) -> Result<()> {
        let depth = self.stack.len();
        if depth == 0 {
            if empty || !is(namespace, local, OFFICE_NAMESPACE, DOCUMENT_CONTENT_ELEMENT) {
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
        if depth == PROLOGUE_DEPTH
            && is(namespace, local, OFFICE_NAMESPACE, AUTOMATIC_STYLES_ELEMENT)
        {
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
        if depth == PROLOGUE_DEPTH && !empty && is(namespace, local, OFFICE_NAMESPACE, BODY_ELEMENT)
        {
            self.body_open = Some(XmlSpan { start, end });
            return Ok(());
        }
        if depth == BODY_CHILD_DEPTH
            && !empty
            && is(namespace, local, OFFICE_NAMESPACE, PRESENTATION_ELEMENT)
        {
            self.presentation_open = Some(XmlSpan { start, end });
            return Ok(());
        }
        if depth == PRESENTATION_CHILD_DEPTH && self.presentation_open.is_some() {
            if is(namespace, local, DRAW_NAMESPACE, PAGE_ELEMENT) {
                if self.pages.len() >= MAX_PAGES {
                    return Err(invalid(format!(
                        "presentation content.xml holds more than {MAX_PAGES} slides"
                    )));
                }
                self.open_page = Some(start);
            } else if !is_regenerated_child(namespace, local) {
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
    fn on_close(
        &mut self,
        namespace: &Option<Vec<u8>>,
        local: &[u8],
        open_start: usize,
        close_start: usize,
        end: usize,
    ) {
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
        if self.styles_depth == Some(self.stack.len())
            && is(namespace, local, OFFICE_NAMESPACE, AUTOMATIC_STYLES_ELEMENT)
        {
            self.styles_depth = None;
            self.styles_site = Some((close_start, AutomaticStylesSite::Content));
        }
    }

    /// Record every namespace binding declared on the root element.
    fn collect_root_namespaces(&mut self, element: &BytesStart<'_>) -> Result<()> {
        for attribute in element.attributes() {
            let attribute = attribute.map_err(|error| {
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
        for attribute in element.attributes() {
            let attribute =
                attribute.map_err(|error| invalid(format!("invalid style attribute: {error}")))?;
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
    pub(super) fn finish(
        self,
        xml: &str,
        has_byte_order_mark: bool,
    ) -> Option<PresentationContentSource> {
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
            _ => splice,
        };
        Some(PresentationContentSource {
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

/// Whether a presentation child is rebuilt from the typed model on save.
fn is_regenerated_child(namespace: &Option<Vec<u8>>, local: &[u8]) -> bool {
    if is(namespace, local, PRESENTATION_NAMESPACE, SETTINGS_ELEMENT) {
        return true;
    }
    DECLARATION_ELEMENTS
        .iter()
        .any(|element| is(namespace, local, PRESENTATION_NAMESPACE, element))
}

/// Locate the `<` that opened the event ending at `end`.
fn event_start(xml: &str, end: usize) -> Result<usize> {
    xml.get(..end)
        .and_then(|prefix| prefix.rfind('<'))
        .ok_or_else(|| invalid("invalid presentation content.xml event boundary"))
}

fn namespace_of(resolved: &ResolveResult<'_>) -> Option<Vec<u8>> {
    match resolved {
        ResolveResult::Bound(Namespace(uri)) => Some(uri.to_vec()),
        _ => None,
    }
}

fn is(
    namespace: &Option<Vec<u8>>,
    local: &[u8],
    expected_uri: &[u8],
    expected_local: &str,
) -> bool {
    namespace.as_deref() == Some(expected_uri) && local == expected_local.as_bytes()
}
