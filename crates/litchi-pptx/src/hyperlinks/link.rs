//! Click actions on shapes and text runs, read and written in place
//! (record 0321).
//!
//! [`Click`] is a click action as authored (`a:hlinkClick`): an optional
//! relationship, an action URI and a tooltip. A slide resolves it to a
//! [`Link`] through its relationships: an external URL, a slide of the
//! presentation, or a relationship-free action such as the next-slide jump.
//! The writer side ([`ShapeLinks`]) locates the shape-level click
//! (`p:cNvPr`) and each text run's (`a:rPr`) of one shape fragment so the
//! editor can replace or remove them byte-exactly.

use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use std::ops::Range;

use litchi_core::Position;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use crate::{Error, Result};

const DML: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DML: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
const PML: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
pub(crate) const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(crate) const STRICT_REL: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const MAX_DEPTH: usize = 256;
const MAX_TEXT_BYTES: usize = 1 << 20;

/// A click action as authored (`a:hlinkClick`), not yet resolved.
#[derive(Clone, Default, PartialEq, Eq)]
pub struct Click {
    pub(crate) relationship_id: Option<String>,
    action: Option<String>,
    tooltip: Option<String>,
}

impl std::fmt::Debug for Click {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Click")
            .field("has_relationship", &self.relationship_id.is_some())
            .finish_non_exhaustive()
    }
}

impl Click {
    /// The click action of a text run.
    #[must_use]
    pub fn from_text(link: &litchi_drawingml::text::body::Hyperlink) -> Self {
        Self {
            relationship_id: litchi_drawingml::text::body::raw::hyperlink_relationship_id(link)
                .filter(|id| !id.is_empty())
                .map(str::to_owned),
            action: link.action.clone(),
            tooltip: link.tooltip.clone(),
        }
    }

    /// The shape-level click action of a shape (`p:cNvPr/a:hlinkClick`),
    /// if it has one.
    ///
    /// # Errors
    ///
    /// Returns an error for a malformed shape fragment.
    pub fn of_shape(shape: crate::shape::Shape<'_>) -> Result<Option<Self>> {
        // Most shapes carry no click at all: skip the standalone parse.
        if !mentions_click(shape.xml()?) {
            return Ok(None);
        }
        let fragment = shape.self_contained_xml()?;
        Ok(ShapeLinks::scan(&fragment)?.shape_click)
    }

    /// The action URI (`action`), such as `ppaction://hlinksldjump`.
    #[must_use]
    pub fn action(&self) -> Option<&str> {
        self.action.as_deref()
    }

    /// The tooltip (`tooltip`).
    #[must_use]
    pub fn tooltip(&self) -> Option<&str> {
        self.tooltip.as_deref()
    }
}

/// Where a click action leads.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Target {
    /// An external target (`TargetMode="External"`): a URL, an e-mail
    /// address or a file, verbatim.
    Url(String),
    /// A slide of this presentation, by zero-based position.
    Slide(Position),
    /// A relationship-free action, such as
    /// `ppaction://hlinkshowjump?jump=nextslide`.
    Action(String),
}

/// A resolved click action.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Link {
    /// Where the click leads.
    pub target: Target,
    /// The tooltip.
    pub tooltip: Option<String>,
}

impl Link {
    /// Resolve `click` with its owner's relationship lookup: `relationship`
    /// answers a relationship ID with `(external, target)`, the target
    /// being the URL of an external relationship or the part name of an
    /// internal one; `slide` answers a part name with its slide position.
    ///
    /// # Errors
    ///
    /// Propagates the lookups' errors.
    pub(crate) fn resolve(
        click: &Click,
        relationship: impl FnOnce(&str) -> Result<Option<(bool, String)>>,
        slide: impl FnOnce(&str) -> Result<Option<Position>>,
    ) -> Result<Option<Self>> {
        let tooltip = click.tooltip.clone().filter(|tooltip| !tooltip.is_empty());
        let target = match click.relationship_id.as_deref() {
            None => match click.action.as_deref() {
                Some(action) if !action.is_empty() => Target::Action(action.to_owned()),
                _ => return Ok(None),
            },
            Some(id) => match relationship(id)? {
                None => return Ok(None),
                Some((true, url)) => Target::Url(url),
                Some((false, part)) => match slide(&part)? {
                    Some(position) => Target::Slide(position),
                    None => return Ok(None),
                },
            },
        };
        Ok(Some(Self { target, tooltip }))
    }
}

/// One element of a shape fragment the writer edits.
#[derive(Debug, Clone, Default)]
pub(crate) struct Element {
    /// The start tag (the whole element when it is empty).
    pub(crate) tag: Range<usize>,
    pub(crate) empty: bool,
    /// The start of the end tag, when the element is not empty.
    pub(crate) end: Option<usize>,
    /// The qualified name's prefix (`a`), empty for a default namespace.
    pub(crate) prefix: Vec<u8>,
    /// The local name.
    pub(crate) local: Vec<u8>,
    /// The spans of the element's direct `a:hlinkClick` children.
    pub(crate) clicks: Vec<Range<usize>>,
    /// Where a new `a:hlinkClick` goes: before the first direct child that
    /// must follow it, else before the end tag.
    pub(crate) insert: Option<usize>,
    /// The first direct `a:hlinkClick`'s action, as authored.
    pub(crate) click: Option<Click>,
}

impl Element {
    /// Where a new click goes in a non-empty element.
    pub(crate) fn insertion(&self) -> Option<usize> {
        self.insert.or(self.end)
    }
}

/// A run-like child of a paragraph (`a:r`, `a:br`, `a:fld`), as the text
/// reader counts them.
#[derive(Debug, Clone)]
pub(crate) struct RunSlot {
    /// Whether it is a text run (`a:r`).
    pub(crate) text: bool,
    /// The end of the run's start tag (where a new `a:rPr` goes).
    pub(crate) open_end: usize,
    /// The run's prefix.
    pub(crate) prefix: Vec<u8>,
    /// Its character properties.
    pub(crate) properties: Option<Element>,
}

/// The click-action slots of one shape fragment.
#[derive(Debug, Clone, Default)]
pub(crate) struct ShapeLinks {
    /// The shape's `p:cNvPr`.
    pub(crate) non_visual: Option<Element>,
    /// The shape-level click action.
    pub(crate) shape_click: Option<Click>,
    /// The run-like children of each paragraph of the shape's text body.
    pub(crate) paragraphs: Vec<Vec<RunSlot>>,
    /// The DrawingML prefix the fragment binds, if any element used one.
    pub(crate) drawing_prefix: Option<Vec<u8>>,
    /// Whether the fragment uses Strict namespaces.
    pub(crate) strict: bool,
    /// The end of the fragment root's start tag.
    pub(crate) root_open_end: usize,
}

fn is(namespace: &ResolveResult<'_>, wanted: [&[u8]; 2]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if wanted.contains(value))
}

fn invalid(message: &str) -> Error {
    Error::Invalid(format!("shape hyperlink: {message}"))
}

/// The children of `a:rPr` that follow `a:hlinkClick` (CT_TextCharacterProperties).
const AFTER_RUN_CLICK: [&[u8]; 3] = [b"hlinkMouseOver", b"rtl", b"extLst"];
/// The children of `p:cNvPr` that follow `a:hlinkClick` (CT_NonVisualDrawingProps).
const AFTER_SHAPE_CLICK: [&[u8]; 2] = [b"hlinkHover", b"extLst"];

fn attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    local: &[u8],
    relationship: bool,
) -> Result<Option<String>> {
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let (namespace, name) = reader.resolver().resolve_attribute(attribute.key);
        let wanted = if relationship {
            matches!(namespace, ResolveResult::Bound(Namespace(value))
                if value == REL.as_bytes() || value == STRICT_REL.as_bytes())
        } else {
            matches!(namespace, ResolveResult::Unbound)
        };
        if wanted && name.as_ref() == local {
            let value = attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| Error::Xml(error.to_string()))?;
            if value.len() > MAX_TEXT_BYTES {
                return Err(Error::Limit {
                    kind: crate::opened::LimitKind::Codec(
                        crate::opened::CodecLimitKind::XmlAttributeBytes,
                    ),
                    observed: value.len(),
                    maximum: MAX_TEXT_BYTES,
                });
            }
            return Ok(Some(value.into_owned()));
        }
    }
    Ok(None)
}

impl ShapeLinks {
    /// Scan a standalone shape fragment (its root is the shape element).
    pub(crate) fn scan(fragment: &[u8]) -> Result<Self> {
        let origin = litchi_core::xml::ReaderOrigin::of(fragment);
        let mut reader = NsReader::from_reader(fragment);
        reader.config_mut().check_end_names = true;
        let mut out = Self::default();
        // (local name, is DrawingML, is PresentationML, element index into
        // `open` when it is one the writer tracks).
        let mut stack: Vec<(Vec<u8>, bool, bool, Option<Slot>)> = Vec::new();
        let mut in_body = false;
        loop {
            let before = origin
                .offset(reader.buffer_position())
                .ok_or_else(|| invalid("offset overflow"))?;
            let (namespace, event) = reader
                .read_resolved_event()
                .map_err(|error| Error::Xml(error.to_string()))?;
            let dml = is(&namespace, [DML, STRICT_DML]);
            let pml = is(&namespace, [PML, STRICT_PML]);
            let strict = is(&namespace, [STRICT_DML, STRICT_PML]);
            let after = origin
                .offset(reader.buffer_position())
                .ok_or_else(|| invalid("offset overflow"))?;
            match event {
                Event::Start(ref element) | Event::Empty(ref element) => {
                    let empty = matches!(event, Event::Empty(_));
                    let local = element.local_name().as_ref().to_vec();
                    let prefix = element
                        .name()
                        .prefix()
                        .map(|prefix| prefix.into_inner().to_vec())
                        .unwrap_or_default();
                    if stack.is_empty() {
                        out.root_open_end = after;
                        out.strict = strict;
                    }
                    if dml && out.drawing_prefix.is_none() {
                        out.drawing_prefix = Some(prefix.clone());
                    }
                    let depth = stack.len();
                    let parent = stack
                        .last()
                        .map(|(name, dml, pml, slot)| (name.clone(), *dml, *pml, *slot));
                    let mut slot = None;
                    // A direct child that must follow a new click.
                    if let Some((_, _, _, Some(parent_slot))) = &parent {
                        let target = out.element_mut(*parent_slot);
                        let follows = match parent_slot {
                            Slot::NonVisual => AFTER_SHAPE_CLICK.contains(&local.as_slice()),
                            Slot::Properties(..) => AFTER_RUN_CLICK.contains(&local.as_slice()),
                        };
                        let click = if dml && local == b"hlinkClick" {
                            Some(Click {
                                relationship_id: attribute(&reader, element, b"id", true)?
                                    .filter(|id| !id.is_empty()),
                                action: attribute(&reader, element, b"action", false)?,
                                tooltip: attribute(&reader, element, b"tooltip", false)?,
                            })
                        } else {
                            None
                        };
                        if let Some(target) = target {
                            if let Some(click) = &click {
                                // Spans are completed at the end event.
                                target.clicks.push(before..after);
                                if target.click.is_none() {
                                    target.click = Some(click.clone());
                                }
                            } else if follows && target.insert.is_none() {
                                target.insert = Some(before);
                            }
                        }
                        if *parent_slot == Slot::NonVisual
                            && out.shape_click.is_none()
                            && let Some(click) = click
                        {
                            out.shape_click = Some(click);
                        }
                    }
                    if pml && local == b"cNvPr" && out.non_visual.is_none() && depth == 2 {
                        out.non_visual = Some(Element {
                            tag: before..after,
                            empty,
                            prefix: prefix.clone(),
                            local: local.clone(),
                            ..Element::default()
                        });
                        slot = Some(Slot::NonVisual);
                    }
                    if pml && local == b"txBody" && depth == 1 {
                        in_body = true;
                    }
                    if in_body && dml {
                        match (local.as_slice(), &parent) {
                            (b"p", Some((name, _, true, _))) if name.as_slice() == b"txBody" => {
                                out.paragraphs.push(Vec::new());
                            },
                            (b"r" | b"br" | b"fld", Some((name, true, _, _)))
                                if name.as_slice() == b"p" && depth == 3 =>
                            {
                                if let Some(runs) = out.paragraphs.last_mut() {
                                    runs.push(RunSlot {
                                        text: local == b"r",
                                        open_end: after,
                                        prefix: prefix.clone(),
                                        properties: None,
                                    });
                                }
                            },
                            (b"rPr", Some((name, true, _, _)))
                                if depth == 4
                                    && matches!(name.as_slice(), b"r" | b"br" | b"fld") =>
                            {
                                let paragraph = out.paragraphs.len().saturating_sub(1);
                                let run = out
                                    .paragraphs
                                    .last()
                                    .map_or(0, |runs| runs.len().saturating_sub(1));
                                if let Some(slot_run) =
                                    out.paragraphs.last_mut().and_then(|runs| runs.last_mut())
                                {
                                    slot_run.properties = Some(Element {
                                        tag: before..after,
                                        empty,
                                        prefix: prefix.clone(),
                                        local: local.clone(),
                                        ..Element::default()
                                    });
                                    slot = Some(Slot::Properties(paragraph, run));
                                }
                            },
                            _ => {},
                        }
                    }
                    if !empty {
                        if stack.len() >= MAX_DEPTH {
                            return Err(invalid("the shape is nested too deeply"));
                        }
                        stack.push((local, dml, pml, slot));
                    }
                },
                Event::End(_) => {
                    let Some((local, dml, pml, slot)) = stack.pop() else {
                        return Err(invalid("unbalanced shape fragment"));
                    };
                    if let Some(slot) = slot
                        && let Some(element) = out.element_mut(slot)
                    {
                        element.end = Some(before);
                    }
                    if dml
                        && local == b"hlinkClick"
                        && let Some((_, _, _, Some(parent))) = stack.last()
                        && let Some(element) = out.element_mut(*parent)
                        && let Some(click) = element.clicks.last_mut()
                    {
                        click.end = after;
                    }
                    if pml && local == b"txBody" && stack.len() == 1 {
                        in_body = false;
                    }
                    if stack.is_empty() {
                        break;
                    }
                },
                Event::Eof => break,
                _ => {},
            }
        }
        Ok(out)
    }

    fn element_mut(&mut self, slot: Slot) -> Option<&mut Element> {
        match slot {
            Slot::NonVisual => self.non_visual.as_mut(),
            Slot::Properties(paragraph, run) => self
                .paragraphs
                .get_mut(paragraph)?
                .get_mut(run)?
                .properties
                .as_mut(),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Slot {
    NonVisual,
    Properties(usize, usize),
}

/// Whether shape markup names a click action anywhere (`hlinkClick`); a
/// shape without one has no click to read.
#[must_use]
pub fn mentions_click(xml: &[u8]) -> bool {
    const NEEDLE: &[u8] = b"hlinkClick";
    xml.windows(NEEDLE.len()).any(|window| window == NEEDLE)
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    const SHAPE: &str = concat!(
        r#"<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" "#,
        r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
        r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
        r#"<p:nvSpPr><p:cNvPr id="2" name="Box"><a:hlinkClick r:id="rId4" tooltip="Go"/></p:cNvPr>"#,
        r#"<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/>"#,
        r#"<p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"><a:latin typeface="X"/><a:rtl val="0"/></a:rPr><a:t>a</a:t></a:r>"#,
        r#"<a:br/><a:r><a:t>b</a:t></a:r></a:p></p:txBody></p:sp>"#
    );

    #[test]
    fn scan_finds_click_slots_in_schema_order() {
        let links = ShapeLinks::scan(SHAPE.as_bytes()).unwrap();
        let click = links.shape_click.clone().unwrap();
        assert_eq!(click.relationship_id.as_deref(), Some("rId4"));
        assert_eq!(click.tooltip(), Some("Go"));
        let non_visual = links.non_visual.clone().unwrap();
        assert_eq!(non_visual.clicks.len(), 1);
        assert_eq!(
            &SHAPE[non_visual.clicks[0].clone()],
            r#"<a:hlinkClick r:id="rId4" tooltip="Go"/>"#
        );
        assert_eq!(links.paragraphs.len(), 1);
        let runs = &links.paragraphs[0];
        assert_eq!(runs.len(), 3);
        assert!(runs[0].text && !runs[1].text && runs[2].text);
        let properties = runs[0].properties.clone().unwrap();
        assert_eq!(
            &SHAPE[properties.insertion().unwrap()..properties.insertion().unwrap() + 6],
            "<a:rtl"
        );
        assert!(runs[2].properties.is_none());
        assert_eq!(links.drawing_prefix.as_deref(), Some(&b"a"[..]));
    }

    #[test]
    fn resolve_maps_relationships_and_actions() {
        let click = Click {
            relationship_id: Some("rId1".into()),
            action: Some("ppaction://hlinksldjump".into()),
            tooltip: None,
        };
        let link = Link::resolve(
            &click,
            |_| Ok(Some((false, "/ppt/slides/slide2.xml".into()))),
            |_| Ok(Some(Position::new(1))),
        )
        .unwrap()
        .unwrap();
        assert_eq!(link.target, Target::Slide(Position::new(1)));
        let action = Click {
            relationship_id: None,
            action: Some("ppaction://hlinkshowjump?jump=nextslide".into()),
            tooltip: Some("Next".into()),
        };
        let link = Link::resolve(&action, |_| unreachable!(), |_| unreachable!())
            .unwrap()
            .unwrap();
        assert_eq!(
            link.target,
            Target::Action("ppaction://hlinkshowjump?jump=nextslide".into())
        );
        assert_eq!(link.tooltip.as_deref(), Some("Next"));
    }
}
