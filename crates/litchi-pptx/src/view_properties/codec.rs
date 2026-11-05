//! Bounded `PresentationML` view-properties XML codec.
//!
//! This module owns the package-independent view-properties model, bounded XML
//! codec, and inert presentation relationship loading.

use super::model::{
    CommonSlideView, CommonView, GridSpacing, Guide, GuideOrientation, NormalView, OutlineSlide,
    OutlineView, Point, Ratio, RestoredPane, SimpleView, SlideLikeView, SorterView, SplitterState,
    ViewKind, ViewProperties,
};
use crate::leniency::{Defect, MarkupDefect, ToleranceReport, ViewDefect};
use crate::{Error, Result};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use litchi_opc::SourceLimits;
use quick_xml::{
    Reader, XmlVersion,
    encoding::Decoder,
    events::{BytesStart, Event},
};
use std::collections::TryReserveError;

const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const PS: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const AS: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";

/// A two-pass XML sink.  The first pass counts exact bytes without allocating;
/// the second reserves that exact capacity and appends through the same
/// bounded path.  A failed append is retained until the writer boundary so
/// every serializer helper remains small and cannot bypass the output limit.
struct XmlSink {
    value: Option<String>,
    len: usize,
    limit: usize,
    failure: Option<SinkFailure>,
}

enum SinkFailure {
    Limit(usize),
    Overflow,
    Allocation(TryReserveError),
}

impl XmlSink {
    fn count(limit: usize) -> Self {
        Self {
            value: None,
            len: 0,
            limit,
            failure: None,
        }
    }

    fn output(capacity: usize, limit: usize) -> Result<Self> {
        let mut value = String::new();
        value
            .try_reserve_exact(capacity)
            .map_err(|source| Error::Allocation {
                resource: "view-properties XML output",
                source,
            })?;
        Ok(Self {
            value: Some(value),
            len: 0,
            limit,
            failure: None,
        })
    }

    fn push(&mut self, value: char) {
        let mut encoded = [0; 4];
        self.push_str(value.encode_utf8(&mut encoded));
    }

    fn push_str(&mut self, value: &str) {
        if self.failure.is_some() {
            return;
        }
        let Some(new_len) = self.len.checked_add(value.len()) else {
            self.failure = Some(SinkFailure::Overflow);
            return;
        };
        if new_len > self.limit {
            self.failure = Some(SinkFailure::Limit(new_len));
            return;
        }
        if let Some(out) = self.value.as_mut() {
            if let Err(source) = out.try_reserve(value.len()) {
                self.failure = Some(SinkFailure::Allocation(source));
                return;
            }
            out.push_str(value);
        }
        self.len = new_len;
    }

    #[must_use]
    fn len(&self) -> usize {
        self.len
    }

    fn check(&mut self) -> Result<()> {
        match self.failure.take() {
            None => Ok(()),
            Some(SinkFailure::Limit(observed)) => Err(limit(
                crate::error::CodecLimitKind::OutputBytes,
                observed,
                self.limit,
            )),
            Some(SinkFailure::Overflow) => Err(Error::CapacityOverflow {
                resource: "view-properties XML output",
            }),
            Some(SinkFailure::Allocation(source)) => Err(Error::Allocation {
                resource: "view-properties XML output",
                source,
            }),
        }
    }

    fn finish(mut self) -> Result<Vec<u8>> {
        self.check()?;
        let value = self.value.take().expect("output sink always owns a string");
        Ok(value.into_bytes())
    }
}
/// Caller-selected bounds for one view-properties XML operation.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[must_use]
pub struct Limits {
    max_input_bytes: usize,
    max_output_bytes: usize,
    max_depth: usize,
    max_nodes: usize,
    max_guides: usize,
    max_slides: usize,
    max_string_bytes: usize,
    max_attribute_bytes: usize,
}

impl Limits {
    /// The historical view-properties safety profile.
    pub const DEFAULT: Self = Self {
        max_input_bytes: 8 * 1024 * 1024,
        max_output_bytes: 8 * 1024 * 1024,
        max_depth: 128,
        max_nodes: 100_000,
        max_guides: 4096,
        max_slides: 100_000,
        max_string_bytes: 1024 * 1024,
        max_attribute_bytes: 1024 * 1024,
    };

    /// Derive the view-properties policy from the retained OPC source profile.
    #[must_use = "the derived view-properties bounds must be retained for the operation"]
    pub fn from_source_limits(source: SourceLimits) -> Self {
        let default = Self::DEFAULT;
        let part_bytes = usize::try_from(source.max_part_bytes()).unwrap_or(usize::MAX);
        let events = source.max_events();
        let depth = source.max_depth();
        let strings = source.max_xml_attribute_bytes();
        Self {
            max_input_bytes: default.max_input_bytes.min(part_bytes),
            max_output_bytes: default.max_output_bytes.min(part_bytes),
            max_depth: default.max_depth.min(depth),
            max_nodes: default.max_nodes.min(events),
            max_guides: default.max_guides.min(events),
            max_slides: default.max_slides.min(events),
            max_string_bytes: default.max_string_bytes.min(strings),
            max_attribute_bytes: default.max_attribute_bytes.min(strings),
        }
    }

    /// Construct a tighter policy without widening the historical ceilings.
    #[must_use]
    pub const fn new(
        max_input_bytes: usize,
        max_output_bytes: usize,
        max_depth: usize,
        max_nodes: usize,
        max_guides: usize,
        max_slides: usize,
        max_string_bytes: usize,
    ) -> Option<Self> {
        let default = Self::DEFAULT;
        if max_input_bytes == 0
            || max_input_bytes > default.max_input_bytes
            || max_output_bytes == 0
            || max_output_bytes > default.max_output_bytes
            || max_depth == 0
            || max_depth > default.max_depth
            || max_nodes == 0
            || max_nodes > default.max_nodes
            || max_guides == 0
            || max_guides > default.max_guides
            || max_slides == 0
            || max_slides > default.max_slides
            || max_string_bytes == 0
            || max_string_bytes > default.max_string_bytes
        {
            return None;
        }
        Some(Self {
            max_input_bytes,
            max_output_bytes,
            max_depth,
            max_nodes,
            max_guides,
            max_slides,
            max_string_bytes,
            max_attribute_bytes: max_string_bytes,
        })
    }

    #[inline]
    #[must_use]
    pub const fn max_input_bytes(self) -> usize {
        self.max_input_bytes
    }

    #[inline]
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    #[inline]
    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }

    #[inline]
    #[must_use]
    pub const fn max_nodes(self) -> usize {
        self.max_nodes
    }

    #[inline]
    #[must_use]
    pub const fn max_guides(self) -> usize {
        self.max_guides
    }

    #[inline]
    #[must_use]
    pub const fn max_slides(self) -> usize {
        self.max_slides
    }

    #[inline]
    #[must_use]
    pub const fn max_string_bytes(self) -> usize {
        self.max_string_bytes
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

impl ViewProperties {
    /// Parse a view-properties part (`p:viewPr`), applying the repairs of
    /// audit finding 6.11 ([`ViewDefect`]): unknown attributes and child
    /// elements are ignored, children are accepted in any order, and a
    /// repeated child keeps its first occurrence. Use
    /// [`Self::parse_with_tolerance`] to collect the repairs, or
    /// [`Self::parse_strict`] to refuse markup that needs one.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(xml: &[u8], limits: Limits) -> Result<Self> {
        Self::parse_with_tolerance(xml, &mut ToleranceReport::default(), limits)
    }

    /// [`Self::parse`], appending each repair to `tolerance` (no part is
    /// stamped; [`super::load_from_package_with_tolerance`] stamps it).
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_with_tolerance(
        xml: &[u8],
        tolerance: &mut ToleranceReport,
        limits: Limits,
    ) -> Result<Self> {
        Self::parse_inner(
            xml,
            &mut Tol {
                report: Some(tolerance),
                limits,
            },
        )
    }

    /// Parse without any repair: an attribute or child the model does not
    /// know, children out of schema order, or a repeated child is an error.
    /// Edit planners that must account for every byte use this contract.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read, is malformed, or would
    /// need a [`ViewDefect`] repair.
    pub fn parse_strict(xml: &[u8], limits: Limits) -> Result<Self> {
        Self::parse_inner(
            xml,
            &mut Tol {
                report: None,
                limits,
            },
        )
    }

    fn parse_inner(xml: &[u8], tol: &mut Tol<'_>) -> Result<Self> {
        let limits = tol.limits;
        if xml.len() > limits.max_input_bytes() {
            return Err(limit(
                crate::error::CodecLimitKind::InputBytes,
                xml.len(),
                limits.max_input_bytes(),
            ));
        }
        // Processing instructions are inert (audit finding 6.3): skipped
        // from the semantic input, recorded once for the part.
        let (xml, instructions) = crate::markup::without_processing_instructions(xml)?;
        if instructions {
            tol.repair(
                MarkupDefect::ProcessingInstruction,
                "processing instructions are rejected",
            )?;
        }
        let mce_limits = litchi_ooxml_common::mce::Limits {
            max_input_bytes: limits.max_input_bytes(),
            max_output_bytes: limits.max_output_bytes(),
            max_depth: limits.max_depth(),
            max_namespace_bindings: 4096.min(limits.max_nodes()),
            max_directive_tokens: 4096.min(limits.max_nodes()),
            max_choices_per_alternate: 1024.min(limits.max_nodes()),
            max_attributes_per_element:
                litchi_ooxml_common::mce::DEFAULT_MAX_ATTRIBUTES_PER_ELEMENT,
        };
        let capabilities = litchi_ooxml_common::mce::Capabilities::ooxml_baseline();
        let x = litchi_ooxml_common::mce::process_markup_compatibility(
            &xml,
            &capabilities,
            &mce_limits,
        )?
        .xml;
        if x.len() > limits.max_output_bytes() {
            return Err(limit(
                crate::error::CodecLimitKind::OutputBytes,
                x.len(),
                limits.max_output_bytes(),
            ));
        }
        project(&parse_dom_processed(x.as_ref(), limits)?, tol)
    }
    /// # Errors
    ///
    /// Returns an error if the output cannot be encoded or written.
    pub fn to_xml(&self, strict: bool, limits: Limits) -> Result<Vec<u8>> {
        validate(self, limits)?;
        let mut count = XmlSink::count(limits.max_output_bytes());
        write_document(self, strict, limits, &mut count)?;
        count.check()?;
        let size = count.len();
        let mut output = XmlSink::output(size, limits.max_output_bytes())?;
        write_document(self, strict, limits, &mut output)?;
        if output.len() != size {
            return Err(Error::CapacityOverflow {
                resource: "view-properties XML output changed between passes",
            });
        }
        output.finish()
    }
}

fn write_document(
    value: &ViewProperties,
    strict: bool,
    limits: Limits,
    x: &mut XmlSink,
) -> Result<()> {
    let p = if strict { PS } else { P };
    let a = if strict { AS } else { A };
    let r = if strict { RS } else { R };
    x.push_str(&format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:viewPr xmlns:p="{p}" xmlns:a="{a}" xmlns:r="{r}""#
    ));
    if let Some(v) = value.last_view {
        attr(x, "lastView", view_str(v));
    }
    bool_attr(x, "showComments", value.show_comments);
    x.push('>');
    if let Some(v) = &value.normal {
        write_normal(x, v, strict, limits)?;
    }
    if let Some(v) = &value.slide {
        write_slide_like(x, "slideViewPr", v, strict, limits)?;
    }
    if let Some(v) = &value.outline {
        write_outline(x, v, strict, limits)?;
    }
    if let Some(v) = &value.notes_text {
        write_simple(x, "notesTextViewPr", v, strict, limits)?;
    }
    if let Some(v) = &value.sorter {
        write_sorter(x, v, strict, limits)?;
    }
    if let Some(v) = &value.notes {
        write_slide_like(x, "notesViewPr", v, strict, limits)?;
    }
    if let Some(v) = &value.grid_spacing {
        x.push_str(&format!("<p:gridSpacing cx=\"{}\" cy=\"{}\"/>", v.cx, v.cy));
    }
    if let Some(v) = &value.extension_xml {
        opaque(x, v, strict, limits)?;
    }
    x.push_str("</p:viewPr>");
    x.check()
}

#[derive(Clone)]
struct Attr {
    q: String,
    ns: String,
    l: String,
    v: String,
}
#[derive(Clone)]
enum Content {
    Node(Node),
    Text(String),
    CData(String),
    Comment(String),
}
#[derive(Clone)]
struct Node {
    q: String,
    ns: String,
    l: String,
    attrs: Vec<Attr>,
    bindings: Vec<(String, String)>,
    content: Vec<Content>,
    source_start: usize,
    capture_raw: bool,
    raw: Option<Box<[u8]>>,
}
/// Parse markup emitted by the bounded MCE processor.  Its byte budget has
/// already been admitted against `max_output_bytes`; only structural limits
/// are applied here because valid namespace-context repair can expand the
/// processed stream beyond the original ingress budget.
fn parse_dom_processed(xml: &[u8], limits: Limits) -> Result<Node> {
    if xml.len() > limits.max_output_bytes() {
        return Err(limit(
            crate::error::CodecLimitKind::OutputBytes,
            xml.len(),
            limits.max_output_bytes(),
        ));
    }
    parse_dom_inner(xml, limits)
}

fn parse_dom_inner(xml: &[u8], limits: Limits) -> Result<Node> {
    std::str::from_utf8(xml).map_err(xml_error)?;
    let mut rd = Reader::from_reader(xml);
    let mut stack: Vec<Node> = Vec::new();
    let mut root = None;
    let mut count = 0usize;
    loop {
        let before = usize::try_from(rd.buffer_position())
            .map_err(|_| invalid("view-properties source offset overflows usize"))?;
        let d = rd.decoder();
        match rd.read_event() {
            Ok(Event::Start(e)) => {
                count = count.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::XmlNodes,
                        usize::MAX,
                        limits.max_nodes(),
                    )
                })?;
                if count > limits.max_nodes() {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlNodes,
                        count,
                        limits.max_nodes(),
                    ));
                }
                if stack.len() >= limits.max_depth() {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        stack.len().checked_add(1).unwrap_or(usize::MAX),
                        limits.max_depth(),
                    ));
                }
                let mut node = make(&e, d, &stack, limits)?;
                node.source_start = before;
                stack.push(node);
            },
            Ok(Event::Empty(e)) => {
                count = count.checked_add(1).ok_or_else(|| {
                    limit(
                        crate::error::CodecLimitKind::XmlNodes,
                        usize::MAX,
                        limits.max_nodes(),
                    )
                })?;
                if count > limits.max_nodes() {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlNodes,
                        count,
                        limits.max_nodes(),
                    ));
                }
                if stack.len() >= limits.max_depth() {
                    return Err(limit(
                        crate::error::CodecLimitKind::XmlDepth,
                        stack.len().checked_add(1).unwrap_or(usize::MAX),
                        limits.max_depth(),
                    ));
                }
                let mut n = make(&e, d, &stack, limits)?;
                if n.capture_raw {
                    let after = usize::try_from(rd.buffer_position())
                        .map_err(|_| invalid("view-properties source offset overflows usize"))?;
                    n.raw = Some(
                        xml.get(before..after)
                            .ok_or_else(|| invalid("view-properties source span is invalid"))?
                            .into(),
                    );
                }
                attach(&mut stack, &mut root, n)?;
            },
            Ok(Event::End(_)) => {
                let mut n = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected closing element"))?;
                if n.capture_raw {
                    let after = usize::try_from(rd.buffer_position())
                        .map_err(|_| invalid("view-properties source offset overflows usize"))?;
                    n.raw = Some(
                        xml.get(n.source_start..after)
                            .ok_or_else(|| invalid("view-properties source span is invalid"))?
                            .into(),
                    );
                }
                attach(&mut stack, &mut root, n)?;
            },
            Ok(Event::Text(t)) => {
                let v = t.decode().map_err(xml_error)?.into_owned();
                bounded_string(&v, limits)?;
                if let Some(n) = stack.last_mut() {
                    n.content.push(Content::Text(v));
                } else if !v.trim().is_empty() {
                    return Err(invalid("text outside viewPr"));
                }
            },
            Ok(Event::CData(t)) => {
                let v = t.decode().map_err(xml_error)?.into_owned();
                bounded_string(&v, limits)?;
                if let Some(n) = stack.last_mut() {
                    n.content.push(Content::CData(v));
                } else {
                    return Err(invalid("CDATA outside viewPr"));
                }
            },
            Ok(Event::Comment(t)) => {
                let v = t.decode().map_err(xml_error)?.into_owned();
                bounded_string(&v, limits)?;
                if let Some(n) = stack.last_mut() {
                    n.content.push(Content::Comment(v));
                }
            },
            Ok(Event::GeneralRef(t)) => {
                let v = litchi_ooxml_common::xml::decode_xml_reference(&t)?;
                bounded_string(&v, limits)?;
                if let Some(n) = stack.last_mut() {
                    n.content.push(Content::Text(v));
                } else {
                    return Err(invalid("entity outside viewPr"));
                }
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid("DTD and processing instructions are rejected"));
            },
            Ok(Event::Decl(_)) => {},
            Ok(Event::Eof) => break,
            Err(e) => return Err(xml_error(e)),
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated view-properties XML"));
    }
    root.ok_or_else(|| invalid("missing viewPr root"))
}
fn make(e: &BytesStart<'_>, d: Decoder, stack: &[Node], limits: Limits) -> Result<Node> {
    let q = std::str::from_utf8(e.name().as_ref())
        .map_err(xml_error)?
        .to_string();
    let mut bindings = stack.last().map(|x| x.bindings.clone()).unwrap_or_default();
    let mut raw = Vec::new();
    for a in e.checked_attributes() {
        let a = a.map_err(xml_error)?;
        let value = a
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, d)
            .map_err(xml_error)?
            .into_owned();
        bounded_attribute(&value, limits)?;
        raw.push((
            std::str::from_utf8(a.key.as_ref())
                .map_err(xml_error)?
                .to_string(),
            value,
        ));
    }
    for (k, v) in &raw {
        if k == "xmlns" || k.starts_with("xmlns:") {
            let key = k.strip_prefix("xmlns:").unwrap_or("").to_string();
            if let Some(x) = bindings.iter_mut().find(|x| x.0 == key) {
                x.1.clone_from(v);
            } else {
                bindings.push((key, v.clone()));
            }
        }
    }
    let (pr, lo) = split(&q)?;
    let local = lo.to_string();
    let capture_raw = matches!(
        local.as_str(),
        "ext" | "extLst" | "clrMru" | "penClr" | "laserClr"
    ) || stack
        .last()
        .is_some_and(|parent| matches!(parent.l.as_str(), "clrMru" | "penClr" | "laserClr"));
    let ns = resolve(&bindings, pr)?;
    let mut attrs = Vec::new();
    for (q, v) in raw {
        if q == "xmlns" || q.starts_with("xmlns:") {
            continue;
        }
        let (pr, lo) = split(&q)?;
        let ans = if pr.is_empty() {
            String::new()
        } else {
            resolve(&bindings, pr)?
        };
        let local = lo.to_string();
        attrs.push(Attr {
            q,
            ns: ans,
            l: local,
            v,
        });
    }
    Ok(Node {
        q,
        ns,
        l: local,
        attrs,
        bindings,
        content: Vec::new(),
        source_start: 0,
        capture_raw,
        raw: None,
    })
}
fn attach(stack: &mut [Node], root: &mut Option<Node>, n: Node) -> Result<()> {
    if let Some(p) = stack.last_mut() {
        p.content.push(Content::Node(n));
    } else if root.replace(n).is_some() {
        return Err(invalid("multiple XML roots"));
    }
    Ok(())
}

/// Where a projection records the repairs it applies (audit finding 6.11):
/// a report for the tolerant read, or nothing for the strict read, which
/// refuses any markup that would need a repair.
struct Tol<'r> {
    report: Option<&'r mut ToleranceReport>,
    limits: Limits,
}

impl Tol<'_> {
    fn repair(&mut self, defect: impl Into<Defect>, what: &str) -> Result<()> {
        match self.report.as_deref_mut() {
            Some(report) => {
                report.record(defect, 0);
                Ok(())
            },
            None => Err(invalid(what.to_string())),
        }
    }
}

/// The children of `n` named in `names` (in schema order) and drawn from the
/// `namespaces`, one slot per name. Foreign or unknown children are skipped,
/// a repeated child keeps its first occurrence, and any order is accepted;
/// each such deviation is a repair.
fn slots<'n, const N: usize>(
    n: &'n Node,
    names: [&str; N],
    namespaces: [&str; 2],
    tol: &mut Tol<'_>,
) -> Result<[Option<&'n Node>; N]> {
    let mut found = [None; N];
    let mut order = 0;
    for c in kids(n)? {
        let position = if namespaces.contains(&c.ns.as_str()) {
            names.iter().position(|name| *name == c.l)
        } else {
            None
        };
        let Some(position) = position else {
            tol.repair(
                ViewDefect::UnknownElement,
                &format!("unexpected {} child '{}'", n.l, c.q),
            )?;
            continue;
        };
        if position < order {
            tol.repair(
                ViewDefect::ElementOrder,
                &format!("{} children out of order", n.l),
            )?;
        }
        order = order.max(position);
        if found[position].is_some() {
            tol.repair(
                ViewDefect::DuplicateElement,
                &format!("duplicate {} child '{}'", n.l, c.q),
            )?;
            continue;
        }
        found[position] = Some(c);
    }
    Ok(found)
}

const PNS: [&str; 2] = [P, PS];
const ANS: [&str; 2] = [A, AS];

fn required<'n>(slot: Option<&'n Node>, what: &str) -> Result<&'n Node> {
    slot.ok_or_else(|| invalid(format!("{what} is required")))
}

/// The list items of `n` named `item`; other children are skipped.
fn items<'n>(n: &'n Node, item: &str, tol: &mut Tol<'_>) -> Result<Vec<&'n Node>> {
    let mut list = Vec::new();
    for c in kids(n)? {
        if (c.ns == P || c.ns == PS) && c.l == item {
            list.push(c);
        } else {
            tol.repair(
                ViewDefect::UnknownElement,
                &format!("unexpected {} child '{}'", n.l, c.q),
            )?;
        }
    }
    Ok(list)
}

fn project(n: &Node, tol: &mut Tol<'_>) -> Result<ViewProperties> {
    expect_p(n, "viewPr")?;
    let mut v = ViewProperties {
        last_view: aopt(n, "lastView")?.map(parse_view).transpose()?,
        show_comments: bopt(n, "showComments")?,
        ..Default::default()
    };
    only(n, &["lastView", "showComments"], tol)?;
    let [normal, slide, outline, notes_text, sorter, notes, grid, ext] = slots(
        n,
        [
            "normalViewPr",
            "slideViewPr",
            "outlineViewPr",
            "notesTextViewPr",
            "sorterViewPr",
            "notesViewPr",
            "gridSpacing",
            "extLst",
        ],
        PNS,
        tol,
    )?;
    v.normal = normal.map(|c| parse_normal(c, tol)).transpose()?;
    v.slide = slide.map(|c| parse_slide_like(c, tol)).transpose()?;
    v.outline = outline.map(|c| parse_outline(c, tol)).transpose()?;
    v.notes_text = notes_text.map(|c| parse_simple(c, tol)).transpose()?;
    v.sorter = sorter.map(|c| parse_sorter(c, tol)).transpose()?;
    v.notes = notes.map(|c| parse_slide_like(c, tol)).transpose()?;
    v.grid_spacing = grid
        .map(|c| -> Result<GridSpacing> {
            only(c, &["cx", "cy"], tol)?;
            Ok(GridSpacing {
                cx: u32req(c, "cx")?,
                cy: u32req(c, "cy")?,
            })
        })
        .transpose()?;
    v.extension_xml = ext.map(|c| node_xml(c, false, tol.limits)).transpose()?;
    validate(&v, tol.limits)?;
    Ok(v)
}
fn parse_normal(n: &Node, tol: &mut Tol<'_>) -> Result<NormalView> {
    let show_outline_icons = bopt(n, "showOutlineIcons")?;
    let snap_vertical_splitter = bopt(n, "snapVertSplitter")?;
    let vertical_bar_state = aopt(n, "vertBarState")?.map(parse_splitter).transpose()?;
    let horizontal_bar_state = aopt(n, "horzBarState")?.map(parse_splitter).transpose()?;
    let prefer_single_view = bopt(n, "preferSingleView")?;
    only(
        n,
        &[
            "showOutlineIcons",
            "snapVertSplitter",
            "vertBarState",
            "horzBarState",
            "preferSingleView",
        ],
        tol,
    )?;
    let [left, top, ext] = slots(n, ["restoredLeft", "restoredTop", "extLst"], PNS, tol)?;
    let restored_left = parse_pane(required(left, "normalViewPr restoredLeft")?, tol)?;
    let restored_top = parse_pane(required(top, "normalViewPr restoredTop")?, tol)?;
    let extension_xml = ext.map(|c| node_xml(c, false, tol.limits)).transpose()?;
    Ok(NormalView {
        show_outline_icons,
        snap_vertical_splitter,
        vertical_bar_state,
        horizontal_bar_state,
        prefer_single_view,
        restored_left,
        restored_top,
        extension_xml,
    })
}
fn parse_pane(n: &Node, tol: &mut Tol<'_>) -> Result<RestoredPane> {
    let size = u32req(n, "sz")?;
    if size > 100_000 {
        return Err(invalid("restored pane size exceeds 100000"));
    }
    let auto_adjust = bopt(n, "autoAdjust")?;
    only(n, &["sz", "autoAdjust"], tol)?;
    leaf(n, tol)?;
    Ok(RestoredPane { size, auto_adjust })
}
fn parse_common(n: &Node, tol: &mut Tol<'_>) -> Result<CommonView> {
    expect_p(n, "cViewPr")?;
    let variable_scale = bopt(n, "varScale")?;
    only(n, &["varScale"], tol)?;
    let [scale, origin] = slots(n, ["scale", "origin"], PNS, tol)?;
    let scale = required(scale, "cViewPr scale")?;
    let origin = required(origin, "cViewPr origin")?;
    only(scale, &[], tol)?;
    let [sx, sy] = slots(scale, ["sx", "sy"], ANS, tol)?;
    let scale_x = ratio(required(sx, "scale sx")?, tol)?;
    let scale_y = ratio(required(sy, "scale sy")?, tol)?;
    let point = Point {
        x: i64req(origin, "x")?,
        y: i64req(origin, "y")?,
    };
    only(origin, &["x", "y"], tol)?;
    leaf(origin, tol)?;
    Ok(CommonView {
        variable_scale,
        scale_x,
        scale_y,
        origin: point,
    })
}
fn ratio(n: &Node, tol: &mut Tol<'_>) -> Result<Ratio> {
    if n.ns != A && n.ns != AS {
        return Err(invalid("ratio must use DrawingML namespace"));
    }
    let numerator = i64req(n, "n")?;
    let denominator = i64req(n, "d")?;
    if denominator == 0 {
        return Err(invalid("scale denominator cannot be zero"));
    }
    only(n, &["n", "d"], tol)?;
    leaf(n, tol)?;
    Ok(Ratio {
        numerator,
        denominator,
    })
}
fn parse_common_slide(n: &Node, tol: &mut Tol<'_>) -> Result<CommonSlideView> {
    expect_p(n, "cSldViewPr")?;
    let snap_to_grid = bopt(n, "snapToGrid")?;
    let snap_to_objects = bopt(n, "snapToObjects")?;
    let show_guides = bopt(n, "showGuides")?;
    only(n, &["snapToGrid", "snapToObjects", "showGuides"], tol)?;
    let [view, guide_list] = slots(n, ["cViewPr", "guideLst"], PNS, tol)?;
    let view = parse_common(required(view, "cSldViewPr cViewPr")?, tol)?;
    let mut guides = Vec::new();
    if let Some(guide_list) = guide_list {
        only(guide_list, &[], tol)?;
        for g in items(guide_list, "guide", tol)? {
            if guides.len() >= tol.limits.max_guides() {
                return Err(limit(
                    crate::error::CodecLimitKind::Guides,
                    guides.len().checked_add(1).unwrap_or(usize::MAX),
                    tol.limits.max_guides(),
                ));
            }
            let orientation = aopt(g, "orient")?
                .map(|x| match x.as_str() {
                    "horz" => Ok(GuideOrientation::Horizontal),
                    "vert" => Ok(GuideOrientation::Vertical),
                    _ => Err(invalid("invalid guide orientation")),
                })
                .transpose()?;
            let position = i32opt(g, "pos")?;
            only(g, &["orient", "pos"], tol)?;
            leaf(g, tol)?;
            guides.push(Guide {
                orientation,
                position,
            });
        }
    }
    Ok(CommonSlideView {
        snap_to_grid,
        snap_to_objects,
        show_guides,
        view,
        guides,
    })
}
fn parse_slide_like(n: &Node, tol: &mut Tol<'_>) -> Result<SlideLikeView> {
    only(n, &[], tol)?;
    let [common, ext] = slots(n, ["cSldViewPr", "extLst"], PNS, tol)?;
    let common = parse_common_slide(required(common, "slide-like view cSldViewPr")?, tol)?;
    Ok(SlideLikeView {
        common,
        extension_xml: ext.map(|c| node_xml(c, false, tol.limits)).transpose()?,
    })
}
fn parse_simple(n: &Node, tol: &mut Tol<'_>) -> Result<SimpleView> {
    only(n, &[], tol)?;
    let [view, ext] = slots(n, ["cViewPr", "extLst"], PNS, tol)?;
    Ok(SimpleView {
        view: parse_common(required(view, "view cViewPr")?, tol)?,
        extension_xml: ext.map(|c| node_xml(c, false, tol.limits)).transpose()?,
    })
}
fn parse_sorter(n: &Node, tol: &mut Tol<'_>) -> Result<SorterView> {
    let show_formatting = bopt(n, "showFormatting")?;
    only(n, &["showFormatting"], tol)?;
    let [view, ext] = slots(n, ["cViewPr", "extLst"], PNS, tol)?;
    Ok(SorterView {
        show_formatting,
        view: parse_common(required(view, "sorterViewPr cViewPr")?, tol)?,
        extension_xml: ext.map(|c| node_xml(c, false, tol.limits)).transpose()?,
    })
}
fn parse_outline(n: &Node, tol: &mut Tol<'_>) -> Result<OutlineView> {
    only(n, &[], tol)?;
    let [view, slide_list, ext] = slots(n, ["cViewPr", "sldLst", "extLst"], PNS, tol)?;
    let view = parse_common(required(view, "outlineViewPr cViewPr")?, tol)?;
    let mut slides = Vec::new();
    if let Some(slide_list) = slide_list {
        only(slide_list, &[], tol)?;
        for s in items(slide_list, "sld", tol)? {
            if slides.len() >= tol.limits.max_slides() {
                return Err(limit(
                    crate::error::CodecLimitKind::Slides,
                    slides.len().checked_add(1).unwrap_or(usize::MAX),
                    tol.limits.max_slides(),
                ));
            }
            let id = arel(s, "id")?;
            let collapse = bopt(s, "collapse")?;
            only_rel(s, &["collapse"], "id", tol)?;
            leaf(s, tol)?;
            slides.push(OutlineSlide {
                relationship_id: id,
                collapse,
                target: None,
            });
        }
    }
    Ok(OutlineView {
        view,
        slides,
        extension_xml: ext.map(|c| node_xml(c, false, tol.limits)).transpose()?,
    })
}

fn write_normal(x: &mut XmlSink, v: &NormalView, s: bool, limits: Limits) -> Result<()> {
    x.push_str("<p:normalViewPr");
    for (n, b) in [
        ("showOutlineIcons", v.show_outline_icons),
        ("snapVertSplitter", v.snap_vertical_splitter),
        ("preferSingleView", v.prefer_single_view),
    ] {
        bool_attr(x, n, b);
    }
    if let Some(z) = v.vertical_bar_state {
        attr(x, "vertBarState", splitter_str(z));
    }
    if let Some(z) = v.horizontal_bar_state {
        attr(x, "horzBarState", splitter_str(z));
    }
    x.push('>');
    pane(x, "restoredLeft", &v.restored_left);
    pane(x, "restoredTop", &v.restored_top);
    if let Some(e) = &v.extension_xml {
        opaque(x, e, s, limits)?;
    }
    x.push_str("</p:normalViewPr>");
    Ok(())
}
fn pane(x: &mut XmlSink, n: &str, v: &RestoredPane) {
    x.push_str(&format!("<p:{n} sz=\"{}\"", v.size));
    bool_attr(x, "autoAdjust", v.auto_adjust);
    x.push_str("/>");
}
fn write_common(x: &mut XmlSink, v: &CommonView) {
    x.push_str("<p:cViewPr");
    bool_attr(x, "varScale", v.variable_scale);
    x.push_str("><p:scale>");
    ratio_write(x, "a:sx", &v.scale_x);
    ratio_write(x, "a:sy", &v.scale_y);
    x.push_str(&format!(
        "</p:scale><p:origin x=\"{}\" y=\"{}\"/></p:cViewPr>",
        v.origin.x, v.origin.y
    ));
}
fn ratio_write(x: &mut XmlSink, n: &str, v: &Ratio) {
    x.push_str(&format!(
        "<{n} n=\"{}\" d=\"{}\"/>",
        v.numerator, v.denominator
    ));
}
fn write_common_slide(x: &mut XmlSink, v: &CommonSlideView) {
    x.push_str("<p:cSldViewPr");
    for (n, b) in [
        ("snapToGrid", v.snap_to_grid),
        ("snapToObjects", v.snap_to_objects),
        ("showGuides", v.show_guides),
    ] {
        bool_attr(x, n, b);
    }
    x.push('>');
    write_common(x, &v.view);
    if !v.guides.is_empty() {
        x.push_str("<p:guideLst>");
        for g in &v.guides {
            x.push_str("<p:guide");
            if let Some(o) = g.orientation {
                attr(
                    x,
                    "orient",
                    if o == GuideOrientation::Horizontal {
                        "horz"
                    } else {
                        "vert"
                    },
                );
            }
            if let Some(p) = g.position {
                attr(x, "pos", &p.to_string());
            }
            x.push_str("/>");
        }
        x.push_str("</p:guideLst>");
    }
    x.push_str("</p:cSldViewPr>");
}
fn write_slide_like(
    x: &mut XmlSink,
    n: &str,
    v: &SlideLikeView,
    s: bool,
    limits: Limits,
) -> Result<()> {
    x.push_str(&format!("<p:{n}>"));
    write_common_slide(x, &v.common);
    if let Some(e) = &v.extension_xml {
        opaque(x, e, s, limits)?;
    }
    x.push_str(&format!("</p:{n}>"));
    Ok(())
}
fn write_simple(x: &mut XmlSink, n: &str, v: &SimpleView, s: bool, limits: Limits) -> Result<()> {
    x.push_str(&format!("<p:{n}>"));
    write_common(x, &v.view);
    if let Some(e) = &v.extension_xml {
        opaque(x, e, s, limits)?;
    }
    x.push_str(&format!("</p:{n}>"));
    Ok(())
}
fn write_sorter(x: &mut XmlSink, v: &SorterView, s: bool, limits: Limits) -> Result<()> {
    x.push_str("<p:sorterViewPr");
    bool_attr(x, "showFormatting", v.show_formatting);
    x.push('>');
    write_common(x, &v.view);
    if let Some(e) = &v.extension_xml {
        opaque(x, e, s, limits)?;
    }
    x.push_str("</p:sorterViewPr>");
    Ok(())
}
fn write_outline(x: &mut XmlSink, v: &OutlineView, s: bool, limits: Limits) -> Result<()> {
    x.push_str("<p:outlineViewPr>");
    write_common(x, &v.view);
    if !v.slides.is_empty() {
        x.push_str("<p:sldLst>");
        for slide in &v.slides {
            x.push_str("<p:sld r:id=\"");
            esc(x, &slide.relationship_id);
            x.push('"');
            bool_attr(x, "collapse", slide.collapse);
            x.push_str("/>");
        }
        x.push_str("</p:sldLst>");
    }
    if let Some(e) = &v.extension_xml {
        opaque(x, e, s, limits)?;
    }
    x.push_str("</p:outlineViewPr>");
    Ok(())
}

fn validate(v: &ViewProperties, limits: Limits) -> Result<()> {
    for ext in [
        v.extension_xml.as_deref(),
        v.normal.as_ref().and_then(|x| x.extension_xml.as_deref()),
        v.slide.as_ref().and_then(|x| x.extension_xml.as_deref()),
        v.outline.as_ref().and_then(|x| x.extension_xml.as_deref()),
        v.notes_text
            .as_ref()
            .and_then(|x| x.extension_xml.as_deref()),
        v.sorter.as_ref().and_then(|x| x.extension_xml.as_deref()),
        v.notes.as_ref().and_then(|x| x.extension_xml.as_deref()),
    ]
    .into_iter()
    .flatten()
    {
        parse_dom_processed(ext, limits)?;
    }
    for common in v
        .slide
        .iter()
        .map(|x| &x.common)
        .chain(v.notes.iter().map(|x| &x.common))
    {
        if common.guides.len() > limits.max_guides() {
            return Err(limit(
                crate::error::CodecLimitKind::Guides,
                common.guides.len(),
                limits.max_guides(),
            ));
        }
    }
    if let Some(outline) = v.outline.as_ref() {
        if outline.slides.len() > limits.max_slides() {
            return Err(limit(
                crate::error::CodecLimitKind::Slides,
                outline.slides.len(),
                limits.max_slides(),
            ));
        }
    }
    Ok(())
}

fn opaque(x: &mut XmlSink, b: &[u8], _strict: bool, _limits: Limits) -> Result<()> {
    let s = std::str::from_utf8(b).map_err(xml_error)?;
    x.push_str(s);
    x.check()?;
    Ok(())
}
fn node_xml(n: &Node, _strict: bool, limits: Limits) -> Result<Vec<u8>> {
    if let Some(raw) = &n.raw {
        let bindings = in_scope_bindings(n)?;
        let in_scope = litchi_ooxml_common::mce::InScopeNamespaces::from_declarations(bindings);
        let fragment = in_scope.make_self_contained(raw, &fragment_mce_limits(limits))?;
        return match fragment {
            std::borrow::Cow::Borrowed(value) => clone_fragment(value, limits),
            std::borrow::Cow::Owned(value) => Ok(value),
        };
    }
    let mut count = XmlSink::count(limits.max_output_bytes());
    node_write(&mut count, n)?;
    count.check()?;
    let mut output = XmlSink::output(count.len(), limits.max_output_bytes())?;
    node_write(&mut output, n)?;
    output.finish()
}

fn in_scope_bindings(n: &Node) -> Result<Vec<(Box<[u8]>, Box<[u8]>)>> {
    let mut bindings = Vec::new();
    bindings
        .try_reserve_exact(n.bindings.len())
        .map_err(|source| Error::Allocation {
            resource: "view-properties opaque namespace bindings",
            source,
        })?;
    for (prefix, namespace) in &n.bindings {
        bindings.push((
            boxed_copy(prefix.as_bytes(), "view-properties opaque namespace prefix")?,
            boxed_copy(namespace.as_bytes(), "view-properties opaque namespace URI")?,
        ));
    }
    Ok(bindings)
}

fn boxed_copy(value: &[u8], resource: &'static str) -> Result<Box<[u8]>> {
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation { resource, source })?;
    bytes.extend_from_slice(value);
    Ok(bytes.into_boxed_slice())
}

fn clone_fragment(value: &[u8], limits: Limits) -> Result<Vec<u8>> {
    if value.len() > limits.max_output_bytes() {
        return Err(limit(
            crate::error::CodecLimitKind::OutputBytes,
            value.len(),
            limits.max_output_bytes(),
        ));
    }
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(value.len())
        .map_err(|source| Error::Allocation {
            resource: "view-properties opaque fragment",
            source,
        })?;
    bytes.extend_from_slice(value);
    Ok(bytes)
}

fn fragment_mce_limits(limits: Limits) -> litchi_ooxml_common::mce::Limits {
    litchi_ooxml_common::mce::Limits {
        max_input_bytes: limits.max_input_bytes(),
        max_output_bytes: limits.max_output_bytes(),
        max_depth: limits.max_depth(),
        max_namespace_bindings: 4096.min(limits.max_nodes()),
        max_directive_tokens: 4096.min(limits.max_nodes()),
        max_choices_per_alternate: 1024.min(limits.max_nodes()),
        max_attributes_per_element: litchi_ooxml_common::mce::DEFAULT_MAX_ATTRIBUTES_PER_ELEMENT,
    }
}
fn node_write(x: &mut XmlSink, n: &Node) -> Result<()> {
    x.push('<');
    x.push_str(&n.q);
    for (p, u) in &n.bindings {
        if p.is_empty() {
            x.push_str(" xmlns=\"");
        } else {
            x.push_str(" xmlns:");
            x.push_str(p);
            x.push_str("=\"");
        }
        esc(x, u);
        x.push('"');
    }
    for a in &n.attrs {
        x.push(' ');
        x.push_str(&a.q);
        x.push_str("=\"");
        esc(x, &a.v);
        x.push('"');
    }
    if n.content.is_empty() {
        x.push_str("/>");
        return Ok(());
    }
    x.push('>');
    for c in &n.content {
        match c {
            Content::Node(n) => node_write(x, n)?,
            Content::Text(v) => text(x, v),
            Content::CData(v) => {
                x.push_str("<![CDATA[");
                x.push_str(v);
                x.push_str("]]>");
            },
            Content::Comment(v) => {
                x.push_str("<!--");
                x.push_str(v);
                x.push_str("-->");
            },
        }
    }
    x.push_str("</");
    x.push_str(&n.q);
    x.push('>');
    Ok(())
}
fn kids(n: &Node) -> Result<Vec<&Node>> {
    let mut v = Vec::new();
    for c in &n.content {
        match c {
            Content::Node(x) => v.push(x),
            Content::Text(x) if x.trim().is_empty() => {},
            Content::Comment(_) => {},
            _ => return Err(invalid("unexpected text in typed view properties")),
        }
    }
    Ok(v)
}
fn leaf(n: &Node, tol: &mut Tol<'_>) -> Result<()> {
    for c in kids(n)? {
        tol.repair(
            ViewDefect::UnknownElement,
            &format!("leaf view property {} has child '{}'", n.l, c.q),
        )?;
    }
    Ok(())
}
fn expect_p(n: &Node, l: &str) -> Result<()> {
    if (n.ns == P || n.ns == PS) && n.l == l {
        Ok(())
    } else {
        Err(invalid(format!("expected {l}")))
    }
}
fn aopt(n: &Node, l: &str) -> Result<Option<String>> {
    let mut v = None;
    for a in &n.attrs {
        if a.ns.is_empty() && a.l == l {
            if v.is_some() {
                return Err(invalid("duplicate attribute"));
            }
            v = Some(a.v.clone());
        }
    }
    Ok(v)
}
fn arel(n: &Node, l: &str) -> Result<String> {
    let mut v = None;
    for a in &n.attrs {
        if (a.ns == R || a.ns == RS) && a.l == l {
            if v.is_some() {
                return Err(invalid("duplicate relationship attribute"));
            }
            v = Some(a.v.clone());
        }
    }
    v.ok_or_else(|| invalid(format!("missing relationship attribute '{l}'")))
}
fn bopt(n: &Node, l: &str) -> Result<Option<bool>> {
    match aopt(n, l)?.as_deref() {
        None => Ok(None),
        Some("1" | "true") => Ok(Some(true)),
        Some("0" | "false") => Ok(Some(false)),
        _ => Err(invalid(format!("invalid boolean '{l}'"))),
    }
}
fn i64req(n: &Node, l: &str) -> Result<i64> {
    aopt(n, l)?
        .ok_or_else(|| invalid(format!("missing '{l}'")))?
        .parse()
        .map_err(|_err| invalid(format!("invalid integer '{l}'")))
}
fn u32req(n: &Node, l: &str) -> Result<u32> {
    aopt(n, l)?
        .ok_or_else(|| invalid(format!("missing '{l}'")))?
        .parse()
        .map_err(|_err| invalid(format!("invalid integer '{l}'")))
}
fn i32opt(n: &Node, l: &str) -> Result<Option<i32>> {
    aopt(n, l)?
        .map(|x| {
            x.parse()
                .map_err(|_err| invalid(format!("invalid integer '{l}'")))
        })
        .transpose()
}
/// Ignore (recording) every attribute of `n` outside `allowed`: newer
/// `PowerPoint` releases add attributes to these elements.
fn only(n: &Node, allowed: &[&str], tol: &mut Tol<'_>) -> Result<()> {
    for a in &n.attrs {
        if !a.ns.is_empty() || !allowed.contains(&a.l.as_str()) {
            tol.repair(
                ViewDefect::UnknownAttribute,
                &format!("unexpected attribute '{}'", a.q),
            )?;
        }
    }
    Ok(())
}
fn only_rel(n: &Node, plain: &[&str], rel: &str, tol: &mut Tol<'_>) -> Result<()> {
    for a in &n.attrs {
        if !((a.ns.is_empty() && plain.contains(&a.l.as_str()))
            || ((a.ns == R || a.ns == RS) && a.l == rel))
        {
            tol.repair(
                ViewDefect::UnknownAttribute,
                &format!("unexpected attribute '{}'", a.q),
            )?;
        }
    }
    Ok(())
}
fn split(q: &str) -> Result<(&str, &str)> {
    if let Some((p, l)) = q.split_once(':') {
        if l.is_empty() || l.contains(':') {
            return Err(invalid("invalid QName"));
        }
        Ok((p, l))
    } else {
        Ok(("", q))
    }
}
fn resolve(b: &[(String, String)], p: &str) -> Result<String> {
    b.iter()
        .rev()
        .find(|x| x.0 == p)
        .map(|x| x.1.clone())
        .ok_or_else(|| invalid(format!("unbound prefix '{p}'")))
}
fn attr(x: &mut XmlSink, n: &str, v: &str) {
    x.push(' ');
    x.push_str(n);
    x.push_str("=\"");
    esc(x, v);
    x.push('"');
}
fn bool_attr(x: &mut XmlSink, n: &str, v: Option<bool>) {
    if let Some(v) = v {
        attr(x, n, if v { "1" } else { "0" });
    }
}
fn esc(x: &mut XmlSink, v: &str) {
    for c in v.chars() {
        match c {
            '&' => x.push_str("&amp;"),
            '<' => x.push_str("&lt;"),
            '"' => x.push_str("&quot;"),
            '\r' => x.push_str("&#xD;"),
            '\n' => x.push_str("&#xA;"),
            '\t' => x.push_str("&#x9;"),
            _ => x.push(c),
        }
    }
}
fn text(x: &mut XmlSink, v: &str) {
    for c in v.chars() {
        match c {
            '&' => x.push_str("&amp;"),
            '<' => x.push_str("&lt;"),
            '>' => x.push_str("&gt;"),
            _ => x.push(c),
        }
    }
}
macro_rules! en{($p:ident,$w:ident,$t:ty,$($s:literal=>$v:path),+)=>{fn $p(s:String)->Result<$t>{match s.as_str(){$($s=>Ok($v),)+_=>Err(invalid(format!("invalid enumeration '{s}'")))}}fn $w(v:$t)->&'static str{match v{$($v=>$s,)+}}}}
en!(parse_view,view_str,ViewKind,"sldView"=>ViewKind::Slide,"sldMasterView"=>ViewKind::SlideMaster,"notesView"=>ViewKind::Notes,"handoutView"=>ViewKind::Handout,"notesMasterView"=>ViewKind::NotesMaster,"outlineView"=>ViewKind::Outline,"sldSorterView"=>ViewKind::SlideSorter,"sldThumbnailView"=>ViewKind::SlideThumbnail);
en!(parse_splitter,splitter_str,SplitterState,"minimized"=>SplitterState::Minimized,"restored"=>SplitterState::Restored,"maximized"=>SplitterState::Maximized);
fn invalid(v: impl Into<String>) -> Error {
    Error::Invalid(v.into())
}
fn xml_error(e: impl std::fmt::Display) -> Error {
    Error::Xml(e.to_string())
}

fn bounded_string(value: &str, limits: Limits) -> Result<()> {
    if value.len() > limits.max_string_bytes() {
        return Err(limit(
            crate::error::CodecLimitKind::StringBytes,
            value.len(),
            limits.max_string_bytes(),
        ));
    }
    Ok(())
}

fn bounded_attribute(value: &str, limits: Limits) -> Result<()> {
    if value.len() > limits.max_attribute_bytes {
        return Err(limit(
            crate::error::CodecLimitKind::XmlAttributeBytes,
            value.len(),
            limits.max_attribute_bytes,
        ));
    }
    Ok(())
}

fn limit(kind: crate::error::CodecLimitKind, observed: usize, maximum: usize) -> Error {
    Error::Limit {
        kind: crate::error::LimitKind::Codec(kind),
        observed,
        maximum,
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::super::package::load_from_package;
    use super::*;
    use litchi_opc::{OpcPackage, SourceLimits};
    fn f(b: &[u8]) -> ViewProperties {
        let p = OpcPackage::from_bytes(b, SourceLimits::default()).unwrap();
        load_from_package(&p, Limits::default()).unwrap().unwrap()
    }
    #[test]
    fn poi_prprops_guides_grid_and_strict_roundtrip() {
        let v = f(include_bytes!(
            "../../../../test-data/poi/test-data/slideshow/prProps.pptx"
        ));
        assert_eq!(v.slide.as_ref().unwrap().common.guides.len(), 2);
        assert_eq!(
            v.grid_spacing,
            Some(GridSpacing {
                cx: 72008,
                cy: 72008
            })
        );
        let x = v.to_xml(true, Limits::default()).unwrap();
        let r = ViewProperties::parse(&x, Limits::default()).unwrap();
        assert_eq!(
            r.slide.unwrap().common.view.scale_x,
            Ratio {
                numerator: 66,
                denominator: 100
            }
        );
    }
    #[test]
    fn poi_outline_and_splitters() {
        let v = f(include_bytes!(
            "../../../../test-data/poi/test-data/slideshow/45545_Comment.pptx"
        ));
        let n = v.normal.unwrap();
        assert_eq!(n.vertical_bar_state, Some(SplitterState::Minimized));
        assert_eq!(n.horizontal_bar_state, Some(SplitterState::Maximized));
        assert!(v.outline.is_some());
        assert_eq!(v.sorter.unwrap().view.origin.y, 1026);
    }
    #[test]
    fn libreoffice_ratio_and_sparse_views() {
        let v = f(include_bytes!(
            "../../../../test-data/libreoffice-core/oox/qa/unit/data/shape-text-alignment.pptx"
        ));
        assert_eq!(
            v.notes_text.unwrap().view.scale_x,
            Ratio {
                numerator: 1,
                denominator: 1
            }
        );
        assert!(v.sorter.is_none());
    }
    #[test]
    fn strict_mce_and_malformed() {
        let x = format!(
            r#"<p:viewPr xmlns:p="{PS}" xmlns:a="{AS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:u" mc:Ignorable="u"><mc:AlternateContent><mc:Choice Requires="u"><u:x/></mc:Choice><mc:Fallback><p:gridSpacing cx="10" cy="20"/></mc:Fallback></mc:AlternateContent></p:viewPr>"#
        );
        assert_eq!(
            ViewProperties::parse(x.as_bytes(), Limits::default())
                .unwrap()
                .grid_spacing,
            Some(GridSpacing { cx: 10, cy: 20 })
        );
        for bad in [
            format!(r#"<p:viewPr xmlns:p="{P}" showComments="maybe"/>"#),
            format!(
                r#"<p:viewPr xmlns:p="{P}" xmlns:a="{A}"><p:notesTextViewPr><p:cViewPr><p:scale><a:sx n="1" d="0"/><a:sy n="1" d="1"/></p:scale><p:origin x="0" y="0"/></p:cViewPr></p:notesTextViewPr></p:viewPr>"#
            ),
            format!(r#"<!DOCTYPE x><p:viewPr xmlns:p="{P}"/>"#),
        ] {
            assert!(
                ViewProperties::parse(bad.as_bytes(), Limits::default()).is_err(),
                "accepted {bad}"
            );
        }
    }

    #[test]
    fn explicit_limits_bound_parse_and_serialize() {
        let limits = Limits::new(1, 1, 1, 1, 1, 1, 1).expect("narrow view policy");
        let xml = format!(r#"<p:viewPr xmlns:p="{P}"/>"#);
        assert!(matches!(
            ViewProperties::parse(xml.as_bytes(), limits),
            Err(Error::Limit { maximum: 1, .. })
        ));
        assert!(matches!(
            ViewProperties::default().to_xml(false, limits),
            Err(Error::Limit { maximum: 1, .. })
        ));
    }

    #[test]
    fn opaque_fragment_preserves_namespace_uri_text_and_attributes() {
        let xml = format!(
            r#"<p:viewPr xmlns:p="{P}" xmlns:u="urn:opaque"><p:extLst>
  <u:fragment  note='urn:opaque' data="urn:opaque"><!-- keep -->literal urn:opaque</u:fragment>
</p:extLst></p:viewPr>"#
        );
        let value = ViewProperties::parse(xml.as_bytes(), Limits::default()).unwrap();
        let expected = format!(
            r#"<p:extLst xmlns:p="{P}" xmlns:u="urn:opaque">
  <u:fragment  note='urn:opaque' data="urn:opaque"><!-- keep -->literal urn:opaque</u:fragment>
</p:extLst>"#
        );
        assert_eq!(value.extension_xml.as_deref(), Some(expected.as_bytes()));
        let output = String::from_utf8(value.to_xml(true, Limits::default()).unwrap()).unwrap();
        assert!(
            output
                .as_bytes()
                .windows(expected.len())
                .any(|window| window == expected.as_bytes())
        );
    }

    #[test]
    fn mce_expansion_uses_output_budget_for_processed_parse() {
        let xml = format!(
            r#"<p:viewPr xmlns:p="{P}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:n="urn:namespace-closure-expands-each-descendant"><p:gridSpacing cx="1" cy="2"/></p:viewPr>"#
        );
        let processed = litchi_ooxml_common::mce::process_markup_compatibility(
            xml.as_bytes(),
            &litchi_ooxml_common::mce::Capabilities::ooxml_baseline(),
            &litchi_ooxml_common::mce::Limits::default(),
        )
        .unwrap()
        .xml;
        assert!(processed.len() > xml.len());
        let limits = Limits::new(
            xml.len(),
            processed.len(),
            Limits::DEFAULT.max_depth(),
            Limits::DEFAULT.max_nodes(),
            Limits::DEFAULT.max_guides(),
            Limits::DEFAULT.max_slides(),
            Limits::DEFAULT.max_string_bytes(),
        )
        .unwrap();
        assert!(ViewProperties::parse(xml.as_bytes(), limits).is_ok());
    }
}
