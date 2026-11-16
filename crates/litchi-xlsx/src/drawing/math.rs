//! Office Math equations in worksheet drawings.
//!
//! An equation is a text box whose paragraph holds the OMML in
//! `a14:m` ([MS-ODRAWXML] §2.3.3.41), inside the `mc:Choice
//! Requires="a14"` branch of a markup-compatibility block (ECMA-376 Part 3)
//! whose `mc:Fallback` shows readers without Office Math the same object
//! without the equation. In a drawing part the block may hold a whole
//! anchor (a child of `xdr:wsDr`), the object of one anchor, or sit inside
//! a paragraph. [`equations`] lists every `a14:m` of a drawing part in
//! document order, with the cell anchor and the `cNvPr` of the shape that
//! holds it, its OMML and its text; [`crate::Package::equations`] and
//! [`crate::workbook::SourceWorksheet::equations`] read a worksheet's
//! drawing. With the `equation` feature, `Equation::latex` converts the
//! OMML to LaTeX, and `drawing::edit::Edit::{add_equation, set_equation,
//! remove_equation}` write equations.

use std::ops::Range;

use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::edit::{LimitKind, Marker, Placement};
use crate::error::{Error, Result, invalid};

#[cfg(feature = "equation")]
pub use litchi_drawingml::math::EquationSource;

const MCE: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const A14: &[u8] = b"http://schemas.microsoft.com/office/drawing/2010/main";
const MATH: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/math";
const XDR: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const STRICT_XDR: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
const DML: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DML: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";
/// The deepest element nesting a drawing part may have here.
const MAX_DEPTH: usize = 256;
/// The most equations one drawing part may hold.
pub const MAX_EQUATIONS: usize = 4096;
/// The default text-body insets in EMU (ECMA-376 Part 1, §21.1.2.1.1).
const DEFAULT_INSETS: [i64; 4] = [91_440, 45_720, 91_440, 45_720];

/// The outcome of an equation write.
#[cfg(feature = "equation")]
#[derive(Debug, Clone, PartialEq, Eq)]
#[non_exhaustive]
pub struct EquationWrite {
    /// The `cNvPr` ID of the shape that holds the equation.
    pub shape_id: u32,
    /// The equation's position among the worksheet's equations
    /// ([`crate::Package::equations`]).
    pub index: usize,
    /// What the LaTeX conversion could not carry into OMML, one message per
    /// loss; empty for OMML input and lossless conversions.
    pub losses: Vec<String>,
}

/// One equation of a worksheet drawing.
#[derive(Clone, PartialEq)]
pub struct Equation {
    shape_id: Option<u32>,
    name: Option<String>,
    placement: Option<Placement>,
    omml: String,
    text: String,
}

impl std::fmt::Debug for Equation {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Equation")
            .field("shape_id", &self.shape_id)
            .field("placement", &self.placement)
            .finish_non_exhaustive()
    }
}

impl Equation {
    /// The `cNvPr` ID of the shape whose text holds the equation.
    #[must_use]
    pub const fn shape_id(&self) -> Option<u32> {
        self.shape_id
    }

    /// The `cNvPr` name of the shape whose text holds the equation.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Where the shape is anchored on the worksheet.
    #[must_use]
    pub const fn placement(&self) -> Option<Placement> {
        self.placement
    }

    /// The A1 reference of the cell the shape's anchor starts in.
    #[must_use]
    pub fn cell(&self) -> Option<String> {
        let from = match self.placement? {
            Placement::TwoCell { from, .. } | Placement::OneCell { from, .. } => from,
            _ => return None,
        };
        Some(a1(from))
    }

    /// The equation's OMML: the exact XML inside its `a14:m` element (an
    /// `m:oMathPara` or `m:oMath`).
    #[must_use]
    pub fn omml(&self) -> &str {
        &self.omml
    }

    /// The equation as LaTeX, converted from its OMML.
    ///
    /// Requires the `equation` feature.
    ///
    /// # Errors
    ///
    /// Returns the conversion error when the OMML is malformed or exceeds
    /// [`litchi_equation::Limits::DEFAULT`].
    #[cfg(feature = "equation")]
    pub fn latex(&self) -> litchi_equation::Result<String> {
        litchi_equation::omml_to_latex(&self.omml, &litchi_equation::Limits::DEFAULT)
    }

    /// The equation's text: every `m:t` in document order.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}

/// The A1 reference of a marker's cell.
fn a1(marker: Marker) -> String {
    let mut letters = Vec::new();
    let mut column = u64::from(marker.col) + 1;
    while column > 0 {
        let rest = (column - 1) % 26;
        letters.push(b'A' + u8::try_from(rest).unwrap_or(0));
        column = (column - 1) / 26;
    }
    letters.reverse();
    format!(
        "{}{}",
        String::from_utf8_lossy(&letters),
        u64::from(marker.row) + 1
    )
}

/// Every equation of a drawing part's XML, in document order.
///
/// # Errors
///
/// Returns an error for malformed XML or exceeded bounds.
pub fn equations(xml: &[u8]) -> Result<Vec<Equation>> {
    let (sites, _) = scan(xml)?;
    Ok(sites
        .into_iter()
        .map(|site| Equation {
            shape_id: site.shape_id,
            name: site.name,
            placement: site.placement,
            omml: String::from_utf8_lossy(&xml[site.content])
                .trim()
                .to_owned(),
            text: site.text,
        })
        .collect())
}

/// Where a markup-compatibility block sits in the drawing part.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub(crate) enum Level {
    /// A child of `xdr:wsDr`: each branch holds a whole anchor.
    Root,
    /// A child of an anchor: each branch holds the anchor's object.
    Anchor,
    /// Anywhere else (inside a shape's text, for example).
    #[default]
    Inner,
}

/// One `mc:AlternateContent` block of a drawing part.
#[derive(Debug, Clone, Default)]
#[cfg_attr(
    not(feature = "equation"),
    allow(dead_code, reason = "the equation edits read it")
)]
pub(crate) struct Block {
    /// The whole block.
    pub(crate) range: Range<usize>,
    pub(crate) level: Level,
    /// The `mc:Choice` branches of the block.
    pub(crate) choices: usize,
    /// Whether a branch holds markup-compatibility content of its own.
    pub(crate) nested: bool,
    /// The start tag of the first `mc:Choice`.
    pub(crate) choice_tag: Option<Range<usize>>,
    /// The content of the first `mc:Choice`, between its tags.
    pub(crate) choice_content: Option<Range<usize>>,
    /// The element children of the first `mc:Choice`.
    pub(crate) choice_children: usize,
    /// The Choice branch's shape (`xdr:sp`, a child of the branch or of the
    /// branch's anchor).
    pub(crate) choice_shape: Option<Range<usize>>,
    /// The anchor that holds an anchor-level block.
    pub(crate) outer_anchor: Option<usize>,
    /// The `mc:Fallback` element, and whether it is empty.
    pub(crate) fallback: Option<(Range<usize>, bool)>,
    /// The start tag of the `mc:Fallback` element.
    pub(crate) fallback_tag: Option<Range<usize>>,
    /// The equations in the Choice branch.
    pub(crate) equations: usize,
    /// Whether the Choice branch holds text besides its equations.
    pub(crate) other_content: bool,
    /// The Choice shape's text anchor (`a:bodyPr@anchor`).
    pub(crate) text_anchor: Option<String>,
    /// The Choice shape's text insets in EMU: left, top, right, bottom.
    pub(crate) insets: [i64; 4],
    /// Whether the Choice shape resizes to its text (`a:spAutoFit`).
    pub(crate) autofit: bool,
    /// The `a:ext` start tag of the Choice shape's frame.
    pub(crate) choice_extent: Option<Range<usize>>,
    /// The `a:ext` start tag of the Fallback shape's frame.
    pub(crate) fallback_extent: Option<Range<usize>>,
    /// The `a:blip` start tag of the Fallback shape's picture fill.
    pub(crate) blip: Option<Range<usize>>,
    /// The whole `a:stretch` element of the Fallback shape's picture fill.
    pub(crate) stretch: Option<Range<usize>>,
}

/// One drawing anchor.
#[derive(Debug, Clone, Default)]
pub(crate) struct Anchor {
    pub(crate) range: Range<usize>,
    kind: Vec<u8>,
    markers: [[i64; 4]; 2],
    extent: Option<(i64, i64)>,
    position: Option<(i64, i64)>,
    /// The `xdr:ext` start tag of a one-cell or absolute anchor.
    pub(crate) extent_tag: Option<Range<usize>>,
    /// The whole `xdr:to` marker of a two-cell anchor.
    pub(crate) to_marker: Option<Range<usize>>,
}

impl Anchor {
    #[cfg_attr(
        not(feature = "equation"),
        allow(dead_code, reason = "the equation edits read it")
    )]
    pub(crate) fn two_cell(&self) -> bool {
        self.kind == b"twoCellAnchor"
    }

    /// The anchor's `from` marker.
    #[cfg_attr(
        not(feature = "equation"),
        allow(dead_code, reason = "the equation edits read it")
    )]
    pub(crate) fn from(&self) -> Marker {
        let values = self.markers[0];
        Marker {
            col: u32::try_from(values[0]).unwrap_or(0),
            col_offset: values[1],
            row: u32::try_from(values[2]).unwrap_or(0),
            row_offset: values[3],
        }
    }

    fn placement(&self) -> Option<Placement> {
        let marker = |values: [i64; 4]| Marker {
            col: u32::try_from(values[0]).unwrap_or(0),
            col_offset: values[1],
            row: u32::try_from(values[2]).unwrap_or(0),
            row_offset: values[3],
        };
        match self.kind.as_slice() {
            b"twoCellAnchor" => Some(Placement::TwoCell {
                from: marker(self.markers[0]),
                to: marker(self.markers[1]),
            }),
            b"oneCellAnchor" => self.extent.map(|(cx, cy)| Placement::OneCell {
                from: marker(self.markers[0]),
                cx,
                cy,
            }),
            b"absoluteAnchor" => self
                .position
                .zip(self.extent)
                .map(|((x, y), (cx, cy))| Placement::Absolute { x, y, cx, cy }),
            _ => None,
        }
    }
}

/// One equation (`a14:m`) of a drawing part.
#[derive(Debug, Clone)]
#[cfg_attr(
    not(feature = "equation"),
    allow(dead_code, reason = "the equation edits read it")
)]
pub(crate) struct Site {
    /// The whole `a14:m` element.
    pub(crate) element: Range<usize>,
    /// The element's content: the OMML.
    pub(crate) content: Range<usize>,
    text: String,
    shape_id: Option<u32>,
    name: Option<String>,
    placement: Option<Placement>,
    anchor: Option<usize>,
    /// The root- or anchor-level block whose Choice shape holds the equation
    /// directly in a paragraph of its text body; `None` for any other
    /// placement.
    pub(crate) block: Option<usize>,
    /// The alignment of its paragraph (`a:pPr@algn`), when given before it.
    pub(crate) align: Option<String>,
}

/// What an open element is, as far as the scan cares.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Tag {
    Block,
    Choice,
    Fallback,
    Anchor,
    Marker(usize),
    MarkerField(usize, usize),
    AnchorExtent,
    Shape,
    ShapeProperties,
    Transform,
    TextBody,
    BodyProperties,
    Paragraph,
    BlipFill,
    Stretch,
    Math,
    MathText,
    Text,
    Other,
}

/// One open element.
#[derive(Debug, Clone, Copy)]
struct Frame {
    tag: Tag,
    start: usize,
    /// The block a [`Tag::Block`] frame is.
    block: usize,
    /// The block branch (`(block, is_choice)`) the element is inside.
    owner: Option<(usize, bool)>,
    /// The innermost anchor the element is inside.
    anchor: Option<usize>,
    /// The innermost shape the element is inside, and whether it is the
    /// shape of its block branch.
    shape: Option<(usize, bool)>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Space {
    Mce,
    A14,
    Math,
    Xdr,
    Dml,
    Other,
}

fn space(namespace: &ResolveResult<'_>) -> Space {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => match *value {
            MCE => Space::Mce,
            A14 => Space::A14,
            MATH => Space::Math,
            XDR | STRICT_XDR => Space::Xdr,
            DML | STRICT_DML => Space::Dml,
            _ => Space::Other,
        },
        _ => Space::Other,
    }
}

fn text_attribute(element: &BytesStart<'_>, name: &[u8]) -> Option<String> {
    element
        .checked_attributes()
        .flatten()
        .find(|attribute| attribute.key.as_ref() == name)
        .and_then(|attribute| String::from_utf8(attribute.value.into_owned()).ok())
}

fn integer_attribute(element: &BytesStart<'_>, name: &[u8]) -> Option<i64> {
    text_attribute(element, name).and_then(|value| value.trim().parse().ok())
}

/// The text of a character-data event or entity reference.
fn decoded(event: &Event<'_>) -> Result<Option<String>> {
    Ok(match event {
        Event::Text(text) => Some(
            text.decode()
                .map_err(|error| invalid(format!("drawing equations: {error}")))?
                .into_owned(),
        ),
        Event::CData(data) => Some(String::from_utf8_lossy(data.as_ref()).into_owned()),
        Event::GeneralRef(reference) => {
            let name = reference
                .decode()
                .map_err(|error| invalid(format!("drawing equations: {error}")))?;
            let resolved = match name.as_ref() {
                "amp" => Some('&'),
                "lt" => Some('<'),
                "gt" => Some('>'),
                "quot" => Some('"'),
                "apos" => Some('\''),
                _ => reference.resolve_char_ref().ok().flatten(),
            };
            resolved.map(String::from)
        },
        _ => None,
    })
}

/// Every equation of a drawing part's XML, in document order, and the
/// markup-compatibility blocks of the part.
///
/// # Errors
///
/// Returns an error for malformed XML or exceeded bounds.
pub(crate) fn scan(xml: &[u8]) -> Result<(Vec<Site>, Vec<Block>)> {
    let (sites, blocks, _) = scan_all(xml)?;
    Ok((sites, blocks))
}

/// [`scan`], with the part's anchors.
#[allow(
    clippy::too_many_lines,
    reason = "one streaming pass records every range the edits need"
)]
pub(crate) fn scan_all(xml: &[u8]) -> Result<(Vec<Site>, Vec<Block>, Vec<Anchor>)> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let position = |reader: &NsReader<&[u8]>| {
        usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid("the drawing is too large"))
    };
    let mut stack: Vec<Frame> = Vec::new();
    let mut sites: Vec<Site> = Vec::new();
    let mut blocks: Vec<Block> = Vec::new();
    let mut anchors: Vec<Anchor> = Vec::new();
    // Each shape's `cNvPr` ID and name.
    let mut shapes: Vec<(Option<u32>, Option<String>)> = Vec::new();
    // The open equation: its index in `sites` and its depth.
    let mut math: Option<(usize, usize)> = None;
    let mut paragraph_align: Option<String> = None;
    loop {
        let start = position(&reader)?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| invalid(format!("drawing equations: {error}")))?;
        let space = space(&namespace);
        let end = position(&reader)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let empty = matches!(event, Event::Empty(_));
                let local = element.local_name();
                let local = local.as_ref();
                let parent = stack.last().copied();
                let parent_tag = parent.map_or(Tag::Other, |frame| frame.tag);
                let grandparent_tag = stack
                    .len()
                    .checked_sub(2)
                    .map_or(Tag::Other, |at| stack[at].tag);
                let mut owner = parent.and_then(|frame| frame.owner);
                let mut anchor = parent.and_then(|frame| frame.anchor);
                let mut shape = parent.and_then(|frame| frame.shape);
                let tag = match (space, local, parent_tag) {
                    (Space::Mce, b"AlternateContent", _) => Tag::Block,
                    (Space::Mce, b"Choice", Tag::Block) => Tag::Choice,
                    (Space::Mce, b"Fallback", Tag::Block) => Tag::Fallback,
                    (Space::Xdr, b"twoCellAnchor" | b"oneCellAnchor" | b"absoluteAnchor", _) => {
                        Tag::Anchor
                    },
                    (Space::Xdr, b"from", Tag::Anchor) => Tag::Marker(0),
                    (Space::Xdr, b"to", Tag::Anchor) => Tag::Marker(1),
                    (Space::Xdr, b"col", Tag::Marker(slot)) => Tag::MarkerField(slot, 0),
                    (Space::Xdr, b"colOff", Tag::Marker(slot)) => Tag::MarkerField(slot, 1),
                    (Space::Xdr, b"row", Tag::Marker(slot)) => Tag::MarkerField(slot, 2),
                    (Space::Xdr, b"rowOff", Tag::Marker(slot)) => Tag::MarkerField(slot, 3),
                    (Space::Xdr, b"ext", Tag::Anchor) => Tag::AnchorExtent,
                    (Space::Xdr, b"sp", _) => Tag::Shape,
                    (Space::Xdr, b"spPr", Tag::Shape) => Tag::ShapeProperties,
                    (Space::Dml, b"xfrm", Tag::ShapeProperties) => Tag::Transform,
                    (Space::Xdr, b"txBody", Tag::Shape) => Tag::TextBody,
                    (Space::Dml, b"bodyPr", Tag::TextBody) => Tag::BodyProperties,
                    (Space::Dml, b"p", Tag::TextBody) => Tag::Paragraph,
                    (Space::Dml, b"blipFill", Tag::ShapeProperties) => Tag::BlipFill,
                    (Space::Dml, b"stretch", Tag::BlipFill) => Tag::Stretch,
                    (Space::A14, b"m", _) if math.is_none() && !empty => Tag::Math,
                    (Space::Math, b"t", _) if math.is_some() && !empty => Tag::MathText,
                    (Space::Dml, b"t", _) => Tag::Text,
                    _ => Tag::Other,
                };
                let mut block_index = 0;
                match tag {
                    Tag::Block => {
                        if let Some((block, _)) = owner {
                            blocks[block].nested = true;
                        }
                        let level = match parent_tag {
                            Tag::Anchor => Level::Anchor,
                            _ if stack.len() == 1 => Level::Root,
                            _ => Level::Inner,
                        };
                        block_index = blocks.len();
                        blocks.push(Block {
                            level,
                            insets: DEFAULT_INSETS,
                            outer_anchor: if level == Level::Anchor { anchor } else { None },
                            ..Block::default()
                        });
                        owner = None;
                    },
                    Tag::Choice | Tag::Fallback => {
                        let block = parent.map_or(0, |frame| frame.block);
                        if tag == Tag::Choice {
                            blocks[block].choices += 1;
                            if blocks[block].choice_tag.is_none() {
                                blocks[block].choice_tag = Some(start..end);
                                blocks[block].choice_content = Some(end..end);
                            }
                        } else {
                            blocks[block].fallback = Some((start..end, empty));
                            blocks[block].fallback_tag = Some(start..end);
                        }
                        owner = Some((block, tag == Tag::Choice));
                    },
                    Tag::Anchor => {
                        anchor = Some(anchors.len());
                        anchors.push(Anchor {
                            range: start..end,
                            kind: local.to_vec(),
                            ..Anchor::default()
                        });
                    },
                    Tag::AnchorExtent => {
                        if let Some(index) = anchor {
                            let found = &mut anchors[index];
                            found.extent = integer_attribute(element, b"cx")
                                .zip(integer_attribute(element, b"cy"));
                            found.extent_tag = Some(start..end);
                        }
                    },
                    Tag::Shape => {
                        // A branch's shape is a child of the branch or of the
                        // branch's anchor.
                        let branch = matches!(parent_tag, Tag::Choice | Tag::Fallback)
                            || (parent_tag == Tag::Anchor
                                && matches!(grandparent_tag, Tag::Choice | Tag::Fallback));
                        shape = Some((shapes.len(), branch && owner.is_some()));
                        shapes.push((None, None));
                    },
                    Tag::Paragraph => paragraph_align = None,
                    _ => {},
                }
                if parent_tag == Tag::Choice
                    && let Some((block, true)) = owner
                    && blocks[block].choice_tag.as_ref().map(|tag| tag.start)
                        == parent.map(|frame| frame.start)
                {
                    blocks[block].choice_children += 1;
                }
                if space == Space::Xdr
                    && local == b"cNvPr"
                    && let Some((index, _)) = shape
                    && shapes[index].0.is_none()
                {
                    shapes[index] = (
                        integer_attribute(element, b"id").and_then(|id| u32::try_from(id).ok()),
                        text_attribute(element, b"name"),
                    );
                }
                if space == Space::Xdr
                    && local == b"pos"
                    && parent_tag == Tag::Anchor
                    && let Some(index) = anchor
                {
                    anchors[index].position =
                        integer_attribute(element, b"x").zip(integer_attribute(element, b"y"));
                }
                // The properties of the block branch's own shape.
                if let (Some((block, choice)), Some((_, true))) = (owner, shape) {
                    let record = &mut blocks[block];
                    match (space, local, parent_tag) {
                        (Space::Dml, b"ext", Tag::Transform)
                            if grandparent_tag == Tag::ShapeProperties =>
                        {
                            let slot = if choice {
                                &mut record.choice_extent
                            } else {
                                &mut record.fallback_extent
                            };
                            if slot.is_none() {
                                *slot = Some(start..end);
                            }
                        },
                        (Space::Dml, b"spAutoFit", Tag::BodyProperties) if choice => {
                            record.autofit = true;
                        },
                        (Space::Dml, b"blip", Tag::BlipFill)
                            if !choice && record.blip.is_none() =>
                        {
                            record.blip = Some(start..end);
                        },
                        (Space::Dml, b"pPr", Tag::Paragraph) if choice => {
                            paragraph_align = text_attribute(element, b"algn");
                        },
                        _ => {},
                    }
                    if tag == Tag::BodyProperties && choice {
                        record.text_anchor = text_attribute(element, b"anchor");
                        for (slot, name) in record.insets.iter_mut().zip([
                            b"lIns".as_slice(),
                            b"tIns",
                            b"rIns",
                            b"bIns",
                        ]) {
                            if let Some(value) = integer_attribute(element, name) {
                                *slot = value;
                            }
                        }
                    }
                }
                if tag == Tag::Math {
                    if let Some((block, true)) = owner {
                        blocks[block].equations += 1;
                    }
                    if sites.len() >= MAX_EQUATIONS {
                        return Err(Error::limit(
                            LimitKind::Elements,
                            sites.len().saturating_add(1),
                            MAX_EQUATIONS,
                        ));
                    }
                    let direct = parent_tag == Tag::Paragraph
                        && grandparent_tag == Tag::TextBody
                        && matches!(shape, Some((_, true)))
                        && stack
                            .len()
                            .checked_sub(3)
                            .is_some_and(|at| stack[at].tag == Tag::Shape);
                    let block = match owner {
                        Some((block, true)) if direct && blocks[block].level != Level::Inner => {
                            Some(block)
                        },
                        _ => None,
                    };
                    let (shape_id, name) = shape
                        .map(|(index, _)| shapes[index].clone())
                        .unwrap_or_default();
                    math = Some((sites.len(), stack.len()));
                    sites.push(Site {
                        element: start..start,
                        content: end..end,
                        text: String::new(),
                        shape_id,
                        name,
                        placement: None,
                        anchor,
                        block,
                        align: paragraph_align.clone(),
                    });
                }
                if !empty {
                    stack.push(Frame {
                        tag,
                        start,
                        block: block_index,
                        owner,
                        anchor,
                        shape,
                    });
                    if stack.len() > MAX_DEPTH {
                        return Err(Error::limit(LimitKind::XmlDepth, stack.len(), MAX_DEPTH));
                    }
                } else if tag == Tag::Stretch
                    && let (Some((block, false)), Some((_, true))) = (owner, shape)
                    && blocks[block].stretch.is_none()
                {
                    blocks[block].stretch = Some(start..end);
                }
            },
            Event::Text(_) | Event::GeneralRef(_) | Event::CData(_) => {
                let Some(frame) = stack.last().copied() else {
                    continue;
                };
                match frame.tag {
                    Tag::MathText => {
                        if let (Some((index, _)), Some(text)) = (math, decoded(&event)?) {
                            sites[index].text.push_str(&text);
                        }
                    },
                    Tag::MarkerField(slot, field) => {
                        if let (Some(index), Some(text)) = (frame.anchor, decoded(&event)?) {
                            anchors[index].markers[slot][field] = text.trim().parse().unwrap_or(0);
                        }
                    },
                    Tag::Text if math.is_none() => {
                        if let Some((block, true)) = frame.owner
                            && decoded(&event)?
                                .is_some_and(|text| text.chars().any(|ch| !ch.is_whitespace()))
                        {
                            blocks[block].other_content = true;
                        }
                    },
                    _ => {},
                }
            },
            Event::End(_) => {
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("unbalanced drawing XML"))?;
                match frame.tag {
                    Tag::Math => {
                        if let Some((index, depth)) = math
                            && depth == stack.len()
                        {
                            let site = &mut sites[index];
                            site.element = frame.start..end;
                            site.content = site.content.start..start;
                            math = None;
                        }
                    },
                    Tag::Block => blocks[frame.block].range = frame.start..end,
                    Tag::Choice => {
                        if let Some((block, _)) = frame.owner
                            && blocks[block].choice_tag.as_ref().map(|tag| tag.start)
                                == Some(frame.start)
                            && let Some(content) = blocks[block].choice_content.as_mut()
                        {
                            content.end = start;
                        }
                    },
                    Tag::Fallback => {
                        if let Some((block, _)) = frame.owner {
                            blocks[block].fallback = Some((frame.start..end, false));
                        }
                    },
                    Tag::Anchor => {
                        if let Some(index) = frame.anchor {
                            anchors[index].range = frame.start..end;
                        }
                    },
                    Tag::Marker(1) => {
                        if let Some(index) = frame.anchor {
                            anchors[index].to_marker = Some(frame.start..end);
                        }
                    },
                    Tag::Shape => {
                        if let (Some((block, true)), Some((_, true))) = (frame.owner, frame.shape)
                            && blocks[block].choice_shape.is_none()
                        {
                            blocks[block].choice_shape = Some(frame.start..end);
                        }
                    },
                    Tag::Stretch => {
                        if let (Some((block, false)), Some((_, true))) = (frame.owner, frame.shape)
                            && blocks[block].stretch.is_none()
                        {
                            blocks[block].stretch = Some(frame.start..end);
                        }
                    },
                    _ => {},
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !stack.is_empty() || math.is_some() {
        return Err(invalid("unbalanced drawing XML"));
    }
    for site in &mut sites {
        site.placement = site.anchor.and_then(|index| anchors[index].placement());
    }
    Ok((sites, blocks, anchors))
}

#[cfg(test)]
mod tests {
    use super::*;

    const ROOT: &str = concat!(
        r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" "#,
        r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
        r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" "#,
        r#"xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">"#
    );

    fn shape(id: u32, body: &str) -> String {
        format!(
            r#"<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="{id}" name="TextBox {id}"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="20"/></a:xfrm></xdr:spPr><xdr:txBody><a:bodyPr anchor="ctr" lIns="0"><a:spAutoFit/></a:bodyPr><a:p><a:pPr algn="r"/>{body}</a:p></xdr:txBody></xdr:sp>"#
        )
    }

    fn one_cell(col: u32, row: u32, object: &str) -> String {
        format!(
            r#"<xdr:oneCellAnchor><xdr:from><xdr:col>{col}</xdr:col><xdr:colOff>5</xdr:colOff><xdr:row>{row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="10" cy="20"/>{object}<xdr:clientData/></xdr:oneCellAnchor>"#
        )
    }

    const MATH_X: &str =
        r#"<a14:m><m:oMath><m:r><m:t>x&amp;</m:t></m:r><m:r><m:t>y</m:t></m:r></m:oMath></a14:m>"#;

    #[test]
    fn equations_are_found_at_every_level() {
        let a14 = r#" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main""#;
        // Around a whole anchor, around an anchor's shape, inside a
        // paragraph.
        let root = format!(
            r#"<mc:AlternateContent><mc:Choice{a14} Requires="a14">{}</mc:Choice><mc:Fallback>{}</mc:Fallback></mc:AlternateContent>"#,
            one_cell(1, 2, &shape(2, MATH_X)),
            one_cell(1, 2, &shape(2, "<a:r><a:t>x</a:t></a:r>")),
        );
        let anchor = one_cell(
            0,
            9,
            &format!(
                r#"<mc:AlternateContent><mc:Choice{a14} Requires="a14">{}</mc:Choice><mc:Fallback/></mc:AlternateContent>"#,
                shape(3, MATH_X)
            ),
        );
        let inner = one_cell(
            27,
            0,
            &shape(
                4,
                &format!(
                    r#"<a:r><a:t>t</a:t></a:r><mc:AlternateContent><mc:Choice{a14} Requires="a14">{MATH_X}</mc:Choice><mc:Fallback/></mc:AlternateContent>"#
                ),
            ),
        );
        let xml = format!("{ROOT}{root}{anchor}{inner}</xdr:wsDr>");
        let found = equations(xml.as_bytes()).unwrap();
        assert_eq!(found.len(), 3);
        assert_eq!(
            found.iter().map(Equation::shape_id).collect::<Vec<_>>(),
            [Some(2), Some(3), Some(4)]
        );
        assert_eq!(
            found.iter().map(Equation::cell).collect::<Vec<_>>(),
            [
                Some("B3".to_owned()),
                Some("A10".to_owned()),
                Some("AB1".to_owned())
            ]
        );
        assert_eq!(found[0].text(), "x&y");
        assert_eq!(found[0].name(), Some("TextBox 2"));
        assert!(
            found[0].omml().starts_with("<m:oMath>"),
            "{}",
            found[0].omml()
        );

        let (sites, blocks) = scan(xml.as_bytes()).unwrap();
        assert_eq!(blocks.len(), 3);
        assert_eq!(
            blocks.iter().map(|block| block.level).collect::<Vec<_>>(),
            [Level::Root, Level::Anchor, Level::Inner]
        );
        assert_eq!(sites[0].block, Some(0));
        assert_eq!(sites[1].block, Some(1));
        assert_eq!(sites[2].block, None);
        assert_eq!(sites[0].align.as_deref(), Some("r"));
        let root = &blocks[0];
        assert_eq!(root.equations, 1);
        assert!(!root.other_content);
        assert!(root.autofit);
        assert_eq!(root.text_anchor.as_deref(), Some("ctr"));
        assert_eq!(root.insets[0], 0);
        assert_eq!(root.choice_children, 1);
        assert!(root.choice_extent.is_some() && root.fallback_extent.is_some());
        // The two anchors of the root-level block come first.
        assert_eq!(blocks[1].outer_anchor, Some(2));
        assert_eq!(blocks[2].equations, 1);
    }

    #[test]
    fn malformed_and_hostile_drawings_are_refused() {
        assert!(equations(b"<xdr:wsDr>").is_err());
        let deep = format!(
            "{}{}",
            "<a>".repeat(MAX_DEPTH + 1),
            "</a>".repeat(MAX_DEPTH + 1)
        );
        assert!(matches!(
            scan(deep.as_bytes()).unwrap_err(),
            Error::Limit { .. }
        ));
        assert!(equations(b"<r/>").unwrap().is_empty());
    }

    #[test]
    fn cells_are_named_in_a1_form() {
        assert_eq!(a1(Marker::cell(0, 0)), "A1");
        assert_eq!(a1(Marker::cell(25, 4)), "Z5");
        assert_eq!(a1(Marker::cell(26, 0)), "AA1");
        assert_eq!(a1(Marker::cell(16_383, 1_048_575)), "XFD1048576");
    }
}
