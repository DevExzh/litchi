//! Where each equation of a slide sits in the slide's XML, for the edits.
//!
//! PowerPoint wraps a whole shape whose text holds an equation in
//! `mc:AlternateContent` ([MS-ODRAWXML] §2.3.3.41 for `a14:m`; ECMA-376
//! Part 3 for markup compatibility): `mc:Choice Requires="a14"` holds the
//! shape with `p:txBody/a:p/a14:m`, and `mc:Fallback` the same shape filled
//! with a picture of its text (`p:spPr/a:blipFill`). [`scan`] finds every
//! equation in the order [`super::equations`] lists them, with the byte
//! ranges an edit replaces: the OMML, the fallback's picture reference, the
//! frames of both branches and the picture's fill rectangle.

use std::ops::Range;

use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::{A14, MATH, MAX_DEPTH, MAX_NODES, MCE, PML, STRICT_PML, invalid};
use crate::{Error, Result};

const DML: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DML: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";

/// The default text-body insets in EMU: left, top, right and bottom
/// (ECMA-376 Part 1, §21.1.2.1.1 `bodyPr`).
pub(crate) const DEFAULT_INSETS: [i64; 4] = [91_440, 45_720, 91_440, 45_720];

/// One shape-level `mc:AlternateContent` block whose Choice holds an
/// equation.
#[derive(Debug, Clone, Default)]
pub(crate) struct Block {
    /// The whole block.
    pub(crate) range: Range<usize>,
    /// Whether the Choice branch is one `p:sp`.
    pub(crate) shape_level: bool,
    /// The equations in the Choice branch's shape.
    pub(crate) equations: usize,
    /// Whether the Choice branch's shape holds text other than whitespace
    /// besides its equations.
    pub(crate) other_content: bool,
    /// Whether the Choice branch's shape holds markup-compatibility content
    /// of its own.
    pub(crate) nested: bool,
    /// The `mc:Choice` branches of the block.
    pub(crate) choices: usize,
    /// The Choice branch's `p:sp`.
    pub(crate) choice_shape: Option<Range<usize>>,
    /// The `mc:Fallback` element, and whether it is empty (`<mc:Fallback/>`).
    pub(crate) fallback: Option<(Range<usize>, bool)>,
    /// The start tag of the `mc:Fallback` element.
    pub(crate) fallback_tag: Option<Range<usize>>,
    /// The Choice shape's text anchor (`a:bodyPr@anchor`).
    pub(crate) anchor: Option<String>,
    /// The `a:ext` start tag of the Choice shape's frame.
    pub(crate) choice_extent: Option<Range<usize>>,
    /// The `a:ext` start tag of the Fallback shape's frame.
    pub(crate) fallback_extent: Option<Range<usize>>,
    /// Whether the Choice shape resizes to its text (`a:spAutoFit`).
    pub(crate) autofit: bool,
    /// The Choice shape's text insets in EMU: left, top, right, bottom.
    pub(crate) insets: [i64; 4],
    /// The `a:blip` start tag of the Fallback shape's picture fill.
    pub(crate) blip: Option<Range<usize>>,
    /// The whole `a:stretch` element of the Fallback shape's picture fill.
    pub(crate) stretch: Option<Range<usize>>,
}

/// One equation (`a14:m`) of a slide.
#[derive(Debug, Clone)]
pub(crate) struct Site {
    /// The whole `a14:m` element.
    pub(crate) element: Range<usize>,
    /// The element's content: the OMML.
    pub(crate) content: Range<usize>,
    /// The shape-level block whose Choice shape holds the equation directly
    /// in a paragraph of its text body; `None` for any other placement.
    pub(crate) block: Option<usize>,
    /// The alignment of its paragraph (`a:pPr@algn`), when given before it.
    pub(crate) align: Option<String>,
}

/// What an open element is, as far as the edits care.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Tag {
    Block,
    Choice,
    Fallback,
    Shape,
    ShapeProperties,
    Transform,
    TextBody,
    BodyProperties,
    Paragraph,
    BlipFill,
    Stretch,
    Math,
    Text,
    Other,
}

/// One open element: what it is, where it starts, the block it is (for
/// [`Tag::Block`]), and the block branch (`(block, is_choice)`) whose shape
/// it is inside.
#[derive(Debug, Clone, Copy)]
struct Frame {
    tag: Tag,
    start: usize,
    block: usize,
    owner: Option<(usize, bool)>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Space {
    Mce,
    A14,
    Math,
    Pml,
    Dml,
    Other,
}

fn space(namespace: &ResolveResult<'_>) -> Space {
    match namespace {
        ResolveResult::Bound(Namespace(value)) => match *value {
            MCE => Space::Mce,
            A14 => Space::A14,
            MATH => Space::Math,
            PML | STRICT_PML => Space::Pml,
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
    element
        .checked_attributes()
        .flatten()
        .find(|attribute| attribute.key.as_ref() == name)
        .and_then(|attribute| {
            std::str::from_utf8(&attribute.value)
                .ok()
                .and_then(|value| value.trim().parse().ok())
        })
}

/// Every equation of a slide's XML, in the order of
/// [`super::equations`], and the shape-level blocks they sit in.
///
/// # Errors
///
/// Returns an error for malformed XML or exceeded bounds.
pub(crate) fn scan(xml: &[u8]) -> Result<(Vec<Site>, Vec<Block>)> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    let position = |reader: &NsReader<&[u8]>| {
        usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid("the slide is too large"))
    };
    let mut stack: Vec<Frame> = Vec::new();
    let mut sites = Vec::new();
    let mut blocks: Vec<Block> = Vec::new();
    // The open equation: its index in `sites` and its depth.
    let mut math: Option<(usize, usize)> = None;
    let mut paragraph_align: Option<String> = None;
    let mut nodes = 0usize;
    loop {
        let start = position(&reader)?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let space = space(&namespace);
        let end = position(&reader)?;
        nodes += 1;
        if nodes > MAX_NODES {
            return Err(Error::Limit {
                kind: crate::opened::LimitKind::Codec(crate::opened::CodecLimitKind::XmlNodes),
                observed: nodes,
                maximum: MAX_NODES,
            });
        }
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let empty = matches!(event, Event::Empty(_));
                let local = element.local_name();
                let local = local.as_ref();
                let parent = stack.last().copied();
                let parent_tag = parent.map_or(Tag::Other, |frame| frame.tag);
                let mut owner = parent.and_then(|frame| frame.owner);
                let tag = match (space, local, parent_tag) {
                    (Space::Mce, b"AlternateContent", _) => {
                        if let Some((block, true)) = owner {
                            blocks[block].nested = true;
                        }
                        Tag::Block
                    },
                    (Space::Mce, b"Choice", Tag::Block) => Tag::Choice,
                    (Space::Mce, b"Fallback", Tag::Block) => Tag::Fallback,
                    (Space::Pml, b"sp", Tag::Choice | Tag::Fallback) => Tag::Shape,
                    (Space::Pml, b"spPr", Tag::Shape) => Tag::ShapeProperties,
                    (Space::Dml, b"xfrm", Tag::ShapeProperties) => Tag::Transform,
                    (Space::Pml, b"txBody", Tag::Shape) => Tag::TextBody,
                    (Space::Dml, b"bodyPr", Tag::TextBody) => Tag::BodyProperties,
                    (Space::Dml, b"p", Tag::TextBody) => Tag::Paragraph,
                    (Space::Dml, b"blipFill", Tag::ShapeProperties) => Tag::BlipFill,
                    (Space::Dml, b"stretch", Tag::BlipFill) => Tag::Stretch,
                    (Space::A14, b"m", _) if math.is_none() && !empty => Tag::Math,
                    (Space::Dml, b"t", _) => Tag::Text,
                    _ => Tag::Other,
                };
                let mut block_index = 0;
                match tag {
                    Tag::Block => {
                        block_index = blocks.len();
                        blocks.push(Block {
                            insets: DEFAULT_INSETS,
                            ..Block::default()
                        });
                        owner = None;
                    },
                    Tag::Choice | Tag::Fallback => {
                        let block = parent.map_or(0, |frame| frame.block);
                        if tag == Tag::Choice {
                            blocks[block].choices += 1;
                        } else {
                            blocks[block].fallback = Some((start..end, empty));
                            blocks[block].fallback_tag = Some(start..end);
                        }
                        owner = None;
                    },
                    Tag::Paragraph => paragraph_align = None,
                    Tag::Shape => {
                        // A branch's parent is its block.
                        let block = stack
                            .iter()
                            .rev()
                            .nth(1)
                            .filter(|frame| frame.tag == Tag::Block)
                            .map(|frame| frame.block)
                            .ok_or_else(|| invalid("unbalanced slide XML"))?;
                        let choice = parent_tag == Tag::Choice;
                        if choice {
                            blocks[block].shape_level = true;
                        }
                        owner = Some((block, choice));
                    },
                    _ => {},
                }
                let in_shape = |depth_from_shape: usize| {
                    stack.len() >= depth_from_shape
                        && stack[stack.len() - depth_from_shape].tag == Tag::Shape
                };
                if let Some((block, choice)) = owner {
                    let record = &mut blocks[block];
                    match (space, local, parent_tag) {
                        (Space::Dml, b"ext", Tag::Transform) if in_shape(3) => {
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
                        _ => {},
                    }
                    if parent_tag == Tag::Paragraph
                        && space == Space::Dml
                        && local == b"pPr"
                        && choice
                    {
                        paragraph_align = text_attribute(element, b"algn");
                    }
                    if tag == Tag::BodyProperties && choice {
                        record.anchor = text_attribute(element, b"anchor");
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
                    if tag == Tag::Math && choice {
                        record.equations += 1;
                    }
                }
                if tag == Tag::Math {
                    let direct = stack.len() >= 5
                        && stack[stack.len() - 1].tag == Tag::Paragraph
                        && stack[stack.len() - 2].tag == Tag::TextBody
                        && stack[stack.len() - 3].tag == Tag::Shape
                        && stack[stack.len() - 4].tag == Tag::Choice
                        && stack[stack.len() - 5].tag == Tag::Block;
                    let block = match owner {
                        Some((block, true)) if direct => Some(block),
                        _ => None,
                    };
                    math = Some((sites.len(), stack.len()));
                    sites.push(Site {
                        element: start..start,
                        content: end..end,
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
                    });
                    if stack.len() > MAX_DEPTH {
                        return Err(invalid("the slide is nested too deeply"));
                    }
                } else if tag == Tag::Stretch
                    && let Some((block, false)) = owner
                    && blocks[block].stretch.is_none()
                {
                    blocks[block].stretch = Some(start..end);
                }
            },
            Event::Text(ref text) => {
                if let Some(frame) = stack.last()
                    && frame.tag == Tag::Text
                    && math.is_none()
                    && let Some((block, true)) = frame.owner
                    && text.iter().any(|byte| !byte.is_ascii_whitespace())
                {
                    blocks[block].other_content = true;
                }
            },
            Event::GeneralRef(_) | Event::CData(_) => {
                if let Some(frame) = stack.last()
                    && frame.tag == Tag::Text
                    && math.is_none()
                    && let Some((block, true)) = frame.owner
                {
                    blocks[block].other_content = true;
                }
            },
            Event::End(_) => {
                let frame = stack.pop().ok_or_else(|| invalid("unbalanced slide XML"))?;
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
                    Tag::Shape => {
                        if let Some((block, true)) = frame.owner
                            && blocks[block].choice_shape.is_none()
                        {
                            blocks[block].choice_shape = Some(frame.start..end);
                        }
                    },
                    Tag::Fallback => {
                        if let Some(block) = stack.last().map(|parent| parent.block) {
                            blocks[block].fallback = Some((frame.start..end, false));
                        }
                    },
                    Tag::Stretch => {
                        if let Some((block, false)) = frame.owner
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
    if !stack.is_empty() {
        return Err(invalid("unbalanced slide XML"));
    }
    Ok((sites, blocks))
}

/// A `p:sp` of the slide outside any markup-compatibility block whose
/// `cNvPr` ID is `id`: its whole element.
///
/// # Errors
///
/// Returns an error for malformed XML.
pub(crate) fn plain_shape(xml: &[u8], id: u32) -> Result<Option<Range<usize>>> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_end_names = true;
    // The open elements: whether each is a block, and for an open shape its
    // depth, start and whether its ID matched.
    let mut depth = 0usize;
    let mut blocks = 0usize;
    let mut block_depths: Vec<usize> = Vec::new();
    let mut shape: Option<(usize, usize, Option<bool>)> = None;
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid("the slide is too large"))?;
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let space = space(&namespace);
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let local = element.local_name();
                if space == Space::Mce && local.as_ref() == b"AlternateContent" {
                    blocks += 1;
                    if matches!(event, Event::Start(_)) {
                        block_depths.push(depth);
                    } else {
                        blocks -= 1;
                    }
                } else if blocks == 0
                    && space == Space::Pml
                    && local.as_ref() == b"sp"
                    && shape.is_none()
                    && matches!(event, Event::Start(_))
                {
                    shape = Some((depth, start, None));
                } else if let Some((_, _, matched @ None)) = shape.as_mut()
                    && space == Space::Pml
                    && local.as_ref() == b"cNvPr"
                {
                    *matched = Some(integer_attribute(element, b"id") == Some(i64::from(id)));
                }
                if matches!(event, Event::Start(_)) {
                    depth += 1;
                    if depth > MAX_DEPTH {
                        return Err(invalid("the slide is nested too deeply"));
                    }
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("unbalanced slide XML"))?;
                if block_depths.last() == Some(&depth) {
                    block_depths.pop();
                    blocks -= 1;
                }
                if let Some((at, shape_start, matched)) = shape
                    && at == depth
                {
                    if matched == Some(true) {
                        let end = usize::try_from(reader.buffer_position())
                            .map_err(|_error| invalid("the slide is too large"))?;
                        return Ok(Some(shape_start..end));
                    }
                    shape = None;
                }
            },
            Event::Eof => return Ok(None),
            _ => {},
        }
    }
}

/// Where new content of a shape's last paragraph goes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum TextPlace {
    /// Before this offset: the paragraph's `a:endParaRPr`, or its end tag.
    Before(usize),
    /// The paragraph is the empty element at this range (`<a:p/>`).
    Empty(Range<usize>),
}

/// The end of a shape's text: where the last paragraph of its text body
/// takes new content, and the text size (`sz`, hundredths of a point) of
/// that paragraph's last run or end properties.
#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct TextEnd {
    pub(crate) place: TextPlace,
    pub(crate) size: Option<u32>,
}

/// The end of the text of a `p:sp` fragment; `None` for a shape without a
/// text paragraph.
///
/// # Errors
///
/// Returns an error for malformed XML.
pub(crate) fn text_end(shape: &[u8]) -> Result<Option<TextEnd>> {
    let mut reader = quick_xml::Reader::from_reader(shape);
    let mut path: Vec<Vec<u8>> = Vec::new();
    // The open last paragraph's `a:endParaRPr` and text size, and the
    // answer so far.
    let mut end_properties: Option<usize> = None;
    let mut size: Option<u32> = None;
    let mut found = None;
    let size_of = |element: &BytesStart<'_>| -> Option<u32> {
        integer_attribute(element, b"sz").and_then(|value| u32::try_from(value).ok())
    };
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid("the shape is too large"))?;
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let end = usize::try_from(reader.buffer_position())
            .map_err(|_error| invalid("the shape is too large"))?;
        let in_body = path.len() == 2 && path[1].as_slice() == b"txBody";
        let in_paragraph =
            path.len() == 3 && path[1].as_slice() == b"txBody" && path[2].as_slice() == b"p";
        let in_run = path.len() == 4
            && path[1].as_slice() == b"txBody"
            && path[2].as_slice() == b"p"
            && path[3].as_slice() == b"r";
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let local = element.local_name().as_ref().to_vec();
                let empty = matches!(event, Event::Empty(_));
                if in_body && local == b"p" {
                    size = None;
                    end_properties = None;
                    if empty {
                        found = Some(TextEnd {
                            place: TextPlace::Empty(start..end),
                            size: None,
                        });
                    }
                }
                if in_paragraph && local == b"endParaRPr" {
                    end_properties = Some(start);
                    size = size_of(element).or(size);
                }
                if in_run && local == b"rPr" {
                    size = size_of(element).or(size);
                }
                if !empty {
                    path.push(local);
                }
            },
            Event::End(_) => {
                if in_paragraph {
                    found = Some(TextEnd {
                        place: TextPlace::Before(end_properties.take().unwrap_or(start)),
                        size,
                    });
                }
                path.pop();
            },
            Event::Eof => return Ok(found),
            _ => {},
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const ROOT: &str = r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="r" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><p:cSld><p:spTree>"#;
    const END: &str = "</p:spTree></p:cSld></p:sld>";
    const MATH: &str = r#"<a14:m><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r></m:oMath></a14:m>"#;

    fn slide(body: &str) -> String {
        format!("{ROOT}{body}{END}")
    }

    #[test]
    fn a_shape_level_block_records_its_edit_ranges() {
        let xml = slide(&format!(
            concat!(
                r#"<mc:AlternateContent><mc:Choice Requires="a14"><p:sp><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm></p:spPr>"#,
                r#"<p:txBody><a:bodyPr lIns="10" bIns="20"><a:spAutoFit/></a:bodyPr><a:p>{}</a:p></p:txBody></p:sp></mc:Choice>"#,
                r#"<mc:Fallback><p:sp><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:blipFill><a:blip r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr>"#,
                r#"<p:txBody><a:p><a:r><a:t> </a:t></a:r></a:p></p:txBody></p:sp></mc:Fallback></mc:AlternateContent>"#
            ),
            MATH
        ));
        let bytes = xml.as_bytes();
        let (sites, blocks) = scan(bytes).unwrap();
        assert_eq!(sites.len(), super::super::equations(bytes).unwrap().len());
        let site = &sites[0];
        assert_eq!(&xml[site.element.clone()], MATH);
        assert!(xml[site.content.clone()].starts_with("<m:oMath "));
        let block = &blocks[site.block.unwrap()];
        assert!(block.shape_level && block.autofit && !block.other_content);
        assert_eq!(block.equations, 1);
        assert_eq!(block.insets, [10, 45_720, 91_440, 20]);
        assert_eq!(
            &xml[block.choice_extent.clone().unwrap()],
            r#"<a:ext cx="3" cy="4"/>"#
        );
        assert_eq!(
            &xml[block.fallback_extent.clone().unwrap()],
            r#"<a:ext cx="3" cy="4"/>"#
        );
        assert_eq!(
            &xml[block.blip.clone().unwrap()],
            r#"<a:blip r:embed="rId9"/>"#
        );
        assert_eq!(
            &xml[block.stretch.clone().unwrap()],
            "<a:stretch><a:fillRect/></a:stretch>"
        );
        assert!(xml[block.range.clone()].starts_with("<mc:AlternateContent>"));
        assert!(xml[block.range.clone()].ends_with("</mc:AlternateContent>"));
    }

    #[test]
    fn other_placements_are_found_but_not_editable() {
        // A block inside a paragraph, and an equation outside any block.
        let inline = format!(
            r#"<p:sp><p:txBody><a:p><mc:AlternateContent><mc:Choice Requires="a14">{MATH}</mc:Choice><mc:Fallback><a:r><a:t>x</a:t></a:r></mc:Fallback></mc:AlternateContent></a:p><a:p>{MATH}</a:p></p:txBody></p:sp>"#
        );
        let xml = slide(&inline);
        let (sites, _) = scan(xml.as_bytes()).unwrap();
        assert_eq!(sites.len(), 2);
        assert!(sites.iter().all(|site| site.block.is_none()));
        // Text or a nested block beside the equation marks the shape.
        for extra in [
            "<a:r><a:t>where</a:t></a:r>".to_owned(),
            format!(
                "<mc:AlternateContent><mc:Choice Requires=\"a14\">{MATH}</mc:Choice></mc:AlternateContent>"
            ),
        ] {
            let xml = slide(&format!(
                r#"<mc:AlternateContent><mc:Choice Requires="a14"><p:sp><p:txBody><a:p>{extra}{MATH}</a:p></p:txBody></p:sp></mc:Choice></mc:AlternateContent>"#
            ));
            let (sites, blocks) = scan(xml.as_bytes()).unwrap();
            let site = sites.last().unwrap();
            let block = &blocks[site.block.unwrap()];
            assert!(block.other_content || block.nested, "{extra}");
            assert_eq!(block.nested, extra.contains("AlternateContent"), "{extra}");
        }
        assert!(scan(b"<p:sld><a14:m>").is_err());
    }
}
