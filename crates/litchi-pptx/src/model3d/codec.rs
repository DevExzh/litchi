//! Bounded `PresentationML` graphic-frame and shared model3d wire codecs.

use std::{borrow::Cow, ops::Range};

use litchi_core::xml::ReaderOrigin;
use litchi_drawingml::model3d as drawing;
use litchi_drawingml::{Error as DrawingError, LimitKind};
use litchi_ooxml_common::xml::unqualified_attribute_value;
use quick_xml::events::Event;
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::model::Scene;
use super::{
    GRAPHIC_URI, MAX_MODELS_PER_SLIDE, MAX_SHAPES_PER_SLIDE, MAX_SLIDE_XML_BYTES, MAX_XML_DEPTH,
};
use crate::{Error, Result};

const PML: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const DML: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DML: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/main";

const SHAPES: &[&[u8]] = &[
    b"sp",
    b"pic",
    b"graphicFrame",
    b"grpSp",
    b"cxnSp",
    b"contentPart",
];

/// Caller-selected bounds for one PPTX model3d XML operation.
///
/// The shared DrawingML codec owns the model3d grammar.  This adapter keeps a
/// distinct PPTX policy type so host callers cannot accidentally pass an
/// unrelated format's limits, while preserving the same byte and tree
/// dimensions at the wire boundary.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[must_use]
pub struct Limits {
    max_xml_bytes: usize,
    max_fragment_bytes: usize,
    max_output_bytes: usize,
    max_children: usize,
    max_raster_children: usize,
    max_depth: usize,
    max_nodes: usize,
}

impl Limits {
    /// The historical PPTX model3d bounds.
    pub const DEFAULT: Self = {
        let shared = drawing::codec::Limits::DEFAULT;
        Self {
            max_xml_bytes: shared.max_xml_bytes,
            max_fragment_bytes: shared.max_fragment_bytes,
            max_output_bytes: shared.max_output_bytes,
            max_children: shared.max_children,
            max_raster_children: shared.max_raster_children,
            max_depth: shared.max_depth,
            max_nodes: shared.max_nodes,
        }
    };

    /// Construct a policy with the historical tree ceilings and selected byte
    /// budgets.
    #[must_use]
    pub const fn new(
        max_xml_bytes: usize,
        max_fragment_bytes: usize,
        max_output_bytes: usize,
    ) -> Option<Self> {
        Self::with_tree(
            max_xml_bytes,
            max_fragment_bytes,
            max_output_bytes,
            Self::DEFAULT.max_children,
            Self::DEFAULT.max_raster_children,
            Self::DEFAULT.max_depth,
            Self::DEFAULT.max_nodes,
        )
    }

    /// Construct a policy with explicit byte and tree ceilings.
    #[must_use]
    pub const fn with_tree(
        max_xml_bytes: usize,
        max_fragment_bytes: usize,
        max_output_bytes: usize,
        max_children: usize,
        max_raster_children: usize,
        max_depth: usize,
        max_nodes: usize,
    ) -> Option<Self> {
        let default = Self::DEFAULT;
        if max_xml_bytes == 0
            || max_xml_bytes > default.max_xml_bytes
            || max_fragment_bytes == 0
            || max_fragment_bytes > default.max_fragment_bytes
            || max_output_bytes == 0
            || max_output_bytes > default.max_output_bytes
            || max_children == 0
            || max_children > default.max_children
            || max_raster_children == 0
            || max_raster_children > default.max_raster_children
            || max_depth == 0
            || max_depth > default.max_depth
            || max_nodes == 0
            || max_nodes > default.max_nodes
        {
            return None;
        }
        Some(Self {
            max_xml_bytes,
            max_fragment_bytes,
            max_output_bytes,
            max_children,
            max_raster_children,
            max_depth,
            max_nodes,
        })
    }

    #[inline]
    #[must_use]
    pub const fn max_xml_bytes(self) -> usize {
        self.max_xml_bytes
    }

    #[inline]
    #[must_use]
    pub const fn max_fragment_bytes(self) -> usize {
        self.max_fragment_bytes
    }

    #[inline]
    #[must_use]
    pub const fn max_output_bytes(self) -> usize {
        self.max_output_bytes
    }

    #[inline]
    #[must_use]
    pub const fn max_children(self) -> usize {
        self.max_children
    }

    #[inline]
    #[must_use]
    pub const fn max_raster_children(self) -> usize {
        self.max_raster_children
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

    fn drawing(self) -> drawing::codec::Limits {
        drawing::codec::Limits {
            max_xml_bytes: self.max_xml_bytes,
            max_fragment_bytes: self.max_fragment_bytes,
            max_output_bytes: self.max_output_bytes,
            max_children: self.max_children,
            max_raster_children: self.max_raster_children,
            max_depth: self.max_depth,
            max_nodes: self.max_nodes,
        }
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// A checked model3d element location and its semantic shape anchor.
#[derive(Debug, Clone)]
pub(crate) struct Location {
    pub(crate) range: Range<usize>,
    pub(crate) shape_index: usize,
    pub(crate) shape_name: Option<Box<str>>,
}

/// A bounded slide inventory used by both reads and snapshot edits.
#[derive(Debug)]
pub(crate) struct Inventory {
    pub(crate) locations: Vec<Location>,
    pub(crate) shape_names: Vec<Option<Box<str>>>,
}

/// Read one shared model3d scene without exposing relationship IDs in the
/// PPTX semantic model.
///
/// # Errors
///
/// Returns an error if the input cannot be read or is malformed.
pub fn read(xml: &[u8], limits: Limits) -> Result<Scene> {
    let prepared = prepare_fragment(xml, limits)?;
    Ok(Scene::from_wire(
        drawing::codec::read(&prepared, limits.drawing()).map_err(super::model::drawing_error)?,
    ))
}

/// Write one shared model3d scene using the common deterministic codec.
///
/// # Errors
///
/// Returns an error if the output cannot be encoded or written.
pub fn write(scene: &Scene, limits: Limits) -> Result<Vec<u8>> {
    // The shared writer supplies missing model and relationship namespace
    // declarations itself. Passing the borrowed wire directly keeps its
    // structural/output preflight ahead of any clone of retained opaque XML.
    drawing::codec::write(&scene.wire, limits.drawing()).map_err(super::model::drawing_error)
}

pub(crate) fn locate(xml: &[u8]) -> Result<Inventory> {
    if xml.len() > MAX_SLIDE_XML_BYTES {
        return Err(limit("model3d slide XML bytes", MAX_SLIDE_XML_BYTES));
    }

    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut shape_stack = Vec::new();
    let mut shape_names = Vec::new();
    let mut graphic_data_depth = None;
    let mut model_stack: Vec<(usize, usize, usize)> = Vec::new();
    let mut locations = Vec::new();
    let mut nodes = 0usize;

    loop {
        let start = position(&reader, origin)?;
        let decoder = reader.decoder();
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(xml_error)?
            .into_owned();
        let end = position(&reader, origin)?;
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        nodes = nodes
            .checked_add(1)
            .ok_or_else(|| limit("model3d slide XML nodes", MAX_SHAPES_PER_SLIDE))?;
        if nodes > MAX_SHAPES_PER_SLIDE.saturating_mul(32) {
            return Err(limit(
                "model3d slide XML nodes",
                MAX_SHAPES_PER_SLIDE.saturating_mul(32),
            ));
        }

        match event {
            Event::Start(element) => {
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("model3d XML depth", MAX_XML_DEPTH))?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit("model3d XML depth", MAX_XML_DEPTH));
                }

                if is_shape(&namespace, element.name()) {
                    if shape_names.len() >= MAX_SHAPES_PER_SLIDE {
                        return Err(limit("model3d slide shapes", MAX_SHAPES_PER_SLIDE));
                    }
                    let index = shape_names.len();
                    shape_names.push(None);
                    shape_stack.push((index, depth));
                }
                if is_graphic_data(&namespace, element.name())
                    && unqualified_attribute_value(&element, b"uri", decoder)?.as_deref()
                        == Some(GRAPHIC_URI)
                {
                    if graphic_data_depth.is_some() {
                        return Err(invalid("nested model3d graphicData"));
                    }
                    graphic_data_depth = Some(depth);
                }
                if is_cnvpr(&namespace, element.name())
                    && let Some((index, _)) = shape_stack.last().copied()
                {
                    shape_names[index] =
                        unqualified_attribute_value(&element, b"name", decoder)?.map(Into::into);
                }
                if is_model(&namespace, element.name()) {
                    let Some(graphic_depth) = graphic_data_depth else {
                        return Err(invalid("model3d element is outside model3d graphicData"));
                    };
                    if graphic_depth.saturating_add(1) != depth {
                        return Err(invalid("model3d element is not a graphicData child"));
                    }
                    let Some((shape_index, _)) = shape_stack.last().copied() else {
                        return Err(invalid("model3d element has no shape anchor"));
                    };
                    if locations.len() >= MAX_MODELS_PER_SLIDE {
                        return Err(limit("model3d instances", MAX_MODELS_PER_SLIDE));
                    }
                    model_stack.push((start, depth, shape_index));
                }
            },
            Event::Empty(element) => {
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("model3d XML depth", MAX_XML_DEPTH))?;
                if child_depth > MAX_XML_DEPTH {
                    return Err(limit("model3d XML depth", MAX_XML_DEPTH));
                }
                if is_cnvpr(&namespace, element.name())
                    && let Some((index, _)) = shape_stack.last().copied()
                {
                    shape_names[index] =
                        unqualified_attribute_value(&element, b"name", decoder)?.map(Into::into);
                }
                if is_shape(&namespace, element.name()) {
                    if shape_names.len() >= MAX_SHAPES_PER_SLIDE {
                        return Err(limit("model3d slide shapes", MAX_SHAPES_PER_SLIDE));
                    }
                    shape_names.push(None);
                }
                if is_model(&namespace, element.name()) {
                    let Some(graphic_depth) = graphic_data_depth else {
                        return Err(invalid("model3d element is outside model3d graphicData"));
                    };
                    if graphic_depth.saturating_add(1) != child_depth {
                        return Err(invalid("model3d element is not a graphicData child"));
                    }
                    let Some((shape_index, _)) = shape_stack.last().copied() else {
                        return Err(invalid("model3d element has no shape anchor"));
                    };
                    locations.push(Location {
                        range: start..end,
                        shape_index,
                        shape_name: shape_names[shape_index].clone(),
                    });
                }
            },
            Event::End(_) => {
                if let Some((model_start, model_depth, shape_index)) = model_stack.last().copied()
                    && model_depth == depth
                {
                    model_stack.pop();
                    locations.push(Location {
                        range: model_start..end,
                        shape_index,
                        shape_name: shape_names[shape_index].clone(),
                    });
                }
                if let Some((_, shape_depth)) = shape_stack.last().copied()
                    && shape_depth == depth
                {
                    shape_stack.pop();
                }
                if graphic_data_depth == Some(depth) {
                    graphic_data_depth = None;
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("model3d XML depth underflow"))?;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("model3d slide XML contains document markup"));
            },
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }

    if depth != 0 || !shape_stack.is_empty() || !model_stack.is_empty() {
        return Err(invalid("model3d slide XML is unterminated"));
    }
    Ok(Inventory {
        locations,
        shape_names,
    })
}

pub(crate) fn replace(xml: &[u8], range: Range<usize>, replacement: &[u8]) -> Result<Vec<u8>> {
    if range.start > range.end || range.end > xml.len() {
        return Err(invalid("model3d replacement range is outside the slide"));
    }
    let new_len = xml
        .len()
        .saturating_sub(range.end.saturating_sub(range.start))
        .saturating_add(replacement.len());
    if new_len > MAX_SLIDE_XML_BYTES {
        return Err(limit(
            "model3d updated slide XML bytes",
            MAX_SLIDE_XML_BYTES,
        ));
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(new_len)
        .map_err(|source| Error::Allocation {
            resource: "model3d updated slide XML",
            source,
        })?;
    output.extend_from_slice(&xml[..range.start]);
    output.extend_from_slice(replacement);
    output.extend_from_slice(&xml[range.end..]);
    Ok(output)
}

fn is_shape(namespace: &ResolveResult<'_>, name: quick_xml::name::QName<'_>) -> bool {
    (is_namespace(namespace, PML) || is_namespace(namespace, STRICT_PML))
        && SHAPES
            .iter()
            .any(|value| *value == name.local_name().as_ref())
}

fn is_cnvpr(namespace: &ResolveResult<'_>, name: quick_xml::name::QName<'_>) -> bool {
    (is_namespace(namespace, PML) || is_namespace(namespace, STRICT_PML))
        && name.local_name().as_ref() == b"cNvPr"
}

fn is_graphic_data(namespace: &ResolveResult<'_>, name: quick_xml::name::QName<'_>) -> bool {
    (is_namespace(namespace, DML) || is_namespace(namespace, STRICT_DML))
        && name.local_name().as_ref() == b"graphicData"
}

fn is_model(namespace: &ResolveResult<'_>, name: quick_xml::name::QName<'_>) -> bool {
    is_namespace(namespace, GRAPHIC_URI.as_bytes()) && name.local_name().as_ref() == b"model3d"
}

fn is_namespace(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected)
}

fn position(reader: &NsReader<&[u8]>, origin: ReaderOrigin) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .ok_or_else(|| invalid("model3d XML offset exceeds usize"))
}

fn xml_error(error: quick_xml::Error) -> Error {
    Error::Xml(error.to_string())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(format!("PPTX model3d {}", message.into()))
}

fn limit(resource: &'static str, limit: usize) -> Error {
    Error::Limit { resource, limit }
}

fn drawing_limit(kind: LimitKind, observed: usize, maximum: usize) -> Error {
    Error::Drawing(DrawingError::Limit {
        kind,
        observed,
        maximum,
    })
}

/// A model element is usually namespace-contextualized by its slide root.
/// The shared codec intentionally accepts one self-contained fragment, so a
/// package adapter supplies the inherited model and relationship declarations
/// when the selected element did not repeat them locally.
fn prepare_fragment<'a>(xml: &'a [u8], limits: Limits) -> Result<Cow<'a, [u8]>> {
    if xml.len() > limits.max_xml_bytes {
        return Err(drawing_limit(
            LimitKind::XmlBytes,
            xml.len(),
            limits.max_xml_bytes,
        ));
    }
    let open_end = opening_tag_end(xml)?;
    let opening = &xml[..open_end];
    let model_prefix =
        element_prefix(opening).ok_or_else(|| invalid("model3d root has no prefix"))?;
    let mut declarations = Vec::new();
    if !has_namespace_declaration(opening, model_prefix) {
        declarations.extend_from_slice(b" xmlns:");
        declarations.extend_from_slice(model_prefix);
        declarations.extend_from_slice(b"=\"");
        declarations.extend_from_slice(GRAPHIC_URI.as_bytes());
        declarations.push(b'\"');
    }
    for local in [b"embed".as_slice(), b"link".as_slice()] {
        let Some(prefix) = attribute_prefix(opening, local) else {
            continue;
        };
        if !has_namespace_declaration(opening, prefix) {
            declarations.extend_from_slice(b" xmlns:");
            declarations.extend_from_slice(prefix);
            declarations.extend_from_slice(b"=\"");
            declarations
                .extend_from_slice(litchi_drawingml::model3d::RELATIONSHIP_NAMESPACE.as_bytes());
            declarations.push(b'\"');
        }
    }
    if declarations.is_empty() {
        return Ok(Cow::Borrowed(xml));
    }
    let prepared_len =
        xml.len()
            .checked_add(declarations.len())
            .ok_or(Error::CapacityOverflow {
                resource: "model3d prepared XML",
            })?;
    if prepared_len > limits.max_xml_bytes {
        return Err(drawing_limit(
            LimitKind::XmlBytes,
            prepared_len,
            limits.max_xml_bytes,
        ));
    }
    let mut prepared = Vec::new();
    prepared
        .try_reserve_exact(prepared_len)
        .map_err(|source| Error::Allocation {
            resource: "model3d prepared XML",
            source,
        })?;
    prepared.extend_from_slice(&xml[..open_end - 1]);
    prepared.extend_from_slice(&declarations);
    prepared.push(b'>');
    prepared.extend_from_slice(&xml[open_end..]);
    Ok(Cow::Owned(prepared))
}

fn opening_tag_end(xml: &[u8]) -> Result<usize> {
    let mut quote = None;
    for (index, byte) in xml.iter().copied().enumerate() {
        match (quote, byte) {
            (None, b'\'' | b'\"') => quote = Some(byte),
            (Some(value), byte) if value == byte => quote = None,
            (None, b'>') => return Ok(index + 1),
            _ => {},
        }
    }
    Err(invalid("model3d root opening tag is unterminated"))
}

fn element_prefix(opening: &[u8]) -> Option<&[u8]> {
    let start = opening
        .iter()
        .position(|byte| *byte == b'<')?
        .saturating_add(1);
    let end = opening[start..]
        .iter()
        .position(|byte| byte.is_ascii_whitespace() || matches!(byte, b'>' | b'/'))?
        .saturating_add(start);
    let name = &opening[start..end];
    let colon = name.iter().position(|byte| *byte == b':')?;
    Some(&name[..colon])
}

fn attribute_prefix<'a>(opening: &'a [u8], local: &[u8]) -> Option<&'a [u8]> {
    let mut offset = 0usize;
    while let Some(relative) = opening[offset..]
        .windows(local.len() + 2)
        .position(|window| {
            window.ends_with(local)
                || window[window.len().saturating_sub(local.len() + 1)..].ends_with(local)
        })
    {
        let index = offset + relative;
        let colon = opening[..index].iter().rposition(|byte| *byte == b':')?;
        let prefix_start = opening[..colon]
            .iter()
            .rposition(u8::is_ascii_whitespace)
            .map_or(1, |value| value.saturating_add(1));
        if prefix_start < colon
            && opening[index..].starts_with(local)
            && opening.get(index.saturating_sub(1)) == Some(&b':')
        {
            return Some(&opening[prefix_start..colon]);
        }
        offset = index.saturating_add(local.len());
    }
    None
}

fn has_namespace_declaration(opening: &[u8], prefix: &[u8]) -> bool {
    if prefix.is_empty() {
        return opening
            .windows(b"xmlns=\"".len())
            .any(|value| value == b"xmlns=\"");
    }
    let mut needle = Vec::with_capacity(7 + prefix.len());
    needle.extend_from_slice(b"xmlns:");
    needle.extend_from_slice(prefix);
    needle.push(b'=');
    opening.windows(needle.len()).any(|value| value == needle)
}
