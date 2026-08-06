//! Bounded, namespace-aware scene indexing.

use std::{borrow::Cow, str};

use litchi_ooxml_common::mce::{Capabilities, process_markup_compatibility};
use litchi_ooxml_common::xml::{
    DRAWINGML_CHART_NAMESPACE, DRAWINGML_NAMESPACE, STRICT_DRAWINGML_CHART_NAMESPACE,
    STRICT_DRAWINGML_NAMESPACE, decode_xml_reference, unqualified_attribute_value,
};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, QName, ResolveResult},
    reader::NsReader,
};
use thiserror::Error as ThisError;

use crate::{Error, Result};

use super::model::{
    Bounds, Common, Kind, PlaceholderRecord, Record, Shape, Shapes, Span, TextSpan,
};

const PML: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const DIAGRAM: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/diagram";
const STRICT_DIAGRAM: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/diagram";
const P14: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const P15: &str = "http://schemas.microsoft.com/office/powerpoint/2012/main";

const TABLE: u8 = 1;
const CHART: u8 = 1 << 1;
const DIAGRAM_MARKER: u8 = 1 << 2;
const OLE: u8 = 1 << 3;

/// Primary safe selector for a shape scene.
///
/// Exact semantic names are the convenient entry point. Numeric pre-order
/// positions remain available for source-order and repair workflows.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Key<'a> {
    Name(&'a str),
    Index(usize),
}

impl<'a> From<&'a str> for Key<'a> {
    fn from(value: &'a str) -> Self {
        Self::Name(value)
    }
}

impl From<usize> for Key<'_> {
    fn from(value: usize) -> Self {
        Self::Index(value)
    }
}

/// Typed, non-panicking shape selection failures.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum LookupError {
    #[error("shape name '{name}' was not found")]
    NameNotFound { name: String },
    #[error("shape name '{name}' is ambiguous ({matches} exact matches)")]
    AmbiguousName { name: String, matches: usize },
    #[error("shape index {index} is outside a scene of length {len}")]
    IndexOutOfBounds { index: usize, len: usize },
}

/// Finite resources used to preprocess and index one slide-like XML owner.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    input_bytes: usize,
    output_bytes: usize,
    depth: usize,
    nodes: usize,
    shapes: usize,
    retained_text_bytes: usize,
}

impl Limits {
    /// Conservative defaults for a slide, layout, master, or notes owner.
    pub const DEFAULT: Self = Self {
        input_bytes: 64 * 1024 * 1024,
        output_bytes: 64 * 1024 * 1024,
        depth: 256,
        nodes: 1_000_000,
        shapes: 100_000,
        retained_text_bytes: 16 * 1024 * 1024,
    };

    /// Construct a finite, nonzero limit set.
    pub const fn new(
        input_bytes: usize,
        output_bytes: usize,
        depth: usize,
        nodes: usize,
        shapes: usize,
        retained_text_bytes: usize,
    ) -> Option<Self> {
        if input_bytes == 0
            || output_bytes == 0
            || depth == 0
            || nodes == 0
            || shapes == 0
            || retained_text_bytes == 0
        {
            None
        } else {
            Some(Self {
                input_bytes,
                output_bytes,
                depth,
                nodes,
                shapes,
                retained_text_bytes,
            })
        }
    }

    #[inline]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    #[inline]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    #[inline]
    pub const fn depth(self) -> usize {
        self.depth
    }

    #[inline]
    pub const fn nodes(self) -> usize {
        self.nodes
    }

    #[inline]
    pub const fn shapes(self) -> usize {
        self.shapes
    }

    #[inline]
    pub const fn retained_text_bytes(self) -> usize {
        self.retained_text_bytes
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// A bounded, borrowed-by-default index over one PresentationML shape tree.
///
/// MCE-free input remains borrowed. If markup-compatibility processing must
/// select or unwrap a branch, the scene owns exactly that processed owner XML;
/// individual shape elements are never copied.
#[derive(Debug)]
pub struct Scene<'a> {
    xml: Cow<'a, [u8]>,
    records: Vec<Record>,
    strings: String,
}

impl<'a> Scene<'a> {
    /// Index a scene with conservative finite limits.
    pub fn read(xml: &'a [u8]) -> Result<Self> {
        Self::read_with(xml, Limits::DEFAULT)
    }

    /// Index a scene with caller-selected finite limits.
    pub fn read_with(xml: &'a [u8], limits: Limits) -> Result<Self> {
        if xml.len() > limits.input_bytes {
            return Err(Error::Limit {
                resource: "shape owner input bytes",
                limit: limits.input_bytes,
            });
        }
        if limits.output_bytes > u32::MAX as usize {
            return Err(Error::Invalid(
                "shape output limit exceeds the compact u32 span domain".into(),
            ));
        }

        let mut capabilities = Capabilities::ooxml_baseline();
        capabilities.understand_namespace(P14);
        capabilities.understand_namespace(P15);
        let mce_limits = litchi_ooxml_common::mce::Limits {
            max_input_bytes: limits.input_bytes,
            max_output_bytes: limits.output_bytes,
            max_depth: limits.depth,
            ..litchi_ooxml_common::mce::Limits::default()
        };
        let output = process_markup_compatibility(xml, &capabilities, &mce_limits)?.xml;
        if output.len() > limits.output_bytes || output.len() > u32::MAX as usize {
            return Err(Error::Limit {
                resource: "processed shape owner bytes",
                limit: limits.output_bytes.min(u32::MAX as usize),
            });
        }

        let (records, strings) = Scanner::new(output.as_ref(), limits).scan()?;
        Ok(Self {
            xml: output,
            records,
            strings,
        })
    }

    /// Processed owner XML against which every [`Span`] is defined.
    #[inline]
    pub fn xml(&self) -> &[u8] {
        self.xml.as_ref()
    }

    /// Whether MCE preprocessing produced a replacement owner buffer.
    #[inline]
    pub const fn is_rewritten(&self) -> bool {
        matches!(self.xml, Cow::Owned(_))
    }

    /// Number of shapes in depth-first pre-order, including grouped children.
    #[inline]
    pub fn len(&self) -> usize {
        self.records.len()
    }

    #[inline]
    pub fn is_empty(&self) -> bool {
        self.records.is_empty()
    }

    /// Iterate all shapes in depth-first pre-order.
    pub fn iter(&self) -> Shapes<'_> {
        Shapes {
            xml: self.xml(),
            records: &self.records,
            strings: &self.strings,
            cursor: 0,
            end: self.records.len(),
            parent: None,
            preorder: true,
        }
    }

    /// Iterate only direct children of the owner shape tree.
    pub fn roots(&self) -> Shapes<'_> {
        Shapes {
            xml: self.xml(),
            records: &self.records,
            strings: &self.strings,
            cursor: 0,
            end: self.records.len(),
            parent: None,
            preorder: false,
        }
    }

    /// Lazily visit placeholder shapes in depth-first pre-order.
    pub fn placeholders(&self) -> impl Iterator<Item = Shape<'_>> + '_ {
        self.iter().filter(|shape| shape.placeholder().is_some())
    }

    /// Select by exact semantic name or checked numeric position.
    ///
    /// A missing name is ordinary absence. Malformed duplicate exact names and
    /// out-of-bounds numeric positions are typed errors.
    pub fn get<'k>(
        &self,
        key: impl Into<Key<'k>>,
    ) -> std::result::Result<Option<Shape<'_>>, LookupError> {
        match key.into() {
            Key::Name(name) => self.get_name(name),
            Key::Index(index) => self.at(index).map(Some),
        }
    }

    fn get_name(&self, name: &str) -> std::result::Result<Option<Shape<'_>>, LookupError> {
        let mut found = None;
        let mut matches = 0usize;
        for (index, record) in self.records.iter().enumerate() {
            if record.name.and_then(|span| span.get(&self.strings)) == Some(name) {
                matches = matches.saturating_add(1);
                found = Some(index);
            }
        }
        if matches > 1 {
            return Err(LookupError::AmbiguousName {
                name: name.to_owned(),
                matches,
            });
        }
        found.map(|index| self.at(index)).transpose()
    }

    /// Select a checked depth-first pre-order position.
    pub fn at(&self, index: usize) -> std::result::Result<Shape<'_>, LookupError> {
        let record = self
            .records
            .get(index)
            .ok_or(LookupError::IndexOutOfBounds {
                index,
                len: self.records.len(),
            })?;
        Ok(Shape::from_common(Common {
            xml: self.xml(),
            records: &self.records,
            strings: &self.strings,
            record,
            index,
        }))
    }

    /// Require a shape selected by semantic name or checked numeric position.
    pub fn shape<'k>(
        &self,
        key: impl Into<Key<'k>>,
    ) -> std::result::Result<Shape<'_>, LookupError> {
        match key.into() {
            Key::Name(name) => self
                .get_name(name)?
                .ok_or_else(|| LookupError::NameNotFound {
                    name: name.to_owned(),
                }),
            Key::Index(index) => self.at(index),
        }
    }
}

impl<'a> IntoIterator for &'a Scene<'_> {
    type Item = Shape<'a>;
    type IntoIter = Shapes<'a>;

    fn into_iter(self) -> Self::IntoIter {
        self.iter()
    }
}

#[derive(Debug)]
struct Active {
    index: u32,
    depth: usize,
    seen_non_visual: bool,
    seen_placeholder: bool,
    x: Option<i64>,
    y: Option<i64>,
    width: Option<i64>,
    height: Option<i64>,
    markers: u8,
    text_depth: Option<usize>,
    text: Option<String>,
    seen_paragraph: bool,
}

struct Scanner<'a> {
    xml: &'a [u8],
    limits: Limits,
    records: Vec<Record>,
    strings: String,
    active: Vec<Active>,
    retained_text: usize,
    depth: usize,
    nodes: usize,
    common_slide_depth: Option<usize>,
    tree_depth: Option<usize>,
    seen_tree: bool,
}

impl<'a> Scanner<'a> {
    fn new(xml: &'a [u8], limits: Limits) -> Self {
        Self {
            xml,
            limits,
            records: Vec::new(),
            strings: String::new(),
            active: Vec::new(),
            retained_text: 0,
            depth: 0,
            nodes: 0,
            common_slide_depth: None,
            tree_depth: None,
            seen_tree: false,
        }
    }

    fn scan(mut self) -> Result<(Vec<Record>, String)> {
        let mut reader = NsReader::from_reader(self.xml);
        loop {
            let start = position(&reader)?;
            let decoder = reader.decoder();
            let event = reader.read_event()?.into_owned();
            let end = position(&reader)?;
            let resolver = reader.resolver().clone();
            let (namespace, event) = resolver.resolve_event(event);
            match event {
                Event::Start(element) => {
                    self.count_node()?;
                    let event_depth = self.enter_depth()?;
                    self.start_element(
                        &namespace,
                        &element,
                        decoder,
                        start,
                        event_depth,
                        false,
                        end,
                    )?;
                    self.depth = event_depth;
                },
                Event::Empty(element) => {
                    self.count_node()?;
                    let event_depth = self.enter_depth()?;
                    self.start_element(
                        &namespace,
                        &element,
                        decoder,
                        start,
                        event_depth,
                        true,
                        end,
                    )?;
                },
                Event::Text(text) => {
                    if self
                        .active
                        .last()
                        .is_some_and(|value| value.text_depth.is_some())
                    {
                        let decoded = text
                            .xml_content(XmlVersion::Explicit1_0)
                            .map_err(|error| Error::Xml(error.to_string()))?;
                        let decoded = quick_xml::escape::unescape(&decoded)
                            .map_err(|error| Error::Xml(error.to_string()))?;
                        self.append_text(&decoded)?;
                    }
                },
                Event::CData(text) => {
                    if self
                        .active
                        .last()
                        .is_some_and(|value| value.text_depth.is_some())
                    {
                        let decoded = text
                            .xml_content(XmlVersion::Explicit1_0)
                            .map_err(|error| Error::Xml(error.to_string()))?;
                        self.append_text(&decoded)?;
                    }
                },
                Event::GeneralRef(reference) => {
                    if self
                        .active
                        .last()
                        .is_some_and(|value| value.text_depth.is_some())
                    {
                        self.append_text(&decode_xml_reference(&reference)?)?;
                    }
                },
                Event::End(element) => self.end_element(&namespace, element.name(), end)?,
                Event::DocType(_) | Event::PI(_) => {
                    return Err(Error::Invalid(
                        "DOCTYPE and processing instructions are forbidden in shape XML".into(),
                    ));
                },
                Event::Eof => break,
                _ => {},
            }
        }
        if self.depth != 0 || !self.active.is_empty() {
            return Err(Error::Invalid(
                "shape XML ended with unclosed elements".into(),
            ));
        }
        Ok((self.records, self.strings))
    }

    #[allow(clippy::too_many_arguments)]
    fn start_element(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: quick_xml::encoding::Decoder,
        start: usize,
        event_depth: usize,
        empty: bool,
        end: usize,
    ) -> Result<()> {
        if is_pml(namespace, element.name(), b"cSld") && self.common_slide_depth.is_none() {
            self.common_slide_depth = Some(event_depth);
        }

        let is_tree = is_pml(namespace, element.name(), b"spTree")
            && (event_depth == 1 || self.common_slide_depth == Some(self.depth));
        if is_tree {
            if self.seen_tree {
                return Err(Error::Invalid(
                    "shape owner contains more than one direct shape tree".into(),
                ));
            }
            self.seen_tree = true;
            self.tree_depth = Some(event_depth);
        }

        let parent = self.direct_shape_parent();
        let in_tree = self.tree_depth == Some(self.depth);
        if (in_tree || parent.is_some())
            && let Some(kind) = classify_shape(namespace, element.name())
        {
            let index = self.begin_shape(kind, parent, element, start, event_depth)?;
            if empty {
                self.finish_shape(index, end)?;
            }
            return Ok(());
        }

        if (in_tree || parent.is_some()) && is_shape_like_extension(namespace, element.name()) {
            let index = self.begin_shape(Kind::Unknown, parent, element, start, event_depth)?;
            if empty {
                self.finish_shape(index, end)?;
            }
            return Ok(());
        }

        let Some(active_offset) = self.active.len().checked_sub(1) else {
            return Ok(());
        };
        let active_depth = self
            .active
            .get(active_offset)
            .ok_or_else(|| Error::Invalid("shape stack became inconsistent".into()))?
            .depth;
        let relative = event_depth.saturating_sub(active_depth);

        if is_pml(namespace, element.name(), b"cNvPr") && relative <= 2 {
            let active = self
                .active
                .get_mut(active_offset)
                .ok_or_else(|| Error::Invalid("shape stack became inconsistent".into()))?;
            if active.seen_non_visual {
                return Err(Error::Invalid(
                    "shape contains more than one direct non-visual property record".into(),
                ));
            }
            active.seen_non_visual = true;
            let index = active.index as usize;
            let name = unqualified_attribute_value(element, b"name", decoder)?;
            let id = unqualified_attribute_value(element, b"id", decoder)?
                .map(|value| {
                    value.parse::<u32>().map_err(|_| {
                        Error::Invalid(format!("invalid non-visual shape ID '{value}'"))
                    })
                })
                .transpose()?;
            let name = name
                .as_deref()
                .map(|value| self.retain(value))
                .transpose()?;
            let record = self.records.get_mut(index).ok_or_else(|| {
                Error::Invalid("shape non-visual metadata lost its record".into())
            })?;
            record.name = name;
            record.id = id;
        } else if is_pml(namespace, element.name(), b"ph") && relative <= 3 {
            let active = self
                .active
                .get_mut(active_offset)
                .ok_or_else(|| Error::Invalid("shape stack became inconsistent".into()))?;
            if active.seen_placeholder {
                return Err(Error::Invalid(
                    "shape contains more than one direct placeholder".into(),
                ));
            }
            active.seen_placeholder = true;
            let record_index = active.index as usize;
            let kind = unqualified_attribute_value(element, b"type", decoder)?;
            let index = match unqualified_attribute_value(element, b"idx", decoder)? {
                Some(value) => value
                    .parse::<u32>()
                    .map_err(|_| Error::Invalid(format!("invalid placeholder index '{value}'"))),
                None => Ok(0),
            }?;
            let kind = kind
                .as_deref()
                .map(|value| self.retain(value))
                .transpose()?;
            let record = self.records.get_mut(record_index).ok_or_else(|| {
                Error::Invalid("shape placeholder metadata lost its record".into())
            })?;
            record.placeholder = Some(PlaceholderRecord { kind, index });
        } else if is_dml(namespace, element.name(), b"off") && relative <= 3 {
            let active = self
                .active
                .get_mut(active_offset)
                .ok_or_else(|| Error::Invalid("shape stack became inconsistent".into()))?;
            if active.x.is_none() && active.y.is_none() {
                active.x = Some(parse_i64(element, b"x", decoder)?);
                active.y = Some(parse_i64(element, b"y", decoder)?);
            }
        } else if is_dml(namespace, element.name(), b"ext") && relative <= 3 {
            let active = self
                .active
                .get_mut(active_offset)
                .ok_or_else(|| Error::Invalid("shape stack became inconsistent".into()))?;
            if active.width.is_none() && active.height.is_none() {
                active.width = Some(parse_nonnegative(element, b"cx", decoder)?);
                active.height = Some(parse_nonnegative(element, b"cy", decoder)?);
            }
        }

        let marker = if is_dml(namespace, element.name(), b"tbl") {
            TABLE
        } else if is_chart(namespace, element.name(), b"chart") {
            CHART
        } else if is_diagram(namespace, element.name(), b"relIds") {
            DIAGRAM_MARKER
        } else if is_pml(namespace, element.name(), b"oleObj") {
            OLE
        } else {
            0
        };
        if marker != 0 {
            let active = self
                .active
                .get_mut(active_offset)
                .ok_or_else(|| Error::Invalid("shape stack became inconsistent".into()))?;
            active.markers |= marker;
        }

        if is_dml(namespace, element.name(), b"p") {
            let needs_separator = self.active.get(active_offset).is_some_and(|active| {
                active.seen_paragraph && active.text.as_ref().is_some_and(|text| !text.is_empty())
            });
            if needs_separator {
                self.append_text("\n")?;
            }
            let active = self
                .active
                .get_mut(active_offset)
                .ok_or_else(|| Error::Invalid("shape stack became inconsistent".into()))?;
            active.seen_paragraph = true;
        } else if is_dml(namespace, element.name(), b"t") && !empty {
            let active = self
                .active
                .get_mut(active_offset)
                .ok_or_else(|| Error::Invalid("shape stack became inconsistent".into()))?;
            if active.text_depth.is_some() {
                return Err(Error::Invalid("nested DrawingML text elements".into()));
            }
            active.text_depth = Some(event_depth);
        } else if is_dml(namespace, element.name(), b"br") {
            self.append_text("\n")?;
        } else if is_dml(namespace, element.name(), b"tab") {
            self.append_text("\t")?;
        }
        Ok(())
    }

    fn end_element(
        &mut self,
        namespace: &ResolveResult<'_>,
        name: QName<'_>,
        end: usize,
    ) -> Result<()> {
        if self.depth == 0 {
            return Err(Error::Invalid(
                "shape XML contains an unmatched end tag".into(),
            ));
        }
        if is_dml(namespace, name, b"t")
            && let Some(active) = self.active.last_mut()
            && active.text_depth == Some(self.depth)
        {
            active.text_depth = None;
        }
        if self
            .active
            .last()
            .is_some_and(|active| active.depth == self.depth)
        {
            let index = self
                .active
                .last()
                .map(|active| active.index)
                .ok_or_else(|| Error::Invalid("shape stack became inconsistent".into()))?;
            self.finish_shape(index, end)?;
        }
        if self.tree_depth == Some(self.depth) {
            self.tree_depth = None;
        }
        if self.common_slide_depth == Some(self.depth) {
            self.common_slide_depth = None;
        }
        self.depth = self
            .depth
            .checked_sub(1)
            .ok_or_else(|| Error::Invalid("shape XML depth underflow".into()))?;
        Ok(())
    }

    fn direct_shape_parent(&self) -> Option<u32> {
        self.active.last().and_then(|active| {
            let index = active.index as usize;
            let record = self.records.get(index)?;
            (active.depth == self.depth && record.kind == Kind::Group).then_some(active.index)
        })
    }

    fn begin_shape(
        &mut self,
        kind: Kind,
        parent: Option<u32>,
        element: &BytesStart<'_>,
        start: usize,
        depth: usize,
    ) -> Result<u32> {
        if self.records.len() >= self.limits.shapes {
            return Err(Error::Limit {
                resource: "shapes",
                limit: self.limits.shapes,
            });
        }
        self.records
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "shape records",
                source,
            })?;
        self.active
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "shape nesting stack",
                source,
            })?;
        let index = u32::try_from(self.records.len())
            .map_err(|_| Error::Invalid("shape count exceeds the compact u32 domain".into()))?;
        let start = u32::try_from(start)
            .map_err(|_| Error::Invalid("shape offset exceeds the compact u32 domain".into()))?;
        let qualified_name = element.name();
        let source_name = str::from_utf8(qualified_name.as_ref())
            .map_err(|_| Error::Invalid("shape element name is not UTF-8".into()))?;
        let source_name = Some(self.retain(source_name)?);
        self.records.push(Record {
            span: Span { start, len: 0 },
            subtree_end: index,
            parent,
            kind,
            name: None,
            id: None,
            bounds: None,
            placeholder: None,
            text: None,
            source_name,
        });
        self.active.push(Active {
            index,
            depth,
            seen_non_visual: false,
            seen_placeholder: false,
            x: None,
            y: None,
            width: None,
            height: None,
            markers: 0,
            text_depth: None,
            text: None,
            seen_paragraph: false,
        });
        Ok(index)
    }

    fn finish_shape(&mut self, expected: u32, end: usize) -> Result<()> {
        let active = self
            .active
            .pop()
            .ok_or_else(|| Error::Invalid("shape stack ended unexpectedly".into()))?;
        if active.index != expected {
            return Err(Error::Invalid("shape nesting stack is inconsistent".into()));
        }
        if active.text_depth.is_some() {
            return Err(Error::Invalid("shape ended inside DrawingML text".into()));
        }
        let index = usize::try_from(active.index)
            .map_err(|_| Error::Invalid("shape index does not fit usize".into()))?;
        let start = self
            .records
            .get(index)
            .ok_or_else(|| Error::Invalid("finished shape lost its record".into()))?
            .span
            .start;
        let start_usize = usize::try_from(start)
            .map_err(|_| Error::Invalid("shape offset does not fit usize".into()))?;
        let len = end
            .checked_sub(start_usize)
            .ok_or_else(|| Error::Invalid("shape end precedes its start".into()))?;
        let len = u32::try_from(len)
            .map_err(|_| Error::Invalid("shape length exceeds the compact u32 domain".into()))?;
        let subtree_end = u32::try_from(self.records.len())
            .map_err(|_| Error::Invalid("shape count exceeds the compact u32 domain".into()))?;
        let text = active
            .text
            .as_deref()
            .filter(|value| !value.is_empty())
            .map(|value| self.retain(value))
            .transpose()?;
        let bounds = match (active.x, active.y, active.width, active.height) {
            (Some(x), Some(y), Some(width), Some(height)) => Some(Bounds::new(x, y, width, height)),
            _ => None,
        };
        let record = self
            .records
            .get_mut(index)
            .ok_or_else(|| Error::Invalid("finished shape lost its record".into()))?;
        record.span.len = len;
        record.subtree_end = subtree_end;
        record.bounds = bounds;
        record.text = text;
        if record.kind == Kind::Frame {
            record.kind = match active.markers {
                value if value & OLE != 0 => Kind::Ole,
                TABLE => Kind::Table,
                CHART => Kind::Chart,
                DIAGRAM_MARKER => Kind::Diagram,
                _ => Kind::Frame,
            };
        }
        Ok(())
    }

    fn retain(&mut self, value: &str) -> Result<TextSpan> {
        let start = self.strings.len();
        let end = start.checked_add(value.len()).ok_or(Error::Limit {
            resource: "shape retained strings",
            limit: self.limits.retained_text_bytes,
        })?;
        if end > self.limits.retained_text_bytes || end > u32::MAX as usize {
            return Err(Error::Limit {
                resource: "shape retained strings",
                limit: self.limits.retained_text_bytes.min(u32::MAX as usize),
            });
        }
        self.strings
            .try_reserve(value.len())
            .map_err(|source| Error::Allocation {
                resource: "shape retained strings",
                source,
            })?;
        self.strings.push_str(value);
        Ok(TextSpan {
            start: u32::try_from(start)
                .map_err(|_| Error::Invalid("shape string offset exceeds u32".into()))?,
            len: u32::try_from(value.len())
                .map_err(|_| Error::Invalid("shape string length exceeds u32".into()))?,
        })
    }

    fn append_text(&mut self, value: &str) -> Result<()> {
        let next = self
            .retained_text
            .checked_add(value.len())
            .ok_or(Error::Limit {
                resource: "shape decoded text",
                limit: self.limits.retained_text_bytes,
            })?;
        if next > self.limits.retained_text_bytes {
            return Err(Error::Limit {
                resource: "shape decoded text",
                limit: self.limits.retained_text_bytes,
            });
        }
        let active = self
            .active
            .last_mut()
            .ok_or_else(|| Error::Invalid("text appeared outside a shape".into()))?;
        let text = active.text.get_or_insert_with(String::new);
        text.try_reserve(value.len())
            .map_err(|source| Error::Allocation {
                resource: "shape decoded text",
                source,
            })?;
        text.push_str(value);
        self.retained_text = next;
        Ok(())
    }

    fn count_node(&mut self) -> Result<()> {
        self.nodes = self.nodes.checked_add(1).ok_or(Error::Limit {
            resource: "shape XML elements",
            limit: self.limits.nodes,
        })?;
        if self.nodes > self.limits.nodes {
            Err(Error::Limit {
                resource: "shape XML elements",
                limit: self.limits.nodes,
            })
        } else {
            Ok(())
        }
    }

    fn enter_depth(&self) -> Result<usize> {
        let depth = self.depth.checked_add(1).ok_or(Error::Limit {
            resource: "shape XML nesting depth",
            limit: self.limits.depth,
        })?;
        if depth > self.limits.depth {
            Err(Error::Limit {
                resource: "shape XML nesting depth",
                limit: self.limits.depth,
            })
        } else {
            Ok(depth)
        }
    }
}

fn classify_shape(namespace: &ResolveResult<'_>, name: QName<'_>) -> Option<Kind> {
    if is_pml(namespace, name, b"sp") {
        Some(Kind::Auto)
    } else if is_pml(namespace, name, b"pic") {
        Some(Kind::Picture)
    } else if is_pml(namespace, name, b"graphicFrame") {
        Some(Kind::Frame)
    } else if is_pml(namespace, name, b"grpSp") {
        Some(Kind::Group)
    } else if is_pml(namespace, name, b"cxnSp") {
        Some(Kind::Connector)
    } else if is_pml(namespace, name, b"contentPart")
        || (name.local_name().as_ref() == b"contentPart"
            && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == P14.as_bytes() || *value == P15.as_bytes()))
    {
        Some(Kind::Content)
    } else {
        None
    }
}

fn is_shape_like_extension(namespace: &ResolveResult<'_>, name: QName<'_>) -> bool {
    const MICROSOFT_POWERPOINT: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/";
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if value.starts_with(MICROSOFT_POWERPOINT))
        && matches!(
            name.local_name().as_ref(),
            b"sp" | b"pic" | b"graphicFrame" | b"grpSp" | b"cxnSp" | b"contentPart"
        )
}

fn is_pml(namespace: &ResolveResult<'_>, name: QName<'_>, local: &[u8]) -> bool {
    name.local_name().as_ref() == local
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == PML || *value == STRICT_PML)
}

fn is_dml(namespace: &ResolveResult<'_>, name: QName<'_>, local: &[u8]) -> bool {
    name.local_name().as_ref() == local
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == DRAWINGML_NAMESPACE || *value == STRICT_DRAWINGML_NAMESPACE)
}

fn is_chart(namespace: &ResolveResult<'_>, name: QName<'_>, local: &[u8]) -> bool {
    name.local_name().as_ref() == local
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == DRAWINGML_CHART_NAMESPACE || *value == STRICT_DRAWINGML_CHART_NAMESPACE)
}

fn is_diagram(namespace: &ResolveResult<'_>, name: QName<'_>, local: &[u8]) -> bool {
    name.local_name().as_ref() == local
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == DIAGRAM || *value == STRICT_DIAGRAM)
}

fn parse_i64(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
) -> Result<i64> {
    let value = unqualified_attribute_value(element, name, decoder)?.ok_or_else(|| {
        Error::Invalid(format!(
            "DrawingML coordinate is missing '{}'",
            String::from_utf8_lossy(name)
        ))
    })?;
    value.parse::<i64>().map_err(|_| {
        Error::Invalid(format!(
            "invalid DrawingML coordinate '{}' for '{}'",
            value,
            String::from_utf8_lossy(name)
        ))
    })
}

fn parse_nonnegative(
    element: &BytesStart<'_>,
    name: &[u8],
    decoder: quick_xml::encoding::Decoder,
) -> Result<i64> {
    let value = parse_i64(element, name, decoder)?;
    if value < 0 {
        Err(Error::Invalid(format!(
            "DrawingML extent '{}' cannot be negative",
            String::from_utf8_lossy(name)
        )))
    } else {
        Ok(value)
    }
}

fn position(reader: &NsReader<&[u8]>) -> Result<usize> {
    usize::try_from(reader.buffer_position())
        .map_err(|_| Error::Invalid("shape XML position does not fit usize".into()))
}
