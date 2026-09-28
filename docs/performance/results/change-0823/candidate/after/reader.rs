//! Bounded, namespace-aware scene indexing.

use std::{borrow::Cow, str};

use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::mce::{Capabilities, process_markup_compatibility};
use litchi_ooxml_common::xml::{
    DRAWINGML_CHART_NAMESPACE, DRAWINGML_NAMESPACE, STRICT_DRAWINGML_CHART_NAMESPACE,
    STRICT_DRAWINGML_NAMESPACE, decode_xml_reference, unqualified_attribute_value,
};
#[cfg(test)]
use quick_xml::reader::NsReader;
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
    name::{Namespace, NamespaceResolver, QName, ResolveResult},
    reader::Reader,
};
use thiserror::Error as ThisError;

use crate::{Error, Result};

use super::model::{
    Bounds, Common, Kind, PLACEHOLDER_TYPE_EXTENSION_URI, PlaceholderRecord,
    PlaceholderTypeExtension, Record, Shape, Shapes, Span, TextSpan,
};
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;

const PML: &[u8] = b"http://schemas.openxmlformats.org/presentationml/2006/main";
const STRICT_PML: &[u8] = b"http://purl.oclc.org/ooxml/presentationml/main";
const DIAGRAM: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/diagram";
const STRICT_DIAGRAM: &[u8] = b"http://purl.oclc.org/ooxml/drawingml/diagram";
const P14: &str = "http://schemas.microsoft.com/office/powerpoint/2010/main";
const P15: &str = "http://schemas.microsoft.com/office/powerpoint/2012/main";
const P232: &str = "http://schemas.microsoft.com/office/powerpoint/2023/02/main";

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
    #[must_use]
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
    #[must_use]
    pub const fn input_bytes(self) -> usize {
        self.input_bytes
    }

    #[inline]
    #[must_use]
    pub const fn output_bytes(self) -> usize {
        self.output_bytes
    }

    #[inline]
    #[must_use]
    pub const fn depth(self) -> usize {
        self.depth
    }

    #[inline]
    #[must_use]
    pub const fn nodes(self) -> usize {
        self.nodes
    }

    #[inline]
    #[must_use]
    pub const fn shapes(self) -> usize {
        self.shapes
    }

    #[inline]
    #[must_use]
    pub const fn retained_text_bytes(self) -> usize {
        self.retained_text_bytes
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

/// A bounded, borrowed-by-default index over one `PresentationML` shape tree.
///
/// MCE-free input remains borrowed. If markup-compatibility processing must
/// select or unwrap a branch, the scene owns exactly that processed owner XML;
/// individual shape elements are never copied.
#[derive(Debug)]
pub struct Scene<'a> {
    xml: Cow<'a, [u8]>,
    records: Vec<Record>,
    strings: String,
    limits: Limits,
}

impl<'a> Scene<'a> {
    /// Index a scene with conservative finite limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn read(xml: &'a [u8]) -> Result<Self> {
        Self::read_with(xml, Limits::DEFAULT)
    }

    /// Index a scene with caller-selected finite limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
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
        capabilities.understand_namespace(P232);
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
            limits,
        })
    }

    /// The finite limits this scene was read under. A successful read proves
    /// the owner XML is within them.
    pub(crate) const fn limits(&self) -> Limits {
        self.limits
    }

    /// Processed owner XML against which every [`Span`] is defined.
    #[inline]
    #[must_use]
    pub fn xml(&self) -> &[u8] {
        self.xml.as_ref()
    }

    /// Whether MCE preprocessing produced a replacement owner buffer.
    #[inline]
    #[must_use]
    pub const fn is_rewritten(&self) -> bool {
        matches!(self.xml, Cow::Owned(_))
    }

    /// Number of shapes in depth-first pre-order, including grouped children.
    #[inline]
    #[must_use]
    pub fn len(&self) -> usize {
        self.records.len()
    }

    #[inline]
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.records.is_empty()
    }

    /// Iterate all shapes in depth-first pre-order.
    #[must_use]
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
    #[must_use]
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
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
    placeholder_depth: Option<usize>,
    placeholder_ext_lst_depth: Option<usize>,
    placeholder_ext_depth: Option<usize>,
    placeholder_type_ext_depth: Option<usize>,
    placeholder_type_depth: Option<usize>,
    placeholder_variant_depth: Option<usize>,
    seen_placeholder_ext_lst: bool,
    placeholder_ext_child_count: u8,
    placeholder_ext_has_other_child: bool,
    placeholder_ext_has_forbidden_attribute: bool,
    placeholder_ext_has_typed_child: bool,
    placeholder_ext_uri: Option<String>,
    seen_placeholder_type_ext: bool,
    seen_placeholder_type: bool,
    seen_placeholder_variant: bool,
    placeholder_type_extension: Option<PlaceholderTypeExtension>,
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
    pml_nv_pr_stack: Vec<bool>,
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
            pml_nv_pr_stack: Vec::new(),
            common_slide_depth: None,
            tree_depth: None,
            seen_tree: false,
        }
    }

    fn scan(mut self) -> Result<(Vec<Record>, String)> {
        let mut reader = Reader::from_reader(self.xml);
        let mut resolver = NamespaceResolver::default();
        let mut pending_pop = false;
        // Spans are byte offsets into the owner XML, whose leading
        // byte-order mark precedes reader position zero.
        let origin = ReaderOrigin::of(self.xml);
        loop {
            let start = position(&reader, origin)?;
            let decoder = reader.decoder();
            // Match NsReader's deferred scope pop: an Empty or End scope is
            // still active while its returned event is inspected. NsReader
            // performs this immediately before its underlying read.
            if pending_pop {
                resolver.pop();
                pending_pop = false;
            }
            let event = reader.read_event()?;
            // NsReader pushes declarations while producing Start/Empty and
            // marks Empty/End for the pop before the next read. Do that work
            // before the post-read position so a resolver error has the same
            // precedence as the original implementation.
            match &event {
                Event::Start(element) => {
                    resolver.push(element).map_err(quick_xml::Error::from)?;
                },
                Event::Empty(element) => {
                    resolver.push(element).map_err(quick_xml::Error::from)?;
                    pending_pop = true;
                },
                Event::End(_) => pending_pop = true,
                _ => {},
            }
            let end = position(&reader, origin)?;
            match event {
                Event::Start(element) => {
                    let namespace = resolver.resolve_element(element.name()).0;
                    self.count_node()?;
                    let event_depth = self.enter_depth()?;
                    let parent_is_nv_pr = self.pml_nv_pr_stack.last().copied().unwrap_or(false);
                    self.start_element(
                        &namespace,
                        &element,
                        decoder,
                        start,
                        event_depth,
                        parent_is_nv_pr,
                        false,
                        end,
                    )?;
                    self.pml_nv_pr_stack.push(is_pml(
                        &namespace,
                        element.local_name().into_inner(),
                        b"nvPr",
                    ));
                    self.depth = event_depth;
                },
                Event::Empty(element) => {
                    let namespace = resolver.resolve_element(element.name()).0;
                    self.count_node()?;
                    let event_depth = self.enter_depth()?;
                    let parent_is_nv_pr = self.pml_nv_pr_stack.last().copied().unwrap_or(false);
                    self.start_element(
                        &namespace,
                        &element,
                        decoder,
                        start,
                        event_depth,
                        parent_is_nv_pr,
                        true,
                        end,
                    )?;
                },
                Event::Text(text) => {
                    self.reject_placeholder_character_data(Some(text.as_ref()), false)?;
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
                    self.reject_placeholder_character_data(Some(text.as_ref()), true)?;
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
                    self.reject_placeholder_character_data(None, true)?;
                    if self
                        .active
                        .last()
                        .is_some_and(|value| value.text_depth.is_some())
                    {
                        self.append_text(&decode_xml_reference(&reference)?)?;
                    }
                },
                Event::End(element) => {
                    // Resolve before the deferred pop. The closing element's
                    // prefix is governed by the scope opened by its parent
                    // (and, for a self-declared element, by that element's
                    // own declaration), exactly as NsReader does.
                    let namespace = resolver.resolve_element(element.name()).0;
                    self.end_element(&namespace, element.name(), end)?;
                    self.pml_nv_pr_stack.pop().ok_or_else(|| {
                        Error::Invalid("shape XML element stack became inconsistent".into())
                    })?;
                },
                Event::DocType(_) | Event::PI(_) => {
                    return Err(Error::Invalid(
                        "DOCTYPE and processing instructions are forbidden in shape XML".into(),
                    ));
                },
                Event::Eof => break,
                _ => {},
            }
        }
        if self.depth != 0 || !self.active.is_empty() || !self.pml_nv_pr_stack.is_empty() {
            return Err(Error::Invalid(
                "shape XML ended with unclosed elements".into(),
            ));
        }
        Ok((self.records, self.strings))
    }

    #[cfg(test)]
    fn scan_with_nsreader_oracle(mut self) -> Result<(Vec<Record>, String)> {
        let mut reader = NsReader::from_reader(self.xml);
        // Spans are byte offsets into the owner XML, whose leading
        // byte-order mark precedes reader position zero.
        let origin = ReaderOrigin::of(self.xml);
        loop {
            let start = ns_position(&reader, origin)?;
            let decoder = reader.decoder();
            // A slice reader's events borrow the input, not the reader, and the
            // resolved namespace borrows the reader only until the next read.
            // Neither needs a per-event copy: owning the event or cloning the
            // resolver would reproduce exactly these bytes and bindings.
            let event = reader.read_event()?;
            let end = ns_position(&reader, origin)?;
            let (namespace, event) = reader.resolver().resolve_event(event);
            match event {
                Event::Start(element) => {
                    self.count_node()?;
                    let event_depth = self.enter_depth()?;
                    let parent_is_nv_pr = self.pml_nv_pr_stack.last().copied().unwrap_or(false);
                    self.start_element(
                        &namespace,
                        &element,
                        decoder,
                        start,
                        event_depth,
                        parent_is_nv_pr,
                        false,
                        end,
                    )?;
                    self.pml_nv_pr_stack.push(is_pml(
                        &namespace,
                        element.local_name().into_inner(),
                        b"nvPr",
                    ));
                    self.depth = event_depth;
                },
                Event::Empty(element) => {
                    self.count_node()?;
                    let event_depth = self.enter_depth()?;
                    let parent_is_nv_pr = self.pml_nv_pr_stack.last().copied().unwrap_or(false);
                    self.start_element(
                        &namespace,
                        &element,
                        decoder,
                        start,
                        event_depth,
                        parent_is_nv_pr,
                        true,
                        end,
                    )?;
                },
                Event::Text(text) => {
                    self.reject_placeholder_character_data(Some(text.as_ref()), false)?;
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
                    self.reject_placeholder_character_data(Some(text.as_ref()), true)?;
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
                    self.reject_placeholder_character_data(None, true)?;
                    if self
                        .active
                        .last()
                        .is_some_and(|value| value.text_depth.is_some())
                    {
                        self.append_text(&decode_xml_reference(&reference)?)?;
                    }
                },
                Event::End(element) => {
                    self.end_element(&namespace, element.name(), end)?;
                    self.pml_nv_pr_stack.pop().ok_or_else(|| {
                        Error::Invalid("shape XML element stack became inconsistent".into())
                    })?;
                },
                Event::DocType(_) | Event::PI(_) => {
                    return Err(Error::Invalid(
                        "DOCTYPE and processing instructions are forbidden in shape XML".into(),
                    ));
                },
                Event::Eof => break,
                _ => {},
            }
        }
        if self.depth != 0 || !self.active.is_empty() || !self.pml_nv_pr_stack.is_empty() {
            return Err(Error::Invalid(
                "shape XML ended with unclosed elements".into(),
            ));
        }
        Ok((self.records, self.strings))
    }

    fn reject_placeholder_character_data(
        &self,
        bytes: Option<&[u8]>,
        explicit_markup: bool,
    ) -> Result<()> {
        let Some(active) = self.active.last() else {
            return Ok(());
        };
        if active.placeholder_type_ext_depth.is_some()
            || active.placeholder_type_depth.is_some()
            || active.placeholder_variant_depth.is_some()
        {
            if explicit_markup || !bytes.is_some_and(|value| value.iter().all(is_xml_whitespace)) {
                return Err(Error::Invalid(
                    "p232 phTypeExt/type and empty tokens allow XML whitespace only".into(),
                ));
            }
        }
        if active.placeholder_ext_depth == Some(self.depth)
            && (explicit_markup
                || bytes.is_some_and(|value| value.iter().any(|byte| !is_xml_whitespace(byte))))
        {
            return Err(Error::Invalid(
                "p:ext owner cannot contain direct character data".into(),
            ));
        }
        Ok(())
    }

    #[allow(
        clippy::too_many_arguments,
        reason = "element handler threads one slot per shape-reader field"
    )]
    fn start_element(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        decoder: quick_xml::encoding::Decoder,
        start: usize,
        event_depth: usize,
        parent_is_nv_pr: bool,
        empty: bool,
        end: usize,
    ) -> Result<()> {
        // Every classification below compares this element's local name;
        // split the qualified name once rather than once per comparison.
        let local_name = element.local_name().into_inner();
        if is_pml(namespace, local_name, b"cSld") && self.common_slide_depth.is_none() {
            self.common_slide_depth = Some(event_depth);
        }

        let is_tree = is_pml(namespace, local_name, b"spTree")
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
            && let Some(kind) = classify_shape(namespace, local_name)
        {
            let index = self.begin_shape(kind, parent, element, start, event_depth)?;
            if empty {
                self.finish_shape(index, end)?;
            }
            return Ok(());
        }

        if (in_tree || parent.is_some()) && is_shape_like_extension(namespace, local_name) {
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

        if is_pml(namespace, local_name, b"cNvPr") && relative <= 2 {
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
                    value.parse::<u32>().map_err(|_err| {
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
        } else if is_pml(namespace, local_name, b"ph") && parent_is_nv_pr {
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
                    .map_err(|_err| Error::Invalid(format!("invalid placeholder index '{value}'"))),
                None => Ok(0),
            }?;
            let kind = kind
                .as_deref()
                .map(|value| self.retain(value))
                .transpose()?;
            let record = self.records.get_mut(record_index).ok_or_else(|| {
                Error::Invalid("shape placeholder metadata lost its record".into())
            })?;
            record.placeholder = Some(PlaceholderRecord {
                kind,
                index,
                type_extension: None,
            });
            if !empty {
                let active = self
                    .active
                    .get_mut(active_offset)
                    .ok_or_else(|| Error::Invalid("shape stack became inconsistent".into()))?;
                active.placeholder_depth = Some(event_depth);
            }
        } else if is_dml(namespace, local_name, b"off") && relative <= 3 {
            let active = self
                .active
                .get_mut(active_offset)
                .ok_or_else(|| Error::Invalid("shape stack became inconsistent".into()))?;
            if active.x.is_none() && active.y.is_none() {
                active.x = Some(parse_i64(element, b"x", decoder)?);
                active.y = Some(parse_i64(element, b"y", decoder)?);
            }
        } else if is_dml(namespace, local_name, b"ext") && relative <= 3 {
            let active = self
                .active
                .get_mut(active_offset)
                .ok_or_else(|| Error::Invalid("shape stack became inconsistent".into()))?;
            if active.width.is_none() && active.height.is_none() {
                active.width = Some(parse_nonnegative(element, b"cx", decoder)?);
                active.height = Some(parse_nonnegative(element, b"cy", decoder)?);
            }
        }

        self.scan_placeholder_extension(
            namespace,
            element,
            local_name,
            decoder,
            event_depth,
            empty,
            active_offset,
        )?;

        let marker = if is_dml(namespace, local_name, b"tbl") {
            TABLE
        } else if is_chart(namespace, local_name, b"chart") {
            CHART
        } else if is_diagram(namespace, local_name, b"relIds") {
            DIAGRAM_MARKER
        } else if is_pml(namespace, local_name, b"oleObj") {
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

        if is_dml(namespace, local_name, b"p") {
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
        } else if is_dml(namespace, local_name, b"t") {
            let active = self
                .active
                .get_mut(active_offset)
                .ok_or_else(|| Error::Invalid("shape stack became inconsistent".into()))?;
            // A text element inside an open one is refused whether it is a
            // start tag or an empty one, as the semantic text reader refuses
            // both: an accepted empty `a:t` inside `a:t` gave the text-run
            // rewrite overlapping spans.
            if active.text_depth.is_some() {
                return Err(Error::Invalid("nested DrawingML text elements".into()));
            }
            if !empty {
                active.text_depth = Some(event_depth);
            }
        } else if is_dml(namespace, local_name, b"br") {
            self.append_text("\n")?;
        } else if is_dml(namespace, local_name, b"tab") {
            self.append_text("\t")?;
        }
        Ok(())
    }

    fn scan_placeholder_extension(
        &mut self,
        namespace: &ResolveResult<'_>,
        element: &BytesStart<'_>,
        local_name: &[u8],
        decoder: quick_xml::encoding::Decoder,
        event_depth: usize,
        empty: bool,
        active_offset: usize,
    ) -> Result<()> {
        let Some(active) = self.active.get_mut(active_offset) else {
            return Ok(());
        };
        let parent_depth = event_depth.checked_sub(1);

        if active
            .placeholder_variant_depth
            .is_some_and(|depth| event_depth > depth)
        {
            return Err(Error::Invalid(
                "p232 placeholder type elements must be empty".into(),
            ));
        }

        if is_pml(namespace, local_name, b"extLst") && active.placeholder_depth == parent_depth {
            if active.seen_placeholder_ext_lst {
                return Err(Error::Invalid(
                    "placeholder contains more than one p:extLst".into(),
                ));
            }
            active.seen_placeholder_ext_lst = true;
            if !empty {
                active.placeholder_ext_lst_depth = Some(event_depth);
            }
        } else if is_pml(namespace, local_name, b"ext")
            && active.placeholder_ext_lst_depth == parent_depth
        {
            let uri = validate_extension_attributes(element, decoder)?;
            if empty {
                return Err(Error::Invalid(
                    "p:ext owner requires exactly one child element".into(),
                ));
            }
            if !empty {
                active.placeholder_ext_depth = Some(event_depth);
                active.placeholder_ext_child_count = 0;
                active.placeholder_ext_has_other_child = false;
                active.placeholder_ext_has_forbidden_attribute = false;
                active.placeholder_ext_has_typed_child = false;
                active.placeholder_ext_uri = Some(uri);
            }
        } else if is_p232(namespace, local_name, b"phTypeExt")
            && active.placeholder_ext_depth == parent_depth
            && active.placeholder_ext_uri.as_deref() == Some(PLACEHOLDER_TYPE_EXTENSION_URI)
        {
            if active.placeholder_ext_child_count != 0 {
                return Err(Error::Invalid(
                    "p:ext owner requires exactly one direct child element".into(),
                ));
            }
            active.placeholder_ext_child_count = 1;
            if has_forbidden_attribute(element, decoder)? {
                return Err(Error::Invalid(
                    "p232:phTypeExt does not allow attributes".into(),
                ));
            }
            if active.placeholder_ext_has_other_child
                || active.placeholder_ext_has_forbidden_attribute
            {
                return Err(Error::Invalid(
                    "p232:phTypeExt p:ext owner contains attributes or extra children".into(),
                ));
            }
            if active.placeholder_ext_uri.as_deref() != Some(PLACEHOLDER_TYPE_EXTENSION_URI) {
                return Err(Error::Invalid(
                    "p232:phTypeExt p:ext owner has an unexpected uri".into(),
                ));
            }
            if active.seen_placeholder_type_ext {
                return Err(Error::Invalid(
                    "placeholder contains more than one p232:phTypeExt".into(),
                ));
            }
            if empty {
                return Err(Error::Invalid(
                    "p232:phTypeExt is missing its required type child".into(),
                ));
            }
            active.seen_placeholder_type_ext = true;
            active.placeholder_ext_has_typed_child = true;
            active.placeholder_type_ext_depth = Some(event_depth);
        } else if is_p232(namespace, local_name, b"type")
            && active.placeholder_type_ext_depth == parent_depth
        {
            if has_forbidden_attribute(element, decoder)? {
                return Err(Error::Invalid("p232:type does not allow attributes".into()));
            }
            if active.seen_placeholder_type {
                return Err(Error::Invalid(
                    "p232:phTypeExt contains more than one type child".into(),
                ));
            }
            if empty {
                return Err(Error::Invalid(
                    "p232:type is missing its required type token".into(),
                ));
            }
            active.seen_placeholder_type = true;
            active.placeholder_type_depth = Some(event_depth);
        } else if active.placeholder_type_ext_depth == parent_depth {
            return Err(Error::Invalid(
                "p232:phTypeExt contains only its required type child".into(),
            ));
        } else if active.placeholder_type_depth == parent_depth {
            if has_forbidden_attribute(element, decoder)? {
                return Err(Error::Invalid(
                    "p232 placeholder type token does not allow attributes".into(),
                ));
            }
            let extension = if is_p232(namespace, local_name, b"cameo") {
                Some(PlaceholderTypeExtension::Cameo)
            } else if is_p232(namespace, local_name, b"unknown") {
                Some(PlaceholderTypeExtension::Unknown)
            } else {
                return Err(Error::Invalid(
                    "p232:type must contain cameo or unknown".into(),
                ));
            };
            if active.seen_placeholder_variant {
                return Err(Error::Invalid(
                    "p232:type contains more than one placeholder token".into(),
                ));
            }
            active.seen_placeholder_variant = true;
            active.placeholder_type_extension = extension;
            if !empty {
                active.placeholder_variant_depth = Some(event_depth);
            }
        } else if active.placeholder_ext_depth == parent_depth {
            if active.placeholder_ext_child_count != 0 {
                return Err(Error::Invalid(
                    "p:ext owner requires exactly one direct child element".into(),
                ));
            }
            active.placeholder_ext_child_count = 1;
            active.placeholder_ext_has_other_child = true;
        }
        Ok(())
    }

    fn finish_placeholder_extension(
        &mut self,
        namespace: &ResolveResult<'_>,
        local_name: &[u8],
    ) -> Result<()> {
        let Some(active) = self.active.last_mut() else {
            return Ok(());
        };
        if is_p232(namespace, local_name, b"cameo") || is_p232(namespace, local_name, b"unknown") {
            if active.placeholder_variant_depth == Some(self.depth) {
                active.placeholder_variant_depth = None;
            }
        } else if is_p232(namespace, local_name, b"type") {
            if active.placeholder_type_depth == Some(self.depth) {
                if !active.seen_placeholder_variant {
                    return Err(Error::Invalid(
                        "p232:type is missing its required type token".into(),
                    ));
                }
                active.placeholder_type_depth = None;
            }
        } else if is_p232(namespace, local_name, b"phTypeExt") {
            if active.placeholder_type_ext_depth == Some(self.depth) {
                if !active.seen_placeholder_type || !active.seen_placeholder_variant {
                    return Err(Error::Invalid(
                        "p232:phTypeExt has incomplete placeholder type metadata".into(),
                    ));
                }
                active.placeholder_type_ext_depth = None;
            }
        } else if is_pml(namespace, local_name, b"ext") {
            if active.placeholder_ext_depth == Some(self.depth) {
                if active.placeholder_ext_child_count != 1 {
                    return Err(Error::Invalid(
                        "p:ext owner requires exactly one child element".into(),
                    ));
                }
                if active.placeholder_ext_uri.as_deref() == Some(PLACEHOLDER_TYPE_EXTENSION_URI)
                    && !active.placeholder_ext_has_typed_child
                {
                    return Err(Error::Invalid(
                        "p232 placeholder owner uri is missing its phTypeExt child".into(),
                    ));
                }
                active.placeholder_ext_depth = None;
                active.placeholder_ext_child_count = 0;
                active.placeholder_ext_has_other_child = false;
                active.placeholder_ext_has_forbidden_attribute = false;
                active.placeholder_ext_has_typed_child = false;
                active.placeholder_ext_uri = None;
            }
        } else if is_pml(namespace, local_name, b"extLst") {
            if active.placeholder_ext_lst_depth == Some(self.depth) {
                active.placeholder_ext_lst_depth = None;
            }
        } else if is_pml(namespace, local_name, b"ph")
            && active.placeholder_depth == Some(self.depth)
        {
            if active.seen_placeholder_type_ext
                && (!active.seen_placeholder_type
                    || !active.seen_placeholder_variant
                    || active.placeholder_type_extension.is_none())
            {
                return Err(Error::Invalid(
                    "p232 placeholder type metadata is incomplete".into(),
                ));
            }
            active.placeholder_depth = None;
            active.placeholder_ext_lst_depth = None;
            active.placeholder_ext_depth = None;
            active.seen_placeholder_ext_lst = false;
            active.placeholder_ext_child_count = 0;
            active.placeholder_ext_has_other_child = false;
            active.placeholder_ext_has_forbidden_attribute = false;
            active.placeholder_ext_has_typed_child = false;
            active.placeholder_ext_uri = None;
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
        let local_name = name.local_name().into_inner();
        self.finish_placeholder_extension(namespace, local_name)?;
        if is_dml(namespace, local_name, b"t")
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
            .map_err(|_err| Error::Invalid("shape count exceeds the compact u32 domain".into()))?;
        let start = u32::try_from(start)
            .map_err(|_err| Error::Invalid("shape offset exceeds the compact u32 domain".into()))?;
        let qualified_name = element.name();
        let source_name = str::from_utf8(qualified_name.as_ref())
            .map_err(|_err| Error::Invalid("shape element name is not UTF-8".into()))?;
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
            placeholder_depth: None,
            placeholder_ext_lst_depth: None,
            placeholder_ext_depth: None,
            placeholder_type_ext_depth: None,
            placeholder_type_depth: None,
            placeholder_variant_depth: None,
            seen_placeholder_ext_lst: false,
            placeholder_ext_child_count: 0,
            placeholder_ext_has_other_child: false,
            placeholder_ext_has_forbidden_attribute: false,
            placeholder_ext_has_typed_child: false,
            placeholder_ext_uri: None,
            seen_placeholder_type_ext: false,
            seen_placeholder_type: false,
            seen_placeholder_variant: false,
            placeholder_type_extension: None,
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
            .map_err(|_err| Error::Invalid("shape index does not fit usize".into()))?;
        let start = self
            .records
            .get(index)
            .ok_or_else(|| Error::Invalid("finished shape lost its record".into()))?
            .span
            .start;
        let start_usize = usize::try_from(start)
            .map_err(|_err| Error::Invalid("shape offset does not fit usize".into()))?;
        let len = end
            .checked_sub(start_usize)
            .ok_or_else(|| Error::Invalid("shape end precedes its start".into()))?;
        let len = u32::try_from(len)
            .map_err(|_err| Error::Invalid("shape length exceeds the compact u32 domain".into()))?;
        let subtree_end = u32::try_from(self.records.len())
            .map_err(|_err| Error::Invalid("shape count exceeds the compact u32 domain".into()))?;
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
        if let Some(placeholder) = record.placeholder.as_mut() {
            placeholder.type_extension = active.placeholder_type_extension;
        }
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
                .map_err(|_err| Error::Invalid("shape string offset exceeds u32".into()))?,
            len: u32::try_from(value.len())
                .map_err(|_err| Error::Invalid("shape string length exceeds u32".into()))?,
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

fn classify_shape(namespace: &ResolveResult<'_>, local_name: &[u8]) -> Option<Kind> {
    if is_pml(namespace, local_name, b"sp") {
        Some(Kind::Auto)
    } else if is_pml(namespace, local_name, b"pic") {
        Some(Kind::Picture)
    } else if is_pml(namespace, local_name, b"graphicFrame") {
        Some(Kind::Frame)
    } else if is_pml(namespace, local_name, b"grpSp") {
        Some(Kind::Group)
    } else if is_pml(namespace, local_name, b"cxnSp") {
        Some(Kind::Connector)
    } else if is_pml(namespace, local_name, b"contentPart")
        || (local_name == b"contentPart"
            && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == P14.as_bytes() || *value == P15.as_bytes()))
    {
        Some(Kind::Content)
    } else {
        None
    }
}

fn is_shape_like_extension(namespace: &ResolveResult<'_>, local_name: &[u8]) -> bool {
    const MICROSOFT_POWERPOINT: &[u8] = b"http://schemas.microsoft.com/office/powerpoint/";
    matches!(namespace, ResolveResult::Bound(Namespace(value)) if value.starts_with(MICROSOFT_POWERPOINT))
        && matches!(
            local_name,
            b"sp" | b"pic" | b"graphicFrame" | b"grpSp" | b"cxnSp" | b"contentPart"
        )
}

// Each classifier takes the element's local name, split once per element by
// the caller, and compares the namespace only when the local name matches.

fn is_pml(namespace: &ResolveResult<'_>, local_name: &[u8], local: &[u8]) -> bool {
    local_name == local
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == PML || *value == STRICT_PML)
}

fn is_p232(namespace: &ResolveResult<'_>, local_name: &[u8], local: &[u8]) -> bool {
    local_name == local
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == P232.as_bytes())
}

fn is_dml(namespace: &ResolveResult<'_>, local_name: &[u8], local: &[u8]) -> bool {
    local_name == local
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == DRAWINGML_NAMESPACE || *value == STRICT_DRAWINGML_NAMESPACE)
}

fn is_chart(namespace: &ResolveResult<'_>, local_name: &[u8], local: &[u8]) -> bool {
    local_name == local
        && matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == DRAWINGML_CHART_NAMESPACE || *value == STRICT_DRAWINGML_CHART_NAMESPACE)
}

fn is_diagram(namespace: &ResolveResult<'_>, local_name: &[u8], local: &[u8]) -> bool {
    local_name == local
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
    value.parse::<i64>().map_err(|_err| {
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

/// Validate the generic `p:ext` owner used by `CT_Extension`.
///
/// The owner always carries exactly one unqualified, nonempty XML Schema
/// `token` URI. Namespace declarations are source namespace context and are not
/// schema attributes. The p232 owner has no normative GUID in the local
/// Microsoft specification, so the reader and writer use the crate's stable
/// `urn:litchi:pptx:p232:phTypeExt` owner contract.
fn validate_extension_attributes(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<String> {
    let uri = unqualified_attribute_value(element, b"uri", decoder)?;
    let mut seen_non_namespace = false;
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let key = attribute.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            continue;
        }
        if key != b"uri" || seen_non_namespace {
            return Err(Error::Invalid(
                "p:ext allows exactly one unqualified uri attribute".into(),
            ));
        }
        seen_non_namespace = true;
    }
    let Some(uri) = uri else {
        return Err(Error::Invalid(
            "p:ext is missing its required uri attribute".into(),
        ));
    };
    collapse_xsd_token(&uri)
        .ok_or_else(|| Error::Invalid("p:ext uri must be a nonempty XML Schema token".into()))
}

fn is_xml_whitespace(byte: &u8) -> bool {
    matches!(*byte, b' ' | b'\t' | b'\r' | b'\n')
}

/// Apply the XML Schema `token` whitespace facet.  `token` uses XML whitespace
/// (`#x20`, tab, carriage return, and line feed), replacing runs with one
/// ordinary space and trimming the ends.  The source bytes remain untouched;
/// this normalized value is only used for typed-owner recognition.
fn collapse_xsd_token(value: &str) -> Option<String> {
    let mut collapsed = String::new();
    let mut pending_space = false;
    for character in value.chars() {
        if matches!(character, ' ' | '\t' | '\r' | '\n') {
            if !collapsed.is_empty() {
                pending_space = true;
            }
            continue;
        }
        if pending_space {
            collapsed.push(' ');
            pending_space = false;
        }
        collapsed.push(character);
    }
    (!collapsed.is_empty()).then_some(collapsed)
}

/// Return whether a p232 element has an attribute other than an XML namespace
/// declaration. Namespace declarations are part of the source namespace
/// context and are therefore not schema attributes on the typed p232 tokens.
fn has_forbidden_attribute(
    element: &BytesStart<'_>,
    _decoder: quick_xml::encoding::Decoder,
) -> Result<bool> {
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let key = attribute.key.as_ref();
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            continue;
        }
        return Ok(true);
    }
    Ok(false)
}

fn position(reader: &Reader<&[u8]>, origin: ReaderOrigin) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .ok_or_else(|| Error::Invalid("shape XML position does not fit usize".into()))
}

#[cfg(test)]
fn ns_position(reader: &NsReader<&[u8]>, origin: ReaderOrigin) -> Result<usize> {
    origin
        .offset(reader.buffer_position())
        .ok_or_else(|| Error::Invalid("shape XML position does not fit usize".into()))
}

#[cfg(test)]
mod tests {
    use std::borrow::Cow;

    use super::*;

    const PML_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const DML_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

    fn outcome(result: Result<(Vec<Record>, String)>) -> String {
        match result {
            Ok((records, strings)) => format!("ok:{records:?}|{strings:?}"),
            Err(error) => format!("err:{error:?}"),
        }
    }

    fn assert_scan_parity(xml: &[u8], limits: Limits) {
        let candidate = outcome(Scanner::new(xml, limits).scan());
        let oracle = outcome(Scanner::new(xml, limits).scan_with_nsreader_oracle());
        assert_eq!(candidate, oracle, "candidate/oracle mismatch for {xml:?}");
    }

    fn rich_scene() -> Vec<u8> {
        format!(
            r#"<p:spTree xmlns:p="{PML_NS}" xmlns:a="{DML_NS}" xmlns="urn:outer-default">
                <p:nvGrpSpPr/><p:grpSpPr/>
                <p:sp><p:nvSpPr><p:cNvPr id="1" name="First"/></p:nvSpPr><p:spPr/>
                    <a:txBody><a:p><a:r><text-wrapper xmlns:a="urn:shadow"><a:t xmlns:a="{DML_NS}">text</a:t></text-wrapper></a:r></a:p></a:txBody>
                </p:sp>
                <holder xmlns="" xmlns:a="urn:shadow"><a:empty/><inner xmlns:a=""><a:empty/></inner></holder>
                <p:sp><p:nvSpPr><p:cNvPr id="2" name="Second"/></p:nvSpPr>
                    <p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm></p:spPr>
                    <a:txBody><a:p><a:r><a:t>second</a:t></a:r></a:p></a:txBody>
                </p:sp>
            </p:spTree>"#
        )
        .into_bytes()
    }

    fn assert_rich_semantics(xml: &[u8]) {
        let (records, strings) = Scanner::new(xml, Limits::DEFAULT)
            .scan()
            .expect("rich scene");
        let scene = Scene {
            xml: Cow::Borrowed(xml),
            records,
            strings,
            limits: Limits::DEFAULT,
        };
        assert_eq!(scene.len(), 2);
        assert_eq!(scene.roots().count(), 2);
        let first = scene.at(0).expect("first shape");
        assert_eq!(first.id(), Some(1));
        assert_eq!(first.name(), Some("First"));
        assert_eq!(first.text(), Some("text"));
        assert_eq!(first.bounds(), None);
        let second = scene.at(1).expect("second shape");
        assert_eq!(second.id(), Some(2));
        assert_eq!(second.name(), Some("Second"));
        assert_eq!(second.text(), Some("second"));
        assert_eq!(
            second
                .bounds()
                .map(|bounds| (bounds.x(), bounds.y(), bounds.width(), bounds.height(),)),
            Some((1, 2, 3, 4))
        );
    }

    #[test]
    fn borrowed_resolver_matches_nsreader_for_rich_scopes_and_end_resolution() {
        let xml = rich_scene();
        assert_scan_parity(&xml, Limits::DEFAULT);
        assert_rich_semantics(&xml);
    }

    #[test]
    fn borrowed_resolver_preserves_bom_offsets_and_deferred_empty_pop() {
        let mut xml = b"\xEF\xBB\xBF".to_vec();
        xml.extend_from_slice(&rich_scene());
        assert_scan_parity(&xml, Limits::DEFAULT);
        let (records, _) = Scanner::new(&xml, Limits::DEFAULT)
            .scan()
            .expect("BOM scene");
        assert!(!records.is_empty());
        assert!(records.iter().all(|record| record.span.start >= 3));
    }

    #[test]
    fn borrowed_resolver_matches_nsreader_on_invalid_names_and_malformed_xml() {
        let cases = [
            format!(r#"<p:spTree xmlns:p="{PML_NS}"><u:sp/></p:spTree>"#).into_bytes(),
            format!(r#"<p:spTree xmlns:p="{PML_NS}"><p:sp xmlns:xml="urn:invalid"/></p:spTree>"#)
                .into_bytes(),
            format!(
                r#"<p:spTree xmlns:p="{PML_NS}"><p:sp xmlns:xmlns="urn:invalid"></p:sp></p:spTree>"#
            )
            .into_bytes(),
            format!(r#"<p:spTree xmlns:p="{PML_NS}" xmlns:xml="urn:invalid"/>"#).into_bytes(),
            format!(r#"<p:spTree xmlns:p="{PML_NS}" xmlns:xmlns="urn:invalid"/>"#).into_bytes(),
            format!(r#"<p:spTree xmlns:p="{PML_NS}"><p:sp></p:spTree>"#).into_bytes(),
            format!(r#"<p:spTree xmlns:p="{PML_NS}""#).into_bytes(),
            format!(r#"<p:spTree xmlns:p="{PML_NS}"><p:sp bad="/></p:spTree>"#).into_bytes(),
            format!(r#"<?shape-pi?><p:spTree xmlns:p="{PML_NS}"/>"#).into_bytes(),
            format!(r#"<!DOCTYPE p:spTree><p:spTree xmlns:p="{PML_NS}"/>"#).into_bytes(),
        ];
        for xml in cases {
            assert_scan_parity(&xml, Limits::DEFAULT);
            if xml.starts_with(b"<?") || xml.starts_with(b"<!DOCTYPE") {
                assert!(Scanner::new(&xml, Limits::DEFAULT).scan().is_err());
            }
        }
    }

    #[test]
    fn borrowed_resolver_matches_nsreader_for_default_strict_and_unknown_roots() {
        let default_namespace = format!(
            r#"<spTree xmlns="{PML_NS}"><nvGrpSpPr/><grpSpPr/><sp><nvSpPr><cNvPr id="7" name="Default"/></nvSpPr></sp></spTree>"#
        )
        .into_bytes();
        let strict_namespace = br#"<p:spTree xmlns:p="http://purl.oclc.org/ooxml/presentationml/main"><p:nvGrpSpPr/><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="8" name="Strict"/></p:nvSpPr></p:sp></p:spTree>"#.to_vec();
        let unknown_root =
            format!(r#"<u:spTree xmlns:p="{PML_NS}"><p:nvGrpSpPr/><p:grpSpPr/><p:sp/></u:spTree>"#)
                .into_bytes();
        for xml in [default_namespace, strict_namespace, unknown_root] {
            assert_scan_parity(&xml, Limits::DEFAULT);
        }
    }

    #[test]
    fn borrowed_resolver_matches_nsreader_for_cdata_entities_and_duplicate_attributes() {
        let cdata = format!(
            r#"<p:spTree xmlns:p="{PML_NS}" xmlns:a="{DML_NS}"><p:nvGrpSpPr/><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="1" name="CData"/></p:nvSpPr><a:txBody><a:p><a:r><a:t><![CDATA[A < B]]></a:t></a:r></a:p></a:txBody></p:sp></p:spTree>"#
        )
        .into_bytes();
        let entity = format!(
            r#"<p:spTree xmlns:p="{PML_NS}" xmlns:a="{DML_NS}"><p:nvGrpSpPr/><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="1" name="Entity"/></p:nvSpPr><a:txBody><a:p><a:r><a:t>A &amp; B &#65;</a:t></a:r></a:p></a:txBody></p:sp></p:spTree>"#
        )
        .into_bytes();
        let duplicate_attribute = format!(
            r#"<p:spTree xmlns:p="{PML_NS}"><p:nvGrpSpPr/><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="1" id="2" name="Duplicate"/></p:nvSpPr></p:sp></p:spTree>"#
        )
        .into_bytes();
        for xml in [cdata, entity, duplicate_attribute] {
            assert_scan_parity(&xml, Limits::DEFAULT);
        }
    }

    #[test]
    fn borrowed_resolver_restores_shadowed_default_drawingml_scope() {
        let xml = format!(
            r#"<p:spTree xmlns="{DML_NS}" xmlns:p="{PML_NS}"><p:nvGrpSpPr/><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="9" name="DefaultDml"/></p:nvSpPr><p:txBody><p><r><t>A</t><t xmlns="urn:shadow"/><t>B</t><scope xmlns="urn:shadow"><t xmlns="{DML_NS}">C</t></scope><t>D</t></r></p></p:txBody></p:sp></p:spTree>"#
        )
        .into_bytes();
        assert_scan_parity(&xml, Limits::DEFAULT);
        let (records, strings) = Scanner::new(&xml, Limits::DEFAULT)
            .scan()
            .expect("default DrawingML scene");
        let scene = Scene {
            xml: Cow::Borrowed(&xml),
            records,
            strings,
            limits: Limits::DEFAULT,
        };
        assert_eq!(scene.len(), 1);
        assert_eq!(scene.at(0).expect("shape").text(), Some("ABCD"));
    }

    #[test]
    fn borrowed_resolver_matches_nsreader_at_scanner_resource_boundaries() {
        let xml = rich_scene();
        for nodes in [1, 2, 4, 8, 16, 32, 64] {
            let limits = Limits::new(
                64 * 1024 * 1024,
                64 * 1024 * 1024,
                256,
                nodes,
                100_000,
                16 * 1024 * 1024,
            )
            .expect("finite limits");
            assert_scan_parity(&xml, limits);
        }
        for depth in [1, 2, 3, 4, 8, 256] {
            let limits = Limits::new(
                64 * 1024 * 1024,
                64 * 1024 * 1024,
                depth,
                1_000_000,
                100_000,
                16 * 1024 * 1024,
            )
            .expect("finite limits");
            assert_scan_parity(&xml, limits);
        }
        for shapes in [1, 2, 3, 100_000] {
            let limits = Limits::new(
                64 * 1024 * 1024,
                64 * 1024 * 1024,
                256,
                1_000_000,
                shapes,
                16 * 1024 * 1024,
            )
            .expect("finite limits");
            assert_scan_parity(&xml, limits);
        }
        for retained_text_bytes in [1, 4, 5, 16, 16 * 1024 * 1024] {
            let limits = Limits::new(
                64 * 1024 * 1024,
                64 * 1024 * 1024,
                256,
                1_000_000,
                100_000,
                retained_text_bytes,
            )
            .expect("finite limits");
            assert_scan_parity(&xml, limits);
        }
    }

    #[test]
    fn borrowed_resolver_matches_nsreader_at_namespace_declaration_ceiling() {
        for additional in [255, 256] {
            let mut declarations = format!(r#"xmlns:p="{PML_NS}""#);
            for index in 0..additional {
                declarations.push_str(&format!(r##" xmlns:n{index}="urn:n{index}""##));
            }
            let empty = format!(r#"<p:spTree {declarations}/>"#).into_bytes();
            let closed = format!(r#"<p:spTree {declarations}></p:spTree>"#).into_bytes();
            let succeeds = additional == 255;
            for xml in [empty, closed] {
                assert_scan_parity(&xml, Limits::DEFAULT);
                assert_eq!(
                    Scanner::new(&xml, Limits::DEFAULT).scan().is_ok(),
                    succeeds,
                    "unexpected declaration-boundary result for {additional} additional declarations"
                );
            }
        }
    }
}
