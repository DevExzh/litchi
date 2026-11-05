//! One bounded pass that records the direct block children of a container.

use super::{BlockEntry, BlockIndex, BlockKind, ContainerRoot, Element};
use crate::error::{Error, Result};
use crate::namespace::{NamespaceBindings, NamespaceCapture, is_wordprocessing_namespace};
use crate::parts::memo::RawPositions;
use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, PrefixDeclaration, ResolveResult};
use quick_xml::reader::NsReader;

/// Maximum number of block children one index retains, the bound the
/// paragraph index already uses.
const MAX_BLOCK_ENTRIES: usize = 1_000_000;
/// Maximum element nesting, as the ordered block view accepts.
const MAX_DEPTH: usize = 256;
/// Maximum number of elements read in one pass, as the ordered block view
/// accepts.
const MAX_NODES: usize = 1_000_000;
/// Transitional and Strict Office Math namespaces.
const MATH_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/math";
const STRICT_MATH_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/math";
/// Sentinel for an entry without a direct position.
pub(super) const NO_DIRECT_POSITION: u32 = u32::MAX;

/// What a child of a block parent turned out to be.
enum Child {
    Block(Element),
    /// A `w:customXml` wrapper whose children are flattened into the
    /// container.
    Wrapper,
    /// A property element of the container or of a wrapper.
    Properties,
}

/// A block child whose start tag has been read and whose end has not.
struct Pending {
    depth: usize,
    start: usize,
    element: Element,
    wrapped: bool,
    namespaces: NamespaceBindings,
    /// Element children of a paragraph other than `w:pPr`, for the
    /// display-math rule.
    content_children: usize,
    /// Whether the first such child is an `m:oMathPara`.
    first_is_math_paragraph: bool,
}

struct Builder {
    origin: usize,
    entries: Vec<BlockEntry>,
    kind_counts: [u32; BlockKind::COUNT],
    direct_paragraphs: u32,
    direct_tables: u32,
    direct_controls: u32,
}

impl Builder {
    fn push(
        &mut self,
        kind: BlockKind,
        element: Element,
        wrapped: bool,
        range: (usize, usize),
        namespaces: NamespaceBindings,
    ) -> Result<()> {
        if self.entries.len() >= MAX_BLOCK_ENTRIES {
            return Err(Error::BlockReadLimit {
                resource: "entries",
                actual: self.entries.len().saturating_add(1),
                maximum: MAX_BLOCK_ENTRIES,
            });
        }
        self.entries
            .try_reserve(1)
            .map_err(|source| Error::Allocation {
                resource: "block index",
                source,
            })?;
        let start = self
            .origin
            .checked_add(range.0)
            .and_then(|start| u32::try_from(start).ok())
            .ok_or_else(|| Error::InvalidFormat("block offset does not fit u32".into()))?;
        let length = range
            .1
            .checked_sub(range.0)
            .and_then(|length| u32::try_from(length).ok())
            .ok_or_else(|| Error::InvalidFormat("block length does not fit u32".into()))?;
        let slot = &mut self.kind_counts[kind.slot()];
        let ordinal = *slot;
        *slot = slot
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("block ordinal overflow".into()))?;
        let counter = if wrapped {
            None
        } else {
            match element {
                Element::Paragraph => Some(&mut self.direct_paragraphs),
                Element::Table => Some(&mut self.direct_tables),
                Element::ContentControl => Some(&mut self.direct_controls),
                Element::MathParagraph
                | Element::SectionProperties
                | Element::AltChunk
                | Element::Other => None,
            }
        };
        let direct = match counter {
            Some(counter) => {
                let value = *counter;
                *counter = counter
                    .checked_add(1)
                    .filter(|next| *next != NO_DIRECT_POSITION)
                    .ok_or_else(|| Error::InvalidFormat("block position overflow".into()))?;
                value
            },
            None => NO_DIRECT_POSITION,
        };
        self.entries.push(BlockEntry {
            start,
            length,
            ordinal,
            direct,
            kind,
            element,
            wrapped,
            namespaces,
        });
        Ok(())
    }
}

fn depth_error() -> Error {
    Error::BlockReadLimit {
        resource: "depth",
        actual: MAX_DEPTH + 1,
        maximum: MAX_DEPTH,
    }
}

fn offset(origin: ReaderOrigin, position: u64) -> Result<usize> {
    origin
        .offset(position)
        .ok_or_else(|| Error::InvalidFormat("block container offset does not fit usize".into()))
}

fn is_math(namespace: &ResolveResult<'_>) -> bool {
    matches!(
        namespace,
        ResolveResult::Bound(Namespace(value))
            if *value == MATH_NAMESPACE || *value == STRICT_MATH_NAMESPACE
    )
}

struct Scan<'a> {
    reader: NsReader<&'a [u8]>,
    root: ContainerRoot,
    /// Depths of the open container and of every open flattened wrapper.
    parents: Vec<usize>,
    container_seen: bool,
    pending: Option<Pending>,
    capture: NamespaceCapture,
    builder: Builder,
}

impl Scan<'_> {
    fn classify(&self, element: &BytesStart<'_>, parent_depth: usize) -> Child {
        let (namespace, local) = self.reader.resolver().resolve_element(element.name());
        if is_wordprocessing_namespace(&namespace) {
            let in_wrapper = self.parents.first() != Some(&parent_depth);
            return match local.as_ref() {
                b"p" => Child::Block(Element::Paragraph),
                b"tbl" => Child::Block(Element::Table),
                b"sdt" => Child::Block(Element::ContentControl),
                b"sectPr" => Child::Block(Element::SectionProperties),
                b"altChunk" => Child::Block(Element::AltChunk),
                b"customXml" => Child::Wrapper,
                b"customXmlPr" if in_wrapper => Child::Properties,
                name if !in_wrapper && self.root.properties() == Some(name) => Child::Properties,
                _ => Child::Block(Element::Other),
            };
        }
        if is_math(&namespace) && local.as_ref() == b"oMathPara" {
            return Child::Block(Element::MathParagraph);
        }
        Child::Block(Element::Other)
    }

    /// Validate the root element and open it when it is the container.
    fn root(&mut self, element: &BytesStart<'_>, has_children: bool) -> Result<()> {
        let (namespace, local) = self.reader.resolver().resolve_element(element.name());
        let is_word = is_wordprocessing_namespace(&namespace);
        let local = local.as_ref();
        match self.root {
            ContainerRoot::Document if is_word && local == b"document" => Ok(()),
            ContainerRoot::ContentControl if is_word && local == b"sdt" => Ok(()),
            ContainerRoot::Element { names, .. } if is_word && names.contains(&local) => {
                self.container_seen = true;
                if has_children {
                    self.parents.push(1);
                }
                Ok(())
            },
            ContainerRoot::Document
            | ContainerRoot::ContentControl
            | ContainerRoot::Element { .. } => Err(Error::InvalidFormat(
                "block container has an unexpected root element".into(),
            )),
        }
    }

    /// Whether `element`, a child of the root at `depth`, opens the container
    /// of a root that holds its blocks in a child.
    fn opens_container(&mut self, element: &BytesStart<'_>, depth: usize) -> Result<bool> {
        let Some(name) = self.root.child_container() else {
            return Ok(false);
        };
        if depth != 2 {
            return Ok(false);
        }
        let (namespace, local) = self.reader.resolver().resolve_element(element.name());
        if !is_wordprocessing_namespace(&namespace) || local.as_ref() != name {
            return Ok(false);
        }
        if self.container_seen {
            if matches!(self.root, ContainerRoot::Document) {
                return Err(Error::InvalidFormat(
                    "main document contains multiple bodies".into(),
                ));
            }
            return Err(Error::InvalidFormat(
                "content control contains multiple content containers".into(),
            ));
        }
        self.container_seen = true;
        Ok(true)
    }

    /// Note a direct child of the open block when the block is a paragraph.
    /// Returns whether a block is open, in which case nothing else applies.
    fn observe_block_child(&mut self, element: &BytesStart<'_>, depth: usize) -> bool {
        let Some(pending) = self.pending.as_ref() else {
            return false;
        };
        if depth != pending.depth + 1 || pending.element != Element::Paragraph {
            return true;
        }
        let (namespace, local) = self.reader.resolver().resolve_element(element.name());
        let properties = is_wordprocessing_namespace(&namespace) && local.as_ref() == b"pPr";
        let math_paragraph = is_math(&namespace) && local.as_ref() == b"oMathPara";
        if let Some(pending) = self.pending.as_mut()
            && !properties
        {
            pending.content_children = pending.content_children.saturating_add(1);
            if pending.content_children == 1 {
                pending.first_is_math_paragraph = math_paragraph;
            }
        }
        true
    }

    fn observe_text(&mut self, depth: usize) {
        if let Some(pending) = self.pending.as_mut()
            && pending.element == Element::Paragraph
            && pending.depth == depth
        {
            pending.content_children = pending.content_children.saturating_add(1);
            pending.first_is_math_paragraph = false;
        }
    }

    /// The parent of a direct block child at `depth`, if `depth` is one.
    fn block_parent(&self, depth: usize) -> Option<usize> {
        self.parents
            .last()
            .copied()
            .filter(|parent| parent.checked_add(1) == Some(depth))
    }

    fn start(&mut self, element: &BytesStart<'_>, depth: usize, start: usize) -> Result<()> {
        if depth == 1 {
            return self.root(element, true);
        }
        if self.observe_block_child(element, depth) {
            return Ok(());
        }
        if self.parents.is_empty() {
            if self.opens_container(element, depth)? {
                self.parents.push(depth);
            }
            return Ok(());
        }
        let Some(parent) = self.block_parent(depth) else {
            return Ok(());
        };
        match self.classify(element, parent) {
            Child::Wrapper => self.parents.push(depth),
            Child::Properties => {},
            Child::Block(kind) => {
                let namespaces = self.capture.capture(self.reader.resolver())?;
                self.pending = Some(Pending {
                    depth,
                    start,
                    element: kind,
                    wrapped: self.parents.len() > 1,
                    namespaces,
                    content_children: 0,
                    first_is_math_paragraph: false,
                });
            },
        }
        Ok(())
    }

    fn empty(
        &mut self,
        element: &BytesStart<'_>,
        depth: usize,
        range: (usize, usize),
    ) -> Result<()> {
        if depth == 1 {
            return self.root(element, false);
        }
        if self.observe_block_child(element, depth) {
            return Ok(());
        }
        if self.parents.is_empty() {
            // An empty container holds no blocks; it is only recorded.
            self.opens_container(element, depth)?;
            return Ok(());
        }
        let Some(parent) = self.block_parent(depth) else {
            return Ok(());
        };
        match self.classify(element, parent) {
            Child::Wrapper | Child::Properties => Ok(()),
            Child::Block(kind) => {
                let namespaces = self.capture.capture(self.reader.resolver())?;
                let wrapped = self.parents.len() > 1;
                self.builder
                    .push(kind.default_kind(), kind, wrapped, range, namespaces)
            },
        }
    }

    fn end(&mut self, depth: usize, end: usize) -> Result<()> {
        if self
            .pending
            .as_ref()
            .is_some_and(|pending| pending.depth == depth)
        {
            let Some(pending) = self.pending.take() else {
                return Ok(());
            };
            let kind = if pending.element == Element::Paragraph
                && pending.content_children == 1
                && pending.first_is_math_paragraph
            {
                BlockKind::Equation
            } else {
                pending.element.default_kind()
            };
            return self.builder.push(
                kind,
                pending.element,
                pending.wrapped,
                (pending.start, end),
                pending.namespaces,
            );
        }
        if self.parents.last() == Some(&depth) {
            self.parents.pop();
        }
        Ok(())
    }
}

/// Index the direct block children of the container `fragment` holds.
///
/// `fragment` is the container element cut from its part at byte `origin`
/// (or the whole part for [`ContainerRoot::Document`]); recorded ranges are
/// offsets into that part. `inherited` are the namespace bindings in scope
/// where the fragment starts. The whole fragment is read, so a malformed tail
/// is reported here, as the ordered block view reports it.
pub(crate) fn index(
    fragment: &[u8],
    origin: usize,
    inherited: &[(Option<Vec<u8>>, Vec<u8>)],
    root: ContainerRoot,
    positions: RawPositions,
) -> Result<BlockIndex> {
    let mut reader = NsReader::from_reader(fragment);
    for (prefix, namespace) in inherited {
        let prefix = prefix
            .as_deref()
            .map_or(PrefixDeclaration::Default, PrefixDeclaration::Named);
        reader
            .resolver_mut()
            .add(prefix, Namespace(namespace))
            .map_err(|_| Error::Xml("malformed block XML".into()))?;
    }
    let reader_origin = ReaderOrigin::of(fragment);
    let mut scan = Scan {
        reader,
        root,
        parents: Vec::new(),
        container_seen: false,
        pending: None,
        capture: NamespaceCapture::default(),
        builder: Builder {
            origin,
            entries: Vec::new(),
            kind_counts: [0; BlockKind::COUNT],
            direct_paragraphs: 0,
            direct_tables: 0,
            direct_controls: 0,
        },
    };
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut saw_root = false;
    let mut root_closed = false;

    loop {
        let start = offset(reader_origin, scan.reader.buffer_position())?;
        let event = scan
            .reader
            .read_event()
            .map_err(|_| Error::Xml("malformed block XML".into()))?;
        let end = offset(reader_origin, scan.reader.buffer_position())?;
        if matches!(event, Event::Start(_) | Event::Empty(_)) {
            nodes = nodes.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("block container element counter overflow".into())
            })?;
            if nodes > MAX_NODES {
                return Err(Error::BlockReadLimit {
                    resource: "elements",
                    actual: nodes,
                    maximum: MAX_NODES,
                });
            }
            if depth == 0 {
                if saw_root || root_closed {
                    return Err(Error::InvalidFormat(
                        "block container has multiple roots".into(),
                    ));
                }
                saw_root = true;
            }
        }
        match event {
            Event::Start(element) => {
                validate_element(&scan.reader, &element)?;
                depth = depth.checked_add(1).ok_or_else(depth_error)?;
                if depth > MAX_DEPTH {
                    return Err(depth_error());
                }
                scan.start(&element, depth, start)?;
            },
            Event::Empty(element) => {
                validate_element(&scan.reader, &element)?;
                let child_depth = depth.checked_add(1).ok_or_else(depth_error)?;
                if child_depth > MAX_DEPTH {
                    return Err(depth_error());
                }
                if depth == 0 {
                    root_closed = true;
                }
                scan.empty(&element, child_depth, (start, end))?;
            },
            Event::End(_) => {
                scan.end(depth, end)?;
                depth = depth.checked_sub(1).ok_or_else(|| {
                    Error::InvalidFormat("invalid block container nesting".into())
                })?;
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::Eof if !saw_root || depth != 0 || scan.pending.is_some() => {
                return Err(Error::InvalidFormat(
                    "block container does not contain exactly one complete root".into(),
                ));
            },
            Event::Eof => break,
            Event::Text(text) => {
                validate_text(
                    &text
                        .xml10_content()
                        .map_err(|_| Error::Xml("invalid block text".into()))?,
                )?;
                if !crate::parts::document_part::is_xml_outer_whitespace(text.as_ref()) {
                    scan.observe_text(depth);
                }
                if depth == 0
                    && !crate::parts::document_part::is_xml_outer_whitespace(text.as_ref())
                {
                    return Err(Error::InvalidFormat(
                        "block container has character data outside its root".into(),
                    ));
                }
            },
            Event::CData(_) if depth == 0 => {
                return Err(Error::InvalidFormat(
                    "block container has CDATA outside its root".into(),
                ));
            },
            Event::CData(text) => {
                validate_text(
                    &text
                        .decode()
                        .map_err(|_| Error::Xml("invalid block CDATA".into()))?,
                )?;
                if !crate::parts::document_part::is_xml_outer_whitespace(text.as_ref()) {
                    scan.observe_text(depth);
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(Error::InvalidFormat(
                    "block XML contains a prohibited declaration or instruction".into(),
                ));
            },
            Event::GeneralRef(reference) => {
                scan.observe_text(depth);
                if depth == 0 {
                    return Err(Error::InvalidFormat(
                        "block XML has a reference outside the root".into(),
                    ));
                }
                validate_text(
                    &litchi_ooxml_common::xml::decode_xml_reference(&reference)
                        .map_err(|_| Error::Xml("invalid block XML reference".into()))?,
                )?;
            },
            Event::Decl(_)
                if saw_root || (start != 0 && fragment.get(..start) != Some(b"\xEF\xBB\xBF")) =>
            {
                return Err(Error::InvalidFormat(
                    "misplaced block XML declaration".into(),
                ));
            },
            Event::Comment(_) | Event::Decl(_) => {},
        }
    }
    BlockIndex::from_entries(scan.builder.entries, scan.builder.kind_counts, positions)
}

/// Validate attributes even on opaque blocks, without echoing authored XML.
pub(super) fn validate_element(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<()> {
    if matches!(
        reader.resolver().resolve_element(element.name()).0,
        ResolveResult::Unknown(_)
    ) {
        return Err(Error::Xml("undeclared block element prefix".into()));
    }
    for attribute in element.checked_attributes() {
        let attribute = attribute.map_err(|_| Error::Xml("malformed block attribute".into()))?;
        if attribute.value.len() > 1024 * 1024 {
            return Err(Error::BlockReadLimit {
                resource: "attribute bytes",
                actual: attribute.value.len(),
                maximum: 1024 * 1024,
            });
        }
        if matches!(
            reader.resolver().resolve_attribute(attribute.key).0,
            ResolveResult::Unknown(_)
        ) {
            return Err(Error::Xml("undeclared block attribute prefix".into()));
        }
        validate_text(
            &attribute
                .decoded_and_normalized_value(quick_xml::XmlVersion::Explicit1_0, reader.decoder())
                .map_err(|_| Error::Xml("invalid block attribute value".into()))?,
        )?;
    }
    Ok(())
}

pub(super) fn validate_text(text: &str) -> Result<()> {
    if text.chars().any(|c| !matches!(c, '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')) {
        return Err(Error::Xml("invalid XML character in block text".into()));
    }
    Ok(())
}
