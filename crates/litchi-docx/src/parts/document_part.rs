#![expect(
    clippy::arbitrary_source_item_ordering,
    reason = "items remain grouped by OOXML schema family and package lifecycle"
)]
#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
#![expect(
    clippy::shadow_unrelated,
    reason = "local parser names mirror the OOXML role currently being decoded"
)]
//! DocumentPart - the main document.xml part of a Word document.

use crate::alt::{Chunk, active, scan};
use crate::error::Result;
use crate::namespace::{is_wordprocessing_namespace, scan_word_element_ranges};
use crate::paragraph::{Paragraph, extract_word_text};
use crate::table::Table;
use litchi_opc::part::Part;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;
use smallvec::SmallVec;
use std::collections::BTreeSet;
use std::sync::Arc;

/// Maximum number of paragraph ranges retained by the reusable semantic
/// index.  Documents beyond this bound continue to use the established
/// streaming selectors; the cache is an optimization and never changes the
/// accepted document surface.
const MAX_PARAGRAPH_INDEX_RANGES: usize = 1_000_000;

/// One byte range for a visible `w:p` element.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct ParagraphRange {
    pub(crate) start: u32,
    pub(crate) length: u32,
}

/// Bounded, source-independent paragraph offsets for one visible XML view.
///
/// The index retains offsets only.  It never owns or exposes XML bytes, so a
/// cached lookup cannot extend the lifetime of a payload or bypass the
/// existing semantic/lossless ownership boundaries.
#[derive(Debug, Clone)]
pub(crate) struct ParagraphIndex {
    ranges: Arc<[ParagraphRange]>,
}

impl ParagraphIndex {
    /// Build an index using the same bounded namespace scanner as the legacy
    /// paragraph selectors.
    pub(crate) fn from_xml(xml: &[u8]) -> Result<Self> {
        let mut ranges = Vec::new();
        scan_word_element_ranges(xml, &[b"p".as_slice()], |_, start, length| {
            if ranges.len() >= MAX_PARAGRAPH_INDEX_RANGES {
                return Err(crate::Error::InvalidFormat(format!(
                    "document paragraph index exceeds {MAX_PARAGRAPH_INDEX_RANGES} ranges"
                )));
            }
            ranges
                .try_reserve(1)
                .map_err(|source| crate::Error::Allocation {
                    resource: "document paragraph index",
                    source,
                })?;
            ranges.push(ParagraphRange { start, length });
            Ok(())
        })?;
        Ok(Self {
            ranges: ranges.into(),
        })
    }

    pub(crate) fn len(&self) -> usize {
        self.ranges.len()
    }

    pub(crate) fn get(&self, index: usize) -> Option<ParagraphRange> {
        self.ranges.get(index).copied()
    }

    pub(crate) fn iter(&self) -> impl Iterator<Item = ParagraphRange> + '_ {
        self.ranges.iter().copied()
    }
}

const MAX_DOCUMENT_SEMANTIC_VALUES: usize = 1_000_000;

fn reserve_document_value<T>(values: &mut Vec<T>, resource: &'static str) -> Result<()> {
    if values.len() >= MAX_DOCUMENT_SEMANTIC_VALUES {
        return Err(crate::Error::InvalidFormat(format!(
            "document semantic value count exceeds {MAX_DOCUMENT_SEMANTIC_VALUES}"
        )));
    }
    values
        .try_reserve(1)
        .map_err(|source| crate::Error::Allocation { resource, source })
}

fn push_document_smallvec<T, const N: usize>(
    inline: &mut SmallVec<[T; N]>,
    spill: &mut Option<Vec<T>>,
    value: T,
    resource: &'static str,
) -> Result<()> {
    if let Some(values) = spill {
        reserve_document_value(values, resource)?;
        values.push(value);
        return Ok(());
    }
    if inline.len() < N {
        inline.push(value);
        return Ok(());
    }
    if inline.len() >= MAX_DOCUMENT_SEMANTIC_VALUES {
        return Err(crate::Error::InvalidFormat(format!(
            "document semantic value count exceeds {MAX_DOCUMENT_SEMANTIC_VALUES}"
        )));
    }
    let capacity = N.checked_add(1).ok_or_else(|| {
        crate::Error::InvalidFormat("document semantic value capacity overflow".into())
    })?;
    let mut values = Vec::new();
    values
        .try_reserve_exact(capacity)
        .map_err(|source| crate::Error::Allocation { resource, source })?;
    values.extend(inline.drain(..));
    values.push(value);
    *spill = Some(values);
    Ok(())
}

pub(crate) fn is_xml_outer_whitespace(bytes: &[u8]) -> bool {
    bytes
        .iter()
        .all(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
}

/// The main document part of a Word document.
///
/// This corresponds to the `/word/document.xml` part in the package.
/// It contains the main document content including paragraphs, tables,
/// sections, and other block-level elements.
pub struct DocumentPart<'a> {
    /// Reference to the underlying part
    part: &'a dyn Part,
    xml: Arc<Vec<u8>>,
    /// Best-effort cache built from the validated visible XML. A malformed
    /// or over-bound payload leaves this empty so the legacy query path keeps
    /// reporting the same error at query time.
    paragraph_index: Option<Arc<ParagraphIndex>>,
}

/// Select markup-compatibility branches for a main-document payload.
///
/// Both materialized OPC parts and source-backed pinned documents use this
/// path so their paragraph semantics remain identical.
pub(crate) fn visible_document_xml(raw: Arc<Vec<u8>>) -> Result<Arc<Vec<u8>>> {
    // Word 2010 paragraph/drawing extensions are declared ignorable by
    // producers. Keep that namespace understood at the package boundary so
    // `w14:paraId`, `w14:textId`, `w14:noSpellErr`, and unrelated w14 markup
    // survive the MCE visibility pass for typed or opaque source-backed
    // readers.
    let mut capabilities = litchi_ooxml_common::mce::Capabilities::default();
    capabilities.understand_namespace(crate::paragraph::extensions::WORD_2010_NAMESPACE);
    match litchi_ooxml_common::mce::process_markup_compatibility(
        raw.as_slice(),
        &capabilities,
        &litchi_ooxml_common::mce::Limits::default(),
    )?
    .xml
    {
        std::borrow::Cow::Borrowed(_) => Ok(raw),
        std::borrow::Cow::Owned(value) => Ok(Arc::new(value)),
    }
}

/// Extract all visible paragraphs from a normalized main-document payload.
pub(crate) fn document_paragraphs(xml: Arc<Vec<u8>>) -> Result<SmallVec<[Paragraph; 32]>> {
    let mut inline = SmallVec::new();
    let mut spill = None;
    scan_word_element_ranges(xml.as_slice(), &[b"p".as_slice()], |_, start, length| {
        push_document_smallvec(
            &mut inline,
            &mut spill,
            Paragraph::from_arc_range(Arc::clone(&xml), start, length),
            "document paragraph views",
        )
    })?;
    match spill {
        Some(values) => Ok(SmallVec::from_vec(values)),
        None => Ok(inline),
    }
}

/// Materialize paragraph views from an already validated range index.
pub(crate) fn document_paragraphs_from_index(
    xml: Arc<Vec<u8>>,
    index: &ParagraphIndex,
) -> Result<SmallVec<[Paragraph; 32]>> {
    let mut inline = SmallVec::new();
    let mut spill = None;
    for range in index.iter() {
        push_document_smallvec(
            &mut inline,
            &mut spill,
            Paragraph::from_arc_range(Arc::clone(&xml), range.start, range.length),
            "document paragraph views",
        )?;
    }
    match spill {
        Some(values) => Ok(SmallVec::from_vec(values)),
        None => Ok(inline),
    }
}

/// Select one paragraph from an already validated range index.
pub(crate) fn document_paragraph_from_index(
    xml: Arc<Vec<u8>>,
    index: &ParagraphIndex,
    position: usize,
) -> Option<Paragraph> {
    index
        .get(position)
        .map(|range| Paragraph::from_arc_range(xml, range.start, range.length))
}

/// Select one visible paragraph without materializing every paragraph view.
///
/// The scanner still consumes the complete payload so malformed trailing XML
/// and the shared depth/node limits retain their established error timing.
pub(crate) fn document_paragraph(xml: Arc<Vec<u8>>, index: usize) -> Result<Option<Paragraph>> {
    let mut position = 0usize;
    let mut paragraph = None;
    scan_word_element_ranges(xml.as_slice(), &[b"p".as_slice()], |_, start, length| {
        if position == index {
            paragraph = Some(Paragraph::from_arc_range(Arc::clone(&xml), start, length));
        }
        position = position.checked_add(1).ok_or_else(|| {
            crate::Error::InvalidFormat("document paragraph counter overflow".into())
        })?;
        Ok(())
    })?;
    Ok(paragraph)
}

/// Count visible paragraphs without allocating paragraph range objects.
pub(crate) fn document_paragraph_count(xml: &[u8]) -> Result<usize> {
    let mut count = 0usize;
    scan_word_element_ranges(xml, &[b"p".as_slice()], |_, _, _| {
        count = count.checked_add(1).ok_or_else(|| {
            crate::Error::InvalidFormat("document paragraph counter overflow".into())
        })?;
        Ok(())
    })?;
    Ok(count)
}

/// Select the active supported block ranges from original document XML.
///
/// This remains available to the mutable writer, whose source-preserving body
/// codec operates on untouched package bytes. Read-only semantic block queries
/// use [`body_block_ranges`] after MCE branch selection so they can retain an
/// `Unknown` fallback for every visible direct body child.
pub(crate) fn active_block_ranges(xml: &[u8]) -> Result<Vec<(usize, u32, u32)>> {
    let mut ranges = Vec::new();
    scan_word_element_ranges(
        xml,
        &[b"p".as_slice(), b"tbl".as_slice(), b"altChunk".as_slice()],
        |target, start, length| {
            reserve_document_value(&mut ranges, "active document block ranges")?;
            ranges.push((target, start, length));
            Ok(())
        },
    )?;
    let mut starts = Vec::new();
    starts
        .try_reserve_exact(ranges.len())
        .map_err(|source| crate::Error::Allocation {
            resource: "active document block offsets",
            source,
        })?;
    for &(_, start, _) in &ranges {
        starts.push(start);
    }
    let selected = active(xml, &starts)?.into_iter().collect::<BTreeSet<_>>();
    ranges.retain(|&(_, start, _)| selected.contains(&start));
    Ok(ranges)
}

/// Select active direct children of the main document body in source order.
///
/// `DocumentPart::from_part` has already selected MCE branches for this
/// source, so every returned range addresses the visible XML and unmodeled
/// body children can remain lossless instead of being silently discarded.
pub(crate) fn body_block_ranges(xml: &[u8]) -> Result<Vec<(usize, u32, u32)>> {
    const PARAGRAPH: usize = 0;
    const TABLE: usize = 1;
    const ALT: usize = 2;
    const UNKNOWN: usize = 3;
    const MAX_DEPTH: usize = 256;
    const MAX_NODES: usize = 1_000_000;

    let mut reader = NsReader::from_reader(xml);
    let mut ranges = Vec::new();
    let mut body_depth = None;
    let mut pending = None::<(usize, usize)>;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut saw_root = false;
    let mut root_closed = false;

    loop {
        let start = usize::try_from(reader.buffer_position()).map_err(|_source_error| {
            crate::Error::InvalidFormat("document XML offset does not fit usize".into())
        })?;
        let event = reader
            .read_event()
            .map_err(|error| crate::Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        let end = usize::try_from(reader.buffer_position()).map_err(|_source_error| {
            crate::Error::InvalidFormat("document XML offset does not fit usize".into())
        })?;

        if matches!(event, Event::Start(_) | Event::Empty(_)) {
            nodes = nodes.checked_add(1).ok_or_else(|| {
                crate::Error::InvalidFormat("document XML element counter overflow".into())
            })?;
            if nodes > MAX_NODES {
                return Err(crate::Error::InvalidFormat(format!(
                    "document XML exceeds {MAX_NODES} elements"
                )));
            }
        }

        match event {
            Event::Start(element) => {
                if depth == 0 {
                    if saw_root || root_closed {
                        return Err(crate::Error::InvalidFormat(
                            "document XML has multiple roots".into(),
                        ));
                    }
                    saw_root = true;
                }
                depth = depth.checked_add(1).ok_or_else(|| {
                    crate::Error::InvalidFormat("document XML nesting is too deep".into())
                })?;
                if depth > MAX_DEPTH {
                    return Err(crate::Error::InvalidFormat(format!(
                        "document XML nesting exceeds the {MAX_DEPTH} depth limit"
                    )));
                }
                let is_word = is_wordprocessing_namespace(&namespace);
                let local = element.local_name();
                if body_depth.is_none() && is_word && local.as_ref() == b"body" {
                    body_depth = Some(depth);
                } else if body_depth.is_some_and(|body| depth == body + 1) {
                    let kind = if is_word && local.as_ref() == b"p" {
                        PARAGRAPH
                    } else if is_word && local.as_ref() == b"tbl" {
                        TABLE
                    } else if is_word && local.as_ref() == b"altChunk" {
                        ALT
                    } else {
                        UNKNOWN
                    };
                    pending = Some((kind, start));
                }
            },
            Event::Empty(element) => {
                if depth == 0 {
                    if saw_root || root_closed {
                        return Err(crate::Error::InvalidFormat(
                            "document XML has multiple roots".into(),
                        ));
                    }
                    saw_root = true;
                    root_closed = true;
                }
                let child_depth = depth.checked_add(1).ok_or_else(|| {
                    crate::Error::InvalidFormat("document XML nesting is too deep".into())
                })?;
                if child_depth > MAX_DEPTH {
                    return Err(crate::Error::InvalidFormat(format!(
                        "document XML nesting exceeds the {MAX_DEPTH} depth limit"
                    )));
                }
                if body_depth.is_some_and(|body| child_depth == body + 1) {
                    let local = element.local_name();
                    let kind = if is_wordprocessing_namespace(&namespace) && local.as_ref() == b"p"
                    {
                        PARAGRAPH
                    } else if is_wordprocessing_namespace(&namespace) && local.as_ref() == b"tbl" {
                        TABLE
                    } else if is_wordprocessing_namespace(&namespace)
                        && local.as_ref() == b"altChunk"
                    {
                        ALT
                    } else {
                        UNKNOWN
                    };
                    let range_start = u32::try_from(start).map_err(|_source_error| {
                        crate::Error::InvalidFormat("document XML offset does not fit u32".into())
                    })?;
                    let length = u32::try_from(end.checked_sub(start).ok_or_else(|| {
                        crate::Error::InvalidFormat("document XML range underflow".into())
                    })?)
                    .map_err(|_source_error| {
                        crate::Error::InvalidFormat("document XML range does not fit u32".into())
                    })?;
                    reserve_document_value(&mut ranges, "document body block ranges")?;
                    ranges.push((kind, range_start, length));
                }
            },
            Event::End(element) => {
                if pending.is_some_and(|_| body_depth.is_some_and(|body| depth == body + 1)) {
                    let (kind, range_start) = pending.take().ok_or_else(|| {
                        crate::Error::InvalidFormat("missing document body block".into())
                    })?;
                    let start = u32::try_from(range_start).map_err(|_source_error| {
                        crate::Error::InvalidFormat("document XML offset does not fit u32".into())
                    })?;
                    let length = u32::try_from(end.checked_sub(range_start).ok_or_else(|| {
                        crate::Error::InvalidFormat("document XML range underflow".into())
                    })?)
                    .map_err(|_source_error| {
                        crate::Error::InvalidFormat("document XML range does not fit u32".into())
                    })?;
                    reserve_document_value(&mut ranges, "document body block ranges")?;
                    ranges.push((kind, start, length));
                }
                if body_depth == Some(depth)
                    && is_wordprocessing_namespace(&namespace)
                    && element.local_name().as_ref() == b"body"
                {
                    body_depth = None;
                }
                depth = depth.checked_sub(1).ok_or_else(|| {
                    crate::Error::InvalidFormat("invalid document XML nesting".into())
                })?;
                if depth == 0 {
                    root_closed = true;
                }
            },
            Event::Eof if !saw_root || depth != 0 || pending.is_some() => {
                return Err(crate::Error::InvalidFormat(
                    "document XML does not contain exactly one root".into(),
                ));
            },
            Event::Eof => break,
            Event::Text(text) => {
                if depth == 0 && !is_xml_outer_whitespace(text.as_ref()) {
                    return Err(crate::Error::InvalidFormat(
                        "document XML has character data outside its root".into(),
                    ));
                }
            },
            Event::CData(_) if depth == 0 => {
                return Err(crate::Error::InvalidFormat(
                    "document XML has CDATA outside its root".into(),
                ));
            },
            Event::CData(_)
            | Event::Comment(_)
            | Event::Decl(_)
            | Event::PI(_)
            | Event::DocType(_)
            | Event::GeneralRef(_) => {},
        }
    }
    Ok(ranges)
}

/// Extract all visible table views from a normalized main-document payload.
pub(crate) fn document_tables(xml: Arc<Vec<u8>>) -> Result<SmallVec<[Table; 8]>> {
    let mut inline = SmallVec::new();
    let mut spill = None;
    scan_word_element_ranges(xml.as_slice(), &[b"tbl".as_slice()], |_, start, length| {
        push_document_smallvec(
            &mut inline,
            &mut spill,
            Table::from_arc_range(Arc::clone(&xml), start, length),
            "document table views",
        )
    })?;
    match spill {
        Some(values) => Ok(SmallVec::from_vec(values)),
        None => Ok(inline),
    }
}

/// Extract all visible blocks in direct body order from a normalized payload.
///
/// Every active direct body child is represented. Typed paragraphs, tables,
/// and `altChunk` anchors retain their existing semantic models while
/// unmodeled children remain inert [`crate::OpaqueBlock`] values. The helper
/// is shared by eager and source-backed documents so MCE-visible ordering and
/// unknown-block retention cannot drift between the two paths.
pub(crate) fn document_blocks(xml: Arc<Vec<u8>>) -> Result<Vec<crate::Block>> {
    use crate::Block;

    let mut alts = scan(xml.as_slice())?;
    let mut elements = Vec::new();
    for (target, start, length) in body_block_ranges(xml.as_slice())? {
        reserve_document_value(&mut elements, "document block views")?;
        let block_source = Arc::clone(&xml);
        elements.push(if target == 0 {
            Block::Paragraph(Box::new(Paragraph::from_arc_range(
                block_source,
                start,
                length,
            )))
        } else if target == 1 {
            Block::Table(Box::new(Table::from_arc_range(block_source, start, length)))
        } else if target == 2 {
            let chunk = alts.remove(&start).ok_or_else(|| {
                crate::error::Error::InvalidFormat(
                    "ordered altChunk lacks parsed anchor metadata".into(),
                )
            })?;
            Block::Alt(Box::new(chunk))
        } else {
            Block::Unknown(Box::new(crate::OpaqueBlock::from_arc_range(
                block_source,
                start,
                length,
            )))
        });
    }
    Ok(elements)
}

/// Extract visible paragraph/table/unknown elements in direct body order.
/// Alternative-format anchors remain available through [`document_blocks`]
/// but are intentionally omitted from this historical element view.
pub(crate) fn document_elements(xml: Arc<Vec<u8>>) -> Result<Vec<crate::Element>> {
    use crate::Element;

    let blocks = document_blocks(xml)?;
    let mut elements = Vec::new();
    for block in blocks {
        let element = match block {
            crate::Block::Paragraph(paragraph) => Some(Element::Paragraph(paragraph)),
            crate::Block::Table(table) => Some(Element::Table(table)),
            crate::Block::Alt(_) => None,
            crate::Block::Unknown(value) => Some(Element::Unknown(value)),
        };
        if let Some(element) = element {
            reserve_document_value(&mut elements, "document element views")?;
            elements.push(element);
        }
    }
    Ok(elements)
}

impl<'a> DocumentPart<'a> {
    /// Return the original OPC part; semantic reads use the cached MCE view.
    #[must_use]
    pub fn part(&self) -> &'a dyn Part {
        self.part
    }

    /// Create a `DocumentPart` from a Part.
    ///
    /// # Arguments
    ///
    /// * `part` - The part containing the document.xml content
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn from_part(part: &'a dyn Part) -> Result<Self> {
        let xml = visible_document_xml(part.blob_arc())?;
        // Index construction is deliberately best-effort. `DocumentPart`
        // historically deferred structural XML errors until the first
        // paragraph query; preserving that timing is more important than
        // caching a malformed payload.
        let paragraph_index = ParagraphIndex::from_xml(xml.as_slice()).ok().map(Arc::new);
        Ok(Self {
            part,
            xml,
            paragraph_index,
        })
    }

    /// Get the shared Arc of XML bytes (zero-copy from Part).
    #[inline]
    fn get_xml_arc(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.xml)
    }

    /// Get the XML bytes of the document.
    #[inline]
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        self.xml.as_slice()
    }

    /// Extract all paragraph text from the document.
    ///
    /// This performs a quick extraction of all text content by finding
    /// `<w:t>` elements in the XML.
    ///
    /// # Performance
    ///
    /// Uses `quick-xml` for efficient streaming XML parsing with pre-allocated
    /// buffers and validated text decoding.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn extract_text(&self) -> Result<String> {
        extract_word_text(self.xml_bytes())
    }

    /// Count the number of paragraphs in the document.
    ///
    /// Counts `<w:p>` elements in the document body.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn paragraph_count(&self) -> Result<usize> {
        if let Some(index) = self.paragraph_index.as_deref() {
            return Ok(index.len());
        }
        document_paragraph_count(self.xml_bytes())
    }

    /// Count the number of tables in the document.
    ///
    /// Counts `<w:tbl>` elements in the document body.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn table_count(&self) -> Result<usize> {
        let mut count = 0;
        scan_word_element_ranges(self.xml_bytes(), &[b"tbl".as_slice()], |_, _, _| {
            count += 1;
            Ok(())
        })?;
        Ok(count)
    }

    /// Get all paragraphs in the document.
    ///
    /// Extracts all `<w:p>` elements from the document body.
    ///
    /// # Performance
    ///
    /// Uses namespace-aware streaming XML parsing and shared byte ranges.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn paragraphs(&self) -> Result<SmallVec<[Paragraph; 32]>> {
        if let Some(index) = self.paragraph_index.as_deref() {
            return document_paragraphs_from_index(self.get_xml_arc(), index);
        }
        document_paragraphs(self.get_xml_arc())
    }

    /// Get one paragraph by zero-based index without allocating the complete
    /// paragraph collection.
    ///
    /// # Errors
    ///
    /// Returns an error if the document XML is malformed or exceeds bounds.
    pub fn paragraph(&self, index: usize) -> Result<Option<Paragraph>> {
        if let Some(paragraph_index) = self.paragraph_index.as_deref() {
            return Ok(document_paragraph_from_index(
                self.get_xml_arc(),
                paragraph_index,
                index,
            ));
        }
        document_paragraph(self.get_xml_arc(), index)
    }

    /// Get all tables in the document.
    ///
    /// Extracts all `<w:tbl>` elements from the document body.
    ///
    /// # Performance
    ///
    /// Uses namespace-aware streaming XML parsing and shared byte ranges.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn tables(&self) -> Result<SmallVec<[Table; 8]>> {
        document_tables(self.get_xml_arc())
    }

    /// Get all document elements (paragraphs and tables) in document order.
    ///
    /// This method parses the XML once and extracts both paragraphs and tables,
    /// returning an ordered vector that preserves the document structure.
    /// This is more efficient than calling `paragraphs()` and `tables()` separately,
    /// and it maintains the correct order of elements for sequential processing.
    ///
    /// # Performance
    ///
    /// Uses a single-pass XML parser that extracts both `<w:p>` and `<w:tbl>` elements
    /// in document order, which is significantly faster than parsing the XML twice.
    ///
    /// # Performance
    ///
    /// Uses one-pass, namespace-aware zero-copy parsing.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn elements(&self) -> Result<Vec<crate::Element>> {
        document_elements(self.get_xml_arc())
    }

    /// Get paragraphs, tables, and alternative-format anchors in document order.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn blocks(&self) -> Result<Vec<crate::Block>> {
        document_blocks(self.get_xml_arc())
    }

    /// Return all alternative-format anchors in XML order.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn alts(&self) -> Result<Vec<Chunk>> {
        let mut alts = Vec::new();
        for block in self.blocks()? {
            if let crate::Block::Alt(chunk) = block {
                reserve_document_value(&mut alts, "document alternative-format anchors")?;
                alts.push(*chunk);
            }
        }
        Ok(alts)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::Element;
    use litchi_opc::packuri::PackURI;
    use litchi_opc::part::BlobPart;

    fn document_part(xml: &[u8]) -> BlobPart {
        BlobPart::new(
            PackURI::new("/word/document.xml").unwrap(),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
                .to_string(),
            xml.to_vec(),
        )
    }

    #[test]
    fn extracts_aliased_word_elements_in_document_order_without_copying_text() {
        let xml = br#"<wp:document xmlns:wp="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:false="urn:not-wordprocessingml"><wp:body><false:p><false:r><false:t>ignored</false:t></false:r></false:p><wp:p><wp:r><wp:t><![CDATA[A < B]]></wp:t></wp:r></wp:p><wp:tbl><wp:tr><wp:tc><wp:p><wp:r><wp:t>cell</wp:t></wp:r></wp:p></wp:tc></wp:tr></wp:tbl><wp:p><wp:r><wp:t>tail</wp:t></wp:r></wp:p><wp:p/><false:tbl/></wp:body></wp:document>"#;
        let part = document_part(xml);
        let document = DocumentPart::from_part(&part).unwrap();

        assert_eq!(document.paragraph_count().unwrap(), 4);
        assert_eq!(document.table_count().unwrap(), 1);
        assert_eq!(document.tables().unwrap().len(), 1);

        let paragraphs = document.paragraphs().unwrap();
        assert_eq!(paragraphs.len(), 4);
        assert_eq!(paragraphs[0].text().unwrap(), "A < B");
        assert_eq!(paragraphs[0].runs().unwrap()[0].text().unwrap(), "A < B");
        assert_eq!(paragraphs[1].text().unwrap(), "cell");
        assert_eq!(paragraphs[2].text().unwrap(), "tail");
        assert_eq!(paragraphs[3].text().unwrap(), "");
        assert_eq!(
            document.paragraph(0).unwrap().unwrap().text().unwrap(),
            "A < B"
        );
        assert_eq!(
            document.paragraph(2).unwrap().unwrap().text().unwrap(),
            "tail"
        );
        assert!(document.paragraph(4).unwrap().is_none());

        let elements = document.elements().unwrap();
        assert_eq!(elements.len(), 6);
        assert!(matches!(elements[0], Element::Unknown(_)));
        assert!(matches!(elements[1], Element::Paragraph(_)));
        assert!(matches!(elements[2], Element::Table(_)));
        assert!(matches!(elements[3], Element::Paragraph(_)));
        assert!(matches!(elements[4], Element::Paragraph(_)));
        assert!(matches!(elements[5], Element::Unknown(_)));

        let blocks = document.blocks().unwrap();
        assert!(
            matches!(&blocks[0], crate::Block::Unknown(value) if value.xml_bytes() == br#"<false:p><false:r><false:t>ignored</false:t></false:r></false:p>"#)
        );
    }

    #[test]
    fn accepts_strict_wordprocessingml_and_self_closing_blocks() {
        let xml = br#"<s:document xmlns:s="http://purl.oclc.org/ooxml/wordprocessingml/main"><s:body><s:p/><s:tbl/></s:body></s:document>"#;
        let part = document_part(xml);
        let document = DocumentPart::from_part(&part).unwrap();

        assert_eq!(document.paragraph_count().unwrap(), 1);
        assert_eq!(document.table_count().unwrap(), 1);
        assert_eq!(document.elements().unwrap().len(), 2);
    }

    #[test]
    fn rejects_unterminated_selected_elements() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r/>"#;
        let part = document_part(xml);
        let document = DocumentPart::from_part(&part).unwrap();

        assert!(document.paragraphs().is_err());
        assert!(document.paragraph(0).is_err());
        assert!(document.elements().is_err());
    }

    #[test]
    fn paragraph_index_matches_independent_text_oracle() {
        let xml = br#"<wp:document xmlns:wp="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><wp:body><wp:p><wp:r><wp:t>outer</wp:t></wp:r></wp:p><wp:tbl><wp:tr><wp:tc><wp:p><wp:r><wp:t>cell</wp:t></wp:r></wp:p></wp:tc></wp:tr></wp:tbl><wp:p><wp:r><wp:t>tail</wp:t></wp:r></wp:p><wp:p/></wp:body></wp:document>"#;
        let index = ParagraphIndex::from_xml(xml).unwrap();
        let source = Arc::new(xml.to_vec());
        let expected = ["outer", "cell", "tail", ""];

        assert_eq!(index.len(), expected.len());
        for (position, expected_text) in expected.into_iter().enumerate() {
            let paragraph = document_paragraph_from_index(Arc::clone(&source), &index, position)
                .expect("indexed paragraph is present");
            assert_eq!(paragraph.text().unwrap(), expected_text);
        }
        assert!(document_paragraph_from_index(source, &index, expected.len()).is_none());
    }
}
