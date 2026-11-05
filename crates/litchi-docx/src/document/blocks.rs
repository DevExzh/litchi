//! Indexed direct block children of `WordprocessingML` containers.
//!
//! One bounded pass over a container's visible XML records, for each direct
//! child of the document body, a table cell, a block content control, or a
//! header, footer, note or comment root:
//!
//! - its [`BlockKind`];
//! - its byte range and the namespace bindings in scope at it, so the value
//!   built from it parses in context without rescanning its part;
//! - its ordinal among the children of the same kind, which is what a path
//!   such as `/body/p[N]` counts;
//! - its position among the container's direct children of the same element
//!   in the raw part, which is what the document transaction addresses.
//!
//! `w:customXml` wrappers are flattened: their children are indexed as
//! children of the container, in order, and marked as wrapped. Property
//! elements (`w:tcPr` of a cell, `w:customXmlPr` of a wrapper) are not
//! blocks. A paragraph whose only child besides `w:pPr` is one `m:oMathPara`,
//! and a bare `m:oMathPara`, are display [equations](BlockKind::Equation); a
//! paragraph counts as a paragraph otherwise.
//!
//! Indexes are built from the markup-compatibility (MCE) projection that
//! every reader sees. Each index also records whether that projection keeps
//! every element where the raw part has it, so a block's
//! [`transaction_position`](IndexedBlock::transaction_position) is offered
//! only when the raw and visible views provably agree on it.

mod control;
mod scan;
#[cfg(test)]
mod tests;

use crate::Block;
use crate::document::OpaqueBlock;
use crate::error::{Error, Result};
use crate::namespace::NamespaceBindings;
use crate::paragraph::{Id, Paragraph};
use crate::parts::memo::{MemoValue, RawPositions};
use crate::table::Table;
use litchi_core::Position;
use std::fmt;
use std::mem::size_of;
use std::sync::{Arc, OnceLock};

pub(crate) use scan::index;

/// What a direct child of a block container is.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BlockKind {
    /// A `w:p` paragraph that is not a display equation.
    Paragraph,
    /// A `w:tbl` table.
    Table,
    /// A block-level `w:sdt` content control.
    ContentControl,
    /// Display math: an `m:oMathPara`, or a `w:p` whose only child besides
    /// `w:pPr` is one `m:oMathPara`.
    Equation,
    /// Section properties (`w:sectPr`).
    SectionProperties,
    /// An alternative-format import anchor (`w:altChunk`).
    AltChunk,
    /// Any other child, retained as inert XML.
    Unknown,
}

impl BlockKind {
    const COUNT: usize = 7;

    const fn slot(self) -> usize {
        match self {
            Self::Paragraph => 0,
            Self::Table => 1,
            Self::ContentControl => 2,
            Self::Equation => 3,
            Self::SectionProperties => 4,
            Self::AltChunk => 5,
            Self::Unknown => 6,
        }
    }
}

/// The element a block child is, which decides its transaction address.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Element {
    /// `w:p`, including a paragraph that holds a display equation.
    Paragraph,
    /// `w:tbl`.
    Table,
    /// `w:sdt`.
    ContentControl,
    /// A bare `m:oMathPara`.
    MathParagraph,
    /// `w:sectPr`.
    SectionProperties,
    /// `w:altChunk`.
    AltChunk,
    /// Anything else.
    Other,
}

impl Element {
    const fn default_kind(self) -> BlockKind {
        match self {
            Self::Paragraph => BlockKind::Paragraph,
            Self::Table => BlockKind::Table,
            Self::ContentControl => BlockKind::ContentControl,
            Self::MathParagraph => BlockKind::Equation,
            Self::SectionProperties => BlockKind::SectionProperties,
            Self::AltChunk => BlockKind::AltChunk,
            Self::Other => BlockKind::Unknown,
        }
    }
}

/// Where a scanned fragment keeps its block children.
#[derive(Debug, Clone, Copy)]
pub(crate) enum ContainerRoot {
    /// A main document: the blocks are the children of its `w:body`.
    Document,
    /// A `w:sdt`: the blocks are the children of its `w:sdtContent`.
    ContentControl,
    /// The root element holds the blocks itself.
    Element {
        /// Accepted local names of the root.
        names: &'static [&'static [u8]],
        /// The root's property element, which is not a block.
        properties: Option<&'static [u8]>,
    },
}

impl ContainerRoot {
    /// A table cell (`w:tc`).
    pub(crate) const CELL: Self = Self::Element {
        names: &[b"tc"],
        properties: Some(b"tcPr"),
    };
    /// A footnote or endnote (`w:footnote`, `w:endnote`).
    pub(crate) const NOTE: Self = Self::Element {
        names: &[b"footnote", b"endnote"],
        properties: None,
    };
    /// A comment (`w:comment`).
    pub(crate) const COMMENT: Self = Self::Element {
        names: &[b"comment"],
        properties: None,
    };
    /// A header or footer part root (`w:hdr`, `w:ftr`).
    pub(crate) const HEADER_FOOTER: Self = Self::Element {
        names: &[b"hdr", b"ftr"],
        properties: None,
    };

    const fn child_container(self) -> Option<&'static [u8]> {
        match self {
            Self::Document => Some(b"body"),
            Self::ContentControl => Some(b"sdtContent"),
            Self::Element { .. } => None,
        }
    }

    const fn properties(self) -> Option<&'static [u8]> {
        match self {
            Self::Element { properties, .. } => properties,
            Self::Document | Self::ContentControl => None,
        }
    }
}

/// One recorded block child.
#[derive(Clone)]
pub(crate) struct BlockEntry {
    start: u32,
    length: u32,
    /// Position among the container's children of the same kind.
    ordinal: u32,
    /// Position among the container's direct children of the same element,
    /// or [`scan::NO_DIRECT_POSITION`].
    direct: u32,
    kind: BlockKind,
    element: Element,
    /// Whether a `w:customXml` wrapper holds the block.
    wrapped: bool,
    namespaces: NamespaceBindings,
}

/// The block children of one container, without the bytes they index.
pub(crate) struct BlockIndex {
    entries: Box<[BlockEntry]>,
    /// Entry positions grouped by kind, each group in document order.
    by_kind: Box<[u32]>,
    /// Where each kind's group starts in `by_kind`; the last value is the
    /// total.
    kind_starts: [u32; BlockKind::COUNT + 1],
    positions: RawPositions,
    /// `w14:paraId` → entry position, sorted by identifier then position,
    /// built on the first identity lookup.
    para_ids: OnceLock<Box<[(u32, u32)]>>,
}

impl BlockIndex {
    fn from_entries(
        entries: Vec<BlockEntry>,
        kind_counts: [u32; BlockKind::COUNT],
        positions: RawPositions,
    ) -> Result<Self> {
        let mut kind_starts = [0u32; BlockKind::COUNT + 1];
        for (slot, count) in kind_counts.iter().enumerate() {
            kind_starts[slot + 1] = kind_starts[slot]
                .checked_add(*count)
                .ok_or_else(|| Error::InvalidFormat("block kind count overflow".into()))?;
        }
        let mut by_kind = Vec::new();
        by_kind
            .try_reserve_exact(entries.len())
            .map_err(|source| Error::Allocation {
                resource: "block kind index",
                source,
            })?;
        by_kind.resize(entries.len(), 0);
        let mut next = kind_starts;
        for (position, entry) in entries.iter().enumerate() {
            let slot = &mut next[entry.kind.slot()];
            let target = usize::try_from(*slot)
                .ok()
                .and_then(|target| by_kind.get_mut(target))
                .ok_or_else(|| Error::InvalidFormat("block kind index overflow".into()))?;
            *target = u32::try_from(position)
                .map_err(|_| Error::InvalidFormat("block position does not fit u32".into()))?;
            *slot = slot.saturating_add(1);
        }
        Ok(Self {
            entries: entries.into_boxed_slice(),
            by_kind: by_kind.into_boxed_slice(),
            kind_starts,
            positions,
            para_ids: OnceLock::new(),
        })
    }

    fn of_kind(&self, kind: BlockKind) -> &[u32] {
        let slot = kind.slot();
        let start = usize::try_from(self.kind_starts[slot]).unwrap_or(usize::MAX);
        let end = usize::try_from(self.kind_starts[slot + 1]).unwrap_or(usize::MAX);
        self.by_kind.get(start..end).unwrap_or_default()
    }

    /// Whether two indexes hold the same blocks; a debug cross-check for
    /// reused memos.
    #[cfg(debug_assertions)]
    pub(crate) fn same_blocks(&self, other: &Self) -> bool {
        self.positions == other.positions
            && self.entries.len() == other.entries.len()
            && self.entries.iter().zip(other.entries.iter()).all(|(a, b)| {
                a.start == b.start
                    && a.length == b.length
                    && a.ordinal == b.ordinal
                    && a.direct == b.direct
                    && a.kind == b.kind
                    && a.element == b.element
                    && a.wrapped == b.wrapped
                    && a.namespaces == b.namespaces
            })
    }
}

impl MemoValue for BlockIndex {
    fn weight_bytes(&self) -> usize {
        // The identity table is filled lazily, so its largest size is
        // charged up front: one pair per indexed block at most.
        self.entries
            .len()
            .saturating_mul(size_of::<BlockEntry>().saturating_add(size_of::<(u32, u32)>()))
            .saturating_add(self.by_kind.len().saturating_mul(size_of::<u32>()))
            .saturating_add(self.entries.iter().fold(0usize, |total, entry| {
                entry.namespaces.iter().fold(
                    total.saturating_add(entry.namespaces.len().saturating_mul(size_of::<(
                        Option<Vec<u8>>,
                        Vec<u8>,
                    )>(
                    ))),
                    |weight, (prefix, uri)| {
                        weight
                            .saturating_add(prefix.as_ref().map_or(0, Vec::capacity))
                            .saturating_add(uri.capacity())
                    },
                )
            }))
    }
}

fn range_bytes(source: &[u8], start: u32, length: u32) -> &[u8] {
    let Ok(start) = usize::try_from(start) else {
        return &[];
    };
    let Ok(length) = usize::try_from(length) else {
        return &[];
    };
    start
        .checked_add(length)
        .and_then(|end| source.get(start..end))
        .unwrap_or_default()
}

/// A retained range with every inherited declaration in `namespaces` written
/// onto its root, so it parses standalone. Uses the bindings captured by the
/// index instead of rescanning the part.
fn self_contained(fragment: &[u8], namespaces: &NamespaceBindings) -> Result<Vec<u8>> {
    let mut declarations: Vec<(Box<[u8]>, Box<[u8]>)> = Vec::new();
    for (prefix, namespace) in namespaces.iter() {
        let prefix: Box<[u8]> = prefix.as_deref().unwrap_or_default().into();
        let namespace: Box<[u8]> = namespace.as_slice().into();
        match declarations
            .iter()
            .position(|(candidate, _)| *candidate == prefix)
        {
            Some(index) => declarations[index] = (prefix, namespace),
            None => {
                declarations
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation {
                        resource: "block namespace declarations",
                        source,
                    })?;
                declarations.push((prefix, namespace));
            },
        }
    }
    declarations.retain(|(_, namespace)| !namespace.is_empty());
    let scope = litchi_ooxml_common::mce::InScopeNamespaces::from_declarations(declarations);
    Ok(scope
        .make_self_contained(fragment, &litchi_ooxml_common::mce::Limits::default())?
        .into_owned())
}

/// The direct block children of one container — the document body, a table
/// cell, a block content control, or a header, footer, note or comment —
/// indexed in visible order.
///
/// A value is cheap to clone: it shares the container's XML and the index.
/// Lookups by position, by kind and ordinal, and by paragraph identifier do
/// not read the XML again; only the values they return parse their own range.
///
/// ```no_run
/// use litchi_docx::{BlockKind, Package};
///
/// let package = Package::open("report.docx")?;
/// let document = package.document()?;
/// let body = document.body_blocks()?;
/// // `/body/p[3]`: the third body paragraph. Tables and display equations
/// // are other kinds; `w:customXml` wrappers are flattened.
/// if let Some(paragraph) = body.paragraph(litchi_core::Position::new(2)) {
///     println!("{}", paragraph.text()?);
/// }
/// println!("{} tables", body.count(BlockKind::Table));
/// # Ok::<(), litchi_docx::Error>(())
/// ```
#[derive(Clone)]
pub struct Blocks {
    source: Arc<Vec<u8>>,
    index: Arc<BlockIndex>,
}

impl fmt::Debug for Blocks {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Blocks")
            .field("len", &self.index.entries.len())
            .field("raw_positions_proven", &self.raw_positions_proven())
            .finish_non_exhaustive()
    }
}

impl Blocks {
    pub(crate) const fn new(source: Arc<Vec<u8>>, index: Arc<BlockIndex>) -> Self {
        Self { source, index }
    }

    /// Index the container element at `origin..origin + length` of `part`,
    /// whose inherited namespace bindings are `inherited`.
    pub(crate) fn index(
        part: Arc<Vec<u8>>,
        origin: u32,
        length: u32,
        inherited: &[(Option<Vec<u8>>, Vec<u8>)],
        root: ContainerRoot,
        positions: RawPositions,
    ) -> Result<Self> {
        let fragment = range_bytes(&part, origin, length);
        let origin = usize::try_from(origin)
            .map_err(|_| Error::InvalidFormat("block container offset overflow".into()))?;
        let index = index(fragment, origin, inherited, root, positions)?;
        Ok(Self::new(part, Arc::new(index)))
    }

    /// The number of blocks.
    #[must_use]
    pub fn len(&self) -> usize {
        self.index.entries.len()
    }

    /// Whether the container holds no blocks.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.index.entries.is_empty()
    }

    /// The block at `position` in document order, counting every kind.
    #[must_use]
    pub fn get(&self, position: Position) -> Option<IndexedBlock> {
        let position = position.get();
        self.index.entries.get(position).map(|entry| IndexedBlock {
            blocks: self.clone(),
            position,
            entry: entry.clone(),
        })
    }

    /// Every block in document order.
    pub fn iter(&self) -> impl ExactSizeIterator<Item = IndexedBlock> + '_ {
        self.index
            .entries
            .iter()
            .enumerate()
            .map(|(position, entry)| IndexedBlock {
                blocks: self.clone(),
                position,
                entry: entry.clone(),
            })
    }

    /// The number of blocks of `kind`.
    #[must_use]
    pub fn count(&self, kind: BlockKind) -> usize {
        self.index.of_kind(kind).len()
    }

    /// The block of `kind` at zero-based `ordinal` among the blocks of that
    /// kind, as a path step such as `p[ordinal + 1]` counts them.
    #[must_use]
    pub fn nth(&self, kind: BlockKind, ordinal: Position) -> Option<IndexedBlock> {
        let position = self.index.of_kind(kind).get(ordinal.get())?;
        self.get(Position::new(usize::try_from(*position).ok()?))
    }

    /// The paragraph at zero-based `ordinal` among the container's
    /// [`BlockKind::Paragraph`] blocks.
    #[must_use]
    pub fn paragraph(&self, ordinal: Position) -> Option<Paragraph> {
        self.nth(BlockKind::Paragraph, ordinal)?.paragraph()
    }

    /// The table at zero-based `ordinal` among the container's tables.
    #[must_use]
    pub fn table(&self, ordinal: Position) -> Option<Table> {
        self.nth(BlockKind::Table, ordinal)?.table()
    }

    /// The content control at zero-based `ordinal` among the container's
    /// block content controls.
    #[must_use]
    pub fn content_control(&self, ordinal: Position) -> Option<BlockContentControl> {
        self.nth(BlockKind::ContentControl, ordinal)?
            .content_control()
    }

    /// The first paragraph-element block (a paragraph or an equation held in
    /// a paragraph) whose `w14:paraId` is `id`.
    ///
    /// The first lookup reads the start tag of every paragraph element once
    /// and keeps a sorted identifier table with the index; later lookups are
    /// a binary search. A duplicated identifier resolves to its first block.
    ///
    /// # Errors
    ///
    /// Returns the error [`Paragraph::para_id`] reports for a paragraph whose
    /// identifier attributes are malformed; nothing is retained then.
    pub(crate) fn by_para_id(&self, id: Id) -> Result<Option<IndexedBlock>> {
        let table = self.para_id_table()?;
        let key = id.get();
        let first = table.partition_point(|(candidate, _)| *candidate < key);
        Ok(match table.get(first) {
            Some((candidate, position)) if *candidate == key => usize::try_from(*position)
                .ok()
                .and_then(|position| self.get(Position::new(position))),
            _ => None,
        })
    }

    fn para_id_table(&self) -> Result<&[(u32, u32)]> {
        if let Some(table) = self.index.para_ids.get() {
            return Ok(table);
        }
        let mut table = Vec::new();
        for (position, entry) in self.index.entries.iter().enumerate() {
            if entry.element != Element::Paragraph {
                continue;
            }
            let bytes = range_bytes(&self.source, entry.start, entry.length);
            let Some(id) =
                crate::paragraph::extensions::start_tag_ids(bytes, &entry.namespaces)?.para_id()
            else {
                continue;
            };
            table.try_reserve(1).map_err(|source| Error::Allocation {
                resource: "paragraph identifier table",
                source,
            })?;
            table.push((
                id.get(),
                u32::try_from(position)
                    .map_err(|_| Error::InvalidFormat("block position does not fit u32".into()))?,
            ));
        }
        table.sort_unstable();
        Ok(self.index.para_ids.get_or_init(|| table.into_boxed_slice()))
    }

    /// Whether the raw part the document transaction edits keeps every
    /// indexed block at the position this index records.
    ///
    /// True when the markup-compatibility projection the index was built from
    /// selected no `mc:AlternateContent` branch and unwrapped no
    /// `mc:ProcessContent` content: such a projection only drops attributes
    /// and whole ignorable subtrees, so every `WordprocessingML` element keeps
    /// its parent and its order. Every part the document transaction admits
    /// (record 0011) qualifies. A container read from a value that does not
    /// carry this proof — a [`crate::Cell`], a note, a comment or a header —
    /// reports `false`; the body index of the same document proves it for
    /// every container in the main document.
    #[must_use]
    pub fn raw_positions_proven(&self) -> bool {
        self.index.positions.is_preserved()
    }

    #[cfg(test)]
    pub(crate) fn shares_index_with(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.index, &other.index)
    }
}

/// One direct child of a block container, as a [`Blocks`] index records it.
#[derive(Clone)]
pub struct IndexedBlock {
    blocks: Blocks,
    position: usize,
    entry: BlockEntry,
}

impl fmt::Debug for IndexedBlock {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("IndexedBlock")
            .field("kind", &self.entry.kind)
            .field("position", &self.position)
            .field("ordinal", &self.entry.ordinal)
            .field("wrapped", &self.entry.wrapped)
            .finish_non_exhaustive()
    }
}

impl IndexedBlock {
    /// What the block is.
    #[must_use]
    pub const fn kind(&self) -> BlockKind {
        self.entry.kind
    }

    /// The block's zero-based position among all blocks of its container.
    #[must_use]
    pub const fn position(&self) -> Position {
        Position::new(self.position)
    }

    /// The block's zero-based position among the blocks of its kind in its
    /// container, as a path step such as `p[ordinal + 1]` counts it.
    #[must_use]
    pub fn ordinal(&self) -> Position {
        Position::new(usize::try_from(self.entry.ordinal).unwrap_or(usize::MAX))
    }

    /// Whether a `w:customXml` wrapper holds the block.
    #[must_use]
    pub const fn is_wrapped(&self) -> bool {
        self.entry.wrapped
    }

    /// The block's zero-based position among the direct children of its
    /// container that are the same element (`w:p`, `w:tbl` or `w:sdt`), in
    /// the visible view; `None` for a wrapped block and for other elements.
    ///
    /// This is the document transaction's position for the block whenever
    /// [`Blocks::raw_positions_proven`] holds for the document the container
    /// belongs to. [`Self::transaction_position`] checks that for you when
    /// the index carries the proof.
    #[must_use]
    pub fn direct_position(&self) -> Option<Position> {
        if self.entry.wrapped || self.entry.direct == scan::NO_DIRECT_POSITION {
            return None;
        }
        usize::try_from(self.entry.direct).ok().map(Position::new)
    }

    /// The position the document transaction addresses this block by: a
    /// paragraph (or a paragraph holding an equation) among the container's
    /// direct `w:p` children, a table among its direct `w:tbl` children, or a
    /// content control among its direct `w:sdt` children, in the raw part.
    ///
    /// # Errors
    ///
    /// Refuses with [`Error::Refused`] instead of guessing when the block
    /// sits inside a `w:customXml` wrapper, when its kind has no transaction
    /// address, or when the raw and visible positions are not proven to agree
    /// ([`Blocks::raw_positions_proven`]).
    pub fn transaction_position(&self) -> Result<Position> {
        let refuse = |kind, reason| Err(Error::refused(kind, "block transaction position", reason));
        if self.entry.wrapped {
            return refuse(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "the block sits inside a w:customXml wrapper, which document transactions do not address",
            );
        }
        let Some(position) = self.direct_position() else {
            return refuse(
                litchi_core::edit::RefusalKind::UnsupportedTarget,
                "document transactions do not address this kind of block",
            );
        };
        if !self.blocks.raw_positions_proven() {
            return refuse(
                litchi_core::edit::RefusalKind::UnsupportedContent,
                "markup compatibility selected or unwrapped content, so raw and visible block positions are not proven to agree",
            );
        }
        Ok(position)
    }

    /// The block's XML exactly as the visible view holds it.
    ///
    /// The fragment's namespace declarations are those in scope where it
    /// appears; use [`Self::self_contained_xml`] to parse it on its own.
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        range_bytes(&self.blocks.source, self.entry.start, self.entry.length)
    }

    /// The block's XML with every namespace declaration it inherits written
    /// onto its root element, so it parses standalone.
    ///
    /// # Errors
    ///
    /// Returns an error when a bound on the namespace rewrite is reached.
    pub fn self_contained_xml(&self) -> Result<Vec<u8>> {
        self_contained(self.xml_bytes(), &self.entry.namespaces)
    }

    /// The block as a paragraph, for a paragraph and for an equation held in
    /// a paragraph.
    #[must_use]
    pub fn paragraph(&self) -> Option<Paragraph> {
        (self.entry.element == Element::Paragraph).then(|| {
            Paragraph::from_arc_range_with_context(
                Arc::clone(&self.blocks.source),
                self.entry.start,
                self.entry.length,
                Arc::clone(&self.entry.namespaces),
            )
        })
    }

    /// The block as a table.
    #[must_use]
    pub fn table(&self) -> Option<Table> {
        (self.entry.element == Element::Table).then(|| {
            Table::from_arc_range_with_context(
                Arc::clone(&self.blocks.source),
                self.entry.start,
                self.entry.length,
                Arc::clone(&self.entry.namespaces),
            )
        })
    }

    /// The block as a content control.
    #[must_use]
    pub fn content_control(&self) -> Option<BlockContentControl> {
        (self.entry.element == Element::ContentControl).then(|| {
            BlockContentControl::new(
                Arc::clone(&self.blocks.source),
                self.entry.start,
                self.entry.length,
                Arc::clone(&self.entry.namespaces),
                self.blocks.index.positions,
            )
        })
    }

    /// The block as the ordered block view [`crate::Document::blocks`]
    /// returns it: section properties, bare equations and unknown children
    /// as [`Block::Unknown`], an equation held in a paragraph as
    /// [`Block::Paragraph`].
    ///
    /// # Errors
    ///
    /// Returns an error when an `altChunk` anchor is malformed.
    pub fn to_block(&self) -> Result<Block> {
        Ok(match self.entry.element {
            Element::Paragraph => {
                Block::Paragraph(Box::new(Paragraph::from_arc_range_with_context(
                    Arc::clone(&self.blocks.source),
                    self.entry.start,
                    self.entry.length,
                    Arc::clone(&self.entry.namespaces),
                )))
            },
            Element::Table => Block::Table(Box::new(Table::from_arc_range_with_context(
                Arc::clone(&self.blocks.source),
                self.entry.start,
                self.entry.length,
                Arc::clone(&self.entry.namespaces),
            ))),
            Element::ContentControl => Block::Sdt(Box::new(BlockContentControl::new(
                Arc::clone(&self.blocks.source),
                self.entry.start,
                self.entry.length,
                Arc::clone(&self.entry.namespaces),
                self.blocks.index.positions,
            ))),
            Element::AltChunk => Block::Alt(Box::new(crate::alt::parse_element(
                self.xml_bytes(),
                &self.entry.namespaces,
            )?)),
            Element::MathParagraph | Element::SectionProperties | Element::Other => {
                Block::Unknown(Box::new(OpaqueBlock::from_arc_range(
                    Arc::clone(&self.blocks.source),
                    self.entry.start,
                    self.entry.length,
                )))
            },
        })
    }
}

/// A block-level content control (`w:sdt`) with its identity and the blocks
/// of its content.
///
/// The value shares the visible XML of the story it was read from. Its
/// identity is read from `w:sdtPr` on first request; its content is indexed
/// by [`Self::blocks`] with the same scanner as the document body.
#[derive(Clone)]
pub struct BlockContentControl {
    source: Arc<Vec<u8>>,
    start: u32,
    length: u32,
    namespaces: NamespaceBindings,
    positions: RawPositions,
    identity: OnceLock<control::Identity>,
}

impl fmt::Debug for BlockContentControl {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BlockContentControl")
            .field("start", &self.start)
            .field("length", &self.length)
            .finish_non_exhaustive()
    }
}

impl BlockContentControl {
    pub(crate) const fn new(
        source: Arc<Vec<u8>>,
        start: u32,
        length: u32,
        namespaces: NamespaceBindings,
        positions: RawPositions,
    ) -> Self {
        Self {
            source,
            start,
            length,
            namespaces,
            positions,
            identity: OnceLock::new(),
        }
    }

    fn identity(&self) -> Result<&control::Identity> {
        if let Some(identity) = self.identity.get() {
            return Ok(identity);
        }
        let identity = control::parse(self.xml_bytes(), &self.namespaces)?;
        Ok(self.identity.get_or_init(|| identity))
    }

    /// The control's `w:id`, if it has one.
    ///
    /// # Errors
    ///
    /// Returns an error when `w:sdtPr` is malformed: a repeated `w:id`,
    /// `w:tag` or `w:alias`, one without `w:val`, or an identifier that is
    /// not a 32-bit signed integer.
    pub(crate) fn id(&self) -> Result<Option<i32>> {
        Ok(self.identity()?.id)
    }

    /// The control's `w:tag`, if it has one.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed or duplicate content-control properties.
    pub fn tag(&self) -> Result<Option<&str>> {
        Ok(self.identity()?.tag.as_deref())
    }

    /// The control's `w:alias`, its friendly name, if it has one.
    ///
    /// # Errors
    ///
    /// Returns an error for malformed or duplicate content-control properties.
    pub fn alias(&self) -> Result<Option<&str>> {
        Ok(self.identity()?.alias.as_deref())
    }

    /// The blocks of the control's `w:sdtContent`, indexed like the body.
    ///
    /// # Errors
    ///
    /// Returns an error when the control's XML is malformed or a bound on
    /// the index is reached.
    pub fn blocks(&self) -> Result<Blocks> {
        Blocks::index(
            Arc::clone(&self.source),
            self.start,
            self.length,
            &self.namespaces,
            ContainerRoot::ContentControl,
            self.positions,
        )
    }

    /// The visible text of the control's content.
    ///
    /// # Errors
    ///
    /// Returns an error when the control's XML is malformed.
    pub fn text(&self) -> Result<String> {
        crate::paragraph::extract_word_text(&self.self_contained_xml()?)
    }

    /// The control's XML exactly as the visible view holds it.
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        range_bytes(&self.source, self.start, self.length)
    }

    /// The control as the inert block the historical element view returns.
    pub(crate) fn to_opaque_block(&self) -> OpaqueBlock {
        OpaqueBlock::from_arc_range(Arc::clone(&self.source), self.start, self.length)
    }

    /// The control's XML with every inherited namespace declaration written
    /// onto its `w:sdt`, so it parses standalone.
    ///
    /// # Errors
    ///
    /// Returns an error when a bound on the namespace rewrite is reached.
    pub fn self_contained_xml(&self) -> Result<Vec<u8>> {
        self_contained(self.xml_bytes(), &self.namespaces)
    }
}

impl fmt::Debug for BlockEntry {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str("BlockEntry { .. }")
    }
}

impl fmt::Debug for BlockIndex {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str("BlockIndex { .. }")
    }
}
