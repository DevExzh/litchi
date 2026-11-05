//! Per-story reads through a memoized story index (M1 L-DOC-R1, R2, R3).
//!
//! A Word binary document keeps its text as seven stories laid out back to
//! back in one character-position (CP) space ([MS-DOC] §2.3.1): the main
//! story, then footnotes, headers and footers, comments, endnotes, textboxes
//! and header textboxes. [`Document::paragraphs`] and [`Document::elements`]
//! read all of them as one untagged sequence and rebuild every paragraph on
//! every call. The readers here address one story at a time:
//!
//! - **Index.** The first read of a story scans that story's text once,
//!   starting at its own first character, and records where each paragraph
//!   ends and how the paragraphs group into top-level blocks: paragraphs
//!   outside tables, and top-level tables with every paragraph of the tables
//!   nested in them (the grouping of [`Document::elements`], one story at a
//!   time). The scan reads the paragraph marks and each paragraph's table
//!   depth from its own PAPX, and no character property, style, image or
//!   formula.
//! - **Values.** A block is built from its own paragraphs only, by the same
//!   extraction [`Document::paragraphs`] performs, so its cost is the
//!   block's size, and every paragraph, run, table, row and cell it returns
//!   carries its CP range.
//! - **Memo.** The document keeps the indexes it builds in a weighted
//!   clean-value cache in the sense of ADR 0005 ("clean parsed values are
//!   evictable; active handles pin them"), shaped after the DOCX block index
//!   and PPTX scene memos rather than a per-story `OnceLock`. The key is the
//!   story: a document's text and property tables never change after open,
//!   so its stories have one generation each. Every index is weighed; the
//!   cache retains at most [`STORY_INDEX_CACHE_MAX_BYTES`], evicting the
//!   least recently used index first, and returns an index heavier than that
//!   ceiling without retaining it. Eviction drops only the cache's share: a
//!   view that holds an index keeps it alive. A failed build is never
//!   retained, so it fails again, the same way, on the next request; a hit
//!   returns the index a miss would build, which test and debug builds
//!   re-derive and compare on every hit. [`Document::trim_story_indexes`]
//!   releases every index no view holds.
//!
//! Concurrent first reads of one story may each build its index; the first
//! to publish wins and the others return the published index.

use std::fmt;
use std::ops::{ControlFlow, Range};
use std::sync::{Mutex, MutexGuard};

use super::super::validation::{BlockGrouper, Placement, table_from_paragraphs};
use super::prelude::*;
use crate::body_text::Story;
use crate::cp_range::CpRange;

/// Logical ceiling on the story indexes one document retains: 16 MiB.
pub(crate) const STORY_INDEX_CACHE_MAX_BYTES: usize = 16 * 1024 * 1024;
/// Modeled bookkeeping per retained index (its slot, `Arc` header and the
/// index's own header), added to the index's tables.
const STORY_INDEX_ENTRY_OVERHEAD: usize = 128;

/// What a top-level block of a story is.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum StoryBlockKind {
    /// A paragraph outside every table.
    Paragraph,
    /// A top-level table, holding the paragraphs of any table nested in it.
    Table,
}

impl StoryBlockKind {
    const fn slot(self) -> usize {
        match self {
            Self::Paragraph => 0,
            Self::Table => 1,
        }
    }
}

/// One top-level block of a story, as the story's index records it.
///
/// Reading it builds nothing; [`StoryBlocks::element`] builds the block.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StoryBlock {
    kind: StoryBlockKind,
    position: usize,
    ordinal: usize,
    paragraphs: Range<usize>,
    cp_range: CpRange,
}

impl StoryBlock {
    /// Whether the block is a paragraph or a table.
    #[must_use]
    pub const fn kind(&self) -> StoryBlockKind {
        self.kind
    }

    /// The block's position among the story's blocks of every kind.
    #[must_use]
    pub const fn position(&self) -> usize {
        self.position
    }

    /// The block's position among the story's blocks of its kind: `N - 1`
    /// for `/body/p[N]` or `/body/tbl[N]`.
    #[must_use]
    pub const fn ordinal(&self) -> usize {
        self.ordinal
    }

    /// The block's paragraphs, as positions among every paragraph of the
    /// story ([`Document::story_paragraphs`]): one for a paragraph, every
    /// paragraph of a table and of the tables nested in it for a table.
    #[must_use]
    pub fn paragraph_range(&self) -> Range<usize> {
        self.paragraphs.clone()
    }

    /// The block's character positions: a paragraph with its mark, or a
    /// table through its last row-end mark.
    #[must_use]
    pub const fn cp_range(&self) -> CpRange {
        self.cp_range
    }
}

/// One recorded top-level block.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct BlockRecord {
    /// The block's first paragraph.
    first: u32,
    /// One past the block's last paragraph.
    end: u32,
    /// The block's position among the blocks of its kind.
    ordinal: u32,
    kind: StoryBlockKind,
}

/// The structure of one story: where its paragraphs end and how they group
/// into top-level blocks. It holds positions only, never text.
#[derive(Debug, PartialEq, Eq)]
pub(crate) struct StoryIndex {
    story: Story,
    /// The story's CP range, `None` when the FIB gives the story no text.
    range: Option<CpRange>,
    /// One past each paragraph's mark (or the end of its trailing text).
    /// Paragraphs tile the story: the first starts at the story's start and
    /// each later one where the previous one ends.
    paragraph_ends: Box<[u32]>,
    blocks: Box<[BlockRecord]>,
    /// Block positions of each kind, in story order, by
    /// [`StoryBlockKind::slot`].
    by_kind: [Box<[u32]>; 2],
}

impl StoryIndex {
    /// Scan `story` of `document` once.
    fn build(document: &Document, story: Story) -> Result<Self> {
        let range = document.story_range(story)?;
        let mut paragraph_ends = Vec::new();
        let mut blocks: Vec<BlockRecord> = Vec::new();
        let mut by_kind: [Vec<u32>; 2] = [Vec::new(), Vec::new()];
        if let Some(range) = range {
            let extractor = document.story_extractor(range);
            let pap = document.0.pap_bin_table.as_ref();
            let mut grouper = BlockGrouper::default();
            let mut failure = None;
            extractor.for_each_span::<()>(|start, end| {
                let paragraph = paragraph_ends.len();
                debug_assert_eq!(
                    start,
                    paragraph_ends.last().copied().unwrap_or(range.start),
                    "story paragraphs tile the story"
                );
                let depth = pap
                    .and_then(|table| table.properties_at(start))
                    .map_or(0, |properties| properties.table_depth());
                let (Ok(first), Ok(next)) =
                    (u32::try_from(paragraph), u32::try_from(paragraph + 1))
                else {
                    failure = Some("story paragraph count exceeds u32");
                    return ControlFlow::Break(());
                };
                paragraph_ends.push(end);
                let kind = match grouper.place(depth) {
                    Placement::Paragraph => StoryBlockKind::Paragraph,
                    Placement::TableStart => StoryBlockKind::Table,
                    Placement::TableContinue => {
                        if let Some(table) = blocks.last_mut() {
                            table.end = next;
                        }
                        return ControlFlow::Continue(());
                    },
                    Placement::Orphan => return ControlFlow::Continue(()),
                };
                let (Ok(position), Ok(ordinal)) = (
                    u32::try_from(blocks.len()),
                    u32::try_from(by_kind[kind.slot()].len()),
                ) else {
                    failure = Some("story block count exceeds u32");
                    return ControlFlow::Break(());
                };
                by_kind[kind.slot()].push(position);
                blocks.push(BlockRecord {
                    first,
                    end: next,
                    ordinal,
                    kind,
                });
                ControlFlow::Continue(())
            });
            if let Some(failure) = failure {
                return Err(PackageError::Corrupted(failure.to_string()));
            }
        }
        let [paragraph_blocks, table_blocks] = by_kind;
        Ok(Self {
            story,
            range,
            paragraph_ends: paragraph_ends.into_boxed_slice(),
            blocks: blocks.into_boxed_slice(),
            by_kind: [
                paragraph_blocks.into_boxed_slice(),
                table_blocks.into_boxed_slice(),
            ],
        })
    }

    /// The logical bytes the cache retains for this index.
    fn weight(&self) -> usize {
        size_of::<Self>()
            .saturating_add(self.paragraph_ends.len().saturating_mul(size_of::<u32>()))
            .saturating_add(self.blocks.len().saturating_mul(size_of::<BlockRecord>()))
            .saturating_add(
                self.by_kind
                    .iter()
                    .map(|positions| positions.len())
                    .sum::<usize>()
                    .saturating_mul(size_of::<u32>()),
            )
    }

    /// The CPs of paragraph `index`, its mark included.
    fn paragraph_span(&self, index: usize) -> Option<(u32, u32)> {
        let end = *self.paragraph_ends.get(index)?;
        let start = match index.checked_sub(1) {
            Some(previous) => *self.paragraph_ends.get(previous)?,
            None => self.range?.start,
        };
        Some((start, end))
    }

    fn block(&self, position: usize) -> Option<StoryBlock> {
        let record = *self.blocks.get(position)?;
        let first = usize::try_from(record.first).ok()?;
        let end = usize::try_from(record.end).ok()?;
        let (start_cp, _) = self.paragraph_span(first)?;
        let (_, end_cp) = self.paragraph_span(end.checked_sub(1)?)?;
        Some(StoryBlock {
            kind: record.kind,
            position,
            ordinal: usize::try_from(record.ordinal).ok()?,
            paragraphs: first..end,
            cp_range: CpRange::new(start_cp, end_cp)?,
        })
    }

    fn of_kind(&self, kind: StoryBlockKind) -> &[u32] {
        &self.by_kind[kind.slot()]
    }
}

/// The top-level blocks of one story — its paragraphs outside tables and its
/// top-level tables — read through the story's memoized index.
///
/// A view is cheap to clone: it borrows the document and shares the index.
/// Counting and locating blocks reads nothing; [`Self::paragraph`],
/// [`Self::table`] and [`Self::element`] build only the block they return.
///
/// ```no_run
/// use litchi_core::Position;
/// use litchi_doc::Package;
///
/// let package = Package::open("report.doc")?;
/// let document = package.document();
/// let body = document.body_blocks()?;
/// // `/body/p[3]`: the third paragraph outside the main story's tables.
/// if let Some(paragraph) = body.paragraph(Position::new(2))? {
///     println!("{}", paragraph.text()?);
/// }
/// // `/body/tbl[1]`: the first top-level table.
/// if let Some(table) = body.table(Position::new(0))? {
///     println!("{} rows", table.row_count()?);
/// }
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
#[derive(Clone)]
pub struct StoryBlocks<'d> {
    document: &'d Document,
    index: Arc<StoryIndex>,
}

impl fmt::Debug for StoryBlocks<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("StoryBlocks")
            .field("story", &self.index.story)
            .field("cp_range", &self.index.range)
            .field("blocks", &self.index.blocks.len())
            .field("paragraphs", &self.index.paragraph_ends.len())
            .finish_non_exhaustive()
    }
}

impl<'d> StoryBlocks<'d> {
    /// The story these blocks belong to.
    #[must_use]
    pub fn story(&self) -> Story {
        self.index.story
    }

    /// The story's character positions, `None` when the FIB gives the story
    /// no text.
    #[must_use]
    pub fn cp_range(&self) -> Option<CpRange> {
        self.index.range
    }

    /// The number of blocks of every kind.
    #[must_use]
    pub fn len(&self) -> usize {
        self.index.blocks.len()
    }

    /// Whether the story holds no block.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.index.blocks.is_empty()
    }

    /// The number of blocks of `kind`.
    #[must_use]
    pub fn count(&self, kind: StoryBlockKind) -> usize {
        self.index.of_kind(kind).len()
    }

    /// The number of paragraphs in the story, table paragraphs included:
    /// the length of [`Document::story_paragraphs`].
    #[must_use]
    pub fn paragraph_count(&self) -> usize {
        self.index.paragraph_ends.len()
    }

    /// The block at `position` in story order, counting every kind.
    #[must_use]
    pub fn get(&self, position: Position) -> Option<StoryBlock> {
        self.index.block(position.get())
    }

    /// Every block in story order.
    pub fn iter(&self) -> impl Iterator<Item = StoryBlock> + '_ {
        (0..self.index.blocks.len()).filter_map(|position| self.index.block(position))
    }

    /// The block of `kind` at `ordinal` among the blocks of that kind.
    #[must_use]
    pub fn get_kind(&self, kind: StoryBlockKind, ordinal: Position) -> Option<StoryBlock> {
        let position = *self.index.of_kind(kind).get(ordinal.get())?;
        self.index.block(usize::try_from(position).ok()?)
    }

    /// The text of the paragraph block at `ordinal`, its paragraph mark
    /// excluded, borrowed from the document: an index lookup that builds
    /// nothing. `None` when there is no such block or the block is a trailing
    /// fragment without a paragraph mark.
    #[must_use]
    pub fn paragraph_text(&self, ordinal: Position) -> Option<&'d str> {
        let block = self.get_kind(StoryBlockKind::Paragraph, ordinal)?;
        let (start, end) = self.index.paragraph_span(block.paragraphs.start)?;
        self.document
            .0
            .text_extractor
            .text_at_range(start, end)
            .strip_suffix('\r')
    }

    /// The paragraph block at `ordinal`: `/body/p[ordinal + 1]` of the main
    /// story, which counts only paragraphs outside tables.
    ///
    /// # Errors
    ///
    /// Returns the errors [`Document::paragraphs`] returns for the
    /// paragraph's properties and runs.
    pub fn paragraph(&self, ordinal: Position) -> Result<Option<Paragraph>> {
        match self.get_kind(StoryBlockKind::Paragraph, ordinal) {
            Some(block) => self.build_paragraph(block.paragraphs.start).map(Some),
            None => Ok(None),
        }
    }

    /// The top-level table at `ordinal`: `/body/tbl[ordinal + 1]` of the main
    /// story.
    ///
    /// # Errors
    ///
    /// Returns the errors [`Document::tables`] returns for the table's
    /// paragraphs.
    pub fn table(&self, ordinal: Position) -> Result<Option<Table>> {
        match self.get_kind(StoryBlockKind::Table, ordinal) {
            Some(block) => self.build_table(block.paragraphs).map(Some),
            None => Ok(None),
        }
    }

    /// The block at `position`, built.
    ///
    /// # Errors
    ///
    /// Returns the errors of [`Self::paragraph`] and [`Self::table`].
    pub fn element(&self, position: Position) -> Result<Option<crate::Element>> {
        let Some(block) = self.get(position) else {
            return Ok(None);
        };
        self.build(block).map(Some)
    }

    /// Every block, built, in story order: the story's share of
    /// [`Document::elements`].
    ///
    /// # Errors
    ///
    /// Returns the errors of [`Self::element`].
    pub fn elements(&self) -> Result<Vec<crate::Element>> {
        let mut elements = Vec::with_capacity(self.len());
        for block in self.iter() {
            elements.push(self.build(block)?);
        }
        Ok(elements)
    }

    /// Every top-level table, built, in story order.
    ///
    /// # Errors
    ///
    /// Returns the errors of [`Self::table`].
    pub fn tables(&self) -> Result<Vec<Table>> {
        let mut tables = Vec::with_capacity(self.count(StoryBlockKind::Table));
        for ordinal in 0..self.count(StoryBlockKind::Table) {
            if let Some(table) = self.table(Position::new(ordinal))? {
                tables.push(table);
            }
        }
        Ok(tables)
    }

    fn build(&self, block: StoryBlock) -> Result<crate::Element> {
        Ok(match block.kind {
            StoryBlockKind::Paragraph => {
                crate::Element::Paragraph(Box::new(self.build_paragraph(block.paragraphs.start)?))
            },
            StoryBlockKind::Table => {
                crate::Element::Table(Box::new(self.build_table(block.paragraphs)?))
            },
        })
    }

    fn build_paragraph(&self, index: usize) -> Result<Paragraph> {
        let mut paragraphs = self
            .document
            .build_story_paragraphs(&self.index, index..index + 1)?;
        paragraphs.pop().ok_or_else(|| {
            PackageError::Corrupted("story index names a missing paragraph".to_string())
        })
    }

    fn build_table(&self, paragraphs: Range<usize>) -> Result<Table> {
        let paragraphs = self
            .document
            .build_story_paragraphs(&self.index, paragraphs)?;
        table_from_paragraphs(paragraphs, 1)
            .ok_or_else(|| PackageError::Corrupted("story table block formed no row".to_string()))
    }
}

impl Document {
    /// The character positions of `story`, or `None` when the FIB gives the
    /// story no text.
    ///
    /// # Errors
    ///
    /// Returns [`PackageError::Corrupted`] when the FIB's story lengths
    /// overflow the CP space.
    pub fn story_range(&self, story: Story) -> Result<Option<CpRange>> {
        let counts = self.0.fib.character_counts();
        let slot = story.character_count_slot();
        let overflow = || PackageError::Corrupted("DOC story CP range overflows".to_string());
        let start = counts[..slot]
            .iter()
            .try_fold(0u32, |total, &count| total.checked_add(count))
            .ok_or_else(overflow)?;
        let length = counts[slot];
        if length == 0 {
            return Ok(None);
        }
        let end = start.checked_add(length).ok_or_else(overflow)?;
        Ok(CpRange::new(start, end))
    }

    /// The top-level blocks of `story`: its paragraphs outside tables and its
    /// top-level tables, read through the story's memoized index.
    ///
    /// The first read of a story scans that story's text once; later reads,
    /// from this view or another, find the index in the document's cache
    /// (see [`Self::story_index_diagnostics`]). A story the FIB gives no text
    /// has no block. A story that lies past the decoded text holds no
    /// paragraph either, where [`Self::paragraphs`] returns one fallback
    /// paragraph holding the whole text.
    ///
    /// # Errors
    ///
    /// Returns the errors of [`Self::story_range`].
    pub fn story_blocks(&self, story: Story) -> Result<StoryBlocks<'_>> {
        Ok(StoryBlocks {
            document: self,
            index: self.story_index(story)?,
        })
    }

    /// Every paragraph of `story` in order, table paragraphs included, each
    /// with its CP range: the story's share of [`Self::paragraphs`].
    ///
    /// # Errors
    ///
    /// Returns the errors of [`Self::story_blocks`] and of
    /// [`Self::paragraphs`].
    pub fn story_paragraphs(&self, story: Story) -> Result<Vec<Paragraph>> {
        let index = self.story_index(story)?;
        self.build_story_paragraphs(&index, 0..index.paragraph_ends.len())
    }

    /// The top-level blocks of the main story; see [`Self::story_blocks`].
    ///
    /// # Errors
    ///
    /// Returns the errors of [`Self::story_blocks`].
    pub fn body_blocks(&self) -> Result<StoryBlocks<'_>> {
        self.story_blocks(Story::Main)
    }

    /// The number of top-level blocks of the main story.
    ///
    /// # Errors
    ///
    /// Returns the errors of [`Self::story_blocks`].
    pub fn body_block_count(&self) -> Result<usize> {
        Ok(self.body_blocks()?.len())
    }

    /// The main story's paragraph outside tables at zero-based `index`:
    /// `/body/p[index + 1]`.
    ///
    /// Unlike [`Self::paragraph_at`], which counts every paragraph of every
    /// story (table cells included), this counts only the body's own
    /// paragraphs. It is the paragraph space of the main-story text editor
    /// ([`crate::body_text`]); the paragraph's
    /// [`content_cp_range`](Paragraph::content_cp_range) is the range that
    /// editor addresses.
    ///
    /// ```no_run
    /// use litchi_core::Position;
    /// use litchi_doc::Package;
    ///
    /// let package = Package::open("report.doc")?;
    /// let document = package.document();
    /// if let Some(paragraph) = document.body_paragraph(Position::new(0))? {
    ///     println!("{}: {:?}", paragraph.text()?, paragraph.cp_range());
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns the errors of [`StoryBlocks::paragraph`].
    pub fn body_paragraph(&self, index: Position) -> Result<Option<Paragraph>> {
        self.body_blocks()?.paragraph(index)
    }

    /// The main story's top-level table at zero-based `index`:
    /// `/body/tbl[index + 1]`.
    ///
    /// # Errors
    ///
    /// Returns the errors of [`StoryBlocks::table`].
    pub fn body_table(&self, index: Position) -> Result<Option<Table>> {
        self.body_blocks()?.table(index)
    }

    /// Counters and gauges of this document's story-index memo; see
    /// [`Self::story_blocks`]. Observing them reads nothing.
    #[must_use]
    pub fn story_index_diagnostics(&self) -> StoryIndexDiagnostics {
        self.0.story_indexes.diagnostics()
    }

    /// Release every memoized story index that no [`StoryBlocks`] view
    /// holds, returning how many were released.
    ///
    /// Views keep the indexes they hold. A later read of a released story
    /// scans it again; releasing changes no value.
    pub fn trim_story_indexes(&self) -> usize {
        self.0.story_indexes.trim_clean()
    }

    /// The memoized index of `story`.
    fn story_index(&self, story: Story) -> Result<Arc<StoryIndex>> {
        let index = self
            .0
            .story_indexes
            .get_or_build(story, || StoryIndex::build(self, story))?;
        #[cfg(debug_assertions)]
        debug_assert!(
            StoryIndex::build(self, story).is_ok_and(|fresh| fresh == *index),
            "a memoized story index differs from a fresh scan"
        );
        Ok(index)
    }

    /// A paragraph extractor over the story range `range`.
    fn story_extractor(&self, range: CpRange) -> ParagraphExtractor<'_> {
        ParagraphExtractor::new_with_story_range(
            self.0.text_extractor.text_shared(),
            self.0.text_extractor.cp_to_byte_shared(),
            self.0.pap_bin_table.as_ref(),
            self.0.chp_bin_table.as_ref(),
            (range.start, range.end),
            self.0.stylesheet.as_ref(),
        )
    }

    /// Build the paragraphs at `paragraphs` of the story `index` describes.
    fn build_story_paragraphs(
        &self,
        index: &StoryIndex,
        paragraphs: Range<usize>,
    ) -> Result<Vec<Paragraph>> {
        let Some(range) = index.range else {
            return Ok(Vec::new());
        };
        let extractor = self.story_extractor(range);
        let mut extracted = Vec::with_capacity(paragraphs.len());
        for paragraph in paragraphs {
            let (start, end) = index.paragraph_span(paragraph).ok_or_else(|| {
                PackageError::Corrupted("story index names a missing paragraph".to_string())
            })?;
            extracted.push(extractor.extract_span(start, end)?);
        }
        let mut output = Vec::with_capacity(extracted.len());
        self.convert_to_paragraphs(extracted, &mut output)?;
        Ok(output)
    }
}

/// Content-free counters and gauges of a document's story-index memo; see
/// [`Document::story_index_diagnostics`].
///
/// Counters accumulate over the document's life. Weights are logical: the
/// index tables plus a modeled per-entry overhead, not allocator
/// measurements.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub struct StoryIndexDiagnostics {
    /// Index requests answered from a retained index.
    pub hits: u64,
    /// Index requests that found no retained index and scanned the story.
    pub misses: u64,
    /// Indexes built after a miss and offered to the cache.
    pub builds: u64,
    /// Built indexes discarded because a concurrent read had already
    /// published one for the same story; the caller received that one.
    pub concurrent_builds: u64,
    /// Retained indexes evicted while no view held them.
    pub clean_evictions: u64,
    /// Retained indexes evicted while a view still held them; the index
    /// lives on with that view.
    pub pinned_evictions: u64,
    /// Retained indexes released by [`Document::trim_story_indexes`].
    pub trimmed: u64,
    /// Built indexes heavier than the whole ceiling, returned unretained.
    pub oversized_bypasses: u64,
    /// Indexes currently retained.
    pub retained_entries: usize,
    /// Logical weight of the retained indexes.
    pub retained_bytes: usize,
    /// Retained indexes a view also holds.
    pub pinned_entries: usize,
    /// The retention ceiling in logical bytes.
    pub max_bytes: usize,
}

/// One retained index.
struct Slot {
    index: Arc<StoryIndex>,
    weight: usize,
    /// The cache clock at the slot's last use.
    used: u64,
}

impl Slot {
    fn is_clean(&self) -> bool {
        Arc::strong_count(&self.index) == 1
    }
}

#[derive(Default)]
struct CacheState {
    slots: [Option<Slot>; 7],
    retained_bytes: usize,
    clock: u64,
    hits: u64,
    misses: u64,
    builds: u64,
    concurrent_builds: u64,
    clean_evictions: u64,
    pinned_evictions: u64,
    trimmed: u64,
    oversized_bypasses: u64,
}

impl CacheState {
    fn tick(&mut self) -> u64 {
        self.clock = self.clock.wrapping_add(1);
        self.clock
    }

    fn take(&mut self, slot: usize) -> Option<Slot> {
        let removed = self.slots.get_mut(slot)?.take()?;
        self.retained_bytes = self.retained_bytes.saturating_sub(removed.weight);
        Some(removed)
    }

    /// Evict the least recently used index; `false` when none is retained.
    fn evict_oldest(&mut self) -> bool {
        let oldest = self
            .slots
            .iter()
            .enumerate()
            .filter_map(|(slot, entry)| entry.as_ref().map(|entry| (slot, entry.used)))
            .min_by_key(|&(_, used)| used)
            .map(|(slot, _)| slot);
        let Some(removed) = oldest.and_then(|slot| self.take(slot)) else {
            return false;
        };
        if removed.is_clean() {
            self.clean_evictions = self.clean_evictions.saturating_add(1);
        } else {
            self.pinned_evictions = self.pinned_evictions.saturating_add(1);
        }
        true
    }
}

/// The story-index memo of one document; see the module documentation.
pub(in crate::document) struct StoryIndexCache {
    /// Logical byte ceiling; `0` retains nothing.
    max_bytes: usize,
    state: Mutex<CacheState>,
}

impl StoryIndexCache {
    pub(in crate::document) fn new() -> Self {
        Self::with_ceiling(STORY_INDEX_CACHE_MAX_BYTES)
    }

    fn with_ceiling(max_bytes: usize) -> Self {
        Self {
            max_bytes,
            state: Mutex::new(CacheState::default()),
        }
    }

    fn lock(&self) -> MutexGuard<'_, CacheState> {
        // Every critical section leaves the state consistent before any call
        // that could panic, so a poisoned lock still holds a usable state.
        self.state
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
    }

    /// The index of `story`, built by `build` on a miss.
    fn get_or_build(
        &self,
        story: Story,
        build: impl FnOnce() -> Result<StoryIndex>,
    ) -> Result<Arc<StoryIndex>> {
        let slot = story_slot(story);
        {
            let mut state = self.lock();
            let now = state.tick();
            if let Some(entry) = state.slots[slot].as_mut() {
                entry.used = now;
                let index = Arc::clone(&entry.index);
                state.hits = state.hits.saturating_add(1);
                return Ok(index);
            }
            state.misses = state.misses.saturating_add(1);
        }
        // Built outside the lock: a build reads only the immutable document,
        // and a failure is returned without being retained.
        let index = Arc::new(build()?);
        let mut state = self.lock();
        state.builds = state.builds.saturating_add(1);
        let now = state.tick();
        if let Some(entry) = state.slots[slot].as_mut() {
            entry.used = now;
            let published = Arc::clone(&entry.index);
            state.concurrent_builds = state.concurrent_builds.saturating_add(1);
            return Ok(published);
        }
        let weight = index.weight().saturating_add(STORY_INDEX_ENTRY_OVERHEAD);
        if weight > self.max_bytes {
            state.oversized_bypasses = state.oversized_bypasses.saturating_add(1);
            return Ok(index);
        }
        while state.retained_bytes.saturating_add(weight) > self.max_bytes {
            if !state.evict_oldest() {
                break;
            }
        }
        state.retained_bytes = state.retained_bytes.saturating_add(weight);
        state.slots[slot] = Some(Slot {
            index: Arc::clone(&index),
            weight,
            used: now,
        });
        Ok(index)
    }

    /// Release every retained index no view holds; returns how many.
    fn trim_clean(&self) -> usize {
        let mut state = self.lock();
        let mut released = 0usize;
        for slot in 0..state.slots.len() {
            if state.slots[slot].as_ref().is_some_and(Slot::is_clean) {
                state.take(slot);
                released = released.saturating_add(1);
            }
        }
        state.trimmed = state
            .trimmed
            .saturating_add(u64::try_from(released).unwrap_or(u64::MAX));
        released
    }

    fn diagnostics(&self) -> StoryIndexDiagnostics {
        let state = self.lock();
        let retained = state.slots.iter().flatten();
        StoryIndexDiagnostics {
            hits: state.hits,
            misses: state.misses,
            builds: state.builds,
            concurrent_builds: state.concurrent_builds,
            clean_evictions: state.clean_evictions,
            pinned_evictions: state.pinned_evictions,
            trimmed: state.trimmed,
            oversized_bypasses: state.oversized_bypasses,
            retained_entries: retained.clone().count(),
            retained_bytes: state.retained_bytes,
            pinned_entries: retained.filter(|entry| !entry.is_clean()).count(),
            max_bytes: self.max_bytes,
        }
    }
}

/// The cache slot of `story`.
const fn story_slot(story: Story) -> usize {
    match story {
        Story::Main => 0,
        Story::Footnote => 1,
        Story::Header => 2,
        Story::Comment => 3,
        Story::Endnote => 4,
        Story::Textbox => 5,
        Story::HeaderTextbox => 6,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn index(story: Story, paragraphs: usize) -> StoryIndex {
        let end = u32::try_from(paragraphs).unwrap();
        StoryIndex {
            story,
            range: CpRange::new(0, end),
            paragraph_ends: (1..=end).collect(),
            blocks: Box::new([]),
            by_kind: [Box::new([]), Box::new([])],
        }
    }

    #[test]
    fn hits_reuse_the_published_index_and_failures_are_not_retained() {
        let cache = StoryIndexCache::new();
        let failed = cache.get_or_build(Story::Main, || {
            Err(PackageError::Corrupted("first build fails".to_string()))
        });
        assert!(failed.is_err());
        let first = cache
            .get_or_build(Story::Main, || Ok(index(Story::Main, 3)))
            .unwrap();
        let second = cache
            .get_or_build(Story::Main, || {
                Err(PackageError::Corrupted(
                    "a hit must not rebuild".to_string(),
                ))
            })
            .unwrap();
        assert!(Arc::ptr_eq(&first, &second));
        let other = cache
            .get_or_build(Story::Footnote, || Ok(index(Story::Footnote, 1)))
            .unwrap();
        assert_eq!(other.story, Story::Footnote);
        let diagnostics = cache.diagnostics();
        assert_eq!(
            (diagnostics.hits, diagnostics.misses, diagnostics.builds),
            (1, 3, 2)
        );
        assert_eq!(diagnostics.retained_entries, 2);
        // The test still holds both indexes.
        assert_eq!(diagnostics.pinned_entries, 2);
        assert_eq!(
            diagnostics.retained_bytes,
            first.weight() + other.weight() + 2 * STORY_INDEX_ENTRY_OVERHEAD
        );
    }

    #[test]
    fn trimming_releases_only_indexes_no_view_holds() {
        let cache = StoryIndexCache::new();
        let held = cache
            .get_or_build(Story::Main, || Ok(index(Story::Main, 2)))
            .unwrap();
        drop(
            cache
                .get_or_build(Story::Header, || Ok(index(Story::Header, 2)))
                .unwrap(),
        );
        assert_eq!(cache.trim_clean(), 1);
        let diagnostics = cache.diagnostics();
        assert_eq!(diagnostics.retained_entries, 1);
        assert_eq!(diagnostics.trimmed, 1);
        let again = cache
            .get_or_build(Story::Main, || {
                Err(PackageError::Corrupted(
                    "the held index is retained".to_string(),
                ))
            })
            .unwrap();
        assert!(Arc::ptr_eq(&held, &again));
        drop((held, again));
        assert_eq!(cache.trim_clean(), 1);
        assert_eq!(cache.diagnostics().retained_bytes, 0);
    }

    #[test]
    fn the_ceiling_bounds_retention_without_refusing_a_read() {
        let small = index(Story::Main, 4).weight() + STORY_INDEX_ENTRY_OVERHEAD;
        let cache = StoryIndexCache::with_ceiling(small);
        // Heavier than the whole ceiling: returned, never retained.
        let oversized = cache
            .get_or_build(Story::Endnote, || Ok(index(Story::Endnote, 1000)))
            .unwrap();
        assert_eq!(oversized.paragraph_ends.len(), 1000);
        assert_eq!(cache.diagnostics().oversized_bypasses, 1);
        assert_eq!(cache.diagnostics().retained_entries, 0);
        // Two indexes that each fit but not together: the older one goes.
        let first = cache
            .get_or_build(Story::Main, || Ok(index(Story::Main, 4)))
            .unwrap();
        drop(first);
        let _second = cache
            .get_or_build(Story::Comment, || Ok(index(Story::Comment, 4)))
            .unwrap();
        let diagnostics = cache.diagnostics();
        assert_eq!(diagnostics.retained_entries, 1);
        assert_eq!(diagnostics.clean_evictions, 1);
        assert!(diagnostics.retained_bytes <= diagnostics.max_bytes);
    }

    #[test]
    fn story_slots_are_distinct() {
        let mut slots = Story::ALL.map(story_slot);
        slots.sort_unstable();
        assert_eq!(slots, [0, 1, 2, 3, 4, 5, 6]);
    }
}
