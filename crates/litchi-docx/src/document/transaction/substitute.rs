//! Formatting-preserving text substitution in WordprocessingML stories
//! (records 0274 and 0285): the one leaf primitive behind template merge and
//! find/replace.
//!
//! A [`Story`] is the text of the paragraphs of one story part, as a caller
//! plans substitutions over it: the main document (its body blocks, or the
//! elements [`Edit::story_within`] names), a header, a footer, the footnotes,
//! the endnotes or the comments. Each [`Paragraph`] reads as one string.
//! `w:t` text is replaceable; every other run content is fixed: a tab is
//! `\t`, a break `\n`, and anything without text (a drawing, a symbol, a note
//! reference, a field marker) is [`OBJECT`]. Field markers bracket each
//! field's [`Field::range`]. A paragraph's text is also split into
//! [`Segment`]s, one per `w:t` element, grouped so that a search never joins
//! text across a boundary the markup draws.
//!
//! Fields: the text of a field result (a complex field after its `separate`
//! marker, a `w:fldSimple`'s runs) is ordinary text, as OfficeCLI's find and
//! Word's own find read it. The text inside a field instruction (the cached
//! result of a field nested in it) is fixed content, so no substitution can
//! change an instruction. Field state is followed across paragraphs: a table
//! of contents whose result spans paragraphs reads the same in each of them.
//!
//! A [`Replacement`] names a byte range of one paragraph's text and its new
//! text. A range over `w:t` text only is rewritten in place: the new text
//! takes the place of the range's first character and keeps that run's
//! formatting, and the rest of the range is removed from the runs it spans,
//! which keep their formatting and every other child. A range equal to a
//! [`Field::is_replaceable`] field's range replaces the whole field with one
//! run holding the new text, formatted as the field's first result run. A
//! [`Replacement::formatted`] replacement also splits its new text (or, for
//! [`Replacement::keep`], the kept text of each run it covers) into runs of
//! its own that [`Edit::substitute_formatted`]'s formatter changes. Nothing
//! else in the story changes.
//!
//! Replacement text may hold tabs and line feeds; they are written as
//! `w:tab` and `w:br` beside the run's `w:t` elements.

use std::ops::Range;
use std::sync::Arc;

use super::formatting::{ElementPath, ElementStep};
use super::{
    Edit, Error, MAX_DOCUMENT_XML_BYTES, MAX_OPERATIONS, MAX_REPLACEMENT_TEXT_BYTES, Operation,
    ParagraphSplice, Refusal, Result, SplicePolicy, checked_end, checked_start, replace_ranges,
};
use crate::namespace::NamespaceBindings;
use crate::properties;

mod scan;
#[cfg(test)]
mod tests;
mod write;

use scan::{Phase, Scope, Start, xml_error};

/// The character that stands for run content without text: a drawing, an
/// embedded object, a symbol, a note reference, a field marker. It is never
/// part of a replaceable range.
pub const OBJECT: char = '\u{FFFC}';

/// How a field is written.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum FieldKind {
    /// One `w:fldSimple` element holding its instruction and result runs.
    Simple,
    /// Runs delimited by `w:fldChar` begin, separate and end markers, with
    /// the instruction in `w:instrText`.
    Complex,
}

/// One field of a paragraph.
#[derive(Clone)]
pub struct Field {
    range: Range<usize>,
    instruction: String,
    kind: FieldKind,
    replaceable: bool,
    /// The story bytes a replacement rewrites: the `w:fldSimple` element, or
    /// the runs from the begin marker's run to the end marker's run.
    markup: Range<usize>,
    /// The `w:rPr` bytes the replacement run copies.
    properties: Option<Range<usize>>,
    /// The qualified name of a run element (`w:r`), for the replacement.
    run_name: Vec<u8>,
}

impl Field {
    /// The field's range in its paragraph's text, from its first marker
    /// character through its last; a complex field that ends in a later
    /// paragraph runs to the end of the paragraph holding its `begin`.
    #[must_use]
    pub fn range(&self) -> Range<usize> {
        self.range.clone()
    }

    /// The field instruction (`MERGEFIELD Name \* MERGEFORMAT`), decoded and
    /// with its surrounding spaces kept.
    #[must_use]
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// How the field is written.
    #[must_use]
    pub const fn kind(&self) -> FieldKind {
        self.kind
    }

    /// Whether a [`Replacement`] of the field's range may replace it: its
    /// markup lies within one parent, holds no nested field, drawing, note
    /// reference, bookmark or other non-run content, and does not span
    /// paragraphs.
    #[must_use]
    pub const fn is_replaceable(&self) -> bool {
        self.replaceable
    }
}

impl std::fmt::Debug for Field {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Field")
            .field("range", &self.range)
            .field("kind", &self.kind)
            .field("replaceable", &self.replaceable)
            .finish_non_exhaustive()
    }
}

/// One `w:t` element's text in its paragraph, as [`Paragraph::segments`]
/// lists it.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Segment {
    range: Range<usize>,
    group: usize,
    scoped: bool,
}

impl Segment {
    /// The segment's bytes in [`Paragraph::text`].
    #[must_use]
    pub fn range(&self) -> Range<usize> {
        self.range.clone()
    }

    /// Segments with the same group are adjacent text that no boundary
    /// separates: tabs, breaks, field markers, drawings and other run
    /// content, and the edges of hyperlinks, content controls, fields and
    /// other inline containers start a new group. Bookmarks, proofing marks,
    /// comment ranges and rendered page breaks do not. Groups ascend.
    #[must_use]
    pub const fn group(&self) -> usize {
        self.group
    }

    /// Whether the segment lies inside the scope [`Edit::story_within`]
    /// read (always true for a whole story).
    #[must_use]
    pub const fn is_scoped(&self) -> bool {
        self.scoped
    }
}

/// A run of `w:t` text or one fixed character.
#[derive(Clone)]
struct Piece {
    /// Its characters in the paragraph text.
    text: Range<usize>,
    /// The `w:t` element, or `None` for a fixed character.
    node: Option<TextNode>,
}

#[derive(Clone)]
struct TextNode {
    /// The complete element.
    element: Range<usize>,
    /// The end of its start tag (after `>`).
    start_tag_end: usize,
    /// The qualified element name in the start tag.
    name: Range<usize>,
    /// Whether the start tag carries `xml:space`.
    space: bool,
    /// Index into the paragraph's runs.
    run: usize,
    group: usize,
    scoped: bool,
}

/// A run that holds text nodes, for splitting.
#[derive(Clone)]
struct RunSlot {
    /// The `w:r` element.
    element: Range<usize>,
    /// End of the run's start tag.
    open_end: usize,
    /// Start of the run's end tag.
    close_start: usize,
    /// The run's direct `w:rPr`.
    properties: Option<Range<usize>>,
    /// Index into the story's namespace scopes.
    bindings: usize,
}

/// One paragraph's text, fields and the markup behind them.
#[derive(Clone)]
pub struct Paragraph {
    text: String,
    fields: Vec<Field>,
    pieces: Vec<Piece>,
    runs: Vec<RunSlot>,
    /// The main-story body block holding the paragraph.
    block: Option<Block>,
    /// The `w:p` element in the story XML.
    element: Range<usize>,
    /// Index into the story's namespace scopes: the bindings in scope at the
    /// paragraph.
    bindings: usize,
    /// The complex fields open where the paragraph starts.
    entry: Vec<Phase>,
    item: Option<usize>,
    /// Text offsets of the [`OBJECT`]s that stand for alternate content.
    alternate: Vec<usize>,
}

impl Paragraph {
    /// The paragraph text: `w:t` text, `\t` for tabs, `\n` for breaks and
    /// [`OBJECT`] for other run content, field markers and text inside a
    /// field instruction.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// The fields that open in this paragraph, in order.
    #[must_use]
    pub fn fields(&self) -> &[Field] {
        &self.fields
    }

    /// Whether `range` holds only replaceable `w:t` text.
    #[must_use]
    pub fn is_text(&self, range: Range<usize>) -> bool {
        !range.is_empty()
            && range.end <= self.text.len()
            && self.text.is_char_boundary(range.start)
            && self.text.is_char_boundary(range.end)
            && self.pieces[self
                .pieces
                .partition_point(|piece| piece.text.end <= range.start)..]
                .iter()
                .take_while(|piece| piece.text.start < range.end)
                .all(|piece| piece.node.is_some())
    }

    /// The `w:t` text of the paragraph, one [`Segment`] per element, in
    /// order.
    pub fn segments(&self) -> impl Iterator<Item = Segment> + '_ {
        self.pieces.iter().filter_map(|piece| {
            piece.node.as_ref().map(|node| Segment {
                range: piece.text.clone(),
                group: node.group,
                scoped: node.scoped,
            })
        })
    }

    /// The byte offsets in [`Self::text`] of the [`OBJECT`]s that stand for
    /// `mc:AlternateContent` wrappers held by the paragraph's runs (a
    /// drawing or text box with a fallback; record 0307), in order.
    ///
    /// Each wrapper is one opaque object: readers show the text of the
    /// branch the projection selects, which the story never lists, so a
    /// caller that matches text the reader shows can tell where that text
    /// would sit.
    #[must_use]
    pub fn alternate_content(&self) -> &[usize] {
        &self.alternate
    }

    /// For a part story, the position of the part root's child holding the
    /// paragraph among all the root's element children: the comment,
    /// footnote or endnote it belongs to. `None` for the main document.
    #[must_use]
    pub const fn item(&self) -> Option<usize> {
        self.item
    }
}

impl std::fmt::Debug for Paragraph {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Paragraph")
            .field("bytes", &self.text.len())
            .field("fields", &self.fields.len())
            .finish_non_exhaustive()
    }
}

/// A direct child of `w:body` that an [`ElementPath`] addresses.
#[derive(Clone, PartialEq, Eq)]
struct Block {
    step: ElementStep,
    range: Range<usize>,
}

/// The paragraphs of one story, in document order (a paragraph nested in a
/// text box comes before the paragraph holding it).
///
/// A main-document story lists the paragraphs inside the body's
/// paragraphs, tables and block content controls (text boxes included);
/// paragraphs in other body children (custom XML blocks, alternate
/// content) are not listed. A part story lists every paragraph of the part.
/// Neither lists anything inside an `mc:AlternateContent` wrapper, whose
/// branches differ: a run holding one reads it as one [`OBJECT`]
/// ([`Paragraph::alternate_content`]; record 0307).
#[derive(Clone)]
pub struct Story {
    part: Option<String>,
    xml: Arc<Vec<u8>>,
    paragraphs: Vec<Paragraph>,
    bindings: Vec<NamespaceBindings>,
}

impl Story {
    /// The part name of a header, footer, notes or comments story; `None`
    /// for the main document.
    #[must_use]
    pub fn part(&self) -> Option<&str> {
        self.part.as_deref()
    }

    /// The paragraphs.
    #[must_use]
    pub fn paragraphs(&self) -> &[Paragraph] {
        &self.paragraphs
    }
}

impl std::fmt::Debug for Story {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Story")
            .field("main", &self.part.is_none())
            .field("paragraphs", &self.paragraphs.len())
            .finish_non_exhaustive()
    }
}

/// What a story part holds.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum StoryKind {
    /// A header.
    Header,
    /// A footer.
    Footer,
    /// The footnotes.
    Footnotes,
    /// The endnotes.
    Endnotes,
    /// The comments.
    Comments,
}

/// One story part the main document relates, as
/// [`crate::document::Edit::story_parts`] lists it.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StoryPart {
    name: String,
    kind: StoryKind,
}

impl StoryPart {
    pub(crate) const fn new(name: String, kind: StoryKind) -> Self {
        Self { name, kind }
    }

    /// The part name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// What the part holds.
    #[must_use]
    pub const fn kind(&self) -> StoryKind {
        self.kind
    }
}

/// New text for a range of one paragraph's text.
#[derive(Clone, PartialEq, Eq)]
pub struct Replacement {
    paragraph: usize,
    range: Range<usize>,
    /// `None` keeps the range's text ([`Self::keep`]).
    text: Option<String>,
    formatted: bool,
}

impl Replacement {
    /// Replace `range` of paragraph `paragraph` (an index into
    /// [`Story::paragraphs`]) with `text`.
    #[must_use]
    pub fn new(paragraph: usize, range: Range<usize>, text: impl Into<String>) -> Self {
        Self {
            paragraph,
            range,
            text: Some(text.into()),
            formatted: false,
        }
    }

    /// Keep the text of `range` of paragraph `paragraph` and format it in
    /// place: each run the range covers gives its share a run of its own
    /// for [`Edit::substitute_formatted`]'s formatter (a kept range is
    /// always [`Self::formatted`]). The range must be `w:t` text.
    #[must_use]
    pub const fn keep(paragraph: usize, range: Range<usize>) -> Self {
        Self {
            paragraph,
            range,
            text: None,
            formatted: true,
        }
    }

    /// Split the new text (or, for [`Self::keep`], the kept text of each
    /// run the range covers) into runs of its own, each a copy of the run
    /// it came from, for [`Edit::substitute_formatted`]'s formatter. The
    /// range must be `w:t` text: a whole field is never formatted.
    #[must_use]
    pub const fn formatted(mut self) -> Self {
        self.formatted = true;
        self
    }

    /// The paragraph index.
    #[must_use]
    pub const fn paragraph(&self) -> usize {
        self.paragraph
    }

    /// The replaced range of the paragraph text.
    #[must_use]
    pub fn range(&self) -> Range<usize> {
        self.range.clone()
    }

    /// The new text; `None` for a range whose text is kept.
    #[must_use]
    pub fn text(&self) -> Option<&str> {
        self.text.as_deref()
    }

    fn replaced(&self) -> Option<&str> {
        self.text.as_deref()
    }
}

impl std::fmt::Debug for Replacement {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Replacement")
            .field("paragraph", &self.paragraph)
            .field("range", &self.range)
            .field("bytes", &self.text.as_ref().map(String::len))
            .field("formatted", &self.formatted)
            .finish()
    }
}

/// Content-free counts of one substitution.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Outcome {
    paragraphs: usize,
    texts: usize,
    fields: usize,
    formatted: usize,
}

impl Outcome {
    /// Paragraphs whose markup changed.
    #[must_use]
    pub const fn paragraphs(self) -> usize {
        self.paragraphs
    }

    /// Text ranges replaced or kept.
    #[must_use]
    pub const fn texts(self) -> usize {
        self.texts
    }

    /// Fields replaced by their text.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Runs split out and handed to the formatter.
    #[must_use]
    pub const fn formatted(self) -> usize {
        self.formatted
    }

    /// Whether nothing changed.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.paragraphs == 0
    }
}

// ---------------------------------------------------------------------------
// Reading stories
// ---------------------------------------------------------------------------

fn whole(xml: &[u8], main: bool) -> Start {
    Start {
        range: 0..xml.len(),
        main,
        inherited: Arc::from([]),
        fields: Vec::new(),
        block: None,
        scopes: vec![Scope::All],
        stop: None,
    }
}

/// The paragraphs of a whole story part (tests read them directly).
#[cfg(test)]
pub(crate) fn scan_paragraphs(xml: &[u8], main: bool) -> Result<Vec<Paragraph>> {
    Ok(scan::scan(xml, whole(xml, main))?.paragraphs)
}

/// Scan a story part's payload.
pub(crate) fn part_story(part: &str, xml: Arc<Vec<u8>>) -> Result<Story> {
    if xml.len() > MAX_DOCUMENT_XML_BYTES {
        return Err(Error::Limit {
            kind: crate::document::LimitKind::XmlBytes,
            maximum: MAX_DOCUMENT_XML_BYTES,
            observed: xml.len(),
        });
    }
    let scanned = scan::scan(&xml, whole(&xml, false))?;
    Ok(Story {
        part: Some(part.to_owned()),
        xml,
        paragraphs: scanned.paragraphs,
        bindings: scanned.bindings,
    })
}

/// Read back a rewritten part: every changed paragraph must read as
/// planned, and every other one as before.
fn read_back_part(before: &Story, after: &[u8], expected: &[(usize, String)]) -> Result<()> {
    let rescanned = scan::scan(after, whole(after, false))?.paragraphs;
    if rescanned.len() != before.paragraphs.len() {
        return Err(xml_error("substituted story lost or gained paragraphs"));
    }
    let mut expected = expected.iter().peekable();
    for (index, (old, new)) in before.paragraphs.iter().zip(&rescanned).enumerate() {
        let want = match expected.peek() {
            Some((changed, text)) if *changed == index => {
                expected.next();
                text.as_str()
            },
            _ => old.text.as_str(),
        };
        if new.text != want {
            return Err(xml_error("substituted story failed text readback"));
        }
    }
    Ok(())
}

/// Rewrite a story part's payload, or `None` when nothing changes.
pub(crate) fn substitute_part<F: FnMut(&mut properties::Run)>(
    story: &Story,
    current: &[u8],
    replacements: &[Replacement],
    mut format: F,
) -> Result<(Option<Vec<u8>>, Outcome)> {
    if story.part.is_none() {
        return Err(refused(0, Refusal::StoryReference));
    }
    if current != story.xml.as_slice() {
        return Err(Error::PatchConflict);
    }
    let plan = write::plan(story, replacements, &mut format)?;
    if plan.changes.is_empty() {
        return Ok((None, plan.outcome));
    }
    let mut expected: Vec<(usize, String)> = Vec::with_capacity(plan.changes.len());
    let mut splices = Vec::new();
    for change in plan.changes {
        expected.push((change.paragraph, change.text));
        splices.extend(change.splices);
    }
    splices.sort_by_key(|splice| splice.0);
    let after = replace_ranges(&story.xml, &splices)?;
    expected.sort_by_key(|(paragraph, _)| *paragraph);
    read_back_part(story, &after, &expected)?;
    Ok((Some(after), plan.outcome))
}

/// Above this many changed paragraphs, and a quarter of the story's, the
/// readback rescans the whole document once instead of each paragraph on
/// its own.
const FRAGMENT_READBACKS: usize = 32;

/// `source` with `replacements` (ascending, disjoint) spliced in.
fn replace_shared_ranges(
    source: &[u8],
    replacements: &[(usize, usize, Arc<Vec<u8>>)],
) -> Result<Vec<u8>> {
    let mut length = source.len();
    let mut previous = 0usize;
    for (start, end, bytes) in replacements {
        if *start < previous || start > end || *end > source.len() {
            return Err(xml_error("story rewrite ranges are invalid or overlap"));
        }
        length = length - (end - start) + bytes.len();
        previous = *end;
    }
    if length > MAX_DOCUMENT_XML_BYTES {
        return Err(Error::Limit {
            kind: crate::document::LimitKind::ProjectedXmlBytes,
            maximum: MAX_DOCUMENT_XML_BYTES,
            observed: length,
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(length)
        .map_err(|source| crate::Error::Allocation {
            resource: "document transaction XML",
            source,
        })?;
    let mut cursor = 0usize;
    for (start, end, bytes) in replacements {
        output.extend_from_slice(&source[cursor..*start]);
        output.extend_from_slice(bytes);
        cursor = *end;
    }
    output.extend_from_slice(&source[cursor..]);
    Ok(output)
}

/// Read the rewritten main document back: every changed paragraph must
/// read as planned. Few changes rescan their own paragraphs, with the
/// namespace and field state they start in; many rescan the document once.
fn read_back_main(
    story: &Story,
    xml: &[u8],
    shifts: &write::Shifts,
    checks: &[(usize, String)],
) -> Result<()> {
    if checks.len() > FRAGMENT_READBACKS && checks.len() * 4 > story.paragraphs.len() {
        let mut reread = scan::scan(xml, whole(xml, true))?.paragraphs;
        reread.sort_by_key(|paragraph| paragraph.element.start);
        for (paragraph, text) in checks {
            let start = shifts.shift(story.paragraphs[*paragraph].element.start);
            let found = reread
                .binary_search_by_key(&start, |paragraph| paragraph.element.start)
                .ok()
                .and_then(|at| reread.get(at));
            if found.is_none_or(|found| found.text != *text) {
                return Err(xml_error("substituted story failed text readback"));
            }
        }
        return Ok(());
    }
    for (paragraph, text) in checks {
        let paragraph = &story.paragraphs[*paragraph];
        let start = shifts.shift(paragraph.element.start);
        let end = shifts.shift(paragraph.element.end);
        let inherited = story
            .bindings
            .get(paragraph.bindings)
            .cloned()
            .ok_or_else(|| xml_error("a paragraph lost its namespace scope"))?;
        let reread = scan::scan(
            xml,
            Start {
                range: start..end,
                main: false,
                inherited,
                fields: paragraph.entry.clone(),
                block: None,
                scopes: vec![Scope::Paragraph(start)],
                stop: None,
            },
        )?;
        if reread
            .paragraphs
            .iter()
            .find(|found| found.element.start == start)
            .is_none_or(|found| found.text != *text)
        {
            return Err(xml_error("substituted story failed text readback"));
        }
    }
    Ok(())
}

fn refused(paragraph: usize, reason: Refusal) -> Error {
    Error::Refused {
        position: paragraph,
        reason,
    }
}

/// Whether `bytes` mention a field character.
fn mentions_field_characters(bytes: &[u8]) -> bool {
    match std::str::from_utf8(bytes) {
        Ok(text) => text.contains("fldChar"),
        Err(_) => bytes.windows(7).any(|window| window == b"fldChar"),
    }
}

/// The staged main-document bytes for a story: shared when the snapshot
/// owns them, else copied.
fn story_xml(snapshot: &super::Snapshot) -> Arc<Vec<u8>> {
    snapshot
        .xml
        .shared()
        .map_or_else(|| Arc::new(snapshot.xml_bytes().to_vec()), Arc::clone)
}

/// Whether `story` was read from the staged bytes of `snapshot`.
fn is_current(snapshot: &super::Snapshot, story: &Story) -> bool {
    snapshot
        .xml
        .shared()
        .is_some_and(|xml| Arc::ptr_eq(xml, &story.xml))
        || snapshot.xml_bytes() == story.xml.as_slice()
}

impl Edit {
    /// The main document's story, over the staged document.
    ///
    /// # Errors
    ///
    /// Returns an XML error for malformed markup.
    pub fn story(&self) -> Result<Story> {
        let xml = story_xml(&self.projected);
        let scanned = scan::scan(&xml, whole(&xml, true))?;
        Ok(Story {
            part: None,
            xml,
            paragraphs: scanned.paragraphs,
            bindings: scanned.bindings,
        })
    }

    /// The main document's story within the elements `scopes` name, over
    /// the staged document: a paragraph lists itself (not the text boxes
    /// nested in it); a table, row, cell or block content control lists
    /// every paragraph inside it; a run, hyperlink or inline content control
    /// lists its paragraph with only its own text [`Segment::is_scoped`].
    /// An empty path names the whole body, as [`Self::story`]. A paragraph
    /// named twice is listed once, with the widest scope.
    ///
    /// Only the scoped part of the document is read, with the field state
    /// in effect where it starts.
    ///
    /// # Errors
    ///
    /// [`Refusal::BlockNotFound`] for a path that names nothing, and an XML
    /// error for malformed markup.
    pub fn story_within(&self, scopes: &[ElementPath]) -> Result<Story> {
        if scopes.is_empty() || scopes.iter().any(|path| path.steps().is_empty()) {
            return self.story();
        }
        let snapshot = &self.projected;
        let xml = story_xml(snapshot);
        let mut listing = Vec::with_capacity(scopes.len());
        // The direct-body paragraphs every scope lies in, when it does.
        let mut fragments: Vec<(litchi_core::Position, Range<usize>)> = Vec::new();
        let mut direct = true;
        let mut first = usize::MAX;
        let mut stop = 0usize;
        for path in scopes {
            let steps = path.steps();
            let range = super::formatting::locate_element(snapshot, path)?;
            let block = super::formatting::locate_element(snapshot, &ElementPath::new([steps[0]]))?;
            first = first.min(block.start);
            stop = stop.max(block.end);
            let paragraph_at = steps
                .iter()
                .rposition(|step| matches!(step, ElementStep::Paragraph(_)));
            let scope = match (steps.last(), paragraph_at) {
                (Some(ElementStep::Paragraph(_)), _) => Scope::Paragraph(range.start),
                (Some(ElementStep::Run(_) | ElementStep::Hyperlink(_)), Some(at))
                | (Some(ElementStep::ContentControl(_)), Some(at)) => {
                    let paragraph = super::formatting::locate_element(
                        snapshot,
                        &ElementPath::new(&steps[..=at]),
                    )?;
                    Scope::Inline {
                        paragraph: paragraph.start,
                        range,
                    }
                },
                _ => Scope::Within(range),
            };
            match (steps.first(), paragraph_at) {
                (Some(ElementStep::Paragraph(position)), Some(0)) => {
                    if !fragments.iter().any(|(known, _)| known == position) {
                        fragments.push((*position, block));
                    }
                },
                _ => direct = false,
            }
            listing.push(scope);
        }
        // Scopes inside direct-body paragraphs, with no field character
        // before them: each paragraph is scanned on its own.
        let before = xml.get(..first).unwrap_or_default();
        let scanned = if direct && !mentions_field_characters(before) {
            fragments.sort_by_key(|(_, range)| range.start);
            let mut paragraphs = Vec::new();
            let mut bindings = Vec::new();
            for (position, range) in fragments {
                let inherited = snapshot
                    .paragraph_namespaces
                    .get(position.get())
                    .cloned()
                    .ok_or_else(|| xml_error("a paragraph has no namespace snapshot"))?;
                let mut scanned = scan::scan(
                    &xml,
                    Start {
                        range: range.clone(),
                        main: true,
                        inherited,
                        fields: Vec::new(),
                        block: Some(Block {
                            step: ElementStep::Paragraph(position),
                            range,
                        }),
                        scopes: listing.clone(),
                        stop: None,
                    },
                )?;
                let offset = bindings.len();
                for paragraph in &mut scanned.paragraphs {
                    paragraph.bindings += offset;
                    for run in &mut paragraph.runs {
                        run.bindings += offset;
                    }
                }
                bindings.append(&mut scanned.bindings);
                paragraphs.append(&mut scanned.paragraphs);
            }
            scan::Scanned {
                paragraphs,
                bindings,
            }
        } else {
            let mut start = whole(&xml, true);
            start.scopes = listing;
            start.stop = Some(stop);
            scan::scan(&xml, start)?
        };
        Ok(Story {
            part: None,
            xml,
            paragraphs: scanned.paragraphs,
            bindings: scanned.bindings,
        })
    }

    /// Apply `replacements` to the main document's story, which must have
    /// been read from this edit's staged document. A [`Replacement::formatted`]
    /// replacement isolates its runs without changing their properties; see
    /// [`Self::substitute_formatted`].
    ///
    /// Each changed body block (paragraph, table or block content control)
    /// is recorded as one exact element replacement, all spliced at once.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] for a story of other bytes, a refusal
    /// for an invalid, overlapping or non-text range, a field that is not
    /// replaceable or text a `w:t` cannot hold, and a limit error for too
    /// many changed blocks or too much replacement text. The edit is
    /// unchanged on failure.
    pub fn substitute(&mut self, story: &Story, replacements: &[Replacement]) -> Result<Outcome> {
        self.substitute_formatted(story, replacements, |_run| {})
    }

    /// [`Self::substitute`], with `format` changing the direct properties
    /// of every run a [`Replacement::formatted`] replacement splits out.
    ///
    /// ```
    /// use litchi_docx::document::substitute::Replacement;
    ///
    /// let mut builder = litchi_docx::Builder::new()?;
    /// builder.document()?.add_paragraph_with_text("say hi");
    /// let package = builder.build()?;
    /// let mut edit = package.edit()?;
    /// let story = edit.story()?;
    /// let outcome = edit.substitute_formatted(
    ///     &story,
    ///     &[Replacement::new(0, 4..6, "hello").formatted()],
    ///     |run| run.bold = Some(true),
    /// )?;
    /// assert_eq!(outcome.formatted(), 1);
    /// let package = edit.commit()?.into_package();
    /// assert_eq!(package.document()?.text()?, "say hello");
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// As [`Self::substitute`], and a refusal for run properties the
    /// formatter sets that the run cannot hold.
    pub fn substitute_formatted<F: FnMut(&mut properties::Run)>(
        &mut self,
        story: &Story,
        replacements: &[Replacement],
        mut format: F,
    ) -> Result<Outcome> {
        if story.part.is_some() {
            return Err(refused(0, Refusal::StoryReference));
        }
        self.ensure_unmanaged("substitute")?;
        if !is_current(&self.projected, story) {
            return Err(Error::PatchConflict);
        }
        let plan = write::plan(story, replacements, &mut format)?;
        if plan.changes.is_empty() {
            return Ok(plan.outcome);
        }
        let shifts = write::Shifts::of(&plan.changes);
        // Group the changes by body block, keeping each paragraph's text for
        // the readback.
        let mut checks: Vec<(usize, String)> = Vec::with_capacity(plan.changes.len());
        let mut blocks: Vec<(Block, Vec<(usize, usize, Vec<u8>)>)> = Vec::new();
        let mut known: std::collections::HashMap<usize, usize> = std::collections::HashMap::new();
        for change in plan.changes {
            let block = story.paragraphs[change.paragraph]
                .block
                .as_ref()
                .ok_or_else(|| xml_error("a main-story paragraph has no body block"))?;
            let at = match blocks.last() {
                Some((last, _)) if last.range == block.range => blocks.len() - 1,
                _ => {
                    if let Some(&at) = known.get(&block.range.start) {
                        at
                    } else {
                        known.insert(block.range.start, blocks.len());
                        blocks.push((block.clone(), Vec::new()));
                        blocks.len() - 1
                    }
                },
            };
            let range = blocks[at].0.range.clone();
            for (start, end, bytes) in change.splices {
                if start < range.start || end > range.end {
                    return Err(xml_error("story splice leaves its body block"));
                }
                blocks[at]
                    .1
                    .push((start - range.start, end - range.start, bytes));
            }
            checks.push((change.paragraph, change.text));
        }
        blocks.sort_by_key(|(block, _)| block.range.start);
        if self.operations.len().saturating_add(blocks.len()) > MAX_OPERATIONS {
            return Err(Error::Limit {
                kind: crate::document::LimitKind::Operations,
                maximum: MAX_OPERATIONS,
                observed: self.operations.len().saturating_add(blocks.len()),
            });
        }
        let mut replaced = Vec::with_capacity(blocks.len());
        let mut ranges = Vec::with_capacity(blocks.len());
        let mut text_bytes = self.replacement_text_bytes;
        for (block, mut local) in blocks {
            local.sort_by_key(|splice| splice.0);
            let before = &story.xml[block.range.clone()];
            let after = replace_ranges(before, &local)?;
            if after == before {
                continue;
            }
            text_bytes = text_bytes.checked_add(after.len()).ok_or(Error::Limit {
                kind: crate::document::LimitKind::ReplacementTextBytes,
                maximum: MAX_REPLACEMENT_TEXT_BYTES,
                observed: usize::MAX,
            })?;
            if text_bytes > MAX_REPLACEMENT_TEXT_BYTES {
                return Err(Error::Limit {
                    kind: crate::document::LimitKind::ReplacementTextBytes,
                    maximum: MAX_REPLACEMENT_TEXT_BYTES,
                    observed: text_bytes,
                });
            }
            let after = Arc::new(after);
            ranges.push((block.range.start, block.range.end, Arc::clone(&after)));
            replaced.push((block, Arc::new(before.to_vec()), after));
        }
        if replaced.is_empty() {
            return Ok(Outcome::default());
        }
        let xml = replace_shared_ranges(&story.xml, &ranges)?;
        read_back_main(story, &xml, &shifts, &checks)?;
        // Only direct-body paragraphs changed, a small part of the document:
        // derive the layout. Proving many splices costs more than the
        // rescan they would save.
        let spliced: usize = replaced
            .iter()
            .map(|(_, before, after)| before.len() + after.len())
            .sum();
        let candidate = if spliced < xml.len() / 2
            && replaced
                .iter()
                .all(|(block, _, _)| matches!(block.step, ElementStep::Paragraph(_)))
        {
            let splices: Vec<ParagraphSplice<'_>> = ranges
                .iter()
                .map(|(start, end, replacement)| ParagraphSplice {
                    start: *start,
                    end: *end,
                    replacement: replacement.as_slice(),
                })
                .collect();
            self.projected_with_splices_under(xml, &splices, SplicePolicy::BoundedGrowth)?
        } else {
            self.projected_with_xml(xml)?
        };
        for (block, _, after) in &replaced {
            let located = match block.step {
                ElementStep::Paragraph(position) => {
                    let range = candidate.range(position)?;
                    checked_start(range, "paragraph")?..checked_end(range, "paragraph")?
                },
                step => super::formatting::locate_element(&candidate, &ElementPath::new([step]))?,
            };
            if candidate.xml_bytes().get(located) != Some(after.as_slice()) {
                return Err(xml_error("substituted block failed exact readback"));
            }
        }
        self.operations
            .try_reserve(replaced.len())
            .map_err(|source| crate::Error::Allocation {
                resource: "document operation metadata",
                source,
            })?;
        for (block, before, after) in replaced {
            self.operations.push(Operation::ReplaceElement {
                path: ElementPath::new([block.step]),
                before,
                after,
            });
        }
        self.replacement_text_bytes = text_bytes;
        self.projected = candidate;
        Ok(plan.outcome)
    }
}
