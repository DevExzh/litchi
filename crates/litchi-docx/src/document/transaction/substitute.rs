//! Formatting-preserving text substitution in WordprocessingML stories
//! (record 0274).
//!
//! A [`Story`] is the text of every paragraph of one story part, as a caller
//! plans substitutions over it: the main document (its body blocks), a
//! header, a footer, the footnotes, the endnotes or the comments. Each
//! [`Paragraph`] reads as one string. `w:t` text is replaceable; every other
//! run content is fixed: a tab is `\t`, a break `\n`, and anything without
//! text (a drawing, a symbol, a note reference, a field marker) is
//! [`OBJECT`]. Field markers bracket each field's [`Field::range`].
//!
//! A [`Replacement`] names a byte range of one paragraph's text and its new
//! text. A range over `w:t` text only is rewritten in place: the new text
//! takes the place of the range's first character and keeps that run's
//! formatting, and the rest of the range is removed from the runs it spans,
//! which keep their formatting and every other child. A range equal to a
//! [`Field::is_replaceable`] field's range replaces the whole field with one
//! run holding the new text, formatted as the field's first result run.
//! Nothing else in the story changes.
//!
//! Replacement text may hold tabs and line feeds; they are written as
//! `w:tab` and `w:br` beside the run's `w:t` elements.

use std::ops::Range;
use std::sync::Arc;

use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use super::formatting::{ElementPath, ElementStep};
use super::{
    Edit, Error, MAX_DOCUMENT_DEPTH, MAX_DOCUMENT_XML_BYTES, MAX_OPERATIONS,
    MAX_REPLACEMENT_TEXT_BYTES, Operation, Refusal, Result, decode_text_fragment, replace_ranges,
};
use crate::namespace::is_wordprocessing_namespace;

#[cfg(test)]
mod tests;

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
    /// character through its last.
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
}

/// One paragraph's text, fields and the markup behind them.
#[derive(Clone)]
pub struct Paragraph {
    text: String,
    fields: Vec<Field>,
    pieces: Vec<Piece>,
    /// The main-story body block holding the paragraph.
    block: Option<Block>,
}

impl Paragraph {
    /// The paragraph text: `w:t` text, `\t` for tabs, `\n` for breaks and
    /// [`OBJECT`] for other run content and field markers.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// The fields that open and close in this paragraph, in order.
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

/// The paragraphs of one story, in document order.
///
/// A main-document story lists the paragraphs inside the body's
/// paragraphs, tables and block content controls (text boxes included);
/// paragraphs in other body children (custom XML blocks, alternate
/// content) are not listed. A part story lists every paragraph of the part.
#[derive(Clone)]
pub struct Story {
    part: Option<String>,
    xml: Arc<Vec<u8>>,
    paragraphs: Vec<Paragraph>,
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

/// New text for a range of one paragraph's text.
#[derive(Clone, PartialEq, Eq)]
pub struct Replacement {
    paragraph: usize,
    range: Range<usize>,
    text: String,
}

impl Replacement {
    /// Replace `range` of paragraph `paragraph` (an index into
    /// [`Story::paragraphs`]) with `text`.
    #[must_use]
    pub fn new(paragraph: usize, range: Range<usize>, text: impl Into<String>) -> Self {
        Self {
            paragraph,
            range,
            text: text.into(),
        }
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

    /// The new text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}

impl std::fmt::Debug for Replacement {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Replacement")
            .field("paragraph", &self.paragraph)
            .field("range", &self.range)
            .field("bytes", &self.text.len())
            .finish()
    }
}

/// Content-free counts of one substitution.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Outcome {
    paragraphs: usize,
    texts: usize,
    fields: usize,
}

impl Outcome {
    /// Paragraphs whose markup changed.
    #[must_use]
    pub const fn paragraphs(self) -> usize {
        self.paragraphs
    }

    /// Text ranges replaced.
    #[must_use]
    pub const fn texts(self) -> usize {
        self.texts
    }

    /// Fields replaced by their text.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Whether nothing changed.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.paragraphs == 0
    }
}

// ---------------------------------------------------------------------------
// Scanning
// ---------------------------------------------------------------------------

fn xml_error(message: impl Into<String>) -> Error {
    Error::Document(crate::Error::Xml(message.into()))
}

fn refused(paragraph: usize, reason: Refusal) -> Error {
    Error::Refused {
        position: paragraph,
        reason,
    }
}

/// Run content that reads as one fixed character.
fn fixed_character(local: &[u8]) -> Option<char> {
    Some(match local {
        b"tab" | b"ptab" => '\t',
        b"br" | b"cr" => '\n',
        b"noBreakHyphen" => '\u{2011}',
        b"softHyphen" => '\u{AD}',
        _ => return None,
    })
}

/// Run content without text that a field run may hold and lose.
fn is_silent(local: &[u8]) -> bool {
    matches!(
        local,
        b"lastRenderedPageBreak" | b"delText" | b"delInstrText" | b"annotationRef"
    )
}

struct OpenText {
    start: usize,
    start_tag_end: usize,
    name: Range<usize>,
    space: bool,
}

struct OpenRun {
    start: usize,
    depth: usize,
    /// The start of the run's parent element.
    parent: usize,
    properties: Option<Range<usize>>,
    properties_start: Option<usize>,
    name: Range<usize>,
    has_text: bool,
    /// A complex field whose end marker this run holds.
    closes: Option<OpenComplex>,
}

struct OpenComplex {
    begin_run: usize,
    parent: usize,
    run_depth: usize,
    range_start: usize,
    objects_before: usize,
    instruction: String,
    separated: bool,
    replaceable: bool,
    begin_properties: Option<Range<usize>>,
    result_properties: Option<Option<Range<usize>>>,
    run_name: Range<usize>,
}

struct OpenSimple {
    start: usize,
    depth: usize,
    range_start: usize,
    objects_before: usize,
    instruction: String,
    replaceable: bool,
    properties: Option<Option<Range<usize>>>,
    run_name: Option<Range<usize>>,
    prefix: Vec<u8>,
}

struct OpenParagraph {
    depth: usize,
    block: Option<Block>,
    text: String,
    pieces: Vec<Piece>,
    fields: Vec<Field>,
    objects: usize,
    complex: Vec<OpenComplex>,
    simple: Vec<OpenSimple>,
}

impl OpenParagraph {
    fn push_fixed(&mut self, character: char) {
        let start = self.text.len();
        self.text.push(character);
        if character == OBJECT {
            self.objects += 1;
        }
        self.pieces.push(Piece {
            text: start..self.text.len(),
            node: None,
        });
    }

    /// Mark every open field whose markup a non-run sibling interrupts.
    fn interrupt_at(&mut self, depth: usize) {
        for field in &mut self.complex {
            if field.run_depth == depth {
                field.replaceable = false;
            }
        }
        for field in &mut self.simple {
            if field.depth + 1 == depth {
                field.replaceable = false;
            }
        }
    }
}

/// Where the scan is in the document.
struct Scanner<'x> {
    xml: &'x [u8],
    main: bool,
    depth: usize,
    body_depth: Option<usize>,
    /// The current body child and the counters of addressable children.
    block: Option<(Option<ElementStep>, usize)>,
    counts: [usize; 3],
    paragraphs: Vec<OpenParagraph>,
    runs: Vec<OpenRun>,
    text: Option<OpenText>,
    instruction: Option<usize>,
    /// The start offsets of the open elements.
    stack: Vec<usize>,
    done: Vec<Paragraph>,
}

/// Scan a story part. A main-document scan lists only paragraphs inside
/// addressable body blocks.
pub(crate) fn scan(xml: &[u8], main: bool) -> Result<Vec<Paragraph>> {
    std::str::from_utf8(xml).map_err(|_| xml_error("story XML is not UTF-8"))?;
    let mut scanner = Scanner {
        xml,
        main,
        depth: 0,
        body_depth: None,
        block: None,
        counts: [0; 3],
        paragraphs: Vec::new(),
        runs: Vec::new(),
        text: None,
        instruction: None,
        stack: Vec::new(),
        done: Vec::new(),
    };
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| xml_error("story offset overflow"))?;
        let event = reader
            .read_event()
            .map_err(|error| xml_error(format!("invalid story XML: {error}")))?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| xml_error("story offset overflow"))?;
        let (namespace, event) = reader.resolver().resolve_event(event);
        let word = is_wordprocessing_namespace(&namespace);
        match event {
            Event::Start(element) => {
                scanner.depth += 1;
                if scanner.depth > MAX_DOCUMENT_DEPTH {
                    return Err(xml_error("story XML is too deep"));
                }
                scanner.stack.push(start);
                scanner.open(&reader, &element, word, start, end, false)?;
            },
            Event::Empty(element) => {
                scanner.depth += 1;
                scanner.stack.push(start);
                scanner.open(&reader, &element, word, start, end, true)?;
                scanner.close(element.local_name().as_ref(), word, end)?;
                scanner.stack.pop();
                scanner.depth -= 1;
            },
            Event::End(element) => {
                scanner.close(element.local_name().as_ref(), word, end)?;
                scanner.stack.pop();
                scanner.depth = scanner
                    .depth
                    .checked_sub(1)
                    .ok_or_else(|| xml_error("unbalanced story XML"))?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if scanner.depth != 0 || !scanner.paragraphs.is_empty() {
        return Err(xml_error("unbalanced story XML"));
    }
    Ok(scanner.done)
}

impl Scanner<'_> {
    fn open(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        word: bool,
        start: usize,
        end: usize,
        empty: bool,
    ) -> Result<()> {
        let local = element.local_name();
        let local = local.as_ref();
        let depth = self.depth;
        if word && self.main && self.body_depth.is_none() && local == b"body" && depth == 2 {
            self.body_depth = Some(depth);
            return Ok(());
        }
        if self.main && self.body_depth.is_some_and(|body| depth == body + 1) {
            let step = match (word, local) {
                (true, b"p") => Some(ElementStep::Paragraph(next(&mut self.counts[0]))),
                (true, b"tbl") => Some(ElementStep::Table(next(&mut self.counts[1]))),
                (true, b"sdt") => Some(ElementStep::ContentControl(next(&mut self.counts[2]))),
                _ => None,
            };
            self.block = Some((step, start));
        }
        // A non-run sibling inside a field's markup.
        if let Some(paragraph) = self.paragraphs.last_mut()
            && !(word && matches!(local, b"r" | b"proofErr"))
        {
            paragraph.interrupt_at(depth);
        }
        if !word {
            self.inside_run_child(None, depth);
            return Ok(());
        }
        match local {
            b"p" => {
                let block = if self.main {
                    match &self.block {
                        Some((Some(step), block_start)) if self.body_depth.is_some() => {
                            Some(Block {
                                step: *step,
                                range: *block_start..*block_start,
                            })
                        },
                        _ => None,
                    }
                } else {
                    None
                };
                // A nested paragraph (a text box) interrupts the fields of
                // the paragraph that holds it.
                if let Some(outer) = self.paragraphs.last_mut() {
                    for field in &mut outer.complex {
                        field.replaceable = false;
                    }
                    for field in &mut outer.simple {
                        field.replaceable = false;
                    }
                }
                self.paragraphs.push(OpenParagraph {
                    depth,
                    block,
                    text: String::new(),
                    pieces: Vec::new(),
                    fields: Vec::new(),
                    objects: 0,
                    complex: Vec::new(),
                    simple: Vec::new(),
                });
            },
            b"r" if !self.paragraphs.is_empty() && self.text.is_none() => {
                let parent = self.parent_start();
                self.runs.push(OpenRun {
                    start,
                    depth,
                    parent,
                    properties: None,
                    properties_start: None,
                    name: name_range(start, element),
                    has_text: false,
                    closes: None,
                });
                if let Some(paragraph) = self.paragraphs.last_mut() {
                    for field in &mut paragraph.simple {
                        if field.run_name.is_none() {
                            field.run_name = Some(name_range(start, element));
                        }
                    }
                }
            },
            b"fldSimple" if !self.paragraphs.is_empty() && !self.in_run() => {
                let instruction = word_attribute(reader, element, b"instr")?.unwrap_or_default();
                let prefix = element
                    .name()
                    .prefix()
                    .map(|prefix| prefix.as_ref().to_vec())
                    .unwrap_or_default();
                let Some(paragraph) = self.paragraphs.last_mut() else {
                    return Ok(());
                };
                for outer in &mut paragraph.simple {
                    outer.replaceable = false;
                }
                let objects_before = paragraph.objects;
                let range_start = paragraph.text.len();
                paragraph.push_fixed(OBJECT);
                paragraph.simple.push(OpenSimple {
                    start,
                    depth,
                    range_start,
                    objects_before,
                    instruction,
                    replaceable: true,
                    properties: None,
                    run_name: None,
                    prefix,
                });
            },
            _ => {
                let run_child = self.runs.last().is_some_and(|run| run.depth + 1 == depth);
                if !run_child {
                    self.inside_run_child(Some(local), depth);
                    return Ok(());
                }
                match local {
                    b"rPr" => {
                        if let Some(run) = self.runs.last_mut() {
                            run.properties_start = Some(start);
                        }
                    },
                    b"t" => {
                        if !empty {
                            self.text = Some(OpenText {
                                start,
                                start_tag_end: end,
                                name: name_range(start, element),
                                space: has_space(element),
                            });
                        }
                    },
                    b"fldChar" => {
                        let kind = word_attribute(reader, element, b"fldCharType")?;
                        self.field_char(kind.as_deref());
                    },
                    b"instrText" => {
                        if !empty {
                            self.instruction = Some(start);
                        }
                    },
                    _ if is_silent(local) => {},
                    _ => {
                        let character = fixed_character(local).unwrap_or(OBJECT);
                        if let Some(paragraph) = self.paragraphs.last_mut() {
                            paragraph.push_fixed(character);
                        }
                    },
                }
            },
        }
        Ok(())
    }

    /// Content below a run child (inside a drawing, for instance) adds
    /// nothing to the paragraph text; a non-Word run child is an object.
    fn inside_run_child(&mut self, local: Option<&[u8]>, depth: usize) {
        let _ = local;
        if self.runs.last().is_some_and(|run| run.depth + 1 == depth)
            && let Some(paragraph) = self.paragraphs.last_mut()
        {
            paragraph.push_fixed(OBJECT);
        }
    }

    fn in_run(&self) -> bool {
        match (self.runs.last(), self.paragraphs.last()) {
            (Some(run), Some(paragraph)) => run.depth > paragraph.depth,
            _ => false,
        }
    }

    /// The start offset of the current element's parent.
    fn parent_start(&self) -> usize {
        self.stack
            .len()
            .checked_sub(2)
            .and_then(|index| self.stack.get(index))
            .copied()
            .unwrap_or_default()
    }

    fn field_char(&mut self, kind: Option<&str>) {
        let Some(run) = self.runs.last() else {
            return;
        };
        let (run_start, run_depth, run_parent, run_properties, run_name) = (
            run.start,
            run.depth,
            run.parent,
            run.properties.clone(),
            run.name.clone(),
        );
        let Some(paragraph) = self.paragraphs.last_mut() else {
            return;
        };
        match kind {
            Some("begin") => {
                for outer in &mut paragraph.complex {
                    outer.replaceable = false;
                }
                let objects_before = paragraph.objects;
                let range_start = paragraph.text.len();
                paragraph.push_fixed(OBJECT);
                paragraph.complex.push(OpenComplex {
                    begin_run: run_start,
                    parent: run_parent,
                    run_depth,
                    range_start,
                    objects_before,
                    instruction: String::new(),
                    separated: false,
                    replaceable: true,
                    begin_properties: run_properties,
                    result_properties: None,
                    run_name,
                });
            },
            Some("separate") => {
                paragraph.push_fixed(OBJECT);
                if let Some(field) = paragraph.complex.last_mut() {
                    field.separated = true;
                }
            },
            Some("end") => {
                paragraph.push_fixed(OBJECT);
                if let Some(mut field) = paragraph.complex.pop() {
                    if field.run_depth != run_depth || field.parent != run_parent {
                        field.replaceable = false;
                    }
                    if let Some(run) = self.runs.last_mut() {
                        run.closes = Some(field);
                    }
                }
            },
            _ => paragraph.push_fixed(OBJECT),
        }
    }

    fn close(&mut self, local: &[u8], word: bool, end: usize) -> Result<()> {
        if word {
            self.close_word(local, end)?;
        }
        if self.main && self.body_depth.is_some_and(|body| self.depth == body + 1) {
            self.close_block(end);
        }
        Ok(())
    }

    fn close_word(&mut self, local: &[u8], end: usize) -> Result<()> {
        let depth = self.depth;
        match local {
            b"t" if self.text.is_some()
                && self.runs.last().is_some_and(|run| run.depth + 1 == depth) =>
            {
                let Some(open) = self.text.take() else {
                    return Ok(());
                };
                let value = decode_text_fragment(
                    self.xml
                        .get(open.start..end)
                        .ok_or_else(|| xml_error("story text offset is out of range"))?,
                )
                .map_err(|_| xml_error("invalid story text"))?;
                if let Some(run) = self.runs.last_mut() {
                    run.has_text |= !value.is_empty();
                }
                if let Some(paragraph) = self.paragraphs.last_mut()
                    && !value.is_empty()
                {
                    let text_start = paragraph.text.len();
                    paragraph.text.push_str(&value);
                    paragraph.pieces.push(Piece {
                        text: text_start..paragraph.text.len(),
                        node: Some(TextNode {
                            element: open.start..end,
                            start_tag_end: open.start_tag_end,
                            name: open.name,
                            space: open.space,
                        }),
                    });
                }
            },
            b"instrText" if self.instruction.is_some() => {
                let Some(open) = self.instruction.take() else {
                    return Ok(());
                };
                let value = decode_text_fragment(
                    self.xml
                        .get(open..end)
                        .ok_or_else(|| xml_error("story instruction offset is out of range"))?,
                )
                .map_err(|_| xml_error("invalid field instruction"))?;
                if let Some(field) = self
                    .paragraphs
                    .last_mut()
                    .and_then(|paragraph| paragraph.complex.last_mut())
                    && !field.separated
                {
                    field.instruction.push_str(&value);
                }
            },
            b"rPr" if self.runs.last().is_some_and(|run| run.depth + 1 == depth) => {
                if let Some(run) = self.runs.last_mut()
                    && let Some(properties_start) = run.properties_start.take()
                {
                    run.properties = Some(properties_start..end);
                }
            },
            b"r" if self.runs.last().is_some_and(|run| run.depth == depth) => {
                let Some(run) = self.runs.pop() else {
                    return Ok(());
                };
                self.close_run(run, end);
            },
            b"fldSimple" => {
                if let Some(paragraph) = self.paragraphs.last_mut()
                    && paragraph
                        .simple
                        .last()
                        .is_some_and(|field| field.depth == depth)
                    && let Some(field) = paragraph.simple.pop()
                {
                    paragraph.push_fixed(OBJECT);
                    let inside = paragraph.objects - field.objects_before;
                    let range = field.range_start..paragraph.text.len();
                    let run_name = field
                        .run_name
                        .map(|name| self.xml[name].to_vec())
                        .unwrap_or_else(|| qualified(&field.prefix, b"r"));
                    paragraph.fields.push(Field {
                        range,
                        instruction: field.instruction,
                        kind: FieldKind::Simple,
                        replaceable: field.replaceable && inside == 2,
                        markup: field.start..end,
                        properties: field.properties.flatten(),
                        run_name,
                    });
                }
            },
            b"p" if self
                .paragraphs
                .last()
                .is_some_and(|paragraph| paragraph.depth == depth) =>
            {
                let Some(paragraph) = self.paragraphs.pop() else {
                    return Ok(());
                };
                let listed = !self.main || paragraph.block.is_some();
                if listed {
                    let mut fields = paragraph.fields;
                    fields.sort_by_key(|field| field.range.start);
                    self.done.push(Paragraph {
                        text: paragraph.text,
                        fields,
                        pieces: paragraph.pieces,
                        block: paragraph.block,
                    });
                }
            },
            _ => {},
        }
        Ok(())
    }

    fn close_run(&mut self, run: OpenRun, end: usize) {
        let Some(paragraph) = self.paragraphs.last_mut() else {
            return;
        };
        // The first result run with text supplies the replacement's
        // formatting; a field without one uses its begin run's.
        for field in &mut paragraph.complex {
            if field.separated && field.result_properties.is_none() && run.has_text {
                field.result_properties = Some(run.properties.clone());
            }
        }
        for field in &mut paragraph.simple {
            if field.properties.is_none() && run.has_text {
                field.properties = Some(run.properties.clone());
            }
        }
        if let Some(field) = run.closes {
            let inside = paragraph.objects - field.objects_before;
            let expected = if field.separated { 3 } else { 2 };
            let properties = match field.result_properties {
                Some(properties) => properties,
                None => field.begin_properties,
            };
            paragraph.fields.push(Field {
                range: field.range_start..paragraph.text.len(),
                instruction: field.instruction,
                kind: FieldKind::Complex,
                replaceable: field.replaceable && inside == expected,
                markup: field.begin_run..end,
                properties,
                run_name: self.xml[field.run_name].to_vec(),
            });
        }
    }

    fn close_block(&mut self, end: usize) {
        let Some((step, block_start)) = self.block.take() else {
            return;
        };
        if step.is_none() {
            return;
        }
        let range = block_start..end;
        for paragraph in self.done.iter_mut().rev() {
            match &mut paragraph.block {
                Some(block) if block.range.start == block_start => block.range = range.clone(),
                _ => break,
            }
        }
    }
}

fn next(counter: &mut usize) -> litchi_core::Position {
    let position = litchi_core::Position::new(*counter);
    *counter += 1;
    position
}

/// The qualified name of a start tag, as absolute offsets.
fn name_range(start: usize, element: &BytesStart<'_>) -> Range<usize> {
    let length = element.name().as_ref().len();
    start + 1..start + 1 + length
}

fn qualified(prefix: &[u8], local: &[u8]) -> Vec<u8> {
    let mut name = Vec::with_capacity(prefix.len() + local.len() + 1);
    if !prefix.is_empty() {
        name.extend_from_slice(prefix);
        name.push(b':');
    }
    name.extend_from_slice(local);
    name
}

fn has_space(element: &BytesStart<'_>) -> bool {
    element
        .checked_attributes()
        .flatten()
        .any(|attribute| attribute.key.as_ref() == b"xml:space")
}

fn word_attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<Option<String>> {
    for attribute in element.checked_attributes() {
        let attribute =
            attribute.map_err(|error| xml_error(format!("invalid story attribute: {error}")))?;
        if attribute.key.local_name().as_ref() != name {
            continue;
        }
        let (namespace, _) = reader.resolver().resolve_attribute(attribute.key);
        if !is_wordprocessing_namespace(&namespace) && !matches!(namespace, ResolveResult::Unbound)
        {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| xml_error(format!("invalid story attribute: {error}")))?;
        return Ok(Some(value.into_owned()));
    }
    Ok(None)
}

// ---------------------------------------------------------------------------
// Rewriting
// ---------------------------------------------------------------------------

/// The splices one plan makes, the texts the rewritten paragraphs must read
/// back as, and its counts.
struct Plan {
    splices: Vec<(usize, usize, Vec<u8>)>,
    /// `(paragraph, text)` for every changed paragraph.
    expected: Vec<(usize, String)>,
    outcome: Outcome,
}

fn validate_text(paragraph: usize, text: &str) -> Result<()> {
    if text.chars().any(|character| {
        !matches!(character, '\u{9}' | '\u{A}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')
            || character == OBJECT
    }) {
        return Err(refused(paragraph, Refusal::StructuralText));
    }
    Ok(())
}

fn plan(story: &Story, replacements: &[Replacement]) -> Result<Plan> {
    let mut splices = Vec::new();
    let mut expected = Vec::new();
    let mut outcome = Outcome::default();
    let mut total = 0usize;
    let mut index = 0usize;
    let mut previous: Option<(usize, usize)> = None;
    while index < replacements.len() {
        let paragraph_index = replacements[index].paragraph;
        let mut end = index;
        while end < replacements.len() && replacements[end].paragraph == paragraph_index {
            end += 1;
        }
        let group = &replacements[index..end];
        index = end;
        let paragraph = story
            .paragraphs
            .get(paragraph_index)
            .ok_or_else(|| refused(paragraph_index, Refusal::BlockNotFound))?;
        if previous.is_some_and(|(last, _)| last >= paragraph_index) {
            return Err(refused(
                paragraph_index,
                Refusal::AmbiguousCompositeSelector,
            ));
        }
        previous = Some((paragraph_index, 0));
        let mut cursor = 0usize;
        let mut text_ranges: Vec<&Replacement> = Vec::new();
        let mut new_text = String::with_capacity(paragraph.text.len());
        for replacement in group {
            let range = replacement.range.clone();
            if range.is_empty()
                || range.start < cursor
                || range.end > paragraph.text.len()
                || !paragraph.text.is_char_boundary(range.start)
                || !paragraph.text.is_char_boundary(range.end)
            {
                return Err(refused(
                    paragraph_index,
                    Refusal::AmbiguousCompositeSelector,
                ));
            }
            validate_text(paragraph_index, &replacement.text)?;
            total = total
                .checked_add(replacement.text.len())
                .filter(|total| *total <= MAX_REPLACEMENT_TEXT_BYTES)
                .ok_or(Error::Limit {
                    kind: crate::document::LimitKind::ReplacementTextBytes,
                    maximum: MAX_REPLACEMENT_TEXT_BYTES,
                    observed: total.saturating_add(replacement.text.len()),
                })?;
            new_text.push_str(&paragraph.text[cursor..range.start]);
            new_text.push_str(&replacement.text);
            cursor = range.end;
            if let Some(field) = paragraph.fields.iter().find(|field| field.range == range) {
                if !field.replaceable {
                    return Err(refused(paragraph_index, Refusal::FieldInterior));
                }
                splices.push((
                    field.markup.start,
                    field.markup.end,
                    field_run(&story.xml, field, &replacement.text),
                ));
                outcome.fields += 1;
            } else {
                if !paragraph.is_text(range) {
                    return Err(refused(paragraph_index, Refusal::FieldInterior));
                }
                text_ranges.push(replacement);
                outcome.texts += 1;
            }
        }
        new_text.push_str(&paragraph.text[cursor..]);
        rewrite_nodes(&story.xml, paragraph, &text_ranges, &mut splices);
        outcome.paragraphs += 1;
        expected.push((paragraph_index, new_text));
    }
    splices.sort_by_key(|splice| splice.0);
    if splices.windows(2).any(|pair| pair[0].1 > pair[1].0) {
        return Err(refused(0, Refusal::AmbiguousCompositeSelector));
    }
    Ok(Plan {
        splices,
        expected,
        outcome,
    })
}

/// Rewrite every `w:t` element whose text the replacements change.
fn rewrite_nodes(
    xml: &[u8],
    paragraph: &Paragraph,
    replacements: &[&Replacement],
    splices: &mut Vec<(usize, usize, Vec<u8>)>,
) {
    if replacements.is_empty() {
        return;
    }
    for piece in &paragraph.pieces {
        let Some(node) = &piece.node else {
            continue;
        };
        let (start, end) = (piece.text.start, piece.text.end);
        let mut overlapping = replacements
            .iter()
            .filter(|replacement| replacement.range.start < end && replacement.range.end > start)
            .peekable();
        if overlapping.peek().is_none() {
            continue;
        }
        let mut content = String::new();
        let mut cursor = start;
        for replacement in overlapping {
            if replacement.range.start >= cursor {
                content.push_str(&paragraph.text[cursor..replacement.range.start]);
            }
            if replacement.range.start >= start {
                content.push_str(&replacement.text);
            }
            cursor = replacement.range.end.min(end);
        }
        content.push_str(&paragraph.text[cursor..end]);
        splices.push((
            node.element.start,
            node.element.end,
            text_elements(xml, node, &content),
        ));
    }
}

/// `content` as `w:t` elements with the original start tag, separated by
/// `w:tab` and `w:br` for tabs and line feeds.
fn text_elements(xml: &[u8], node: &TextNode, content: &str) -> Vec<u8> {
    let name = &xml[node.name.clone()];
    let prefix = name
        .iter()
        .position(|byte| *byte == b':')
        .map_or(&b""[..], |colon| &name[..colon]);
    let start_tag = &xml[node.element.start..node.start_tag_end];
    let mut output = Vec::with_capacity(content.len() + 2 * start_tag.len() + 8);
    let mut segments = content.split(['\t', '\n']).peekable();
    let mut separators = content
        .chars()
        .filter(|character| matches!(character, '\t' | '\n'));
    let mut wrote = false;
    while let Some(segment) = segments.next() {
        if !segment.is_empty() || (!wrote && segments.peek().is_none()) {
            write_start_tag(&mut output, start_tag, node.space, segment);
            escape(&mut output, segment);
            output.extend_from_slice(b"</");
            output.extend_from_slice(name);
            output.push(b'>');
            wrote = true;
        }
        if let Some(separator) = separators.next() {
            output.push(b'<');
            output.extend_from_slice(&qualified(
                prefix,
                if separator == '\t' { b"tab" } else { b"br" },
            ));
            output.extend_from_slice(b"/>");
        }
    }
    output
}

fn write_start_tag(output: &mut Vec<u8>, start_tag: &[u8], space: bool, segment: &str) {
    let needs_space = !space && (segment.starts_with(' ') || segment.ends_with(' '));
    if needs_space {
        output.extend_from_slice(&start_tag[..start_tag.len() - 1]);
        output.extend_from_slice(b" xml:space=\"preserve\">");
    } else {
        output.extend_from_slice(start_tag);
    }
}

fn escape(output: &mut Vec<u8>, text: &str) {
    for byte in text.bytes() {
        match byte {
            b'&' => output.extend_from_slice(b"&amp;"),
            b'<' => output.extend_from_slice(b"&lt;"),
            b'>' => output.extend_from_slice(b"&gt;"),
            _ => output.push(byte),
        }
    }
}

/// One run holding `text`, formatted as the field's result.
fn field_run(xml: &[u8], field: &Field, text: &str) -> Vec<u8> {
    if text.is_empty() {
        return Vec::new();
    }
    let run_name = field.run_name.as_slice();
    let prefix = run_name
        .iter()
        .position(|byte| *byte == b':')
        .map_or(&b""[..], |colon| &run_name[..colon]);
    let text_name = qualified(prefix, b"t");
    let mut start_tag = Vec::with_capacity(text_name.len() + 2);
    start_tag.push(b'<');
    start_tag.extend_from_slice(&text_name);
    start_tag.push(b'>');
    let mut output = Vec::new();
    output.push(b'<');
    output.extend_from_slice(run_name);
    output.push(b'>');
    if let Some(properties) = &field.properties {
        output.extend_from_slice(&xml[properties.clone()]);
    }
    // The synthetic `w:t` start tag spans `start_tag`; reuse the element
    // writer with a node over a scratch buffer.
    let mut scratch = start_tag.clone();
    scratch.extend_from_slice(b"</");
    scratch.extend_from_slice(&text_name);
    scratch.push(b'>');
    let node = TextNode {
        element: 0..scratch.len(),
        start_tag_end: start_tag.len(),
        name: 1..1 + text_name.len(),
        space: false,
    };
    output.extend_from_slice(&text_elements(&scratch, &node, text));
    output.extend_from_slice(b"</");
    output.extend_from_slice(run_name);
    output.push(b'>');
    output
}

/// Read back the rewritten story: every changed paragraph must read as
/// planned, and every other one as before.
fn read_back(
    before: &[Paragraph],
    after: &[u8],
    main: bool,
    expected: &[(usize, String)],
) -> Result<()> {
    let rescanned = scan(after, main)?;
    if rescanned.len() != before.len() {
        return Err(xml_error("substituted story lost or gained paragraphs"));
    }
    let mut expected = expected.iter().peekable();
    for (index, (old, new)) in before.iter().zip(&rescanned).enumerate() {
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

/// Scan a story part's payload.
pub(crate) fn part_story(part: &str, xml: Arc<Vec<u8>>) -> Result<Story> {
    if xml.len() > MAX_DOCUMENT_XML_BYTES {
        return Err(Error::Limit {
            kind: crate::document::LimitKind::XmlBytes,
            maximum: MAX_DOCUMENT_XML_BYTES,
            observed: xml.len(),
        });
    }
    let paragraphs = scan(&xml, false)?;
    Ok(Story {
        part: Some(part.to_owned()),
        xml,
        paragraphs,
    })
}

/// Rewrite a story part's payload, or `None` when nothing changes.
pub(crate) fn substitute_part(
    story: &Story,
    current: &[u8],
    replacements: &[Replacement],
) -> Result<(Option<Vec<u8>>, Outcome)> {
    if story.part.is_none() {
        return Err(refused(0, Refusal::StoryReference));
    }
    if current != story.xml.as_slice() {
        return Err(Error::StaleSource);
    }
    let plan = plan(story, replacements)?;
    if plan.splices.is_empty() {
        return Ok((None, plan.outcome));
    }
    let after = replace_ranges(&story.xml, &plan.splices)?;
    read_back(&story.paragraphs, &after, false, &plan.expected)?;
    Ok((Some(after), plan.outcome))
}

impl Edit {
    /// The main document's story, over the staged document.
    ///
    /// # Errors
    ///
    /// Returns an XML error for malformed markup.
    pub fn story(&self) -> Result<Story> {
        let xml = Arc::new(self.projected.xml_bytes().to_vec());
        let paragraphs = scan(&xml, true)?;
        Ok(Story {
            part: None,
            xml,
            paragraphs,
        })
    }

    /// Apply `replacements` to the main document's story, which must have
    /// been read from this edit's staged document.
    ///
    /// Each changed body block (paragraph, table or block content control)
    /// is recorded as one exact element replacement, all spliced and
    /// rescanned once.
    ///
    /// # Errors
    ///
    /// Returns [`Error::StaleSource`] for a story of other bytes, a refusal
    /// for an invalid, overlapping or non-text range, a field that is not
    /// replaceable or text a `w:t` cannot hold, and a limit error for too
    /// many changed blocks or too much replacement text. The edit is
    /// unchanged on failure.
    pub fn substitute(&mut self, story: &Story, replacements: &[Replacement]) -> Result<Outcome> {
        if story.part.is_some() {
            return Err(refused(0, Refusal::StoryReference));
        }
        self.ensure_unmanaged("substitute")?;
        if self.projected.xml_bytes() != story.xml.as_slice() {
            return Err(Error::StaleSource);
        }
        let plan = plan(story, replacements)?;
        if plan.splices.is_empty() {
            return Ok(plan.outcome);
        }
        // Group the splices by body block.
        let mut blocks: Vec<(Block, Vec<(usize, usize, Vec<u8>)>)> = Vec::new();
        let mut splices = plan.splices.into_iter().peekable();
        let mut by_start: Vec<&Block> = story
            .paragraphs
            .iter()
            .filter_map(|paragraph| paragraph.block.as_ref())
            .collect();
        by_start.dedup_by(|left, right| left.range == right.range);
        for block in by_start {
            let mut local = Vec::new();
            while let Some(splice) = splices.peek() {
                if splice.0 >= block.range.end {
                    break;
                }
                if splice.0 < block.range.start || splice.1 > block.range.end {
                    return Err(xml_error("story splice leaves its body block"));
                }
                if let Some((start, end, bytes)) = splices.next() {
                    local.push((start - block.range.start, end - block.range.start, bytes));
                }
            }
            if !local.is_empty() {
                blocks.push((block.clone(), local));
            }
        }
        if splices.next().is_some() {
            return Err(xml_error("story splice lies outside the body blocks"));
        }
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
        for (block, local) in blocks {
            let path = ElementPath::new([block.step]);
            let located = super::formatting::locate_element(&self.projected, &path)?;
            if located != block.range {
                return Err(Error::StaleSource);
            }
            let before = &story.xml[block.range.clone()];
            let after = replace_ranges(before, &local)?;
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
            ranges.push((block.range.start, block.range.end, after.clone()));
            replaced.push((path, Arc::new(before.to_vec()), Arc::new(after)));
        }
        let xml = replace_ranges(self.projected.xml_bytes(), &ranges)?;
        read_back(&story.paragraphs, &xml, true, &plan.expected)?;
        let candidate = self.projected.with_rewritten_xml(xml)?;
        for (path, _, after) in &replaced {
            let located = super::formatting::locate_element(&candidate, path)?;
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
        for (path, before, after) in replaced {
            self.operations.push(Operation::ReplaceElement {
                path,
                before,
                after,
            });
        }
        self.replacement_text_bytes = text_bytes;
        self.projected = candidate;
        Ok(plan.outcome)
    }
}
