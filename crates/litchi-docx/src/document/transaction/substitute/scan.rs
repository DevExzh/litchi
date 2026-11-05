//! The story scanner: one pass over a story part (or a fragment of the main
//! document) that lists its paragraphs' text, text nodes, runs and fields.
//!
//! Field state is tracked across paragraphs: a complex field whose
//! instruction or result spans paragraphs (a table of contents) keeps its
//! state from the paragraph holding its `begin` marker to the one holding its
//! `end` marker (record 0285). Text inside a field instruction (the result of
//! a field nested in the instruction) is fixed content, never text; text of a
//! field result is ordinary text.
//!
//! An `mc:AlternateContent` wrapper (a drawing or text box with a fallback,
//! record 0307) is one opaque object: a run holding it reads it as one
//! [`OBJECT`], and nothing inside it, neither branch's text nor the
//! paragraphs of its text boxes, is listed. The branches differ, and the
//! visible projection shows only one, so neither is text to an edit.

use std::ops::Range;

use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

use super::super::formatting::ElementStep;
use super::super::{Error, MAX_DOCUMENT_DEPTH, Result, decode_text_fragment};
use super::{Block, Field, FieldKind, OBJECT, Paragraph, Piece, RunSlot, TextNode};
use crate::namespace::{NamespaceBindings, NamespaceCapture, is_wordprocessing_namespace};

/// Where a complex field is: inside its instruction (before `separate`) or
/// its result.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Phase {
    Instruction,
    Result,
}

/// Which paragraphs a scan lists.
#[derive(Clone)]
pub(super) enum Scope {
    /// Every paragraph (a main-document scan still lists only paragraphs of
    /// addressable body blocks).
    All,
    /// The paragraph starting at an offset; nothing nested in it.
    Paragraph(usize),
    /// Every paragraph inside a byte range (a table, row, cell or content
    /// control).
    Within(Range<usize>),
    /// The paragraph starting at an offset, with only the text inside a
    /// byte range (a run, hyperlink or inline content control) in scope.
    Inline {
        paragraph: usize,
        range: Range<usize>,
    },
}

impl Scope {
    /// Whether the paragraph at `element` is listed, and which of its text
    /// is in scope: `Some(None)` all of it, `Some(Some(range))` a part.
    fn admits(&self, element: &Range<usize>) -> Option<Option<Range<usize>>> {
        match self {
            Self::All => Some(None),
            Self::Paragraph(start) => (*start == element.start).then_some(None),
            Self::Within(range) => {
                (range.start <= element.start && element.end <= range.end).then_some(None)
            },
            Self::Inline { paragraph, range } => {
                (*paragraph == element.start).then(|| Some(range.clone()))
            },
        }
    }
}

/// How one scan starts.
pub(super) struct Start {
    /// The fragment of the story XML to scan.
    pub(super) range: Range<usize>,
    /// Whether this is the main document: body blocks are tracked and only
    /// paragraphs of addressable blocks are listed. A part scan numbers the
    /// direct children of the part's root instead.
    pub(super) main: bool,
    /// Bindings in scope at the fragment's first element (empty for a whole
    /// part).
    pub(super) inherited: NamespaceBindings,
    /// The complex fields open at the fragment's start, outermost first.
    pub(super) fields: Vec<Phase>,
    /// The body block of a main-document fragment that is one direct-body
    /// paragraph: every paragraph of the fragment belongs to it.
    pub(super) block: Option<Block>,
    /// Which paragraphs to list.
    pub(super) scopes: Vec<Scope>,
    /// Stop once the scan passes this offset with no paragraph open.
    pub(super) stop: Option<usize>,
}

/// What a scan found.
pub(super) struct Scanned {
    pub(super) paragraphs: Vec<Paragraph>,
    /// Namespace scopes the paragraphs and runs refer to.
    pub(super) bindings: Vec<NamespaceBindings>,
}

/// What an open element means to the scanner, for its end tag.
#[derive(Clone, Copy, Default)]
struct Frame {
    /// The element declares namespaces: its end restores the enclosing
    /// binding scope.
    declares: bool,
    /// The element is non-run inline content of a paragraph: entering and
    /// leaving it separates text.
    boundary: bool,
    /// The element is a `w:fldSimple`: its runs are a field result.
    simple: bool,
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
    open_end: usize,
    properties: Option<Range<usize>>,
    properties_start: Option<usize>,
    name: Range<usize>,
    has_text: bool,
    /// The binding scope in effect at the run.
    bindings: usize,
    /// The run's index in its paragraph's runs, once it holds text.
    slot: Option<usize>,
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
    start: usize,
    block: Option<Block>,
    bindings: usize,
    entry: Vec<Phase>,
    item: Option<usize>,
    text: String,
    pieces: Vec<Piece>,
    runs: Vec<RunSlot>,
    fields: Vec<Field>,
    objects: usize,
    /// Text offsets of the objects that stand for `mc:AlternateContent`.
    alternate: Vec<usize>,
    group: usize,
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
        self.group += 1;
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

struct Scanner<'x> {
    xml: &'x [u8],
    main: bool,
    /// A part scan: the root's direct children are numbered items.
    part: bool,
    /// The body block every paragraph of a fragment scan belongs to.
    fragment_block: Option<Block>,
    depth: usize,
    body_depth: Option<usize>,
    /// The current body child and the counters of addressable children.
    block: Option<(Option<ElementStep>, usize)>,
    counts: [usize; 3],
    /// The direct child of a part's root that is open, and how many there
    /// have been.
    item: Option<usize>,
    items: usize,
    paragraphs: Vec<OpenParagraph>,
    runs: Vec<OpenRun>,
    text: Option<OpenText>,
    instruction: Option<usize>,
    /// The start offsets of the open elements and what they mean.
    stack: Vec<(usize, Frame)>,
    /// The complex fields open in the story, outermost first.
    fields: Vec<Phase>,
    /// Binding scopes: the stack of indexes into `bindings`.
    scopes: Vec<usize>,
    bindings: Vec<NamespaceBindings>,
    capture: NamespaceCapture,
    listing: Vec<Scope>,
    done: Vec<Paragraph>,
}

/// Scan `xml[start.range]`.
pub(super) fn scan(xml: &[u8], start: Start) -> Result<Scanned> {
    let fragment = xml
        .get(start.range.clone())
        .ok_or_else(|| xml_error("story fragment is out of range"))?;
    std::str::from_utf8(fragment).map_err(|_| xml_error("story XML is not UTF-8"))?;
    let base = start.range.start;
    let mut reader = NsReader::from_reader(fragment);
    for (prefix, namespace) in start.inherited.iter() {
        let prefix = prefix.as_deref().map_or(
            quick_xml::name::PrefixDeclaration::Default,
            quick_xml::name::PrefixDeclaration::Named,
        );
        reader
            .resolver_mut()
            .add(prefix, Namespace(namespace))
            .map_err(|_| xml_error("story fragment namespace is invalid"))?;
    }
    let mut scanner = Scanner {
        xml,
        main: start.main && start.block.is_none(),
        part: !start.main,
        fragment_block: start.block,
        depth: 0,
        body_depth: None,
        block: None,
        counts: [0; 3],
        item: None,
        items: 0,
        paragraphs: Vec::new(),
        runs: Vec::new(),
        text: None,
        instruction: None,
        stack: Vec::new(),
        fields: start.fields,
        scopes: vec![0],
        bindings: vec![start.inherited],
        capture: NamespaceCapture::default(),
        listing: start.scopes,
        done: Vec::new(),
    };
    let origin = ReaderOrigin::of(fragment);
    // The depth of the open `mc:AlternateContent` whose content is skipped.
    let mut opaque = None::<usize>;
    loop {
        let event_start = origin
            .offset(reader.buffer_position())
            .and_then(|offset| offset.checked_add(base))
            .ok_or_else(|| xml_error("story offset overflow"))?;
        if let Some(stop) = start.stop
            && event_start >= stop
            && scanner.paragraphs.is_empty()
        {
            break;
        }
        let event = reader
            .read_event()
            .map_err(|error| xml_error(format!("invalid story XML: {error}")))?;
        let event_end = origin
            .offset(reader.buffer_position())
            .and_then(|offset| offset.checked_add(base))
            .ok_or_else(|| xml_error("story offset overflow"))?;
        let (namespace, event) = reader.resolver().resolve_event(event);
        if let Some(wrapper) = opaque {
            match event {
                Event::Start(_) => {
                    scanner.depth += 1;
                    if scanner.depth > MAX_DOCUMENT_DEPTH {
                        return Err(xml_error("story XML is too deep"));
                    }
                    continue;
                },
                Event::End(_) if scanner.depth > wrapper => {
                    scanner.depth -= 1;
                    continue;
                },
                Event::End(_) => opaque = None,
                Event::Eof => {},
                _ => continue,
            }
        }
        let word = is_wordprocessing_namespace(&namespace);
        match event {
            Event::Start(element) => {
                scanner.depth += 1;
                if scanner.depth > MAX_DOCUMENT_DEPTH {
                    return Err(xml_error("story XML is too deep"));
                }
                let alternate =
                    !word && is_alternate_content(&namespace, element.local_name().as_ref());
                if alternate {
                    opaque = Some(scanner.depth);
                }
                let declares = declares_namespaces(&xml[event_start..event_end]);
                let frame = scanner.open(&reader, &element, word, event_start, event_end, false)?;
                if alternate {
                    scanner.mark_alternate();
                }
                let frame = Frame { declares, ..frame };
                if declares {
                    let captured = scanner.capture.capture(reader.resolver())?;
                    scanner.bindings.push(captured);
                    scanner.scopes.push(scanner.bindings.len() - 1);
                }
                scanner.stack.push((event_start, frame));
            },
            Event::Empty(element) => {
                scanner.depth += 1;
                let frame = scanner.open(&reader, &element, word, event_start, event_end, true)?;
                if !word && is_alternate_content(&namespace, element.local_name().as_ref()) {
                    scanner.mark_alternate();
                }
                scanner.stack.push((event_start, frame));
                scanner.close(element.local_name().as_ref(), word, event_start, event_end)?;
            },
            Event::End(element) => {
                scanner.close(element.local_name().as_ref(), word, event_start, event_end)?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if start.stop.is_none() && (scanner.depth != 0 || !scanner.paragraphs.is_empty()) {
        return Err(xml_error("unbalanced story XML"));
    }
    Ok(Scanned {
        paragraphs: scanner.done,
        bindings: scanner.bindings,
    })
}

/// Whether an element is an `mc:AlternateContent` wrapper.
fn is_alternate_content(namespace: &ResolveResult<'_>, local: &[u8]) -> bool {
    local == b"AlternateContent"
        && matches!(
            namespace,
            ResolveResult::Bound(Namespace(uri))
                if *uri == litchi_ooxml_common::mce::NAMESPACE.as_bytes()
        )
}

/// Whether a start tag declares a namespace.
fn declares_namespaces(tag: &[u8]) -> bool {
    tag.windows(5).any(|window| window == b"xmlns")
}

pub(super) fn xml_error(message: impl Into<String>) -> Error {
    Error::Document(crate::Error::Xml(message.into()))
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

impl Scanner<'_> {
    #[allow(clippy::too_many_arguments, reason = "one element's event data")]
    fn open(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        word: bool,
        start: usize,
        end: usize,
        empty: bool,
    ) -> Result<Frame> {
        let local = element.local_name();
        let local = local.as_ref();
        let depth = self.depth;
        let mut frame = Frame::default();
        if word && self.main && self.body_depth.is_none() && local == b"body" && depth == 2 {
            self.body_depth = Some(depth);
            return Ok(frame);
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
        if self.part && depth == 2 {
            self.item = Some(self.items);
            self.items += 1;
        }
        let run_child = self.runs.last().is_some_and(|run| run.depth + 1 == depth);
        let inline = !run_child && self.inline_level();
        // A non-run sibling inside a field's markup.
        if let Some(paragraph) = self.paragraphs.last_mut()
            && !(word && matches!(local, b"r" | b"proofErr"))
        {
            paragraph.interrupt_at(depth);
        }
        // Non-run inline content separates text; markers do not.
        if inline && !empty && !(word && matches!(local, b"r" | b"pPr" | b"p")) {
            frame.boundary = true;
            if let Some(paragraph) = self.paragraphs.last_mut() {
                paragraph.group += 1;
            }
        }
        if !word {
            if run_child && let Some(paragraph) = self.paragraphs.last_mut() {
                paragraph.push_fixed(OBJECT);
            }
            return Ok(frame);
        }
        match local {
            b"p" => {
                let block = if let Some(block) = &self.fragment_block {
                    Some(block.clone())
                } else if self.main {
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
                let bindings = self.scope();
                self.paragraphs.push(OpenParagraph {
                    depth,
                    start,
                    block,
                    bindings,
                    entry: self.fields.clone(),
                    item: self.item,
                    text: String::new(),
                    pieces: Vec::new(),
                    runs: Vec::new(),
                    fields: Vec::new(),
                    objects: 0,
                    alternate: Vec::new(),
                    group: 0,
                    complex: Vec::new(),
                    simple: Vec::new(),
                });
            },
            b"r" if !self.paragraphs.is_empty() && self.text.is_none() => {
                let parent = self.parent_start();
                let bindings = self.scope();
                self.runs.push(OpenRun {
                    start,
                    depth,
                    parent,
                    open_end: end,
                    properties: None,
                    properties_start: None,
                    name: name_range(start, element),
                    has_text: false,
                    bindings,
                    slot: None,
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
                if !empty {
                    // Its runs are the field's result.
                    frame.simple = true;
                    self.fields.push(Phase::Result);
                }
                let Some(paragraph) = self.paragraphs.last_mut() else {
                    return Ok(frame);
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
                if !run_child {
                    return Ok(frame);
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
                        if let Some(paragraph) = self.paragraphs.last_mut() {
                            paragraph.group += 1;
                        }
                    },
                    b"lastRenderedPageBreak" => {},
                    _ if is_silent(local) => {
                        if let Some(paragraph) = self.paragraphs.last_mut() {
                            paragraph.group += 1;
                        }
                    },
                    _ => {
                        let character = fixed_character(local).unwrap_or(OBJECT);
                        if let Some(paragraph) = self.paragraphs.last_mut() {
                            paragraph.push_fixed(character);
                        }
                    },
                }
            },
        }
        Ok(frame)
    }

    /// Record the object an `mc:AlternateContent` just opened as a direct
    /// run child pushed: the paragraph's last character.
    fn mark_alternate(&mut self) {
        let depth = self.depth;
        if self.runs.last().is_none_or(|run| run.depth + 1 != depth) {
            return;
        }
        if let Some(paragraph) = self.paragraphs.last_mut()
            && paragraph.text.ends_with(OBJECT)
        {
            let offset = paragraph.text.len() - OBJECT.len_utf8();
            paragraph.alternate.push(offset);
        }
    }

    /// Whether the next element is paragraph content outside any run.
    fn inline_level(&self) -> bool {
        match (self.paragraphs.last(), self.runs.last()) {
            (Some(_), None) => true,
            (Some(paragraph), Some(run)) => run.depth < paragraph.depth,
            (None, _) => false,
        }
    }

    fn in_run(&self) -> bool {
        match (self.runs.last(), self.paragraphs.last()) {
            (Some(run), Some(paragraph)) => run.depth > paragraph.depth,
            _ => false,
        }
    }

    /// The binding scope in effect.
    fn scope(&self) -> usize {
        self.scopes.last().copied().unwrap_or_default()
    }

    /// The start offset of the current element's parent.
    fn parent_start(&self) -> usize {
        self.stack
            .last()
            .map(|(start, _)| *start)
            .unwrap_or_default()
    }

    fn field_char(&mut self, kind: Option<&str>) {
        match kind {
            Some("begin") => self.fields.push(Phase::Instruction),
            Some("separate") => {
                if let Some(phase) = self.fields.last_mut() {
                    *phase = Phase::Result;
                }
            },
            Some("end") => {
                self.fields.pop();
            },
            _ => {},
        }
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

    fn close(&mut self, local: &[u8], word: bool, start: usize, end: usize) -> Result<()> {
        let depth = self.depth;
        let (_, frame) = self
            .stack
            .pop()
            .ok_or_else(|| xml_error("unbalanced story XML"))?;
        if word {
            self.close_word(local, start, end)?;
        }
        if frame.boundary
            && let Some(paragraph) = self.paragraphs.last_mut()
        {
            paragraph.group += 1;
        }
        if frame.simple {
            self.fields.pop();
        }
        if frame.declares {
            self.scopes.pop();
        }
        if self.main && self.body_depth.is_some_and(|body| depth == body + 1) {
            self.close_block(end);
        }
        if self.part && depth == 2 {
            self.item = None;
        }
        self.depth = depth
            .checked_sub(1)
            .ok_or_else(|| xml_error("unbalanced story XML"))?;
        Ok(())
    }

    fn close_word(&mut self, local: &[u8], start: usize, end: usize) -> Result<()> {
        let depth = self.depth;
        match local {
            b"t" if self.text.is_some()
                && self.runs.last().is_some_and(|run| run.depth + 1 == depth) =>
            {
                let Some(open) = self.text.take() else {
                    return Ok(());
                };
                // Plain content (no reference, markup or carriage return)
                // is its own text; anything else takes the full decoder.
                let content = self
                    .xml
                    .get(open.start_tag_end..start)
                    .ok_or_else(|| xml_error("story text offset is out of range"))?;
                let value = if content
                    .iter()
                    .any(|byte| matches!(byte, b'&' | b'<' | b'\r'))
                {
                    decode_text_fragment(
                        self.xml
                            .get(open.start..end)
                            .ok_or_else(|| xml_error("story text offset is out of range"))?,
                    )
                    .map_err(|_| xml_error("invalid story text"))?
                } else {
                    std::str::from_utf8(content)
                        .map_err(|_| xml_error("story XML is not UTF-8"))?
                        .to_owned()
                };
                if value.is_empty() {
                    return Ok(());
                }
                // Text inside a field instruction is fixed content.
                if self.fields.contains(&Phase::Instruction) {
                    if let Some(paragraph) = self.paragraphs.last_mut() {
                        paragraph.push_fixed(OBJECT);
                    }
                    return Ok(());
                }
                let Some(run) = self.runs.last_mut() else {
                    return Ok(());
                };
                run.has_text = true;
                let Some(paragraph) = self.paragraphs.last_mut() else {
                    return Ok(());
                };
                let slot = match run.slot {
                    Some(slot) => slot,
                    None => {
                        paragraph.runs.push(RunSlot {
                            element: run.start..run.start,
                            open_end: run.open_end,
                            close_start: run.open_end,
                            properties: None,
                            bindings: run.bindings,
                        });
                        let slot = paragraph.runs.len() - 1;
                        run.slot = Some(slot);
                        slot
                    },
                };
                let text_start = paragraph.text.len();
                paragraph.text.push_str(&value);
                paragraph.pieces.push(Piece {
                    text: text_start..paragraph.text.len(),
                    node: Some(TextNode {
                        element: open.start..end,
                        start_tag_end: open.start_tag_end,
                        name: open.name,
                        space: open.space,
                        run: slot,
                        group: paragraph.group,
                        scoped: true,
                    }),
                });
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
                self.close_run(run, start, end);
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
                self.close_paragraph(paragraph, end);
            },
            _ => {},
        }
        Ok(())
    }

    fn close_paragraph(&mut self, mut paragraph: OpenParagraph, end: usize) {
        if self.main && paragraph.block.is_none() {
            return;
        }
        let element = paragraph.start..end;
        let mut scoped: Option<Vec<Range<usize>>> = Some(Vec::new());
        let mut listed = false;
        for scope in &self.listing {
            match scope.admits(&element) {
                None => {},
                Some(None) => {
                    listed = true;
                    scoped = None;
                },
                Some(Some(range)) => {
                    listed = true;
                    if let Some(ranges) = &mut scoped {
                        ranges.push(range);
                    }
                },
            }
        }
        if !listed {
            return;
        }
        if let Some(ranges) = &scoped {
            for piece in &mut paragraph.pieces {
                if let Some(node) = &mut piece.node {
                    node.scoped = ranges.iter().any(|range| {
                        range.start <= node.element.start && node.element.end <= range.end
                    });
                }
            }
        }
        // Fields still open here span paragraphs: listed, never replaceable.
        for field in paragraph.complex.drain(..) {
            paragraph.fields.push(Field {
                range: field.range_start..paragraph.text.len(),
                instruction: field.instruction,
                kind: FieldKind::Complex,
                replaceable: false,
                markup: field.begin_run..field.begin_run,
                properties: None,
                run_name: self.xml[field.run_name].to_vec(),
            });
        }
        paragraph.fields.sort_by_key(|field| field.range.start);
        self.done.push(Paragraph {
            text: paragraph.text,
            fields: paragraph.fields,
            pieces: paragraph.pieces,
            runs: paragraph.runs,
            block: paragraph.block,
            element,
            bindings: paragraph.bindings,
            entry: paragraph.entry,
            item: paragraph.item,
            alternate: paragraph.alternate,
        });
    }

    fn close_run(&mut self, run: OpenRun, close_start: usize, end: usize) {
        let Some(paragraph) = self.paragraphs.last_mut() else {
            return;
        };
        if let Some(slot) = run.slot.and_then(|slot| paragraph.runs.get_mut(slot)) {
            slot.element = run.start..end;
            slot.close_start = close_start;
            slot.properties = run.properties.clone();
        }
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

pub(super) fn qualified(prefix: &[u8], local: &[u8]) -> Vec<u8> {
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
