//! Office Math equations of a paragraph: insertion, replacement and removal.
//!
//! A paragraph's equations are counted as its readers count them
//! ([`crate::Paragraph::omml_equations`] and
//! [`crate::Paragraph::paragraph_level_equations`]): every outermost
//! `m:oMath` of the paragraph, inline when it has no `m:oMathPara` ancestor
//! and display when it has one. Each edit rewrites only the bytes of the
//! selected equation (or, for a removal that empties it, its math
//! paragraph) and records them as one exact element replacement of the
//! paragraph, so the rest of the part keeps its bytes.
//!
//! New equation markup is fitted to its insertion point: a namespace
//! declaration that repeats a binding already in scope there is dropped (Word
//! declares `m` and `w` once, on the part's root), and the math and
//! WordprocessingML namespaces take the names of the document's conformance.
//! New markup is admitted only when it holds Office Math elements and the
//! Word run content of a math run: anything else (a reference, a range, a
//! field, a relationship, VML, alternate content, an unknown namespace) is
//! refused with [`Refusal::EquationContent`]. Namespace bindings are tracked in one stack
//! per prefix, so a scan costs the paragraph's size, not its size times
//! the bindings in scope.

use std::collections::HashMap;
use std::ops::Range;
use std::sync::Arc;

use litchi_core::Position;
use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::events::{BytesStart, Event};

use super::super::{
    Removal, buffer, inside_complex_field, invalid, opened, paragraph_dependencies, root_tag,
    staging_add, tag_name,
};
use super::{
    Edit, ElementAnchor, ElementPath, ElementStep, Error, Refusal, Result, bindings_at,
    child_index, container_children, element_children, locate, not_found, refusal,
};
use crate::{OfficeMath, OfficeMathParagraph};

const TRANSITIONAL_MATH: &[u8] = b"http://schemas.openxmlformats.org/officeDocument/2006/math";
const STRICT_MATH: &[u8] = b"http://purl.oclc.org/ooxml/officeDocument/math";
const TRANSITIONAL_WORD: &[u8] = b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";
const MARKUP_COMPATIBILITY: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";
const RELATIONSHIPS: [&[u8]; 2] = [
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    b"http://purl.oclc.org/ooxml/officeDocument/relationships",
];
/// The deepest element nesting and the most elements a paragraph scan
/// accepts; the paragraph readers apply the same bounds.
const MAX_SCAN_DEPTH: usize = 128;
const MAX_SCAN_NODES: usize = 1_000_000;

/// One equation of a paragraph, as an edit selects it.
///
/// ```
/// use litchi_core::Position;
/// use litchi_docx::document::ParagraphEquation;
///
/// // The paragraph's second inline equation.
/// let selected = ParagraphEquation::Inline(Position::new(1));
/// assert_eq!(selected.position(), Position::new(1));
/// ```
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ParagraphEquation {
    /// The equation at this position among the paragraph's inline equations
    /// (the order of [`crate::Paragraph::omml_equations`]).
    Inline(Position),
    /// The equation at this position among the paragraph's display
    /// equations (the order of
    /// [`crate::Paragraph::paragraph_level_equations`]).
    Display(Position),
}

impl ParagraphEquation {
    /// The equation's position among the paragraph's equations of its
    /// kind.
    #[must_use]
    pub const fn position(self) -> Position {
        match self {
            Self::Inline(position) | Self::Display(position) => position,
        }
    }
}

/// Namespace bindings in scope, innermost last.
type Scope = Vec<(Option<Vec<u8>>, Vec<u8>)>;

/// One outermost `m:oMath` of a paragraph fragment.
#[derive(Clone)]
struct Span {
    range: Range<usize>,
    /// The enclosing math paragraph, for a display equation.
    paragraph: Option<usize>,
}

/// The equations of a paragraph fragment.
struct Equations {
    spans: Vec<Span>,
    /// Each math paragraph's range and the number of equations it holds.
    paragraphs: Vec<(Range<usize>, usize)>,
    /// Whether an equation sits inside markup-compatibility alternate
    /// content, whose branches readers project to one: positions then
    /// differ between the raw part and what readers count.
    alternate: bool,
}

impl Equations {
    fn select(&self, selected: ParagraphEquation) -> Result<&Span> {
        if self.alternate {
            return Err(refusal(Refusal::ComplexContent));
        }
        let display = matches!(selected, ParagraphEquation::Display(_));
        self.spans
            .iter()
            .filter(|span| span.paragraph.is_some() == display)
            .nth(selected.position().get())
            .ok_or_else(|| refusal(Refusal::EquationNotFound))
    }
}

/// The namespace declarations of a start tag.
fn declarations(element: &BytesStart<'_>) -> Result<Scope> {
    let mut found = Vec::new();
    for attribute in element.unchecked_attributes() {
        let attribute = attribute.map_err(|_| malformed())?;
        let key = attribute.key.as_ref();
        if key == b"xmlns" {
            found.push((None, attribute.value.to_vec()));
        } else if let Some(prefix) = key.strip_prefix(b"xmlns:") {
            found.push((Some(prefix.to_vec()), attribute.value.to_vec()));
        }
    }
    Ok(found)
}

fn malformed() -> Error {
    Error::Document(crate::Error::Xml("invalid Word paragraph XML".into()))
}

/// The namespace bindings in scope during a scan: one stack of the URIs
/// bound to each prefix, and the prefixes each open element declared, so
/// opening and closing an element costs its own declarations and a lookup
/// costs one hash probe, however many bindings are in scope.
struct Bindings {
    bound: HashMap<Option<Vec<u8>>, Vec<Vec<u8>>>,
    opened: Vec<Vec<Option<Vec<u8>>>>,
}

impl Bindings {
    fn new(seed: &[(Option<Vec<u8>>, Vec<u8>)]) -> Self {
        let mut bindings = Self {
            bound: HashMap::new(),
            opened: Vec::new(),
        };
        bindings.open(seed.to_vec());
        bindings
    }

    /// Enter an element that declares `declarations`.
    fn open(&mut self, declarations: Scope) {
        let mut prefixes = Vec::with_capacity(declarations.len());
        for (prefix, namespace) in declarations {
            self.bound
                .entry(prefix.clone())
                .or_default()
                .push(namespace);
            prefixes.push(prefix);
        }
        self.opened.push(prefixes);
    }

    /// Leave the innermost open element.
    fn close(&mut self) -> Result<()> {
        let prefixes = self.opened.pop().ok_or_else(malformed)?;
        for prefix in prefixes {
            if let Some(stack) = self.bound.get_mut(&prefix) {
                stack.pop();
            }
        }
        Ok(())
    }

    /// The open elements below the seed.
    fn depth(&self) -> usize {
        self.opened.len().saturating_sub(1)
    }

    fn lookup(&self, prefix: Option<&[u8]>) -> Option<&[u8]> {
        self.bound
            .get(&prefix.map(<[u8]>::to_vec))
            .and_then(|stack| stack.last())
            .map(Vec::as_slice)
    }

    /// The bindings in scope, one per prefix.
    fn snapshot(&self) -> Scope {
        self.bound
            .iter()
            .filter_map(|(prefix, stack)| Some((prefix.clone(), stack.last()?.clone())))
            .collect()
    }

    /// The namespace of an element's name.
    fn namespace_of(&self, element: &BytesStart<'_>) -> Option<&[u8]> {
        let name = element.name();
        self.lookup(name.prefix().map(|prefix| prefix.into_inner()))
    }
}

fn is_math(namespace: Option<&[u8]>) -> bool {
    matches!(namespace, Some(TRANSITIONAL_MATH | STRICT_MATH))
}

fn is_word(namespace: Option<&[u8]>) -> bool {
    matches!(namespace, Some(TRANSITIONAL_WORD | STRICT_WORD))
}

/// The bindings in scope at byte `offset` of `xml`, a position between
/// markup (an element's start or a content position), starting from
/// `seed`.
fn scope_at(seed: &[(Option<Vec<u8>>, Vec<u8>)], xml: &[u8], offset: usize) -> Result<Scope> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let mut bindings = Bindings::new(seed);
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("Word XML offset overflow"))?;
        if start >= offset {
            break;
        }
        match reader.read_event().map_err(|_| malformed())? {
            Event::Start(element) => bindings.open(declarations(&element)?),
            Event::End(_) => bindings.close()?,
            Event::Eof => return Err(invalid("scope offset is outside the XML")),
            _ => {},
        }
    }
    Ok(bindings.snapshot())
}

/// The equations of the paragraph fragment `xml`, whose root's parent has
/// the bindings `seed`.
fn equations(xml: &[u8], seed: &[(Option<Vec<u8>>, Vec<u8>)]) -> Result<Equations> {
    #[derive(Clone, Copy, PartialEq, Eq)]
    enum Open {
        Math,
        MathParagraph(usize),
        Alternate,
        Other,
    }
    let mut reader = quick_xml::Reader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let mut bindings = Bindings::new(seed);
    // What each open element is.
    let mut stack: Vec<Open> = Vec::new();
    // The open math paragraph and the number of open alternate-content
    // elements.
    let mut math_paragraph: Option<usize> = None;
    let mut alternates = 0usize;
    let mut found = Equations {
        spans: Vec::new(),
        paragraphs: Vec::new(),
        alternate: false,
    };
    // The open equation: its start, its stack depth and its span fields.
    let mut capture: Option<(usize, usize, Option<usize>)> = None;
    let mut nodes = 0usize;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("Word XML offset overflow"))?;
        let event = reader.read_event().map_err(|_| malformed())?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("Word XML offset overflow"))?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                nodes += 1;
                if nodes > MAX_SCAN_NODES || stack.len() >= MAX_SCAN_DEPTH {
                    return Err(invalid("Office Math scan limit exceeded"));
                }
                bindings.open(declarations(element)?);
                let local = element.local_name();
                let local = local.as_ref();
                let namespace = bindings.namespace_of(element);
                let math = is_math(namespace);
                let open = if capture.is_some() {
                    Open::Other
                } else if math && local == b"oMath" {
                    if alternates > 0 {
                        found.alternate = true;
                    }
                    if let Some(index) = math_paragraph
                        && let Some(paragraph) = found.paragraphs.get_mut(index)
                    {
                        paragraph.1 += 1;
                    }
                    Open::Math
                } else if math && local == b"oMathPara" && math_paragraph.is_none() {
                    found.paragraphs.push((start..start, 0));
                    Open::MathParagraph(found.paragraphs.len() - 1)
                } else if namespace == Some(MARKUP_COMPATIBILITY) && local == b"AlternateContent" {
                    Open::Alternate
                } else {
                    Open::Other
                };
                if matches!(event, Event::Empty(_)) {
                    bindings.close()?;
                    match open {
                        Open::Math => found.spans.push(Span {
                            range: start..end,
                            paragraph: math_paragraph,
                        }),
                        Open::MathParagraph(index) => {
                            if let Some(paragraph) = found.paragraphs.get_mut(index) {
                                paragraph.0 = start..end;
                            }
                        },
                        Open::Alternate | Open::Other => {},
                    }
                } else {
                    match open {
                        Open::Math => {
                            capture = Some((start, stack.len() + 1, math_paragraph));
                        },
                        Open::MathParagraph(index) => math_paragraph = Some(index),
                        Open::Alternate => alternates += 1,
                        Open::Other => {},
                    }
                    stack.push(open);
                }
            },
            Event::End(_) => {
                let depth = stack.len();
                let open = stack.pop().ok_or_else(malformed)?;
                bindings.close()?;
                match open {
                    Open::MathParagraph(index) => {
                        if let Some(paragraph) = found.paragraphs.get_mut(index) {
                            paragraph.0.end = end;
                        }
                        math_paragraph = None;
                    },
                    Open::Alternate => alternates = alternates.saturating_sub(1),
                    Open::Math | Open::Other => {},
                }
                if capture.is_some_and(|(_, captured, _)| captured == depth)
                    && let Some((from, _, paragraph)) = capture.take()
                {
                    found.spans.push(Span {
                        range: from..end,
                        paragraph,
                    });
                }
            },
            Event::Eof => {
                if !stack.is_empty() || bindings.depth() != 0 {
                    return Err(malformed());
                }
                return Ok(found);
            },
            _ => {},
        }
    }
}

/// The equation's markup fitted to a parent with the bindings `scope`:
/// declarations already in scope are dropped, and the math and
/// WordprocessingML namespaces take the document's conformance (Strict
/// names in a Strict document, Transitional ones otherwise). Every other
/// byte is kept.
fn fit_namespaces(
    xml: &[u8],
    scope: &[(Option<Vec<u8>>, Vec<u8>)],
    strict: bool,
) -> Result<Vec<u8>> {
    let convert = |namespace: &[u8]| -> Vec<u8> {
        match (strict, namespace) {
            (true, TRANSITIONAL_MATH) => STRICT_MATH.to_vec(),
            (true, TRANSITIONAL_WORD) => STRICT_WORD.to_vec(),
            (false, STRICT_MATH) => TRANSITIONAL_MATH.to_vec(),
            (false, STRICT_WORD) => TRANSITIONAL_WORD.to_vec(),
            _ => namespace.to_vec(),
        }
    };
    let mut reader = quick_xml::Reader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let mut bindings = Bindings::new(scope);
    let mut output = buffer(xml.len())?;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("Office Math offset overflow"))?;
        let event = reader.read_event().map_err(|_| malformed())?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("Office Math offset overflow"))?;
        let raw = xml.get(start..end).ok_or_else(malformed)?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                let mut kept: Scope = Vec::new();
                let mut tag = Vec::with_capacity(raw.len());
                tag.push(b'<');
                tag.extend_from_slice(element.name().as_ref());
                let mut changed = false;
                for attribute in element.unchecked_attributes() {
                    let attribute = attribute.map_err(|_| malformed())?;
                    let key = attribute.key.as_ref();
                    let prefix = if key == b"xmlns" {
                        Some(None)
                    } else {
                        key.strip_prefix(b"xmlns:").map(Some)
                    };
                    let value = match prefix {
                        Some(prefix) => {
                            let namespace = convert(&attribute.value);
                            changed |= namespace.as_slice() != attribute.value.as_ref();
                            let earlier = kept
                                .iter()
                                .rev()
                                .find(|(bound, _)| bound.as_deref() == prefix)
                                .map(|(_, namespace)| namespace.as_slice());
                            if earlier.or_else(|| bindings.lookup(prefix))
                                == Some(namespace.as_slice())
                            {
                                changed = true;
                                continue;
                            }
                            kept.push((prefix.map(<[u8]>::to_vec), namespace.clone()));
                            namespace
                        },
                        None => attribute.value.to_vec(),
                    };
                    let quote = if value.contains(&b'"') { b'\'' } else { b'"' };
                    tag.push(b' ');
                    tag.extend_from_slice(key);
                    tag.push(b'=');
                    tag.push(quote);
                    tag.extend_from_slice(&value);
                    tag.push(quote);
                }
                let empty = matches!(event, Event::Empty(_));
                tag.extend_from_slice(if empty { b"/>" } else { b">" });
                output.extend_from_slice(if changed { &tag } else { raw });
                if !empty {
                    bindings.open(kept);
                }
            },
            Event::End(_) => {
                if bindings.depth() == 0 {
                    return Err(malformed());
                }
                bindings.close()?;
                output.extend_from_slice(raw);
            },
            Event::Eof => break,
            _ => output.extend_from_slice(raw),
        }
    }
    staging_add(output.len(), 0)?;
    Ok(output)
}

/// What equation markup may not hold, given what an edit does with it.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Use {
    /// The markup is dropped: references it holds would be orphaned.
    Dropped,
    /// The markup is written: it must not bring references, ranges, fields,
    /// relationships or block content of its own.
    Written,
}

/// Refuse equation markup whose removal (`Use::Dropped`) would orphan what
/// it holds, or whose insertion (`Use::Written`) would add what an equation
/// does not own: one end of a range, a story reference, a field, a
/// relationship or block-level Word content.
fn admissible(xml: &[u8], scope: &[(Option<Vec<u8>>, Vec<u8>)], use_: Use) -> Result<()> {
    match use_ {
        Use::Dropped => droppable(xml, scope),
        Use::Written => writable(xml, scope).map_err(|()| refusal(Refusal::EquationContent)),
    }
}

/// Refuse to drop markup that holds one end of a range, a story reference
/// or a relationship.
fn droppable(xml: &[u8], scope: &[(Option<Vec<u8>>, Vec<u8>)]) -> Result<()> {
    paragraph_dependencies(xml, scope)
        .refuse(Removal::Delete)
        .map_err(refusal)?;
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut bindings = Bindings::new(scope);
    loop {
        let event = reader.read_event().map_err(|_| malformed())?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                bindings.open(declarations(element)?);
                for attribute in element.unchecked_attributes() {
                    let attribute = attribute.map_err(|_| malformed())?;
                    if let Some(prefix) = attribute.key.prefix()
                        && prefix.into_inner() != b"xmlns"
                        && bindings
                            .lookup(Some(prefix.into_inner()))
                            .is_some_and(|namespace| RELATIONSHIPS.contains(&namespace))
                    {
                        return Err(refusal(Refusal::StoryReference));
                    }
                }
                if matches!(event, Event::Empty(_)) {
                    bindings.close()?;
                }
            },
            Event::End(_) => bindings.close()?,
            Event::Eof => return Ok(()),
            _ => {},
        }
    }
}

/// The WordprocessingML run content a math run (`m:r`, ECMA-376 Part 1,
/// §22.1.2.87) may hold: `EG_RunInnerContent` without references, fields,
/// drawings and objects. A math run's `w:rPr`, a control's `w:rPr` and the
/// `w:ins`/`w:del` wrappers of math runs are admitted where the schema puts
/// them.
const WORD_IN_MATH: [&[u8]; 8] = [
    b"t",
    b"delText",
    b"sym",
    b"br",
    b"cr",
    b"tab",
    b"noBreakHyphen",
    b"softHyphen",
];

/// The run properties a `w:rPr` holds (`EG_RPrBase`, ECMA-376 Part 1,
/// §17.3.2.28), and the change-tracking markers of a run's properties.
const RUN_PROPERTIES: [&[u8]; 44] = [
    b"rStyle",
    b"rFonts",
    b"b",
    b"bCs",
    b"i",
    b"iCs",
    b"caps",
    b"smallCaps",
    b"strike",
    b"dstrike",
    b"outline",
    b"shadow",
    b"emboss",
    b"imprint",
    b"noProof",
    b"snapToGrid",
    b"vanish",
    b"webHidden",
    b"color",
    b"spacing",
    b"w",
    b"kern",
    b"position",
    b"sz",
    b"szCs",
    b"highlight",
    b"u",
    b"effect",
    b"bdr",
    b"shd",
    b"fitText",
    b"vertAlign",
    b"rtl",
    b"cs",
    b"em",
    b"lang",
    b"eastAsianLayout",
    b"specVanish",
    b"oMath",
    b"ins",
    b"del",
    b"moveFrom",
    b"moveTo",
    b"rPrChange",
];

/// Where an element of new equation markup sits, which decides what it may
/// hold.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Place {
    /// Office Math content or a tracked-change wrapper of math runs.
    Math,
    /// Inside a math run (`m:r`): its properties and run content.
    Run,
    /// Inside a math control's properties (`m:ctrlPr`): run properties and
    /// their tracked-change wrappers.
    Control,
    /// Directly inside a `w:rPr`: run properties.
    Properties,
    /// Inside a `w:rPrChange`: the original `w:rPr`.
    Change,
    /// Inside a run property: nothing.
    Leaf,
}

/// Whether new equation markup holds only what an equation owns: Office
/// Math elements, the WordprocessingML run content [`WORD_IN_MATH`] names,
/// the run properties [`RUN_PROPERTIES`] names inside a `w:rPr` (and the
/// original `w:rPr` of a `w:rPrChange`), and attributes in the math,
/// WordprocessingML or XML namespaces. Everything else (a relationship, a
/// story reference, one end of a range, a field, a drawing, VML,
/// markup-compatibility content or an element of an unknown namespace) is
/// refused.
fn writable(xml: &[u8], scope: &[(Option<Vec<u8>>, Vec<u8>)]) -> std::result::Result<(), ()> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut bindings = Bindings::new(scope);
    // Where each open element's children sit.
    let mut places: Vec<Place> = Vec::new();
    loop {
        let event = reader.read_event().map_err(|_| ())?;
        match event {
            Event::Start(ref element) | Event::Empty(ref element) => {
                bindings.open(declarations(element).map_err(|_| ())?);
                let namespace = bindings.namespace_of(element);
                let local = element.local_name();
                let local = local.as_ref();
                let word = is_word(namespace);
                let place = places.last().copied().unwrap_or(Place::Math);
                // Where this element's own children sit, when it is allowed.
                let math = is_math(namespace);
                let inner = match place {
                    Place::Math | Place::Run if math && local == b"r" => Some(Place::Run),
                    Place::Math | Place::Run if math && local == b"ctrlPr" => Some(Place::Control),
                    Place::Math | Place::Run if math => Some(Place::Math),
                    // Tracked changes of math runs, as Word writes them.
                    Place::Math if word && matches!(local, b"ins" | b"del") => Some(Place::Math),
                    Place::Run | Place::Control if word && local == b"rPr" => {
                        Some(Place::Properties)
                    },
                    Place::Run if word && WORD_IN_MATH.contains(&local) => Some(Place::Leaf),
                    Place::Control if word && matches!(local, b"ins" | b"del") => {
                        Some(Place::Control)
                    },
                    Place::Properties if word && local == b"rPrChange" => Some(Place::Change),
                    Place::Properties if word && RUN_PROPERTIES.contains(&local) => {
                        Some(Place::Leaf)
                    },
                    Place::Change if word && local == b"rPr" => Some(Place::Properties),
                    _ => None,
                };
                let inner = inner.ok_or(())?;
                for attribute in element.unchecked_attributes() {
                    let attribute = attribute.map_err(|_| ())?;
                    let Some(prefix) = attribute.key.prefix() else {
                        continue;
                    };
                    let prefix = prefix.into_inner();
                    if prefix == b"xmlns" || prefix == b"xml" {
                        continue;
                    }
                    let namespace = bindings.lookup(Some(prefix));
                    if !is_math(namespace) && !is_word(namespace) {
                        return Err(());
                    }
                }
                if matches!(event, Event::Empty(_)) {
                    bindings.close().map_err(|_| ())?;
                } else {
                    places.push(inner);
                }
            },
            Event::End(_) => {
                places.pop();
                bindings.close().map_err(|_| ())?;
            },
            Event::Eof => return Ok(()),
            _ => {},
        }
    }
}

impl Edit {
    /// The paragraph at `paragraph`: its bytes, the bindings in scope at
    /// it and its equations.
    fn paragraph_equations(
        &self,
        paragraph: &ElementPath,
    ) -> Result<(Arc<Vec<u8>>, Scope, Equations)> {
        if paragraph.steps().last().map(|step| step.local()) != Some(b"p".as_slice()) {
            return Err(not_found());
        }
        let located = locate(&self.projected, paragraph.steps())?;
        let seed = scope_at(&[], self.projected.xml_bytes(), located.start)?;
        let before = self.element_bytes(located)?;
        let found = equations(&before, &seed)?;
        Ok((before, seed, found))
    }

    /// Replace `range` of the paragraph at `paragraph`, whose bytes are
    /// `before`, with `xml`.
    fn splice_paragraph(
        &mut self,
        paragraph: &ElementPath,
        before: &Arc<Vec<u8>>,
        range: Range<usize>,
        xml: &[u8],
    ) -> Result<&mut Self> {
        let mut after = buffer(staging_add(before.len(), xml.len())?)?;
        after.extend_from_slice(before.get(..range.start).ok_or_else(malformed)?);
        after.extend_from_slice(xml);
        after.extend_from_slice(before.get(range.end..).ok_or_else(malformed)?);
        self.replace_element(paragraph, before, &Arc::new(after))
    }

    /// Replace the markup at `range` of a paragraph with `math` fitted to
    /// its parent, after checking what is dropped and what is written.
    /// Returns whether anything changed.
    fn replace_markup(
        &mut self,
        paragraph: &ElementPath,
        (before, seed): (&Arc<Vec<u8>>, &Scope),
        range: Range<usize>,
        math: &[u8],
    ) -> Result<bool> {
        let old = before.get(range.clone()).ok_or_else(malformed)?;
        let scope = scope_at(seed, before, range.start)?;
        admissible(old, &scope, Use::Dropped)?;
        let xml = fit_namespaces(math, &scope, self.is_strict())?;
        admissible(&xml, &scope, Use::Written)?;
        if xml.as_slice() == old {
            return Ok(false);
        }
        self.splice_paragraph(paragraph, before, range, &xml)?;
        Ok(true)
    }

    /// The range and equation count of the `index`-th math paragraph.
    fn math_paragraph(found: &Equations, index: Position) -> Result<(Range<usize>, usize)> {
        if found.alternate {
            return Err(refusal(Refusal::ComplexContent));
        }
        found
            .paragraphs
            .get(index.get())
            .cloned()
            .ok_or_else(|| refusal(Refusal::EquationNotFound))
    }

    /// Insert `math` as an inline equation of the paragraph at `paragraph`,
    /// before its direct run `at` (after its last child for `None`), and
    /// return its position among the paragraph's inline equations.
    ///
    /// Word writes inline math as an `m:oMath` beside the paragraph's runs
    /// (ECMA-376 Part 1, §22.1.2.77), which is what this inserts.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::OfficeMath;
    /// use litchi_docx::document::{raw, ElementPath, ElementStep};
    ///
    /// let mut edit = raw::Tree::from_xml(
    ///     br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body><w:p><w:r><w:t>Area: </w:t></w:r></w:p></w:body></w:document>"#
    ///         .to_vec(),
    ///     raw::Limits::default(),
    /// )?;
    /// let paragraph = ElementPath::new([ElementStep::Paragraph(Position::new(0))]);
    /// let math = OfficeMath::text("πr²");
    /// assert_eq!(edit.insert_equation_at(&paragraph, None, &math)?, Position::new(0));
    /// let xml = std::str::from_utf8(edit.element_xml(&paragraph)?)?;
    /// assert!(xml.ends_with("</w:r><m:oMath><m:r><m:t>πr²</m:t></m:r></m:oMath></w:p>"));
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::RunNotFound`] for a run position past the end,
    /// [`Refusal::FieldInterior`] inside a complex field,
    /// [`Refusal::ComplexContent`] when an equation of the paragraph sits in
    /// alternate content, [`Refusal::BlockNotFound`] for a path that names
    /// no paragraph, or a resource-limit error.
    pub fn insert_equation_at(
        &mut self,
        paragraph: &ElementPath,
        at: Option<Position>,
        math: &OfficeMath,
    ) -> Result<Position> {
        self.ensure_unmanaged("insert_equation_at")?;
        if paragraph.steps().last().map(|step| step.local()) != Some(b"p".as_slice()) {
            return Err(not_found());
        }
        let located = locate(&self.projected, paragraph.steps())?;
        let seed = scope_at(&[], self.projected.xml_bytes(), located.start)?;
        let inherited = bindings_at(self.projected.xml_bytes(), located.start)?;
        let before = self.element_bytes(located)?;
        // An empty paragraph (`<w:p/>`) is opened first.
        let base = match root_tag(&before)? {
            (_, true) => {
                let mut open = opened(&before)?;
                open.extend_from_slice(b"</");
                open.extend_from_slice(tag_name(&before)?);
                open.push(b'>');
                open
            },
            (_, false) => before.to_vec(),
        };
        let found = equations(&base, &seed)?;
        if found.alternate {
            return Err(refusal(Refusal::ComplexContent));
        }
        let layout = element_children(&base, b"p", &inherited)?;
        let runs = layout.count(b"r");
        let child = match at {
            None => layout.children.len(),
            Some(position) if position.get() == runs => layout.children.len(),
            Some(position) => layout
                .index_of(b"r", position)
                .ok_or_else(|| refusal(Refusal::RunNotFound))?,
        };
        let offset = layout
            .insertion_offset(child)
            .ok_or_else(|| refusal(Refusal::RunNotFound))?;
        if inside_complex_field(&base, offset, &inherited).map_err(refusal)? {
            return Err(refusal(Refusal::FieldInterior));
        }
        let scope = scope_at(&seed, &base, offset)?;
        let xml = fit_namespaces(math.xml().as_bytes(), &scope, self.is_strict())?;
        admissible(&xml, &scope, Use::Written)?;
        let mut after = buffer(staging_add(base.len(), xml.len())?)?;
        after.extend_from_slice(&base[..offset]);
        after.extend_from_slice(&xml);
        after.extend_from_slice(&base[offset..]);
        let position = found
            .spans
            .iter()
            .filter(|span| span.paragraph.is_none() && span.range.start < offset)
            .count();
        self.replace_element(paragraph, &before, &Arc::new(after))?;
        Ok(Position::new(position))
    }

    /// Insert a display-equation paragraph, a `w:p` holding `math`, into the
    /// body, a table cell or a block content control at `anchor` (in the
    /// body, `End` is before the body-final section properties), and return
    /// its position among the container's direct paragraphs.
    ///
    /// Word writes a display equation as a paragraph whose content is one
    /// `m:oMathPara` (ECMA-376 Part 1, §22.1.2.78), which is what this
    /// inserts.
    ///
    /// ```
    /// use litchi_docx::document::{raw, ElementAnchor, ElementPath};
    /// use litchi_docx::{OfficeMath, OfficeMathParagraph};
    ///
    /// let mut edit = raw::Tree::from_xml(
    ///     br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body><w:p/><w:sectPr/></w:body></w:document>"#
    ///         .to_vec(),
    ///     raw::Limits::default(),
    /// )?;
    /// let math = OfficeMathParagraph::from_equation(OfficeMath::text("x=1"));
    /// let position = edit.insert_display_equation_at(&ElementPath::body(), ElementAnchor::End, &math)?;
    /// assert_eq!(position.get(), 1);
    /// let xml = std::str::from_utf8(edit.projected().xml_bytes())?;
    /// assert!(xml.contains("<w:p/><w:p><m:oMathPara><m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath></m:oMathPara></w:p><w:sectPr/>"));
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for a missing container or sibling,
    /// or a resource-limit error.
    pub fn insert_display_equation_at(
        &mut self,
        container: &ElementPath,
        anchor: ElementAnchor,
        math: &OfficeMathParagraph,
    ) -> Result<Position> {
        self.ensure_unmanaged("insert_display_equation_at")?;
        let (children, local) = container_children(&self.projected, container)?;
        if !ElementStep::Paragraph(Position::new(0)).fits(local) {
            return Err(not_found());
        }
        let child = child_index(&children, local, anchor)?;
        let offset = children.insertion_offset(child).ok_or_else(not_found)?;
        let mut scope = scope_at(&[], self.projected.xml_bytes(), offset)?;
        let prefix = children.prefix.clone().filter(|prefix| !prefix.is_empty());
        let declare = prefix
            .is_none()
            .then(|| self.projected.conformance.namespace().as_bytes());
        let prefix = prefix.unwrap_or_else(|| b"w".to_vec());
        let mut xml = Vec::new();
        xml.push(b'<');
        xml.extend_from_slice(&prefix);
        xml.extend_from_slice(b":p");
        if let Some(namespace) = declare {
            xml.extend_from_slice(b" xmlns:");
            xml.extend_from_slice(&prefix);
            xml.extend_from_slice(b"=\"");
            xml.extend_from_slice(namespace);
            xml.push(b'"');
            scope.push((Some(prefix.clone()), namespace.to_vec()));
        }
        xml.push(b'>');
        let fitted = fit_namespaces(math.xml().as_bytes(), &scope, self.is_strict())?;
        admissible(&fitted, &scope, Use::Written)?;
        xml.extend_from_slice(&fitted);
        xml.extend_from_slice(b"</");
        xml.extend_from_slice(&prefix);
        xml.extend_from_slice(b":p>");
        let position = Position::new(children.count_before(b"p", child));
        let xml = Arc::new(xml);
        if container.steps().is_empty() {
            self.insert_paragraph_fragment(Position::new(child), position, &xml)?;
        } else {
            self.insert_element(container, Position::new(child), &xml)?;
        }
        Ok(position)
    }

    /// Replace the equation `equation` of the paragraph at `paragraph` with
    /// `math`, and return whether anything changed. Only the equation's own
    /// `m:oMath` is rewritten: the paragraph's properties, its runs, and a
    /// display equation's math-paragraph properties keep their bytes.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::OfficeMath;
    /// use litchi_docx::document::{raw, ElementPath, ElementStep, ParagraphEquation};
    ///
    /// let mut edit = raw::Tree::from_xml(
    ///     br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body><w:p><w:pPr><w:jc w:val="center"/></w:pPr><m:oMathPara><m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr><m:oMath><m:r><m:t>a</m:t></m:r></m:oMath></m:oMathPara></w:p></w:body></w:document>"#
    ///         .to_vec(),
    ///     raw::Limits::default(),
    /// )?;
    /// let paragraph = ElementPath::new([ElementStep::Paragraph(Position::new(0))]);
    /// let selected = ParagraphEquation::Display(Position::new(0));
    /// assert!(edit.replace_equation_at(&paragraph, selected, &OfficeMath::text("b"))?);
    /// let xml = std::str::from_utf8(edit.element_xml(&paragraph)?)?;
    /// assert_eq!(
    ///     xml,
    ///     r#"<w:p><w:pPr><w:jc w:val="center"/></w:pPr><m:oMathPara><m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr><m:oMath><m:r><m:t>b</m:t></m:r></m:oMath></m:oMathPara></w:p>"#
    /// );
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::EquationNotFound`] for a missing equation,
    /// [`Refusal::ComplexContent`] when an equation of the paragraph sits in
    /// alternate content, [`Refusal::OpenRange`] or
    /// [`Refusal::StoryReference`] when the replaced markup holds one end of
    /// a range, a story reference or a relationship,
    /// [`Refusal::EquationContent`] when `math` holds content other than
    /// Office Math and Word run content, [`Refusal::BlockNotFound`] for a
    /// path that names no paragraph, or a resource-limit error.
    pub fn replace_equation_at(
        &mut self,
        paragraph: &ElementPath,
        equation: ParagraphEquation,
        math: &OfficeMath,
    ) -> Result<bool> {
        self.ensure_unmanaged("replace_equation_at")?;
        let (before, seed, found) = self.paragraph_equations(paragraph)?;
        let range = found.select(equation)?.range.clone();
        self.replace_markup(paragraph, (&before, &seed), range, math.xml().as_bytes())
    }

    /// Replace the `index`-th `m:oMathPara` of the paragraph at `paragraph`,
    /// its properties and every equation it holds, with `math`, and return
    /// whether anything changed.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::{OfficeMath, OfficeMathParagraph};
    /// use litchi_docx::document::{raw, ElementPath, ElementStep};
    ///
    /// let mut edit = raw::Tree::from_xml(
    ///     br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body><w:p><m:oMathPara><m:oMath><m:r><m:t>a</m:t></m:r></m:oMath><m:oMath><m:r><m:t>b</m:t></m:r></m:oMath></m:oMathPara></w:p></w:body></w:document>"#
    ///         .to_vec(),
    ///     raw::Limits::default(),
    /// )?;
    /// let paragraph = ElementPath::new([ElementStep::Paragraph(Position::new(0))]);
    /// let math = OfficeMathParagraph::from_equation(OfficeMath::text("c"));
    /// assert!(edit.replace_math_paragraph_at(&paragraph, Position::new(0), &math)?);
    /// let xml = std::str::from_utf8(edit.element_xml(&paragraph)?)?;
    /// assert_eq!(xml, "<w:p><m:oMathPara><m:oMath><m:r><m:t>c</m:t></m:r></m:oMath></m:oMathPara></w:p>");
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// As [`Self::replace_equation_at`].
    pub fn replace_math_paragraph_at(
        &mut self,
        paragraph: &ElementPath,
        index: Position,
        math: &OfficeMathParagraph,
    ) -> Result<bool> {
        self.ensure_unmanaged("replace_math_paragraph_at")?;
        let (before, seed, found) = self.paragraph_equations(paragraph)?;
        let (range, _) = Self::math_paragraph(&found, index)?;
        self.replace_markup(paragraph, (&before, &seed), range, math.xml().as_bytes())
    }

    /// Replace the equations of the `index`-th `m:oMathPara` of the
    /// paragraph at `paragraph` with the one equation `math`, keeping the
    /// math paragraph's properties (`m:oMathParaPr`), and return whether
    /// anything changed.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::OfficeMath;
    /// use litchi_docx::document::{raw, ElementPath, ElementStep};
    ///
    /// let mut edit = raw::Tree::from_xml(
    ///     br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body><w:p><w:bookmarkStart w:id="0" w:name="eq"/><m:oMathPara><m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr><m:oMath><m:r><m:t>a</m:t></m:r></m:oMath><m:oMath><m:r><m:t>b</m:t></m:r></m:oMath></m:oMathPara><w:bookmarkEnd w:id="0"/></w:p></w:body></w:document>"#
    ///         .to_vec(),
    ///     raw::Limits::default(),
    /// )?;
    /// let paragraph = ElementPath::new([ElementStep::Paragraph(Position::new(0))]);
    /// assert!(edit.replace_math_paragraph_equations_at(&paragraph, Position::new(0), &OfficeMath::text("c"))?);
    /// let xml = std::str::from_utf8(edit.element_xml(&paragraph)?)?;
    /// assert!(xml.contains(r#"<m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr><m:oMath><m:r><m:t>c</m:t></m:r></m:oMath></m:oMathPara>"#));
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// As [`Self::replace_equation_at`].
    pub fn replace_math_paragraph_equations_at(
        &mut self,
        paragraph: &ElementPath,
        index: Position,
        math: &OfficeMath,
    ) -> Result<bool> {
        self.ensure_unmanaged("replace_math_paragraph_equations_at")?;
        let (before, seed, found) = self.paragraph_equations(paragraph)?;
        Self::math_paragraph(&found, index)?;
        let mut held = found
            .spans
            .iter()
            .filter(|span| span.paragraph == Some(index.get()));
        let first = held
            .next()
            .ok_or_else(|| refusal(Refusal::EquationNotFound))?
            .range
            .clone();
        let last = held.next_back().map_or(first.end, |span| span.range.end);
        self.replace_markup(
            paragraph,
            (&before, &seed),
            first.start..last,
            math.xml().as_bytes(),
        )
    }

    /// Remove the `index`-th `m:oMathPara` of the paragraph at `paragraph`
    /// with every equation it holds; the paragraph itself stays.
    ///
    /// # Errors
    ///
    /// As [`Self::replace_equation_at`].
    pub fn remove_math_paragraph_at(
        &mut self,
        paragraph: &ElementPath,
        index: Position,
    ) -> Result<&mut Self> {
        self.ensure_unmanaged("remove_math_paragraph_at")?;
        let (before, seed, found) = self.paragraph_equations(paragraph)?;
        let (range, _) = Self::math_paragraph(&found, index)?;
        let scope = scope_at(&seed, &before, range.start)?;
        admissible(
            before.get(range.clone()).ok_or_else(malformed)?,
            &scope,
            Use::Dropped,
        )?;
        self.splice_paragraph(paragraph, &before, range, &[])
    }

    /// Remove the equation `equation` of the paragraph at `paragraph`. A
    /// display equation that is the only one of its `m:oMathPara` takes the
    /// math paragraph with it; the paragraph itself stays (remove it with
    /// [`Self::remove_element_at`]).
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::document::{raw, ElementPath, ElementStep, ParagraphEquation};
    ///
    /// let mut edit = raw::Tree::from_xml(
    ///     br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body><w:p><w:r><w:t>x</w:t></w:r><m:oMath><m:r><m:t>y</m:t></m:r></m:oMath></w:p></w:body></w:document>"#
    ///         .to_vec(),
    ///     raw::Limits::default(),
    /// )?;
    /// let paragraph = ElementPath::new([ElementStep::Paragraph(Position::new(0))]);
    /// edit.remove_equation_at(&paragraph, ParagraphEquation::Inline(Position::new(0)))?;
    /// let xml = std::str::from_utf8(edit.element_xml(&paragraph)?)?;
    /// assert_eq!(xml, "<w:p><w:r><w:t>x</w:t></w:r></w:p>");
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// As [`Self::replace_equation_at`].
    pub fn remove_equation_at(
        &mut self,
        paragraph: &ElementPath,
        equation: ParagraphEquation,
    ) -> Result<&mut Self> {
        self.ensure_unmanaged("remove_equation_at")?;
        let (before, seed, found) = self.paragraph_equations(paragraph)?;
        let span = found.select(equation)?;
        let range = match span.paragraph.and_then(|index| found.paragraphs.get(index)) {
            Some((range, 1)) => range.clone(),
            _ => span.range.clone(),
        };
        let scope = scope_at(&seed, &before, range.start)?;
        admissible(
            before.get(range.clone()).ok_or_else(malformed)?,
            &scope,
            Use::Dropped,
        )?;
        self.splice_paragraph(paragraph, &before, range, &[])
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const DECLARATIONS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math""#;

    fn scope() -> Scope {
        vec![
            (Some(b"w".to_vec()), TRANSITIONAL_WORD.to_vec()),
            (Some(b"m".to_vec()), TRANSITIONAL_MATH.to_vec()),
        ]
    }

    #[test]
    fn declarations_in_scope_are_dropped_and_others_kept() {
        let xml = format!(
            r#"<m:oMath {DECLARATIONS}><m:r><w:rPr><w:b/></w:rPr><m:t>x</m:t></m:r></m:oMath>"#
        );
        let fitted = fit_namespaces(xml.as_bytes(), &scope(), false).unwrap();
        assert_eq!(
            fitted,
            b"<m:oMath><m:r><w:rPr><w:b/></w:rPr><m:t>x</m:t></m:r></m:oMath>"
        );
        // Without the bindings in scope the fragment keeps its own.
        let kept = fit_namespaces(xml.as_bytes(), &[], false).unwrap();
        assert_eq!(kept, xml.as_bytes());
        // A prefix bound elsewhere keeps the fragment's own declaration.
        let other = vec![(Some(b"m".to_vec()), b"urn:other".to_vec())];
        let fitted = fit_namespaces(xml.as_bytes(), &other, false).unwrap();
        assert!(fitted.starts_with(br#"<m:oMath xmlns:w="#));
        assert!(fitted.windows(9).any(|window| window == b"xmlns:m=\""));
    }

    #[test]
    fn strict_documents_take_the_strict_namespaces() {
        let xml = format!(r#"<m:oMath {DECLARATIONS}><m:r><m:t>x</m:t></m:r></m:oMath>"#);
        let fitted = String::from_utf8(fit_namespaces(xml.as_bytes(), &[], true).unwrap()).unwrap();
        assert_eq!(
            fitted,
            r#"<m:oMath xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math"><m:r><m:t>x</m:t></m:r></m:oMath>"#
        );
        let strict = vec![
            (Some(b"w".to_vec()), STRICT_WORD.to_vec()),
            (Some(b"m".to_vec()), STRICT_MATH.to_vec()),
        ];
        let fitted = fit_namespaces(xml.as_bytes(), &strict, true).unwrap();
        assert_eq!(fitted, b"<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>");
    }

    #[test]
    fn nested_declarations_are_fitted_against_their_parent() {
        let xml = r#"<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara>"#;
        let fitted = fit_namespaces(xml.as_bytes(), &[], false).unwrap();
        assert_eq!(
            fitted,
            br#"<m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara>"#
        );
    }

    #[test]
    fn equations_are_classified_as_the_readers_classify_them() {
        let xml = br#"<w:p><w:r><m:oMath><m:r><m:t>a</m:t></m:r></m:oMath></w:r><m:oMathPara><m:oMath/><m:oMath><m:r><m:t>b</m:t></m:r></m:oMath></m:oMathPara><m:oMath><m:r><m:t>c</m:t></m:r></m:oMath></w:p>"#;
        let found = equations(xml, &scope()).unwrap();
        let kinds: Vec<bool> = found
            .spans
            .iter()
            .map(|span| span.paragraph.is_some())
            .collect();
        assert_eq!(kinds, [false, true, true, false]);
        assert_eq!(found.paragraphs.len(), 1);
        assert_eq!(found.paragraphs[0].1, 2);
        let display = found
            .select(ParagraphEquation::Display(Position::new(1)))
            .unwrap();
        assert!(xml[display.range.clone()].ends_with(b"<m:t>b</m:t></m:r></m:oMath>"));
        assert!(matches!(
            found.select(ParagraphEquation::Inline(Position::new(2))),
            Err(Error::Refused {
                reason: Refusal::EquationNotFound,
                ..
            })
        ));
    }

    #[test]
    fn equations_in_alternate_content_are_refused() {
        let xml = br#"<w:p xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:AlternateContent><mc:Choice Requires="m"><m:oMath/></mc:Choice><mc:Fallback/></mc:AlternateContent></w:p>"#;
        let found = equations(xml, &scope()).unwrap();
        assert!(found.alternate);
        assert!(matches!(
            found.select(ParagraphEquation::Inline(Position::new(0))),
            Err(Error::Refused {
                reason: Refusal::ComplexContent,
                ..
            })
        ));
    }

    #[test]
    fn markup_with_relationships_or_open_ranges_is_not_removable() {
        let relationship = br#"<m:oMath><m:r><w:drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a r:embed="rId1"/></w:drawing></m:r></m:oMath>"#;
        assert!(admissible(relationship, &scope(), Use::Dropped).is_err());
        let range =
            br#"<m:oMath><w:bookmarkStart w:id="1" w:name="x"/><m:r><m:t>x</m:t></m:r></m:oMath>"#;
        assert!(admissible(range, &scope(), Use::Dropped).is_err());
        let plain = br#"<m:oMath><m:r><m:t>x</m:t></m:r></m:oMath>"#;
        assert!(admissible(plain, &scope(), Use::Dropped).is_ok());
        assert!(admissible(plain, &scope(), Use::Written).is_ok());
    }

    fn written(xml: &[u8]) -> Option<Refusal> {
        match admissible(xml, &scope(), Use::Written) {
            Ok(()) => None,
            Err(Error::Refused { reason, .. }) => Some(reason),
            Err(other) => panic!("not a refusal: {other:?}"),
        }
    }

    #[test]
    fn written_markup_brings_no_references_ranges_relationships_or_blocks() {
        let r = r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;
        let cases: [(String, Refusal); 10] = [
            // VML with an Office relationship id.
            (
                r#"<m:oMath xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><m:r><v:shape><v:imagedata o:relid="rId3"/></v:shape></m:r></m:oMath>"#.to_owned(),
                Refusal::EquationContent,
            ),
            // An element of an unknown namespace, and a Word name bound to one.
            (
                r#"<m:oMath><m:r><x:junk xmlns:x="urn:x"/></m:r></m:oMath>"#.to_owned(),
                Refusal::EquationContent,
            ),
            (
                r#"<m:oMath><m:r><w:drawing xmlns:w="urn:junk"/></m:r></m:oMath>"#.to_owned(),
                Refusal::EquationContent,
            ),
            // Markup-compatibility content.
            (
                r#"<m:oMath xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:AlternateContent><mc:Choice Requires="w"><m:r><m:t>x</m:t></m:r></mc:Choice></mc:AlternateContent></m:oMath>"#.to_owned(),
                Refusal::EquationContent,
            ),
            (
                format!(r#"<m:oMath {r}><m:r><w:drawing><a r:embed="rId999"/></w:drawing></m:r></m:oMath>"#),
                Refusal::EquationContent,
            ),
            (
                r#"<m:oMath><w:bookmarkStart w:id="7" w:name="b"/><m:r><m:t>x</m:t></m:r></m:oMath>"#.to_owned(),
                Refusal::EquationContent,
            ),
            // A balanced pair still brings a range of its own.
            (
                r#"<m:oMath><w:bookmarkStart w:id="7" w:name="b"/><w:bookmarkEnd w:id="7"/></m:oMath>"#.to_owned(),
                Refusal::EquationContent,
            ),
            (
                r#"<m:oMath><m:r><w:footnoteReference w:id="99"/></m:r></m:oMath>"#.to_owned(),
                Refusal::EquationContent,
            ),
            (
                r#"<m:oMath><m:e><w:p><w:r><w:t>x</w:t></w:r></w:p></m:e></m:oMath>"#.to_owned(),
                Refusal::EquationContent,
            ),
            (
                r#"<m:oMath><m:r><w:fldChar w:fldCharType="begin"/></m:r><m:r><w:fldChar w:fldCharType="end"/></m:r></m:oMath>"#.to_owned(),
                Refusal::EquationContent,
            ),
        ];
        for (xml, wanted) in cases {
            assert_eq!(written(xml.as_bytes()), Some(wanted), "{xml}");
        }
        // Run properties, Word's math font, tracked insertions and run text
        // are fine.
        assert_eq!(
            written(br#"<m:oMath><m:r><w:rPr><w:rFonts w:ascii="Cambria Math"/><w:b/></w:rPr><m:t>x</m:t></m:r></m:oMath>"#),
            None
        );
        assert_eq!(
            written(br#"<m:oMath><m:ctrlPr><w:rPr><w:i/></w:rPr></m:ctrlPr><w:ins w:id="1" w:author="A"><m:r><w:t xml:space="preserve">y </w:t><w:tab/></m:r></w:ins></m:oMath>"#),
            None
        );
        // Run properties hold run properties only.
        for inside in [
            r#"<w:footnoteReference w:id="9"/>"#,
            r#"<w:bookmarkStart w:id="9" w:name="b"/>"#,
            "<w:drawing/>",
            "<w:p><w:r><w:t>x</w:t></w:r></w:p>",
            "<w:b><w:t>x</w:t></w:b>",
        ] {
            let xml = format!("<m:oMath><m:r><w:rPr>{inside}</w:rPr><m:t>x</m:t></m:r></m:oMath>");
            assert_eq!(
                written(xml.as_bytes()),
                Some(Refusal::EquationContent),
                "{xml}"
            );
        }
        assert_eq!(
            written(br#"<m:oMath><m:r><w:rPr><w:b/><w:rPrChange w:id="2" w:author="A"><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><m:t>x</m:t></m:r></m:oMath>"#),
            None
        );
        // Word run content belongs in a math run.
        for xml in [
            "<m:oMath><w:t>x</w:t></m:oMath>",
            "<m:oMath><w:rPr><w:b/></w:rPr></m:oMath>",
            "<m:oMath><m:f><m:num><w:sym w:char=\"F061\"/></m:num></m:f></m:oMath>",
        ] {
            assert_eq!(
                written(xml.as_bytes()),
                Some(Refusal::EquationContent),
                "{xml}"
            );
        }
        // An attribute of an unknown namespace is not.
        assert_eq!(
            written(br#"<m:oMath xmlns:x="urn:x"><m:r x:junk="1"><m:t>x</m:t></m:r></m:oMath>"#),
            Some(Refusal::EquationContent)
        );
    }

    #[test]
    fn transitional_documents_take_the_transitional_namespaces() {
        let xml = r#"<m:oMath xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math" xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"><m:r><w:rPr><w:b/></w:rPr><m:t>x</m:t></m:r></m:oMath>"#;
        let fitted = fit_namespaces(xml.as_bytes(), &scope(), false).unwrap();
        assert_eq!(
            fitted,
            b"<m:oMath><m:r><w:rPr><w:b/></w:rPr><m:t>x</m:t></m:r></m:oMath>"
        );
    }
}
