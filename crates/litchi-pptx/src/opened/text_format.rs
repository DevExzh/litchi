//! Rich formatting of existing `DrawingML` text bodies and table cells.
//!
//! The edits address a shape by its positional scene index and descend into
//! its text body (`p:txBody`) or, for a table frame, into one `a:tc` cell.
//! Paragraph and run positions count the direct `a:p` children of the body
//! and the `a:r`/`a:br`/`a:fld` children of a paragraph, in document order,
//! exactly as the text-body reader numbers them.
//!
//! Every rewrite is a byte splice: untouched siblings, attributes, extension
//! lists and unknown children keep their exact source bytes. Inserted
//! children follow the schema order of `a:rPr`, `a:pPr` and `a:tcPr` and use
//! the prefix of the element that receives them.

use std::borrow::Cow;
use std::ops::Range;

use litchi_core::Position;
use litchi_drawingml::text::Alignment;
use quick_xml::Reader;
use quick_xml::events::Event;

use crate::{Error, Result};

/// The largest accepted font size, in hundredths of a point (`ST_TextFontSize`).
pub const MAX_FONT_SIZE: u32 = 400_000;
/// The smallest accepted font size, in hundredths of a point (`ST_TextFontSize`).
pub const MIN_FONT_SIZE: u32 = 100;
/// The largest accepted typeface name, in bytes.
pub const MAX_TYPEFACE_BYTES: usize = 1_024;

/// The text element a [`TextEdit`] changes.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum TextTarget {
    /// Every paragraph and run of a shape's text body.
    Shape {
        /// Positional scene index of the shape.
        shape: Position,
    },
    /// One paragraph of a shape's text body.
    Paragraph {
        /// Positional scene index of the shape.
        shape: Position,
        /// Position among the body's `a:p` children.
        paragraph: Position,
    },
    /// One run (`a:r`, `a:br` or `a:fld`) of a paragraph.
    Run {
        /// Positional scene index of the shape.
        shape: Position,
        /// Position among the body's `a:p` children.
        paragraph: Position,
        /// Position among the paragraph's run children.
        run: Position,
    },
    /// One cell of a table graphic frame.
    Cell {
        /// Positional scene index of the table frame.
        shape: Position,
        /// Position among the table's `a:tr` rows.
        row: Position,
        /// Position among the row's `a:tc` cells.
        column: Position,
    },
}

/// An RGB colour written as `a:srgbClr@val`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Rgb(pub [u8; 3]);

/// Optional property assignments for one [`TextEdit`].
///
/// Each field is `None` to leave the property as it is, `Some(None)` to
/// remove the local property, and `Some(Some(value))` to set it.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[allow(
    clippy::option_option,
    reason = "omission, removal and replacement are distinct assignments"
)]
pub struct TextProperties<'a> {
    text: Option<&'a str>,
    bold: Option<Option<bool>>,
    italic: Option<Option<bool>>,
    size: Option<Option<u32>>,
    color: Option<Option<Rgb>>,
    font: Option<Option<&'a str>>,
    latin: Option<Option<&'a str>>,
    align: Option<Option<Alignment>>,
    fill: Option<Option<Rgb>>,
}

#[allow(
    clippy::option_option,
    reason = "omission, removal and replacement are distinct assignments"
)]
impl<'a> TextProperties<'a> {
    /// Replace the target's text. A run keeps its run properties; a
    /// paragraph, shape or cell keeps its first paragraph's properties and
    /// first run's properties. A line feed starts a line break inside a
    /// paragraph and a new paragraph inside a shape or cell; unlike
    /// [`super::Edit::set_shape_texts`], a shape whose body holds no run
    /// (an empty placeholder) is accepted.
    #[must_use]
    pub const fn with_text(mut self, text: &'a str) -> Self {
        self.text = Some(text);
        self
    }
    /// Set or remove `a:rPr@b` on every addressed run.
    #[must_use]
    pub const fn with_bold(mut self, value: Option<bool>) -> Self {
        self.bold = Some(value);
        self
    }
    /// Set or remove `a:rPr@i` on every addressed run.
    #[must_use]
    pub const fn with_italic(mut self, value: Option<bool>) -> Self {
        self.italic = Some(value);
        self
    }
    /// Set or remove `a:rPr@sz`, in hundredths of a point.
    #[must_use]
    pub const fn with_size(mut self, value: Option<u32>) -> Self {
        self.size = Some(value);
        self
    }
    /// Set or remove the run's solid text fill.
    #[must_use]
    pub const fn with_color(mut self, value: Option<Rgb>) -> Self {
        self.color = Some(value);
        self
    }
    /// Set or remove the latin, east-Asian and complex-script typefaces together.
    #[must_use]
    pub const fn with_font(mut self, value: Option<&'a str>) -> Self {
        self.font = Some(value);
        self
    }
    /// Set or remove only the latin typeface (`a:latin`).
    #[must_use]
    pub const fn with_latin(mut self, value: Option<&'a str>) -> Self {
        self.latin = Some(value);
        self
    }
    /// Set or remove `a:pPr@algn` on every addressed paragraph.
    #[must_use]
    pub const fn with_align(mut self, value: Option<Alignment>) -> Self {
        self.align = Some(value);
        self
    }
    /// Set or remove a table cell's solid fill (`a:tcPr`); cells only.
    #[must_use]
    pub const fn with_fill(mut self, value: Option<Rgb>) -> Self {
        self.fill = Some(value);
        self
    }
    /// The replacement text, if any.
    #[must_use]
    pub const fn text(&self) -> Option<&'a str> {
        self.text
    }
    fn has_run_properties(&self) -> bool {
        self.bold.is_some()
            || self.italic.is_some()
            || self.size.is_some()
            || self.color.is_some()
            || self.font.is_some()
            || self.latin.is_some()
    }
    fn is_empty(&self) -> bool {
        self.text.is_none()
            && !self.has_run_properties()
            && self.align.is_none()
            && self.fill.is_none()
    }
}

/// One formatting assignment for [`super::Edit::set_text_properties`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextEdit<'a> {
    target: TextTarget,
    properties: TextProperties<'a>,
}

impl<'a> TextEdit<'a> {
    /// Assign `properties` to `target`.
    #[must_use]
    pub const fn new(target: TextTarget, properties: TextProperties<'a>) -> Self {
        Self { target, properties }
    }
    /// The addressed element.
    #[must_use]
    pub const fn target(&self) -> TextTarget {
        self.target
    }
    /// The assignments.
    #[must_use]
    pub const fn properties(&self) -> &TextProperties<'a> {
        &self.properties
    }
}

/// Stage a same-slide batch of text formatting edits, applied in order.
///
/// Returns the candidate slide XML and the number of edits whose bytes
/// changed, or `None` when nothing changed. The candidate is read back as a
/// complete shape scene whose shape count and identities must match.
pub(crate) fn stage_text_properties(
    xml: &[u8],
    edits: &[TextEdit<'_>],
    max_text_bytes: usize,
    max_output_bytes: usize,
) -> Result<(Option<Vec<u8>>, usize)> {
    if edits.is_empty() {
        return Ok((None, 0));
    }
    if edits.len() > super::MAX_SHAPE_PROPERTIES_REPLACEMENTS {
        return Err(Error::Limit {
            resource: "opened-presentation text property edits",
            limit: super::MAX_SHAPE_PROPERTIES_REPLACEMENTS,
        });
    }
    let mut text_bytes = 0usize;
    for edit in edits {
        validate(edit)?;
        text_bytes = text_bytes.saturating_add(edit.properties.text.map_or(0, str::len));
    }
    if text_bytes > max_text_bytes {
        return Err(Error::Limit {
            resource: "opened-presentation aggregate text property bytes",
            limit: max_text_bytes,
        });
    }
    let original = crate::shape::Scene::read(xml)?;
    if original.is_rewritten() {
        return Err(Error::UnsafeEdit {
            operation: "set_text_properties",
            reason: "text property edits do not support markup-compatibility branch selection",
        });
    }
    let mut current: Cow<'_, [u8]> = Cow::Borrowed(xml);
    let mut changed = 0;
    for edit in edits {
        let next = {
            let scene = crate::shape::Scene::read(&current)?;
            let index = shape_position(edit.target).get();
            let shape = scene.at(index)?;
            let span = shape.common().span()?.range(current.len())?;
            let is_table = matches!(shape, crate::shape::Shape::Table(_));
            let element = &current[span.clone()];
            let replacement = match edit.target {
                TextTarget::Cell { row, column, .. } => {
                    if !is_table {
                        return Err(Error::UnsafeEdit {
                            operation: "set_text_properties",
                            reason: "a table cell edit requires a table graphic frame",
                        });
                    }
                    rewrite_cell_frame(element, row.get(), column.get(), &edit.properties)?
                },
                _ => {
                    if is_table || edit.properties.fill.is_some() {
                        return Err(Error::UnsafeEdit {
                            operation: "set_text_properties",
                            reason: "text formatting requires a shape text body; fill applies to table cells",
                        });
                    }
                    rewrite_shape(element, edit.target, &edit.properties)?
                },
            };
            if replacement.as_slice() == element {
                None
            } else {
                let length = current.len() - span.len() + replacement.len();
                if length > max_output_bytes {
                    return Err(Error::Limit {
                        resource: "opened-presentation text property output bytes",
                        limit: max_output_bytes,
                    });
                }
                let mut output = Vec::new();
                output
                    .try_reserve_exact(length)
                    .map_err(|source| Error::Allocation {
                        resource: "opened-presentation text property output",
                        source,
                    })?;
                output.extend_from_slice(&current[..span.start]);
                output.extend_from_slice(&replacement);
                output.extend_from_slice(&current[span.end..]);
                Some(output)
            }
        };
        if let Some(next) = next {
            current = Cow::Owned(next);
            changed += 1;
        }
    }
    let Cow::Owned(candidate) = current else {
        return Ok((None, 0));
    };
    let published = crate::shape::Scene::read(&candidate)?;
    if published.is_rewritten() || published.len() != original.len() {
        return Err(Error::Invalid(
            "opened-presentation text property batch changed the semantic shape graph".into(),
        ));
    }
    for index in 0..original.len() {
        if original.at(index)?.id() != published.at(index)?.id() {
            return Err(Error::Invalid(
                "opened-presentation text property batch changed shape identities".into(),
            ));
        }
    }
    Ok((Some(candidate), changed))
}

const fn shape_position(target: TextTarget) -> Position {
    match target {
        TextTarget::Shape { shape }
        | TextTarget::Paragraph { shape, .. }
        | TextTarget::Run { shape, .. }
        | TextTarget::Cell { shape, .. } => shape,
    }
}

fn validate(edit: &TextEdit<'_>) -> Result<()> {
    let properties = &edit.properties;
    if properties.is_empty() {
        return Err(Error::Invalid(
            "a text property edit assigns nothing".into(),
        ));
    }
    if let Some(Some(size)) = properties.size
        && !(MIN_FONT_SIZE..=MAX_FONT_SIZE).contains(&size)
    {
        return Err(Error::Invalid(
            "font size must be between 1 and 4000 points".into(),
        ));
    }
    for face in [properties.font, properties.latin]
        .into_iter()
        .flatten()
        .flatten()
    {
        if face.is_empty() || face.len() > MAX_TYPEFACE_BYTES {
            return Err(Error::Invalid(
                "a typeface name must have 1 to 1024 bytes".into(),
            ));
        }
        check_characters(face)?;
    }
    if let Some(text) = properties.text {
        check_characters(text)?;
        if matches!(edit.target, TextTarget::Run { .. }) && text.contains(['\n', '\r']) {
            return Err(Error::Invalid(
                "run text cannot contain a line break".into(),
            ));
        }
    }
    if properties.fill.is_some() && !matches!(edit.target, TextTarget::Cell { .. }) {
        return Err(Error::Invalid("fill applies to table cells only".into()));
    }
    Ok(())
}

fn check_characters(value: &str) -> Result<()> {
    if value.chars().all(|ch| {
        matches!(ch, '\t' | '\n' | '\r' | '\u{20}'..='\u{d7ff}' | '\u{e000}'..='\u{fffd}' | '\u{10000}'..='\u{10ffff}')
    }) {
        Ok(())
    } else {
        Err(Error::Invalid(
            "text contains a character that XML cannot represent".into(),
        ))
    }
}

// ---------------------------------------------------------------------------
// One-level element views over borrowed bytes
// ---------------------------------------------------------------------------

/// One element inside a parent's byte slice.
#[derive(Debug, Clone)]
struct Node {
    /// Whole element, start tag through end tag.
    span: Range<usize>,
    /// End of the start tag.
    open_end: usize,
    /// Start of the end tag (equal to `span.end` for an empty element).
    close_start: usize,
    empty: bool,
    prefix_len: usize,
    name_len: usize,
}

impl Node {
    fn qname<'x>(&self, xml: &'x [u8]) -> &'x [u8] {
        &xml[self.span.start + 1..self.span.start + 1 + self.name_len]
    }
    fn local<'x>(&self, xml: &'x [u8]) -> &'x [u8] {
        let qname = self.qname(xml);
        if self.prefix_len == 0 {
            qname
        } else {
            &qname[self.prefix_len + 1..]
        }
    }
    fn prefix<'x>(&self, xml: &'x [u8]) -> &'x [u8] {
        &self.qname(xml)[..self.prefix_len]
    }
    fn content(&self) -> Range<usize> {
        self.open_end..self.close_start
    }
}

fn invalid(message: &'static str) -> Error {
    Error::Invalid(message.into())
}

fn qname_parts(name: &[u8]) -> (usize, usize) {
    let prefix_len = name.iter().position(|byte| *byte == b':').unwrap_or(0);
    (prefix_len, name.len())
}

/// Parse the root element of `xml` (which must start at offset 0 after
/// optional whitespace) as a [`Node`].
fn root(xml: &[u8]) -> Result<Node> {
    let mut nodes = elements(xml, 0..xml.len())?;
    if nodes.len() != 1 {
        return Err(invalid("text property edit expected one root element"));
    }
    Ok(nodes.remove(0))
}

/// The direct child elements inside `range` of `xml`.
fn elements(xml: &[u8], range: Range<usize>) -> Result<Vec<Node>> {
    let base = range.start;
    let slice = &xml[range];
    let mut reader = Reader::from_reader(slice);
    reader.config_mut().check_end_names = true;
    let mut nodes = Vec::new();
    let mut depth = 0usize;
    let mut open: Option<(usize, usize, usize, usize)> = None;
    loop {
        let before = usize::try_from(reader.buffer_position()).unwrap_or(usize::MAX);
        let event = reader
            .read_event()
            .map_err(|_| invalid("text property edit found malformed XML"))?;
        let after = usize::try_from(reader.buffer_position()).unwrap_or(usize::MAX);
        match event {
            Event::Start(start) => {
                if depth == 0 {
                    let (prefix_len, name_len) = qname_parts(start.name().as_ref());
                    open = Some((before, after, prefix_len, name_len));
                }
                depth += 1;
            },
            Event::Empty(start) => {
                if depth == 0 {
                    let (prefix_len, name_len) = qname_parts(start.name().as_ref());
                    nodes.push(Node {
                        span: base + before..base + after,
                        open_end: base + after,
                        close_start: base + after,
                        empty: true,
                        prefix_len,
                        name_len,
                    });
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("text property edit found an unbalanced end tag"))?;
                if depth == 0 {
                    let (start, open_end, prefix_len, name_len) = open
                        .take()
                        .ok_or_else(|| invalid("text property edit lost an element start"))?;
                    nodes.push(Node {
                        span: base + start..base + after,
                        open_end: base + open_end,
                        close_start: base + before,
                        empty: false,
                        prefix_len,
                        name_len,
                    });
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 {
        return Err(invalid("text property edit found an unclosed element"));
    }
    Ok(nodes)
}

fn children(xml: &[u8], node: &Node) -> Result<Vec<Node>> {
    if node.empty {
        Ok(Vec::new())
    } else {
        elements(xml, node.content())
    }
}

fn child<'n>(xml: &[u8], nodes: &'n [Node], local: &[u8]) -> Option<&'n Node> {
    nodes.iter().find(|node| node.local(xml) == local)
}

/// Splices relative to one element slice, applied back to front.
struct Splices(Vec<(Range<usize>, Vec<u8>)>);

impl Splices {
    fn new() -> Self {
        Self(Vec::new())
    }
    fn replace(&mut self, range: Range<usize>, bytes: Vec<u8>) {
        self.0.push((range, bytes));
    }
    fn insert(&mut self, at: usize, bytes: Vec<u8>) {
        self.0.push((at..at, bytes));
    }
    fn apply(mut self, xml: &[u8]) -> Result<Vec<u8>> {
        self.0.sort_by_key(|(range, _)| (range.start, range.end));
        for pair in self.0.windows(2) {
            if pair[0].0.end > pair[1].0.start {
                return Err(invalid("text property edit produced overlapping splices"));
            }
        }
        let mut output = Vec::with_capacity(xml.len() + 64);
        let mut cursor = 0;
        for (range, bytes) in &self.0 {
            output.extend_from_slice(&xml[cursor..range.start]);
            output.extend_from_slice(bytes);
            cursor = range.end;
        }
        output.extend_from_slice(&xml[cursor..]);
        Ok(output)
    }
}

// ---------------------------------------------------------------------------
// Start-tag attribute editing
// ---------------------------------------------------------------------------

/// Rewrite a start tag, setting (`Some`) or removing (`None`) attributes.
/// Untouched attributes keep their exact source spelling.
fn rewrite_start_tag(
    tag: &[u8],
    changes: &[(&[u8], Option<&[u8]>)],
    empty: bool,
) -> Result<Vec<u8>> {
    let body_end = if tag.ends_with(b"/>") {
        tag.len() - 2
    } else if tag.ends_with(b">") {
        tag.len() - 1
    } else {
        return Err(invalid("text property edit found a malformed start tag"));
    };
    let inner = &tag[1..body_end];
    let name_end = inner
        .iter()
        .position(|byte| byte.is_ascii_whitespace())
        .unwrap_or(inner.len());
    let mut output = Vec::with_capacity(tag.len() + 32);
    output.push(b'<');
    output.extend_from_slice(&inner[..name_end]);
    // Walk raw attributes: name, '=', quoted value.
    let mut cursor = name_end;
    let bytes = inner;
    while cursor < bytes.len() {
        while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if cursor >= bytes.len() {
            break;
        }
        let name_start = cursor;
        while cursor < bytes.len() && bytes[cursor] != b'=' && !bytes[cursor].is_ascii_whitespace()
        {
            cursor += 1;
        }
        let name = &bytes[name_start..cursor];
        while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        if bytes.get(cursor) != Some(&b'=') {
            return Err(invalid("text property edit found a malformed attribute"));
        }
        cursor += 1;
        while cursor < bytes.len() && bytes[cursor].is_ascii_whitespace() {
            cursor += 1;
        }
        let quote = *bytes
            .get(cursor)
            .ok_or_else(|| invalid("text property edit found a malformed attribute"))?;
        if quote != b'"' && quote != b'\'' {
            return Err(invalid("text property edit found an unquoted attribute"));
        }
        let value_start = cursor;
        cursor += 1;
        while cursor < bytes.len() && bytes[cursor] != quote {
            cursor += 1;
        }
        if cursor >= bytes.len() {
            return Err(invalid(
                "text property edit found an unterminated attribute",
            ));
        }
        cursor += 1;
        if changes.iter().any(|(key, _)| *key == name) {
            continue;
        }
        output.push(b' ');
        output.extend_from_slice(name);
        output.push(b'=');
        output.extend_from_slice(&bytes[value_start..cursor]);
    }
    for (key, value) in changes {
        if let Some(value) = value {
            output.push(b' ');
            output.extend_from_slice(key);
            output.extend_from_slice(b"=\"");
            output.extend_from_slice(value);
            output.push(b'"');
        }
    }
    output.extend_from_slice(if empty { b"/>" } else { b">" });
    Ok(output)
}

fn escape_attribute(value: &str) -> Vec<u8> {
    quick_xml::escape::escape(value).as_bytes().to_vec()
}

fn escape_text(value: &str, maximum: usize) -> Result<Vec<u8>> {
    Ok(super::xml::escape_character_data(value, maximum)?
        .as_bytes()
        .to_vec())
}

fn qualified(prefix: &[u8], local: &str) -> Vec<u8> {
    let mut name = Vec::with_capacity(prefix.len() + local.len() + 1);
    if !prefix.is_empty() {
        name.extend_from_slice(prefix);
        name.push(b':');
    }
    name.extend_from_slice(local.as_bytes());
    name
}

fn srgb_fill(prefix: &[u8], color: Rgb) -> Vec<u8> {
    let fill = qualified(prefix, "solidFill");
    let clr = qualified(prefix, "srgbClr");
    let mut output = Vec::with_capacity(64);
    output.push(b'<');
    output.extend_from_slice(&fill);
    output.extend_from_slice(b"><");
    output.extend_from_slice(&clr);
    output.extend_from_slice(
        format!(
            " val=\"{:02X}{:02X}{:02X}\"/></",
            color.0[0], color.0[1], color.0[2]
        )
        .as_bytes(),
    );
    output.extend_from_slice(&fill);
    output.push(b'>');
    output
}

fn typeface(prefix: &[u8], local: &str, face: &str) -> Vec<u8> {
    let mut output = Vec::with_capacity(32 + face.len());
    output.push(b'<');
    output.extend_from_slice(&qualified(prefix, local));
    output.extend_from_slice(b" typeface=\"");
    output.extend_from_slice(&escape_attribute(face));
    output.extend_from_slice(b"\"/>");
    output
}

// ---------------------------------------------------------------------------
// Properties elements with schema-ordered children
// ---------------------------------------------------------------------------

fn rpr_rank(local: &[u8]) -> u8 {
    match local {
        b"ln" => 0,
        b"noFill" | b"solidFill" | b"gradFill" | b"blipFill" | b"pattFill" | b"grpFill" => 1,
        b"effectLst" | b"effectDag" => 2,
        b"highlight" => 3,
        b"uLnTx" | b"uLn" => 4,
        b"uFillTx" | b"uFill" => 5,
        b"latin" => 6,
        b"ea" => 7,
        b"cs" => 8,
        b"sym" => 9,
        b"hlinkClick" => 10,
        b"hlinkMouseOver" => 11,
        b"rtl" => 12,
        _ => 13,
    }
}

fn tcpr_rank(local: &[u8]) -> u8 {
    match local {
        b"lnL" => 0,
        b"lnR" => 1,
        b"lnT" => 2,
        b"lnB" => 3,
        b"lnTlToBr" => 4,
        b"lnBlToTr" => 5,
        b"cell3D" => 6,
        b"noFill" | b"solidFill" | b"gradFill" | b"blipFill" | b"pattFill" | b"grpFill" => 7,
        b"headers" => 8,
        _ => 9,
    }
}

/// A requested change to the children of a properties element: replace
/// every child of `rank` with `bytes` (or remove them for `None`).
struct ChildChange {
    rank: u8,
    bytes: Option<Vec<u8>>,
}

/// Rewrite a properties element (`a:rPr`, `a:pPr`, `a:tcPr`, …) whose bytes
/// are `element`, applying attribute and ranked child changes.
fn rewrite_properties(
    element: &[u8],
    attributes: &[(&[u8], Option<&[u8]>)],
    child_changes: &[ChildChange],
    rank: fn(&[u8]) -> u8,
) -> Result<Vec<u8>> {
    let node = root(element)?;
    let kids = children(element, &node)?;
    let has_new_children = child_changes.iter().any(|change| change.bytes.is_some());
    let remaining = kids
        .iter()
        .filter(|kid| {
            let kid_rank = rank(kid.local(element));
            !child_changes.iter().any(|change| change.rank == kid_rank)
        })
        .count();
    let empty = remaining == 0 && !has_new_children;
    let start_tag = rewrite_start_tag(&element[node.span.start..node.open_end], attributes, empty)?;
    if empty {
        return Ok(start_tag);
    }
    let mut output = Vec::with_capacity(element.len() + 64);
    output.extend_from_slice(&start_tag);
    let mut pending: Vec<&ChildChange> = child_changes.iter().collect();
    pending.sort_by_key(|change| change.rank);
    let mut cursor = node.open_end;
    let mut pending_index = 0;
    for kid in &kids {
        let kid_rank = rank(kid.local(element));
        // Insert new children ranked before this kid.
        while pending_index < pending.len() && pending[pending_index].rank < kid_rank {
            output.extend_from_slice(&element[cursor..kid.span.start]);
            cursor = kid.span.start;
            if let Some(bytes) = &pending[pending_index].bytes {
                output.extend_from_slice(bytes);
            }
            pending_index += 1;
        }
        output.extend_from_slice(&element[cursor..kid.span.start]);
        cursor = kid.span.end;
        if let Some(change) = child_changes.iter().find(|change| change.rank == kid_rank) {
            // The first matching kid takes the replacement; later ones of the
            // same rank are dropped.
            if pending_index < pending.len() && pending[pending_index].rank == kid_rank {
                if let Some(bytes) = &change.bytes {
                    output.extend_from_slice(bytes);
                }
                pending_index += 1;
            }
        } else {
            output.extend_from_slice(&element[kid.span.clone()]);
        }
    }
    let close_start = if node.empty {
        element.len()
    } else {
        node.close_start
    };
    output.extend_from_slice(&element[cursor.min(close_start)..close_start]);
    while pending_index < pending.len() {
        if let Some(bytes) = &pending[pending_index].bytes {
            output.extend_from_slice(bytes);
        }
        pending_index += 1;
    }
    output.extend_from_slice(b"</");
    output.extend_from_slice(node.qname(element));
    output.push(b'>');
    Ok(output)
}

fn bool_value(value: bool) -> &'static [u8] {
    if value { b"1" } else { b"0" }
}

/// Rewrite an `a:rPr` (or `a:endParaRPr`) element's bytes.
fn rewrite_rpr(element: &[u8], prefix: &[u8], properties: &TextProperties<'_>) -> Result<Vec<u8>> {
    let size;
    let mut attributes: Vec<(&[u8], Option<&[u8]>)> = Vec::new();
    if let Some(bold) = properties.bold {
        attributes.push((b"b", bold.map(bool_value)));
    }
    if let Some(italic) = properties.italic {
        attributes.push((b"i", italic.map(bool_value)));
    }
    if let Some(value) = properties.size {
        size = value.map(|value| value.to_string());
        attributes.push((b"sz", size.as_deref().map(str::as_bytes)));
    }
    let mut changes = Vec::new();
    if let Some(color) = properties.color {
        changes.push(ChildChange {
            rank: 1,
            bytes: color.map(|color| srgb_fill(prefix, color)),
        });
    }
    let latin = properties.latin.or(properties.font);
    if let Some(face) = latin {
        changes.push(ChildChange {
            rank: 6,
            bytes: face.map(|face| typeface(prefix, "latin", face)),
        });
    }
    if let Some(face) = properties.font {
        changes.push(ChildChange {
            rank: 7,
            bytes: face.map(|face| typeface(prefix, "ea", face)),
        });
        changes.push(ChildChange {
            rank: 8,
            bytes: face.map(|face| typeface(prefix, "cs", face)),
        });
    }
    rewrite_properties(element, &attributes, &changes, rpr_rank)
}

fn alignment_token(value: Alignment) -> &'static [u8] {
    match value {
        Alignment::Left => b"l",
        Alignment::Center => b"ctr",
        Alignment::Right => b"r",
        Alignment::Justified => b"just",
        Alignment::JustifiedLow => b"justLow",
        Alignment::Distributed => b"dist",
        Alignment::ThaiDistributed => b"thaiDist",
    }
}

// ---------------------------------------------------------------------------
// Runs, paragraphs, bodies
// ---------------------------------------------------------------------------

fn is_run(local: &[u8]) -> bool {
    matches!(local, b"r" | b"br" | b"fld")
}

/// Rewrite one run element (`a:r`, `a:br`, `a:fld`) for run properties and
/// optional replacement text.
fn rewrite_run(run: &[u8], properties: &TextProperties<'_>, text: Option<&str>) -> Result<Vec<u8>> {
    let node = root(run)?;
    let prefix = node.prefix(run).to_vec();
    if text.is_some() && node.local(run) != b"r" {
        return Err(Error::UnsafeEdit {
            operation: "set_text_properties",
            reason: "run text applies to a text run (a:r)",
        });
    }
    let kids = children(run, &node)?;
    let mut splices = Splices::new();
    let mut body = run.to_vec();
    let mut node = node;
    if node.empty {
        // `<a:br/>` → `<a:br></a:br>` so a properties child can be inserted.
        if !properties.has_run_properties() {
            return Ok(run.to_vec());
        }
        let qname = node.qname(run).to_vec();
        let mut expanded = run[..run.len() - 2].to_vec();
        expanded.push(b'>');
        let open_end = expanded.len();
        expanded.extend_from_slice(b"</");
        expanded.extend_from_slice(&qname);
        expanded.push(b'>');
        body = expanded;
        node = Node {
            span: 0..body.len(),
            open_end,
            close_start: open_end,
            empty: false,
            prefix_len: node.prefix_len,
            name_len: node.name_len,
        };
    }
    if properties.has_run_properties() {
        match child(&body, &kids, b"rPr") {
            Some(rpr) if !kids.is_empty() => {
                let replaced = rewrite_rpr(&body[rpr.span.clone()], &prefix, properties)?;
                splices.replace(rpr.span.clone(), replaced);
            },
            _ => {
                let mut empty = b"<".to_vec();
                empty.extend_from_slice(&qualified(&prefix, "rPr"));
                empty.extend_from_slice(b"/>");
                let created = rewrite_rpr(&empty, &prefix, properties)?;
                splices.insert(node.open_end, created);
            },
        }
    }
    if let Some(text) = text {
        let escaped = escape_text(text, text.len().saturating_mul(6).max(16))?;
        let t =
            child(&body, &kids, b"t").ok_or_else(|| invalid("a text run has no a:t element"))?;
        let qname = t.qname(&body).to_vec();
        let mut element = b"<".to_vec();
        element
            .extend_from_slice(&body[t.span.start + 1..t.open_end - if t.empty { 2 } else { 1 }]);
        element.push(b'>');
        element.extend_from_slice(&escaped);
        element.extend_from_slice(b"</");
        element.extend_from_slice(&qname);
        element.push(b'>');
        splices.replace(t.span.clone(), element);
    }
    splices.apply(&body)
}

/// Build a fresh text run carrying `rpr` (raw, possibly empty).
fn new_run(prefix: &[u8], rpr: &[u8], text: &str) -> Result<Vec<u8>> {
    let r = qualified(prefix, "r");
    let t = qualified(prefix, "t");
    let escaped = escape_text(text, text.len().saturating_mul(6).max(16))?;
    let mut output = Vec::with_capacity(32 + rpr.len() + escaped.len());
    output.push(b'<');
    output.extend_from_slice(&r);
    output.push(b'>');
    output.extend_from_slice(rpr);
    output.push(b'<');
    output.extend_from_slice(&t);
    output.push(b'>');
    output.extend_from_slice(&escaped);
    output.extend_from_slice(b"</");
    output.extend_from_slice(&t);
    output.extend_from_slice(b"></");
    output.extend_from_slice(&r);
    output.push(b'>');
    Ok(output)
}

fn new_break(prefix: &[u8], rpr: &[u8]) -> Vec<u8> {
    let br = qualified(prefix, "br");
    let mut output = Vec::with_capacity(16 + rpr.len());
    output.push(b'<');
    output.extend_from_slice(&br);
    if rpr.is_empty() {
        output.extend_from_slice(b"/>");
    } else {
        output.push(b'>');
        output.extend_from_slice(rpr);
        output.extend_from_slice(b"</");
        output.extend_from_slice(&br);
        output.push(b'>');
    }
    output
}

/// The raw `a:rPr` of the paragraph's first text run, or an empty slice.
fn first_run_properties<'x>(paragraph: &'x [u8], kids: &[Node]) -> Result<&'x [u8]> {
    for kid in kids {
        if kid.local(paragraph) == b"r" {
            let run_kids = children(paragraph, kid)?;
            if let Some(rpr) = child(paragraph, &run_kids, b"rPr") {
                return Ok(&paragraph[rpr.span.clone()]);
            }
            return Ok(&[]);
        }
    }
    Ok(&[])
}

/// Rewrite one paragraph: alignment, run properties on the selected runs,
/// and optional replacement text for the whole paragraph or one run.
fn rewrite_paragraph(
    paragraph: &[u8],
    run: Option<usize>,
    properties: &TextProperties<'_>,
) -> Result<Vec<u8>> {
    let node = root(paragraph)?;
    let prefix = node.prefix(paragraph).to_vec();
    let kids = children(paragraph, &node)?;
    let mut splices = Splices::new();
    if let Some(run) = run {
        let target = kids
            .iter()
            .filter(|kid| is_run(kid.local(paragraph)))
            .nth(run)
            .ok_or(Error::IndexOutOfBounds {
                index: run,
                len: kids
                    .iter()
                    .filter(|kid| is_run(kid.local(paragraph)))
                    .count(),
            })?;
        let replaced = rewrite_run(&paragraph[target.span.clone()], properties, properties.text)?;
        splices.replace(target.span.clone(), replaced);
        return splices.apply(paragraph);
    }
    if let Some(align) = properties.align {
        let token = align.map(alignment_token);
        let changes = [(b"algn".as_slice(), token)];
        match child(paragraph, &kids, b"pPr") {
            Some(ppr) => {
                let element = &paragraph[ppr.span.clone()];
                let rewritten = rewrite_properties(element, &changes, &[], |_| 0)?;
                splices.replace(ppr.span.clone(), rewritten);
            },
            None if token.is_some() => {
                let mut empty = b"<".to_vec();
                empty.extend_from_slice(&qualified(&prefix, "pPr"));
                empty.extend_from_slice(b"/>");
                splices.insert(node.open_end, rewrite_start_tag(&empty, &changes, true)?);
            },
            None => {},
        }
    }
    if let Some(text) = properties.text {
        let template = first_run_properties(paragraph, &kids)?;
        let template = if properties.has_run_properties() {
            let base = if template.is_empty() {
                let mut empty = b"<".to_vec();
                empty.extend_from_slice(&qualified(&prefix, "rPr"));
                empty.extend_from_slice(b"/>");
                empty
            } else {
                template.to_vec()
            };
            rewrite_rpr(&base, &prefix, properties)?
        } else {
            template.to_vec()
        };
        let mut runs = Vec::new();
        for (index, line) in text.split('\n').enumerate() {
            if index > 0 {
                runs.extend_from_slice(&new_break(&prefix, &template));
            }
            let line = line.strip_suffix('\r').unwrap_or(line);
            if !line.is_empty() {
                runs.extend_from_slice(&new_run(&prefix, &template, line)?);
            }
        }
        let run_nodes: Vec<&Node> = kids
            .iter()
            .filter(|kid| is_run(kid.local(paragraph)))
            .collect();
        match (run_nodes.first(), run_nodes.last()) {
            (Some(first), Some(last)) => {
                splices.replace(first.span.start..last.span.end, runs);
            },
            _ => {
                // After pPr if present, else at the start; before endParaRPr.
                let at = child(paragraph, &kids, b"pPr").map_or_else(
                    || {
                        if node.empty {
                            node.open_end
                        } else {
                            node.open_end
                        }
                    },
                    |ppr| ppr.span.end,
                );
                if node.empty {
                    return expand_and_insert(paragraph, &node, &splices, runs);
                }
                splices.insert(at, runs);
            },
        }
    } else if properties.has_run_properties() {
        for kid in kids.iter().filter(|kid| is_run(kid.local(paragraph))) {
            let replaced = rewrite_run(&paragraph[kid.span.clone()], properties, None)?;
            splices.replace(kid.span.clone(), replaced);
        }
    }
    splices.apply(paragraph)
}

/// An empty `<a:p/>` gains content: expand it, then apply pending splices.
fn expand_and_insert(
    paragraph: &[u8],
    node: &Node,
    splices: &Splices,
    runs: Vec<u8>,
) -> Result<Vec<u8>> {
    // Pending splices of an empty paragraph can only touch its start tag
    // (a pPr insertion at open_end); rebuild them in order.
    let mut output = paragraph[..paragraph.len() - 2].to_vec();
    output.push(b'>');
    for (_, bytes) in &splices.0 {
        output.extend_from_slice(bytes);
    }
    output.extend_from_slice(&runs);
    output.extend_from_slice(b"</");
    output.extend_from_slice(node.qname(paragraph));
    output.push(b'>');
    Ok(output)
}

/// Apply `properties` to selected paragraphs of a text body element.
fn rewrite_body(
    body: &[u8],
    paragraph: Option<usize>,
    run: Option<usize>,
    properties: &TextProperties<'_>,
) -> Result<Vec<u8>> {
    let node = root(body)?;
    let kids = children(body, &node)?;
    let paragraphs: Vec<&Node> = kids.iter().filter(|kid| kid.local(body) == b"p").collect();
    let mut splices = Splices::new();
    match paragraph {
        Some(index) => {
            let target = paragraphs.get(index).ok_or(Error::IndexOutOfBounds {
                index,
                len: paragraphs.len(),
            })?;
            let replaced = rewrite_paragraph(&body[target.span.clone()], run, properties)?;
            splices.replace(target.span.clone(), replaced);
        },
        None => {
            for target in paragraphs {
                let replaced = rewrite_paragraph(&body[target.span.clone()], None, properties)?;
                splices.replace(target.span.clone(), replaced);
            }
        },
    }
    splices.apply(body)
}

/// Rewrite a `p:sp` element for a shape, paragraph or run target.
fn rewrite_shape(
    shape: &[u8],
    target: TextTarget,
    properties: &TextProperties<'_>,
) -> Result<Vec<u8>> {
    let node = root(shape)?;
    let kids = children(shape, &node)?;
    let body = child(shape, &kids, b"txBody").ok_or(Error::UnsafeEdit {
        operation: "set_text_properties",
        reason: "the shape has no text body",
    })?;
    let (paragraph, run) = match target {
        TextTarget::Shape { .. } => (None, None),
        TextTarget::Paragraph { paragraph, .. } => (Some(paragraph.get()), None),
        TextTarget::Run { paragraph, run, .. } => (Some(paragraph.get()), Some(run.get())),
        TextTarget::Cell { .. } => unreachable!("cells are rewritten through their frame"),
    };
    let element = &shape[body.span.clone()];
    let replaced = match (target, properties.text) {
        (TextTarget::Shape { .. }, Some(text)) => {
            let body_node = root(element)?;
            let body_kids = children(element, &body_node)?;
            let prefix = child(element, &body_kids, b"bodyPr")
                .or_else(|| child(element, &body_kids, b"p"))
                .map(|node| node.prefix(element).to_vec())
                .ok_or_else(|| invalid("the shape text body has no drawing children"))?;
            rewrite_body_text(element, &prefix, text, properties)?
        },
        _ => rewrite_body(element, paragraph, run, properties)?,
    };
    let mut splices = Splices::new();
    splices.replace(body.span.clone(), replaced);
    splices.apply(shape)
}

// ---------------------------------------------------------------------------
// Table cells
// ---------------------------------------------------------------------------

fn descend(xml: &[u8], node: &Node, local: &[u8]) -> Result<Node> {
    let kids = children(xml, node)?;
    child(xml, &kids, local)
        .cloned()
        .ok_or_else(|| invalid("the table frame lacks its graphic table"))
}

/// Rewrite a `p:graphicFrame` holding an `a:tbl` for one cell.
fn rewrite_cell_frame(
    frame: &[u8],
    row: usize,
    column: usize,
    properties: &TextProperties<'_>,
) -> Result<Vec<u8>> {
    let node = root(frame)?;
    let graphic = descend(frame, &node, b"graphic")?;
    let data = descend(frame, &graphic, b"graphicData")?;
    let table = descend(frame, &data, b"tbl")?;
    let rows: Vec<Node> = children(frame, &table)?
        .into_iter()
        .filter(|kid| kid.local(frame) == b"tr")
        .collect();
    let tr = rows.get(row).ok_or(Error::IndexOutOfBounds {
        index: row,
        len: rows.len(),
    })?;
    let cells: Vec<Node> = children(frame, tr)?
        .into_iter()
        .filter(|kid| kid.local(frame) == b"tc")
        .collect();
    let tc = cells.get(column).ok_or(Error::IndexOutOfBounds {
        index: column,
        len: cells.len(),
    })?;
    let replaced = rewrite_cell(&frame[tc.span.clone()], properties)?;
    let mut splices = Splices::new();
    splices.replace(tc.span.clone(), replaced);
    splices.apply(frame)
}

fn rewrite_cell(cell: &[u8], properties: &TextProperties<'_>) -> Result<Vec<u8>> {
    let node = root(cell)?;
    if node.empty {
        return Err(invalid("an empty a:tc cannot hold text properties"));
    }
    let prefix = node.prefix(cell).to_vec();
    let kids = children(cell, &node)?;
    let mut splices = Splices::new();
    let body = child(cell, &kids, b"txBody");
    let formatting = properties.has_run_properties() || properties.align.is_some();
    if properties.text.is_some() || formatting {
        let body = body.ok_or(Error::UnsafeEdit {
            operation: "set_text_properties",
            reason: "the table cell has no text body",
        })?;
        let element = &cell[body.span.clone()];
        let rewritten = match properties.text {
            Some(text) => rewrite_body_text(element, &prefix, text, properties)?,
            None => rewrite_body(element, None, None, properties)?,
        };
        splices.replace(body.span.clone(), rewritten);
    }
    if let Some(fill) = properties.fill {
        let changes = [ChildChange {
            rank: 7,
            bytes: fill.map(|color| srgb_fill(&prefix, color)),
        }];
        match child(cell, &kids, b"tcPr") {
            Some(tcpr) => {
                let rewritten =
                    rewrite_properties(&cell[tcpr.span.clone()], &[], &changes, tcpr_rank)?;
                splices.replace(tcpr.span.clone(), rewritten);
            },
            None if fill.is_some() => {
                let mut empty = b"<".to_vec();
                empty.extend_from_slice(&qualified(&prefix, "tcPr"));
                empty.extend_from_slice(b"/>");
                let created = rewrite_properties(&empty, &[], &changes, tcpr_rank)?;
                // tcPr follows txBody and precedes extLst.
                let at =
                    child(cell, &kids, b"extLst").map_or(node.close_start, |ext| ext.span.start);
                splices.insert(at, created);
            },
            None => {},
        }
    }
    splices.apply(cell)
}

/// Replace a text body's paragraphs with one paragraph per line, keeping the
/// first paragraph's `a:pPr`, its first run's `a:rPr` and its
/// `a:endParaRPr`.
fn rewrite_body_text(
    body: &[u8],
    prefix: &[u8],
    text: &str,
    properties: &TextProperties<'_>,
) -> Result<Vec<u8>> {
    let node = root(body)?;
    let kids = children(body, &node)?;
    let paragraphs: Vec<&Node> = kids.iter().filter(|kid| kid.local(body) == b"p").collect();
    let (ppr, rpr, end) = match paragraphs.first() {
        Some(first) => {
            let paragraph = &body[first.span.clone()];
            let local = root(paragraph)?;
            let first_kids = children(paragraph, &local)?;
            let ppr = child(paragraph, &first_kids, b"pPr")
                .map_or(&[][..], |ppr| &paragraph[ppr.span.clone()])
                .to_vec();
            let rpr = first_run_properties(paragraph, &first_kids)?.to_vec();
            let end = child(paragraph, &first_kids, b"endParaRPr")
                .map_or(&[][..], |end| &paragraph[end.span.clone()])
                .to_vec();
            (ppr, rpr, end)
        },
        None => (Vec::new(), Vec::new(), Vec::new()),
    };
    let p = qualified(prefix, "p");
    let mut generated = Vec::new();
    for line in text.split('\n') {
        let line = line.strip_suffix('\r').unwrap_or(line);
        let mut paragraph = b"<".to_vec();
        paragraph.extend_from_slice(&p);
        paragraph.push(b'>');
        paragraph.extend_from_slice(&ppr);
        if !line.is_empty() {
            paragraph.extend_from_slice(&new_run(prefix, &rpr, line)?);
        }
        paragraph.extend_from_slice(&end);
        paragraph.extend_from_slice(b"</");
        paragraph.extend_from_slice(&p);
        paragraph.push(b'>');
        let paragraph = if properties.has_run_properties() || properties.align.is_some() {
            let formatting = TextProperties {
                text: None,
                fill: None,
                ..*properties
            };
            rewrite_paragraph(&paragraph, None, &formatting)?
        } else {
            paragraph
        };
        generated.extend_from_slice(&paragraph);
    }
    let mut splices = Splices::new();
    match (paragraphs.first(), paragraphs.last()) {
        (Some(first), Some(last)) => splices.replace(first.span.start..last.span.end, generated),
        _ => {
            if node.empty {
                return Err(invalid("an empty cell text body cannot take paragraphs"));
            }
            splices.insert(node.close_start, generated);
        },
    }
    splices.apply(body)
}
