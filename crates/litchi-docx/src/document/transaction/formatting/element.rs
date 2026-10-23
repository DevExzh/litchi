//! Edits of story elements reached through an [`ElementPath`] below `w:body`.
//!
//! The body-position operations of the parent module address direct-body
//! paragraphs and their direct runs. These operations reach every
//! WordprocessingML element of the main story that the read model addresses:
//! paragraphs, runs and hyperlinks inside table cells, block content
//! controls and hyperlinks, and the tables, rows and cells themselves.
//!
//! Each records one of three exact-markup operations:
//!
//! - [`Operation::ReplaceElement`] replaces one element's exact bytes;
//! - [`Operation::InsertElement`] inserts exact bytes as a direct child of a
//!   container;
//! - [`Operation::RemoveElement`] removes one exact direct child.
//!
//! Each is checked against the projected document on replay, so a patch
//! applies only to the bytes it was produced from, and its inverse is exact.
//! A path that addresses a direct-body paragraph is routed to the
//! body-position operation of the same edit, which keeps those edits on their
//! incremental layout path.

use std::sync::Arc;

use litchi_core::Position;
use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::events::Event;
use quick_xml::reader::NsReader;

use super::super::{
    Conformance, Edit, Error, FragmentPrefix, Operation, Refusal, Snapshot, checked_end,
    checked_slice, checked_start, is_fragment_word_element, replace_range, rewrite_text_owner,
    scan_paragraph_text_owner, scan_text_owner, select_direct_child, select_single_control_content,
    validate_authored_text,
};
use super::{
    BodyAnchor, BodyBlock, NewParagraph, NewRun, Removal, Result, buffer, element_len, ensure_utf8,
    inside_complex_field, invalid, opened, paragraph_dependencies, paragraph_markup,
    run_dependencies, run_markup, spliced, staging_add,
};
use crate::properties;

mod grid;
pub use grid::NewCell;

/// One step of an [`ElementPath`]: the position of a direct
/// WordprocessingML child, counted among the container's direct children with
/// the same element name.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ElementStep {
    /// A `w:p`.
    Paragraph(Position),
    /// A `w:r`.
    Run(Position),
    /// A `w:hyperlink`.
    Hyperlink(Position),
    /// A `w:tbl`.
    Table(Position),
    /// A `w:tr` of a table.
    Row(Position),
    /// A `w:tc` of a row.
    Cell(Position),
    /// A block or inline `w:sdt`. A later step selects a child of its single
    /// `w:sdtContent`.
    ContentControl(Position),
    /// The body-final `w:sectPr`: only position 0, and only as the first
    /// step (record 0319).
    SectionProperties(Position),
}

impl ElementStep {
    /// The element's local name.
    const fn local(self) -> &'static [u8] {
        match self {
            Self::Paragraph(_) => b"p",
            Self::Run(_) => b"r",
            Self::Hyperlink(_) => b"hyperlink",
            Self::Table(_) => b"tbl",
            Self::Row(_) => b"tr",
            Self::Cell(_) => b"tc",
            Self::ContentControl(_) => b"sdt",
            Self::SectionProperties(_) => b"sectPr",
        }
    }

    /// The step's position among its same-named siblings.
    #[must_use]
    pub const fn position(self) -> Position {
        match self {
            Self::Paragraph(position)
            | Self::Run(position)
            | Self::Hyperlink(position)
            | Self::Table(position)
            | Self::Row(position)
            | Self::Cell(position)
            | Self::ContentControl(position)
            | Self::SectionProperties(position) => position,
        }
    }

    const fn with_position(self, position: Position) -> Self {
        match self {
            Self::Paragraph(_) => Self::Paragraph(position),
            Self::Run(_) => Self::Run(position),
            Self::Hyperlink(_) => Self::Hyperlink(position),
            Self::Table(_) => Self::Table(position),
            Self::Row(_) => Self::Row(position),
            Self::Cell(_) => Self::Cell(position),
            Self::ContentControl(_) => Self::ContentControl(position),
            Self::SectionProperties(_) => Self::SectionProperties(position),
        }
    }

    /// Whether `self` may be a direct child of a container whose content
    /// root has local name `container`.
    fn fits(self, container: &[u8]) -> bool {
        matches!(
            (container, self),
            (
                b"body" | b"tc" | b"sdtContent",
                Self::Paragraph(_) | Self::Table(_) | Self::ContentControl(_)
            ) | (b"tbl", Self::Row(_))
                | (b"tr", Self::Cell(_))
                | (
                    b"p",
                    Self::Run(_) | Self::Hyperlink(_) | Self::ContentControl(_)
                )
                | (b"sdtContent", Self::Run(_) | Self::Hyperlink(_))
                | (b"hyperlink", Self::Run(_))
        )
    }

    /// The token of the step in a durable target string.
    fn token(self) -> (&'static str, usize) {
        let name = match self {
            Self::Paragraph(_) => "p",
            Self::Run(_) => "r",
            Self::Hyperlink(_) => "hyperlink",
            Self::Table(_) => "tbl",
            Self::Row(_) => "tr",
            Self::Cell(_) => "tc",
            Self::ContentControl(_) => "sdt",
            Self::SectionProperties(_) => "sectPr",
        };
        (name, self.position().get())
    }

    fn from_token(name: &str, position: Position) -> Option<Self> {
        Some(match name {
            "p" => Self::Paragraph(position),
            "r" => Self::Run(position),
            "hyperlink" => Self::Hyperlink(position),
            "tbl" => Self::Table(position),
            "tr" => Self::Row(position),
            "tc" => Self::Cell(position),
            "sdt" => Self::ContentControl(position),
            "sectPr" => Self::SectionProperties(position),
            _ => return None,
        })
    }
}

/// Where an inserted or moved element goes among a container's direct
/// children.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ElementAnchor {
    /// Before the container's first content child (its first block in the
    /// body, a cell or a content control; its first row in a table; its
    /// first cell in a row), after any property element.
    Start,
    /// Immediately before the named sibling.
    Before(ElementStep),
    /// Immediately after the named sibling.
    After(ElementStep),
    /// After the container's last child; in the body, before the body-final
    /// section properties.
    End,
}

/// The address of one main-story element, as the direct-child steps that
/// lead to it from `w:body`. The empty path is the body itself.
///
/// ```
/// use litchi_core::Position;
/// use litchi_docx::document::{ElementPath, ElementStep};
///
/// // The second paragraph of the first cell of the first table.
/// let path = ElementPath::new([
///     ElementStep::Table(Position::new(0)),
///     ElementStep::Row(Position::new(0)),
///     ElementStep::Cell(Position::new(0)),
///     ElementStep::Paragraph(Position::new(1)),
/// ]);
/// assert_eq!(path.steps().len(), 4);
/// assert_eq!(path.to_string(), "tbl:0/tr:0/tc:0/p:1");
/// ```
#[derive(Debug, Clone, PartialEq, Eq, Hash, Default)]
pub struct ElementPath(Arc<[ElementStep]>);

impl ElementPath {
    /// The path of `steps`. Its validity is checked where it is used.
    #[must_use]
    pub fn new(steps: impl Into<Arc<[ElementStep]>>) -> Self {
        Self(steps.into())
    }

    /// The empty path: the body.
    #[must_use]
    pub fn body() -> Self {
        Self::default()
    }

    /// The steps from the body.
    #[must_use]
    pub fn steps(&self) -> &[ElementStep] {
        &self.0
    }

    /// The path of a child of this element.
    #[must_use]
    pub fn child(&self, step: ElementStep) -> Self {
        let mut steps = Vec::with_capacity(self.0.len() + 1);
        steps.extend_from_slice(&self.0);
        steps.push(step);
        Self(steps.into())
    }

    /// The parent's path and this element's last step, or `None` for the
    /// body.
    #[must_use]
    pub fn split_last(&self) -> Option<(Self, ElementStep)> {
        let (last, parent) = self.0.split_last()?;
        Some((Self(parent.into()), *last))
    }

    /// Parse the durable spelling written by `Display`.
    pub(crate) fn parse(text: &str) -> Option<Self> {
        if text.is_empty() {
            return Some(Self::body());
        }
        let mut steps = Vec::new();
        for token in text.split('/') {
            let (name, position) = token.split_once(':')?;
            let position = position.parse::<usize>().ok()?;
            steps.push(ElementStep::from_token(name, Position::new(position))?);
        }
        Some(Self(steps.into()))
    }

    /// The direct-body paragraph this path names, when it is one.
    fn body_paragraph(&self) -> Option<Position> {
        match self.steps() {
            [ElementStep::Paragraph(position)] => Some(*position),
            _ => None,
        }
    }

    /// The direct-body paragraph and direct run this path names, when it
    /// names a direct run of a direct-body paragraph.
    fn body_run(&self) -> Option<(Position, Position)> {
        match self.steps() {
            [ElementStep::Paragraph(paragraph), ElementStep::Run(run)] => Some((*paragraph, *run)),
            _ => None,
        }
    }
}

impl std::fmt::Display for ElementPath {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        for (index, step) in self.0.iter().enumerate() {
            if index > 0 {
                formatter.write_str("/")?;
            }
            let (name, position) = step.token();
            write!(formatter, "{name}:{position}")?;
        }
        Ok(())
    }
}

/// A table to insert: its grid, cell text and optional style.
///
/// Each cell holds one paragraph with the cell's text. The table gets the
/// single 4-eighths-point border on every side and between cells that Word
/// and OfficeCLI give a new table, unless [`Self::with_borders`] removes it.
///
/// ```
/// use litchi_docx::document::NewTable;
///
/// let table = NewTable::new(2, 3)?
///     .with_cell_text(0, 0, "Name")
///     .with_column_widths(vec![2000, 3000, 4000])?;
/// assert_eq!(table.rows(), 2);
/// assert_eq!(table.columns(), 3);
/// # Ok::<(), litchi_docx::document::Error>(())
/// ```
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct NewTable {
    columns: usize,
    cells: Vec<Vec<String>>,
    widths: Vec<u32>,
    style: Option<String>,
    borders: bool,
}

/// The largest grid a new table may have, in rows and in columns.
pub const MAX_NEW_TABLE_DIMENSION: usize = 4096;
/// The total width of a new table's grid when no widths are given, in
/// twentieths of a point (6.5 inches, the Letter text width).
const DEFAULT_TABLE_WIDTH: u32 = 9360;

impl NewTable {
    /// A table of empty cells.
    ///
    /// # Errors
    ///
    /// Returns a typed limit error for a zero or oversized dimension.
    pub fn new(rows: usize, columns: usize) -> Result<Self> {
        for (observed, kind) in [
            (rows, crate::document::LimitKind::TableRows),
            (columns, crate::document::LimitKind::TableColumns),
        ] {
            if observed == 0 || observed > MAX_NEW_TABLE_DIMENSION {
                return Err(Error::Limit {
                    kind,
                    observed,
                    maximum: MAX_NEW_TABLE_DIMENSION,
                });
            }
        }
        let width = DEFAULT_TABLE_WIDTH / u32::try_from(columns).unwrap_or(u32::MAX).max(1);
        Ok(Self {
            columns,
            cells: vec![vec![String::new(); columns]; rows],
            widths: vec![width.max(1); columns],
            style: None,
            borders: true,
        })
    }

    /// The number of rows.
    #[must_use]
    pub fn rows(&self) -> usize {
        self.cells.len()
    }

    /// The number of columns.
    #[must_use]
    pub const fn columns(&self) -> usize {
        self.columns
    }

    /// Set one cell's text; positions outside the grid are ignored.
    #[must_use]
    pub fn with_cell_text(mut self, row: usize, column: usize, text: impl Into<String>) -> Self {
        if let Some(cell) = self.cells.get_mut(row).and_then(|row| row.get_mut(column)) {
            *cell = text.into();
        }
        self
    }

    /// Set the grid's column widths, in twentieths of a point.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::InvalidPropertyValue`] when the count differs from
    /// the column count or a width is zero.
    pub fn with_column_widths(mut self, widths: Vec<u32>) -> Result<Self> {
        if widths.len() != self.columns || widths.contains(&0) {
            return Err(Error::Refused {
                position: 0,
                reason: Refusal::InvalidPropertyValue,
            });
        }
        self.widths = widths;
        Ok(self)
    }

    /// Apply a table style by its style identifier.
    #[must_use]
    pub fn with_style(mut self, style: impl Into<String>) -> Self {
        self.style = Some(style.into());
        self
    }

    /// Keep (the default) or omit the default single borders.
    #[must_use]
    pub const fn with_borders(mut self, borders: bool) -> Self {
        self.borders = borders;
        self
    }

    /// The text of one cell.
    #[must_use]
    pub fn cell_text(&self, row: usize, column: usize) -> Option<&str> {
        self.cells
            .get(row)
            .and_then(|row| row.get(column))
            .map(String::as_str)
    }
}

/// A located element and the content root its children live in.
#[derive(Clone, Copy)]
struct Located {
    start: usize,
    end: usize,
}

/// The direct element children of a container.
struct Children {
    /// Offset of the container's end tag; insertion at the end goes here.
    close_start: usize,
    /// The prefix the container's element name is written with.
    prefix: Option<Vec<u8>>,
    /// `(start, end, word local name)`, in order. Non-WordprocessingML
    /// elements carry `None`.
    children: Vec<(usize, usize, Option<Vec<u8>>)>,
    /// Index of the body-final section properties, for the body.
    section: Option<usize>,
}

impl Children {
    fn insertion_limit(&self) -> usize {
        self.section.unwrap_or(self.children.len())
    }

    /// The child index of the `position`-th child named `local`.
    fn index_of(&self, local: &[u8], position: Position) -> Option<usize> {
        self.children
            .iter()
            .enumerate()
            .filter(|(_, child)| child.2.as_deref() == Some(local))
            .nth(position.get())
            .map(|(index, _)| index)
    }

    /// The number of children named `local` before child index `child`.
    fn count_before(&self, local: &[u8], child: usize) -> usize {
        self.children
            .iter()
            .take(child)
            .filter(|found| found.2.as_deref() == Some(local))
            .count()
    }

    fn count(&self, local: &[u8]) -> usize {
        self.count_before(local, self.children.len())
    }

    fn insertion_offset(&self, child: usize) -> Option<usize> {
        if child > self.insertion_limit() {
            return None;
        }
        Some(
            self.children
                .get(child)
                .map_or(self.close_start, |found| found.0),
        )
    }
}

fn not_found() -> Error {
    Error::Refused {
        position: 0,
        reason: Refusal::BlockNotFound,
    }
}

/// Whether `xml` is a paragraph whose removal loses only text: its
/// properties and runs of text, tabs and breaks (record 0304).
fn is_plain_text_block(xml: &[u8]) -> bool {
    let mut reader = quick_xml::Reader::from_reader(xml);
    // Depth inside a `w:pPr` or `w:rPr` subtree, whose content is formatting.
    let mut properties = 0usize;
    let mut depth = 0usize;
    loop {
        let (local, empty) = match reader.read_event() {
            Ok(Event::Start(element)) => (element.local_name().as_ref().to_vec(), false),
            Ok(Event::Empty(element)) => (element.local_name().as_ref().to_vec(), true),
            Ok(Event::End(_)) => {
                depth = depth.saturating_sub(1);
                if properties > depth {
                    properties = 0;
                }
                continue;
            },
            Ok(Event::Eof) => return true,
            Ok(_) => continue,
            Err(_) => return false,
        };
        if properties == 0 {
            let allowed = match depth {
                0 => local == b"p",
                1 => matches!(local.as_slice(), b"pPr" | b"r" | b"proofErr"),
                _ => matches!(
                    local.as_slice(),
                    b"rPr"
                        | b"t"
                        | b"tab"
                        | b"br"
                        | b"cr"
                        | b"noBreakHyphen"
                        | b"softHyphen"
                        | b"lastRenderedPageBreak"
                ),
            };
            if !allowed {
                return false;
            }
            if !empty && matches!(local.as_slice(), b"pPr" | b"rPr") {
                properties = depth + 1;
            }
        }
        if !empty {
            depth += 1;
        }
    }
}

fn refusal(reason: Refusal) -> Error {
    Error::Refused {
        position: 0,
        reason,
    }
}

/// The `(start, end)` byte ranges of the body's direct element children and
/// the offset of the body's end tag, as the element operations count them
/// (record 0279).
pub(in crate::document::transaction) fn body_children(
    snapshot: &Snapshot,
) -> Result<(Vec<(usize, usize)>, usize)> {
    let (children, _) = container_children(snapshot, &ElementPath::new(Vec::<ElementStep>::new()))?;
    Ok((
        children
            .children
            .iter()
            .map(|(start, end, _)| (*start, *end))
            .collect(),
        children.close_start,
    ))
}

/// The byte range of the element `path` names in `snapshot`, for the
/// story substitution of record 0274.
pub(in crate::document::transaction) fn locate_element(
    snapshot: &Snapshot,
    path: &ElementPath,
) -> Result<std::ops::Range<usize>> {
    let located = locate(snapshot, path.steps())?;
    Ok(located.start..located.end)
}

/// Locate the element `path` names in `snapshot`.
fn locate(snapshot: &Snapshot, path: &[ElementStep]) -> Result<Located> {
    let (first, rest) = path.split_first().ok_or_else(not_found)?;
    let (mut start, mut end) = match *first {
        ElementStep::Paragraph(position) => {
            let range = snapshot.range(position)?;
            (
                checked_start(range, "paragraph")?,
                checked_end(range, "paragraph")?,
            )
        },
        ElementStep::Table(position) => {
            let range = snapshot
                .tables
                .get(position.get())
                .copied()
                .ok_or_else(not_found)?;
            (checked_start(range, "table")?, checked_end(range, "table")?)
        },
        ElementStep::ContentControl(position) => {
            let range = snapshot
                .block_controls
                .get(position.get())
                .copied()
                .ok_or_else(not_found)?;
            (
                checked_start(range, "content control")?,
                checked_end(range, "content control")?,
            )
        },
        ElementStep::SectionProperties(position) if position.get() == 0 && rest.is_empty() => {
            let body = super::body_of_snapshot(snapshot)?;
            let index = body.section.ok_or_else(not_found)?;
            let child = body.children.get(index).ok_or_else(not_found)?;
            (child.start, child.end)
        },
        _ => return Err(not_found()),
    };
    let mut local: &[u8] = first.local();
    for step in rest {
        let (container_start, container_end, container) = if local == b"sdt" {
            let (content_start, content_end) =
                select_single_control_content(snapshot, start, end, 0)?;
            (content_start, content_end, b"sdtContent".as_slice())
        } else {
            (start, end, local)
        };
        if !step.fits(container) {
            return Err(not_found());
        }
        let container_xml = checked_slice(
            snapshot.xml_bytes(),
            container_start,
            container_end,
            "element container",
        )?;
        let (relative_start, relative_end) =
            if matches!(step, ElementStep::Row(_) | ElementStep::Cell(_)) {
                select_grid_child(container_xml, step.local(), step.position())?
            } else {
                let range = select_direct_child(
                    container_xml,
                    container,
                    step.local(),
                    step.position(),
                    Refusal::BlockNotFound,
                )
                .map_err(refusal)?;
                (
                    checked_start(range, "element")?,
                    checked_end(range, "element")?,
                )
            };
        start = container_start
            .checked_add(relative_start)
            .ok_or_else(|| invalid("element offset overflows"))?;
        end = container_start
            .checked_add(relative_end)
            .ok_or_else(|| invalid("element offset overflows"))?;
        local = step.local();
    }
    Ok(Located { start, end })
}

/// The range of the `position`-th direct `child` (a row of a table or a cell
/// of a row) of the fragment `xml`.
///
/// Readers count every row or cell of the grid, including ones inside a
/// row- or cell-level content control or custom-XML wrapper; such a wrapper
/// before the target would make the reader's position differ from the
/// direct-child position, so it is refused. Children are matched by the
/// root's own prefix.
fn select_grid_child(xml: &[u8], child: &[u8], position: Position) -> Result<(usize, usize)> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let mut depth = 0usize;
    let mut root_prefix = None::<Vec<u8>>;
    let mut index = 0usize;
    let mut capture = None::<usize>;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("Word fragment offset overflow"))?;
        let event = reader
            .read_event()
            .map_err(|_| Error::Document(crate::Error::Xml("invalid Word table XML".into())))?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("Word fragment offset overflow"))?;
        let (element, empty) = match event {
            Event::Start(element) => (element, false),
            Event::Empty(element) => (element, true),
            Event::End(_) => {
                depth = depth.saturating_sub(1);
                if depth == 1
                    && let Some(capture_start) = capture
                {
                    return Ok((capture_start, end));
                }
                continue;
            },
            Event::Eof => return Err(not_found()),
            _ => continue,
        };
        let prefix = element
            .name()
            .prefix()
            .map_or_else(Vec::new, |prefix| prefix.into_inner().to_vec());
        if depth == 0 {
            root_prefix = Some(prefix);
        } else if depth == 1 && root_prefix.as_ref() == Some(&prefix) {
            match element.local_name().as_ref() {
                b"sdt" | b"customXml" => return Err(refusal(Refusal::ComplexContent)),
                local if local == child => {
                    if index == position.get() {
                        if empty {
                            return Ok((start, end));
                        }
                        capture = Some(start);
                    }
                    index += 1;
                },
                _ => {},
            }
        }
        if !empty {
            depth += 1;
        }
    }
}

/// The namespace bindings in scope at the element starting at `offset`,
/// including the element's own declarations.
fn bindings_at(xml: &[u8], offset: usize) -> Result<Vec<(Option<Vec<u8>>, Vec<u8>)>> {
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("document offset overflow"))?;
        let event = reader
            .read_event()
            .map_err(|_| Error::Document(crate::Error::Xml("invalid Word document XML".into())))?;
        if start >= offset {
            if !matches!(event, Event::Start(_) | Event::Empty(_)) || start != offset {
                return Err(invalid("element offset is not an element start"));
            }
            break;
        }
        if matches!(event, Event::Eof) {
            return Err(invalid("element offset is outside the document"));
        }
    }
    Ok(reader
        .resolver()
        .bindings()
        .map(|(declaration, namespace)| {
            let prefix = match declaration {
                quick_xml::name::PrefixDeclaration::Default => None,
                quick_xml::name::PrefixDeclaration::Named(prefix) => Some(prefix.to_vec()),
            };
            (prefix, namespace.into_inner().to_vec())
        })
        .collect())
}

/// The direct element children of the element fragment `xml`, whose root
/// is the WordprocessingML element `root`.
fn element_children(
    xml: &[u8],
    root: &[u8],
    inherited: &[(Option<Vec<u8>>, Vec<u8>)],
) -> Result<Children> {
    let mut reader = NsReader::from_reader(xml);
    super::seed_namespaces(&mut reader, inherited)
        .map_err(|()| invalid("invalid inherited Word namespace"))?;
    let origin = ReaderOrigin::of(xml);
    let mut fragment_prefix = FragmentPrefix::Unseen;
    let mut depth = 0usize;
    let mut layout = Children {
        close_start: 0,
        prefix: None,
        children: Vec::new(),
        section: None,
    };
    let mut open = None::<(usize, Option<Vec<u8>>)>;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("Word fragment offset overflow"))?;
        let raw = reader
            .read_event()
            .map_err(|_| Error::Document(crate::Error::Xml("invalid Word fragment".into())))?;
        let (namespace, event) = reader.resolver().resolve_event(raw);
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("Word fragment offset overflow"))?;
        let word_local = |element: &quick_xml::events::BytesStart<'_>, prefix: &FragmentPrefix| {
            is_fragment_word_element(&namespace, prefix)
                .then(|| element.local_name().as_ref().to_vec())
        };
        match event {
            Event::Start(element) if depth == 0 => {
                fragment_prefix = FragmentPrefix::from_name(element.name());
                if word_local(&element, &fragment_prefix).as_deref() != Some(root) {
                    return Err(invalid("Word fragment has an unexpected root"));
                }
                layout.prefix = Some(
                    element
                        .name()
                        .prefix()
                        .map_or_else(Vec::new, |prefix| prefix.into_inner().to_vec()),
                );
                depth = 1;
            },
            Event::Empty(_) if depth == 0 => {
                // An empty container (`<w:tc/>`) has nowhere to splice a
                // child without rewriting its tag.
                return Err(refusal(Refusal::ComplexContent));
            },
            Event::Start(element) => {
                if depth == 1 {
                    open = Some((start, word_local(&element, &fragment_prefix)));
                }
                depth += 1;
            },
            Event::Empty(element) => {
                if depth == 1 {
                    layout
                        .children
                        .push((start, end, word_local(&element, &fragment_prefix)));
                }
            },
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("Word fragment is unbalanced"))?;
                if depth == 1
                    && let Some((child_start, local)) = open.take()
                {
                    layout.children.push((child_start, end, local));
                } else if depth == 0 {
                    layout.close_start = start;
                    return Ok(layout);
                }
            },
            Event::Eof => return Err(invalid("Word fragment is unterminated")),
            _ => {},
        }
    }
}

/// The content root of the container `path` names: its absolute range,
/// local name and inherited namespaces.
struct ContainerRoot {
    start: usize,
    end: usize,
    local: &'static [u8],
}

fn container_root(snapshot: &Snapshot, path: &ElementPath) -> Result<Option<ContainerRoot>> {
    let Some((_, last)) = path.split_last() else {
        return Ok(None);
    };
    let located = locate(snapshot, path.steps())?;
    Ok(Some(if matches!(last, ElementStep::ContentControl(_)) {
        let (start, end) = select_single_control_content(snapshot, located.start, located.end, 0)?;
        ContainerRoot {
            start,
            end,
            local: b"sdtContent",
        }
    } else {
        ContainerRoot {
            start: located.start,
            end: located.end,
            local: last.local(),
        }
    }))
}

/// The direct children of the container `path` names, as absolute ranges.
fn container_children(
    snapshot: &Snapshot,
    path: &ElementPath,
) -> Result<(Children, &'static [u8])> {
    match container_root(snapshot, path)? {
        None => {
            let body = super::body_of_snapshot(snapshot)?;
            let children = body
                .children
                .iter()
                .map(|child| {
                    let local: Option<&[u8]> = match child.kind {
                        super::ChildKind::Paragraph(_) => Some(b"p"),
                        super::ChildKind::Table(_) => Some(b"tbl"),
                        super::ChildKind::ContentControl(_) => Some(b"sdt"),
                        super::ChildKind::Other => None,
                    };
                    (child.start, child.end, local.map(<[u8]>::to_vec))
                })
                .collect();
            Ok((
                Children {
                    close_start: body.body_end,
                    prefix: body.prefix,
                    children,
                    section: body.section,
                },
                b"body",
            ))
        },
        Some(root) => {
            let xml = checked_slice(snapshot.xml_bytes(), root.start, root.end, "container")?;
            let inherited = bindings_at(snapshot.xml_bytes(), root.start)?;
            let mut children = element_children(xml, root.local, &inherited)?;
            children.close_start += root.start;
            for child in &mut children.children {
                child.0 += root.start;
                child.1 += root.start;
            }
            Ok((children, root.local))
        },
    }
}

impl Edit {
    // -----------------------------------------------------------------
    // Replay primitives
    // -----------------------------------------------------------------

    /// Replace the exact bytes of the element at `path`; the replay of
    /// [`Operation::ReplaceElement`].
    pub(in crate::document::transaction) fn replace_element(
        &mut self,
        path: &ElementPath,
        before: &Arc<Vec<u8>>,
        after: &Arc<Vec<u8>>,
    ) -> Result<&mut Self> {
        self.ensure_unmanaged("replace_element")?;
        self.reserve_operation()?;
        let located = locate(&self.projected, path.steps())?;
        if self.projected.xml_bytes().get(located.start..located.end) != Some(before.as_slice()) {
            return Err(Error::SemanticPrecondition);
        }
        if element_len(after)? != after.len() {
            return Err(invalid("element replacement is not one element"));
        }
        ensure_utf8(after)?;
        let replacement_text_bytes = self.checked_text_total(after.len())?;
        let candidate = self.projected_with_xml(replace_range(
            self.projected.xml_bytes(),
            located.start,
            located.end,
            after,
        )?)?;
        let readback = locate(&candidate, path.steps())?;
        if candidate.xml_bytes().get(readback.start..readback.end) != Some(after.as_slice()) {
            return Err(invalid("element replacement failed exact readback"));
        }
        self.operations.push(Operation::ReplaceElement {
            path: path.clone(),
            before: Arc::clone(before),
            after: Arc::clone(after),
        });
        self.replacement_text_bytes = replacement_text_bytes;
        self.projected = candidate;
        Ok(self)
    }

    /// Insert exact element bytes at a container's direct child index; the
    /// replay of [`Operation::InsertElement`].
    pub(in crate::document::transaction) fn insert_element(
        &mut self,
        container: &ElementPath,
        child: Position,
        xml: &Arc<Vec<u8>>,
    ) -> Result<&mut Self> {
        self.ensure_unmanaged("insert_element")?;
        self.reserve_operation()?;
        if element_len(xml)? != xml.len() {
            return Err(invalid("inserted element is not one element"));
        }
        ensure_utf8(xml)?;
        let replacement_text_bytes = self.checked_text_total(xml.len())?;
        let (children, _) = container_children(&self.projected, container)?;
        let offset = children
            .insertion_offset(child.get())
            .ok_or(Error::SemanticPrecondition)?;
        let count = children.children.len();
        let candidate = self.projected_with_xml(replace_range(
            self.projected.xml_bytes(),
            offset,
            offset,
            xml,
        )?)?;
        let (after, _) = container_children(&candidate, container)?;
        let inserted = after
            .children
            .get(child.get())
            .map(|found| (found.0, found.1));
        if after.children.len() != count + 1
            || inserted.and_then(|(start, end)| candidate.xml_bytes().get(start..end))
                != Some(xml.as_slice())
        {
            return Err(invalid("element insertion failed exact readback"));
        }
        self.operations.push(Operation::InsertElement {
            container: container.clone(),
            child,
            xml: Arc::clone(xml),
        });
        self.replacement_text_bytes = replacement_text_bytes;
        self.projected = candidate;
        Ok(self)
    }

    /// Remove the exact element at a container's direct child index; the
    /// replay of [`Operation::RemoveElement`].
    pub(in crate::document::transaction) fn remove_element(
        &mut self,
        container: &ElementPath,
        child: Position,
        xml: &Arc<Vec<u8>>,
    ) -> Result<&mut Self> {
        self.ensure_unmanaged("remove_element")?;
        self.reserve_operation()?;
        let (children, _) = container_children(&self.projected, container)?;
        let (start, end, _) = children
            .children
            .get(child.get())
            .cloned()
            .ok_or(Error::SemanticPrecondition)?;
        if children.section == Some(child.get())
            || self.projected.xml_bytes().get(start..end) != Some(xml.as_slice())
        {
            return Err(Error::SemanticPrecondition);
        }
        let count = children.children.len();
        let candidate =
            self.projected_with_xml(replace_range(self.projected.xml_bytes(), start, end, &[])?)?;
        let (after, _) = container_children(&candidate, container)?;
        if after.children.len().checked_add(1) != Some(count) {
            return Err(invalid("element removal failed readback"));
        }
        self.operations.push(Operation::RemoveElement {
            container: container.clone(),
            child,
            xml: Arc::clone(xml),
        });
        self.projected = candidate;
        Ok(self)
    }

    // -----------------------------------------------------------------
    // Element access
    // -----------------------------------------------------------------

    /// The exact bytes of the element at `path` in the projected document.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for a path that names no element.
    pub fn element_xml(&self, path: &ElementPath) -> Result<&[u8]> {
        let located = locate(&self.projected, path.steps())?;
        checked_slice(
            self.projected.xml_bytes(),
            located.start,
            located.end,
            "element",
        )
    }

    /// The element at `path`, its bytes and its inherited namespaces.
    fn element_with_scope(
        &self,
        path: &ElementPath,
        local: &[u8],
    ) -> Result<(Located, Vec<(Option<Vec<u8>>, Vec<u8>)>)> {
        if path.steps().last().map(|step| step.local()) != Some(local) {
            return Err(not_found());
        }
        let located = locate(&self.projected, path.steps())?;
        let inherited = bindings_at(self.projected.xml_bytes(), located.start)?;
        Ok((located, inherited))
    }

    fn element_bytes(&self, located: Located) -> Result<Arc<Vec<u8>>> {
        let bytes = checked_slice(
            self.projected.xml_bytes(),
            located.start,
            located.end,
            "element",
        )?;
        let mut owned = buffer(bytes.len())?;
        owned.extend_from_slice(bytes);
        Ok(Arc::new(owned))
    }

    // -----------------------------------------------------------------
    // Paragraph and run edits
    // -----------------------------------------------------------------

    /// Change the direct properties of the paragraph at `path`, with the
    /// rules of [`Self::format_paragraph`].
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::document::{raw, ElementPath, ElementStep};
    /// use litchi_docx::properties::Justification;
    ///
    /// let mut edit = raw::Tree::from_xml(
    ///     br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:body></w:document>"#
    ///         .to_vec(),
    ///     raw::Limits::default(),
    /// )?;
    /// let path = ElementPath::new([
    ///     ElementStep::Table(Position::new(0)),
    ///     ElementStep::Row(Position::new(0)),
    ///     ElementStep::Cell(Position::new(0)),
    ///     ElementStep::Paragraph(Position::new(0)),
    /// ]);
    /// edit.format_paragraph_at(&path, |paragraph| {
    ///     paragraph.justification = Some(Justification::Center);
    /// })?;
    /// let bytes = edit.to_bytes(raw::Limits::default())?;
    /// let xml = std::str::from_utf8(&bytes)?;
    /// assert!(xml.contains(r#"<w:tc><w:p><w:pPr><w:jc w:val="center"/></w:pPr>"#));
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// As [`Self::format_paragraph`], and [`Refusal::BlockNotFound`] for a
    /// path that names no paragraph.
    pub fn format_paragraph_at(
        &mut self,
        path: &ElementPath,
        change: impl FnOnce(&mut properties::Paragraph),
    ) -> Result<&mut Self> {
        if let Some(position) = path.body_paragraph() {
            return self.format_paragraph(position, change);
        }
        self.ensure_unmanaged("format_paragraph")?;
        let (located, inherited) = self.element_with_scope(path, b"p")?;
        let before = self.element_bytes(located)?;
        let current = properties::parse_paragraph(&before, &inherited)?;
        let mut target = current.clone();
        change(&mut target);
        let strict = matches!(self.projected.conformance, Conformance::Strict);
        let Some(after) =
            properties::splice_paragraph(&before, &current, &target, strict, &inherited)
                .map_err(spliced(Position::new(0)))?
        else {
            return Ok(self);
        };
        self.replace_element(path, &before, &Arc::new(after))
    }

    /// Change the direct properties (`w:tblPr`) of the table at `path`
    /// (record 0342). The style and the borders are written: only the
    /// properties whose typed value changes are rewritten, and every other
    /// byte of the table is kept. An unchanged value is an exact no-op.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::document::{raw, ElementPath, ElementStep};
    /// use litchi_docx::properties::{Border, BorderStyle, EighthPoints};
    ///
    /// let mut edit = raw::Tree::from_xml(
    ///     br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:body></w:document>"#
    ///         .to_vec(),
    ///     raw::Limits::default(),
    /// )?;
    /// let path = ElementPath::new([ElementStep::Table(Position::new(0))]);
    /// edit.format_table_at(&path, |table| {
    ///     let mut border = Border::new(BorderStyle::from_xml("double").unwrap());
    ///     border.size = Some(EighthPoints(8));
    ///     table.borders.get_or_insert_default().top = Some(border);
    /// })?;
    /// let bytes = edit.to_bytes(raw::Limits::default())?;
    /// let xml = std::str::from_utf8(&bytes)?;
    /// assert!(xml.contains(r#"<w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="double" w:sz="8"/></w:tblBorders>"#));
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::UnmodelledProperty`] for a change of any other
    /// table property, [`Refusal::BlockNotFound`] for a path that names no
    /// table, and the splice refusals of [`Self::format_paragraph`].
    pub fn format_table_at(
        &mut self,
        path: &ElementPath,
        change: impl FnOnce(&mut properties::Table),
    ) -> Result<&mut Self> {
        self.ensure_unmanaged("format_table")?;
        let (located, inherited) = self.element_with_scope(path, b"tbl")?;
        let before = self.element_bytes(located)?;
        let current = properties::parse_table(&before, &inherited)?;
        let mut target = current.clone();
        change(&mut target);
        let Some(after) = properties::splice_table(&before, &current, &target, &inherited)
            .map_err(spliced(Position::new(0)))?
        else {
            return Ok(self);
        };
        self.replace_element(path, &before, &Arc::new(after))
    }

    /// Change the direct properties (`w:trPr`) of the table row at `path`
    /// (record 0342): its height and its `cantSplit`, `tblHeader` and
    /// `hidden` flags, with the rules of [`Self::format_table_at`].
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::document::{raw, ElementPath, ElementStep};
    /// use litchi_docx::properties::{Height, HeightRule, Twips};
    ///
    /// let mut edit = raw::Tree::from_xml(
    ///     br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:body></w:document>"#
    ///         .to_vec(),
    ///     raw::Limits::default(),
    /// )?;
    /// let path = ElementPath::new([
    ///     ElementStep::Table(Position::new(0)),
    ///     ElementStep::Row(Position::new(0)),
    /// ]);
    /// edit.format_row_at(&path, |row| {
    ///     row.header = Some(true);
    ///     row.height = Some(Height::new(Twips(400), Some(HeightRule::Exact)));
    /// })?;
    /// let bytes = edit.to_bytes(raw::Limits::default())?;
    /// let xml = std::str::from_utf8(&bytes)?;
    /// assert!(xml.contains(r#"<w:tr><w:trPr><w:trHeight w:val="400" w:hRule="exact"/><w:tblHeader/></w:trPr>"#));
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::UnmodelledProperty`] for a change of any other
    /// row property, [`Refusal::BlockNotFound`] for a path that names no
    /// row, and the splice refusals of [`Self::format_paragraph`].
    pub fn format_row_at(
        &mut self,
        path: &ElementPath,
        change: impl FnOnce(&mut properties::Row),
    ) -> Result<&mut Self> {
        self.ensure_unmanaged("format_row")?;
        let (located, inherited) = self.element_with_scope(path, b"tr")?;
        let before = self.element_bytes(located)?;
        let current = properties::parse_row(&before, &inherited)?;
        let mut target = current.clone();
        change(&mut target);
        let Some(after) = properties::splice_row(&before, &current, &target, &inherited)
            .map_err(spliced(Position::new(0)))?
        else {
            return Ok(self);
        };
        self.replace_element(path, &before, &Arc::new(after))
    }

    /// Change the direct properties of the run at `path` (a run of a
    /// paragraph, a hyperlink or an inline content control), with the rules
    /// of [`Self::format_run`].
    ///
    /// # Errors
    ///
    /// As [`Self::format_run`], and [`Refusal::BlockNotFound`] for a path
    /// that names no run.
    pub fn format_run_at(
        &mut self,
        path: &ElementPath,
        change: impl FnOnce(&mut properties::Run),
    ) -> Result<&mut Self> {
        if let Some((paragraph, run)) = path.body_run() {
            return self.format_run(paragraph, run, change);
        }
        self.ensure_unmanaged("format_run")?;
        let (located, inherited) = self.element_with_scope(path, b"r")?;
        let before = self.element_bytes(located)?;
        let current = properties::parse_run(&before, &inherited)?;
        let mut target = current.clone();
        change(&mut target);
        let Some(after) = properties::splice_run(&before, &current, &target, &inherited)
            .map_err(spliced(Position::new(0)))?
        else {
            return Ok(self);
        };
        self.replace_element(path, &before, &Arc::new(after))
    }

    /// Replace the text of the paragraph, run or hyperlink at `path`,
    /// keeping its properties and opaque children, with the rules of
    /// [`Self::replace_paragraph_text`] and [`Self::replace_run_text`].
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for text or content the rewrite cannot
    /// represent, [`Refusal::BlockNotFound`] for a path that names no
    /// paragraph, run or hyperlink, or a resource-limit error.
    pub fn replace_text_at(
        &mut self,
        path: &ElementPath,
        authored_text: impl Into<String>,
    ) -> Result<&mut Self> {
        let text = authored_text.into();
        if let Some(position) = path.body_paragraph() {
            return self.replace_paragraph_text(position, text);
        }
        if let Some((paragraph, run)) = path.body_run() {
            return self.replace_run_text(paragraph, run, text);
        }
        self.ensure_unmanaged("replace_text_at")?;
        let local = path
            .steps()
            .last()
            .map(|step| step.local())
            .filter(|local| matches!(*local, b"p" | b"r" | b"hyperlink"))
            .ok_or_else(not_found)?;
        validate_authored_text(&text).map_err(refusal)?;
        let located = locate(&self.projected, path.steps())?;
        let before = self.element_bytes(located)?;
        let owner = if local == b"p" {
            scan_paragraph_text_owner(&before)
        } else {
            scan_text_owner(&before, local)
        }
        .map_err(refusal)?;
        if owner.text == text {
            return Ok(self);
        }
        let after = rewrite_text_owner(&before, &owner, &text)?;
        self.replace_element(path, &before, &Arc::new(after))?;
        let readback = if local == b"p" {
            scan_paragraph_text_owner(self.element_xml(path)?)
        } else {
            scan_text_owner(self.element_xml(path)?, local)
        }
        .map_err(refusal)?;
        if readback.text != text {
            return Err(invalid("element text replacement failed semantic readback"));
        }
        Ok(self)
    }

    /// Replace the content of the table cell at `path` with one paragraph
    /// holding `text`.
    ///
    /// The cell's properties (`w:tcPr`) and its first paragraph's
    /// properties are kept. Every other block of the cell is removed, so the
    /// edit is refused when one of them holds an open range, a story
    /// reference or a section break, and (record 0304) when one of them is
    /// not a paragraph of plain runs: a table, a content control, a
    /// drawing, a field, a bookmark or anything else that is not text would
    /// be lost. The first paragraph is rewritten with the rules of
    /// [`Self::replace_paragraph_text`], so its refusals apply too.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for a missing cell or unsafe content, invalid
    /// text, or a resource-limit error.
    pub fn replace_cell_text_at(
        &mut self,
        path: &ElementPath,
        authored_text: impl Into<String>,
    ) -> Result<&mut Self> {
        self.ensure_unmanaged("replace_cell_text_at")?;
        let text = authored_text.into();
        validate_authored_text(&text).map_err(refusal)?;
        let (located, inherited) = self.element_with_scope(path, b"tc")?;
        let before = self.element_bytes(located)?;
        let layout = element_children(&before, b"tc", &inherited)?;
        let first_paragraph = layout
            .children
            .iter()
            .position(|child| child.2.as_deref() == Some(b"p"));
        // Every block after the cell properties is replaced by the first
        // paragraph, rewritten; the others are removed.
        let first_block = layout
            .children
            .iter()
            .position(|child| child.2.as_deref() != Some(b"tcPr"))
            .unwrap_or(layout.children.len());
        for (index, (start, end, _)) in layout.children.iter().enumerate().skip(first_block) {
            if Some(index) == first_paragraph {
                continue;
            }
            paragraph_dependencies(&before[*start..*end], &inherited)
                .refuse(Removal::Delete)
                .map_err(refusal)?;
            if !is_plain_text_block(&before[*start..*end]) {
                return Err(refusal(Refusal::ComplexContent));
            }
        }
        let paragraph = match first_paragraph {
            Some(index) => {
                let (start, end, _) = &layout.children[index];
                let paragraph = &before[*start..*end];
                match scan_text_owner(paragraph, b"p") {
                    Ok(owner) => rewrite_text_owner(paragraph, &owner, &text)?,
                    // A paragraph with no text slot (an empty one) loses
                    // nothing to a plain replacement; any other paragraph the
                    // text rewrite refuses holds content it would drop
                    // (record 0304).
                    Err(_) if is_plain_text_block(paragraph) => {
                        self.plain_paragraph(&layout, &text)?
                    },
                    Err(reason) => return Err(refusal(reason)),
                }
            },
            None => self.plain_paragraph(&layout, &text)?,
        };
        let removed_start = layout
            .children
            .get(first_block)
            .map_or(layout.close_start, |child| child.0);
        let mut after = buffer(staging_add(
            staging_add(removed_start, paragraph.len())?,
            before.len() - layout.close_start,
        )?)?;
        after.extend_from_slice(&before[..removed_start]);
        after.extend_from_slice(&paragraph);
        after.extend_from_slice(&before[layout.close_start..]);
        if after.as_slice() == before.as_slice() {
            return Ok(self);
        }
        self.replace_element(path, &before, &Arc::new(after))
    }

    fn plain_paragraph(&self, layout: &Children, text: &str) -> Result<Vec<u8>> {
        let paragraph = if text.is_empty() {
            NewParagraph::new()
        } else {
            NewParagraph::new().with_run(NewRun::new(text))
        };
        let strict = matches!(self.projected.conformance, Conformance::Strict);
        paragraph_markup(
            layout.prefix.as_deref(),
            self.projected.conformance,
            &paragraph,
            strict,
        )
        .map_err(spliced(Position::new(0)))
    }

    /// Insert a run into the paragraph, hyperlink or inline content control
    /// at `owner`, before its direct run `at` (or after its last child for
    /// `None`), and return the run's position among the owner's direct runs.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::RunNotFound`] for a run position past the end,
    /// [`Refusal::FieldInterior`] inside a complex field, a typed refusal for
    /// the new run's text or properties, or a resource-limit error.
    pub fn insert_run_at(
        &mut self,
        owner: &ElementPath,
        at: Option<Position>,
        new_run: &NewRun,
    ) -> Result<Position> {
        if let Some(paragraph) = owner.body_paragraph() {
            return match at {
                Some(run) => self.insert_run(paragraph, run, new_run),
                None => self.append_run(paragraph, new_run),
            };
        }
        self.ensure_unmanaged("insert_run_at")?;
        let conformance = self.projected.conformance;
        self.insert_inline(owner, at, b"r", |prefix| {
            run_markup(prefix, conformance, new_run).map_err(spliced(Position::new(0)))
        })
    }

    /// Whether the edited document is Strict OOXML.
    pub(crate) fn is_strict(&self) -> bool {
        matches!(self.projected.conformance, Conformance::Strict)
    }

    /// Insert a run holding an inline picture into the paragraph at
    /// `paragraph`, before its direct run `at` (after its last child for
    /// `None`), and return the run's position among the paragraph's direct
    /// runs. `relationship` is the image relationship the caller stages.
    pub(crate) fn insert_picture_markup_at(
        &mut self,
        paragraph: &ElementPath,
        at: Option<Position>,
        relationship: &str,
        picture: &PictureMarkup<'_>,
    ) -> Result<Position> {
        self.ensure_unmanaged("insert_picture_at")?;
        if paragraph.steps().last().map(|step| step.local()) != Some(b"p".as_slice()) {
            return Err(not_found());
        }
        let conformance = self.projected.conformance;
        let (_, before) = self.insert_inline_counted(paragraph, at, b"r", true, |prefix| {
            picture_markup(prefix, conformance, relationship, picture)
        })?;
        Ok(Position::new(before))
    }

    /// Insert a run holding an inline chart into the paragraph at
    /// `paragraph`, before its direct run `at` (after its last child for
    /// `None`), and return the chart's position among the paragraph's
    /// charts. `relationship` is the chart relationship the caller stages.
    pub(crate) fn insert_chart_markup_at(
        &mut self,
        paragraph: &ElementPath,
        at: Option<Position>,
        relationship: &str,
        chart: &ChartMarkup<'_>,
    ) -> Result<Position> {
        self.ensure_unmanaged("insert_chart_at")?;
        if paragraph.steps().last().map(|step| step.local()) != Some(b"p".as_slice()) {
            return Err(not_found());
        }
        let conformance = self.projected.conformance;
        self.insert_inline_counted(paragraph, at, b"r", false, |prefix| {
            chart_markup(prefix, conformance, relationship, chart)
        })?;
        let located = locate(&self.projected, paragraph.steps())?;
        let bytes = self.element_bytes(located)?;
        let spans = crate::chart::inline::chart_spans(&bytes).map_err(Error::Document)?;
        spans
            .iter()
            .position(|span| span.chart.relationship_id() == relationship)
            .map(Position::new)
            .ok_or_else(|| invalid("the inserted chart did not read back"))
    }

    /// Insert a run holding an inline SmartArt graphic into the paragraph at
    /// `paragraph`, before its direct run `at` (after its last child for
    /// `None`) (record 0295). `relationships` are the data, layout, quick
    /// style and colors relationships the caller stages.
    pub(crate) fn insert_smartart_markup_at(
        &mut self,
        paragraph: &ElementPath,
        at: Option<Position>,
        relationships: [&str; 4],
        smartart: &SmartArtMarkup<'_>,
    ) -> Result<()> {
        self.ensure_unmanaged("insert_smartart_at")?;
        if paragraph.steps().last().map(|step| step.local()) != Some(b"p".as_slice()) {
            return Err(not_found());
        }
        let conformance = self.projected.conformance;
        self.insert_inline_counted(paragraph, at, b"r", false, |prefix| {
            smartart_markup(prefix, conformance, relationships, smartart)
        })?;
        Ok(())
    }

    /// Resize (`Some((width, height))`, in EMUs) or remove (`None`) the
    /// chart at `ordinal` among the paragraph's charts. A removed chart
    /// takes its run with it when the run holds nothing else.
    pub(crate) fn rewrite_chart_at(
        &mut self,
        paragraph: &ElementPath,
        ordinal: Position,
        extent: Option<(i64, i64)>,
    ) -> Result<bool> {
        self.ensure_unmanaged("rewrite_chart_at")?;
        if paragraph.steps().last().map(|step| step.local()) != Some(b"p".as_slice()) {
            return Err(not_found());
        }
        let located = locate(&self.projected, paragraph.steps())?;
        let before = self.element_bytes(located)?;
        let spans = crate::chart::inline::chart_spans(&before).map_err(Error::Document)?;
        let span = spans
            .get(ordinal.get())
            .ok_or_else(|| refusal(Refusal::RunNotFound))?;
        let mut after = before.as_ref().clone();
        match extent {
            Some((width, height)) => {
                let range = span
                    .extent
                    .clone()
                    .ok_or_else(|| invalid("the chart drawing has no extent"))?;
                let tag = String::from_utf8_lossy(&before[range.clone()]).into_owned();
                let prefix = tag
                    .strip_prefix('<')
                    .and_then(|rest| rest.split_whitespace().next())
                    .unwrap_or("wp:extent")
                    .trim_end_matches("/>")
                    .to_owned();
                let replacement = format!(r#"<{prefix} cx="{width}" cy="{height}"/>"#);
                if tag == replacement {
                    return Ok(false);
                }
                after.splice(range, replacement.bytes());
                // A floating frame's own transform, when it has one, follows.
            },
            None => {
                let range = span
                    .sole_run
                    .clone()
                    .unwrap_or_else(|| span.drawing.clone());
                after.drain(range);
            },
        }
        self.replace_element(paragraph, &before, &Arc::new(after))?;
        Ok(true)
    }

    /// Rewrite (`Some`) or remove (`None`) the picture at `ordinal` among
    /// the paragraph's inline pictures (the order of
    /// [`crate::Paragraph::images`], record 0335). A removed picture takes
    /// its run with it when the run holds nothing else. Returns whether
    /// anything changed.
    pub(crate) fn rewrite_picture_at(
        &mut self,
        paragraph: &ElementPath,
        ordinal: Position,
        rewrite: Option<&PictureRewrite<'_>>,
    ) -> Result<bool> {
        self.ensure_unmanaged("rewrite_picture_at")?;
        if paragraph.steps().last().map(|step| step.local()) != Some(b"p".as_slice()) {
            return Err(not_found());
        }
        let located = locate(&self.projected, paragraph.steps())?;
        let before = self.element_bytes(located)?;
        let spans = crate::picture_spans::picture_spans(&before).map_err(Error::Document)?;
        let span = spans
            .get(ordinal.get())
            .ok_or_else(|| refusal(Refusal::RunNotFound))?;
        let mut edits: Vec<(std::ops::Range<usize>, Vec<u8>)> = Vec::new();
        match rewrite {
            None => {
                let range = span
                    .sole_run
                    .clone()
                    .unwrap_or_else(|| span.drawing.clone());
                edits.push((range, Vec::new()));
            },
            Some(rewrite) => {
                let tag = |range: &Option<std::ops::Range<usize>>,
                           changes: &[(&str, &str)],
                           edits: &mut Vec<(std::ops::Range<usize>, Vec<u8>)>|
                 -> Result<()> {
                    let Some(range) = range.clone() else {
                        return Err(invalid("the picture drawing lacks an element to rewrite"));
                    };
                    let mut bytes = before[range.clone()].to_vec();
                    for (key, value) in changes {
                        bytes = crate::picture_spans::with_attribute(&bytes, key, value)
                            .ok_or_else(|| invalid("the picture drawing has an unreadable tag"))?;
                    }
                    edits.push((range, bytes));
                    Ok(())
                };
                if let Some((width, height)) = rewrite.extent {
                    let (width, height) = (width.to_string(), height.to_string());
                    let changes = [("cx", width.as_str()), ("cy", height.as_str())];
                    tag(&span.extent, &changes, &mut edits)?;
                    if span.shape_extent.is_some() {
                        tag(&span.shape_extent, &changes, &mut edits)?;
                    }
                }
                let mut properties: Vec<(&str, &str)> = Vec::new();
                if let Some(name) = rewrite.name {
                    properties.push(("name", name));
                }
                if let Some(description) = rewrite.description {
                    properties.push(("descr", description));
                }
                if !properties.is_empty() {
                    tag(&span.properties, &properties, &mut edits)?;
                }
                if let Some(embed) = rewrite.embed {
                    let range = span
                        .blip
                        .clone()
                        .ok_or_else(|| invalid("the picture names no image"))?;
                    let key = crate::picture_spans::attribute_named(&before[range], "embed")
                        .ok_or_else(|| invalid("the picture names no image"))?;
                    tag(&span.blip, &[(key.as_str(), embed)], &mut edits)?;
                }
            },
        }
        edits.sort_by_key(|(range, _)| std::cmp::Reverse(range.start));
        let mut after = before.as_ref().clone();
        for (range, bytes) in edits {
            after.splice(range, bytes);
        }
        if after.as_slice() == before.as_slice() {
            return Ok(false);
        }
        self.replace_element(paragraph, &before, &Arc::new(after))?;
        Ok(true)
    }

    /// Remove every SmartArt drawing of the main story whose `dgm:relIds`
    /// names the data relationship `data_relationship_id` (record 0335),
    /// with its run when the run holds nothing else. Returns how many were
    /// removed.
    pub(crate) fn remove_smartart_drawings(&mut self, data_relationship_id: &str) -> Result<usize> {
        self.ensure_unmanaged("remove_smartart")?;
        let mut removed = 0;
        let families: [(usize, fn(Position) -> ElementStep); 3] = [
            (self.projected.paragraphs.len(), ElementStep::Paragraph),
            (self.projected.tables.len(), ElementStep::Table),
            (
                self.projected.block_controls.len(),
                ElementStep::ContentControl,
            ),
        ];
        for (count, step) in families {
            for index in 0..count {
                let path = ElementPath::new([step(Position::new(index))]);
                let located = locate(&self.projected, path.steps())?;
                let before = self.element_bytes(located)?;
                let spans =
                    crate::picture_spans::diagram_spans(&before).map_err(Error::Document)?;
                let mut ranges: Vec<std::ops::Range<usize>> = spans
                    .into_iter()
                    .filter(|span| span.data_relationship_id == data_relationship_id)
                    .map(|span| span.sole_run.unwrap_or(span.drawing))
                    .collect();
                if ranges.is_empty() {
                    continue;
                }
                removed += ranges.len();
                ranges.sort_by_key(|range| std::cmp::Reverse(range.start));
                let mut after = before.as_ref().clone();
                for range in ranges {
                    after.drain(range);
                }
                self.replace_element(&path, &before, &Arc::new(after))?;
            }
        }
        Ok(removed)
    }

    /// The largest drawing object identifier (`wp:docPr/@id`) of the main
    /// document, if any.
    pub(crate) fn max_drawing_id(&self) -> Result<Option<u32>> {
        let xml = self.projected.xml_bytes();
        let mut reader = quick_xml::Reader::from_reader(xml);
        let mut maximum = None;
        loop {
            match reader
                .read_event()
                .map_err(|_| Error::Document(crate::Error::Xml("invalid Word XML".into())))?
            {
                Event::Start(element) | Event::Empty(element)
                    if element.local_name().as_ref() == b"docPr" =>
                {
                    let id = element
                        .unchecked_attributes()
                        .filter_map(std::result::Result::ok)
                        .find(|attribute| attribute.key.as_ref() == b"id")
                        .and_then(|attribute| {
                            std::str::from_utf8(&attribute.value)
                                .ok()?
                                .parse::<u32>()
                                .ok()
                        });
                    maximum = maximum.max(id);
                },
                Event::Eof => return Ok(maximum),
                _ => {},
            }
        }
    }

    /// Insert a hyperlink holding `new_run` into the paragraph at
    /// `paragraph`, before its direct run `at`, and return its position
    /// among the paragraph's direct hyperlinks. Exactly one of
    /// `relationship` (an `r:id`) and `anchor` (a `w:anchor`) is given; the
    /// relationship itself is the caller's to stage.
    pub(crate) fn insert_hyperlink_markup_at(
        &mut self,
        paragraph: &ElementPath,
        at: Option<Position>,
        relationship: Option<&str>,
        anchor: Option<&str>,
        new_run: &NewRun,
    ) -> Result<Position> {
        self.ensure_unmanaged("insert_hyperlink_at")?;
        if paragraph.steps().last().map(|step| step.local()) != Some(b"p".as_slice()) {
            return Err(not_found());
        }
        let conformance = self.projected.conformance;
        self.insert_inline(paragraph, at, b"hyperlink", |prefix| {
            hyperlink_markup(prefix, conformance, relationship, anchor, new_run)
        })
    }

    /// Point the hyperlink at `hyperlink` to the relationship `relationship`
    /// or the bookmark `anchor`, keeping its other attributes and its runs.
    pub(crate) fn retarget_hyperlink_at(
        &mut self,
        hyperlink: &ElementPath,
        relationship: Option<&str>,
        anchor: Option<&str>,
    ) -> Result<&mut Self> {
        self.ensure_unmanaged("set_hyperlink_target_at")?;
        let (located, inherited) = self.element_with_scope(hyperlink, b"hyperlink")?;
        let before = self.element_bytes(located)?;
        let after = retarget_markup(
            &before,
            &inherited,
            self.projected.conformance,
            relationship,
            anchor,
        )?;
        if after == before.as_slice() {
            return Ok(self);
        }
        self.replace_element(hyperlink, &before, &Arc::new(after))
    }

    /// Set the screen tip (`w:tooltip`) of the hyperlink at `hyperlink` to
    /// `tooltip`, or remove it for `None` (record 0345). The start tag keeps
    /// every other attribute as written; an unchanged tip is an exact no-op.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::document::{raw, ElementPath, ElementStep};
    ///
    /// let mut edit = raw::Tree::from_xml(
    ///     br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:hyperlink w:anchor="Intro"><w:r><w:t>Intro</w:t></w:r></w:hyperlink></w:p></w:body></w:document>"#
    ///         .to_vec(),
    ///     raw::Limits::default(),
    /// )?;
    /// let link = ElementPath::new([
    ///     ElementStep::Paragraph(Position::new(0)),
    ///     ElementStep::Hyperlink(Position::new(0)),
    /// ]);
    /// edit.set_hyperlink_tooltip_at(&link, Some("Jump to the introduction"))?;
    /// assert!(String::from_utf8_lossy(&edit.to_bytes(raw::Limits::default())?)
    ///     .contains(r#"w:tooltip="Jump to the introduction""#));
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for a missing hyperlink,
    /// [`Refusal::InvalidPropertyValue`] for a tip holding a control
    /// character XML cannot carry, or a resource-limit error.
    pub fn set_hyperlink_tooltip_at(
        &mut self,
        hyperlink: &ElementPath,
        tooltip: Option<&str>,
    ) -> Result<&mut Self> {
        self.ensure_unmanaged("set_hyperlink_tooltip_at")?;
        if tooltip.is_some_and(|tip| {
            tip.chars()
                .any(|c| c.is_control() && !matches!(c, '\t' | '\n' | '\r'))
        }) {
            return Err(refusal(Refusal::InvalidPropertyValue));
        }
        let (located, _) = self.element_with_scope(hyperlink, b"hyperlink")?;
        let before = self.element_bytes(located)?;
        let (tag_end, _) = super::root_tag(&before)?;
        let tag = &before[..tag_end];
        let key = crate::picture_spans::attribute_named(tag, "tooltip").unwrap_or_else(|| {
            let name = String::from_utf8_lossy(super::tag_name(tag).unwrap_or(b"w:hyperlink"));
            match name.split_once(':') {
                Some((prefix, _)) => format!("{prefix}:tooltip"),
                None => "tooltip".to_owned(),
            }
        });
        let rewritten = match tooltip {
            Some(tip) => crate::picture_spans::with_attribute(tag, &key, tip),
            None => crate::picture_spans::without_attribute(tag, &key),
        }
        .ok_or_else(|| invalid("invalid hyperlink start tag"))?;
        if rewritten.as_slice() == tag {
            return Ok(self);
        }
        let mut after = buffer(staging_add(before.len(), rewritten.len())?)?;
        after.extend_from_slice(&rewritten);
        after.extend_from_slice(&before[tag_end..]);
        self.replace_element(hyperlink, &before, &Arc::new(after))
    }

    /// Anchor comment `id` in the paragraph at `paragraph`: a range around
    /// its direct runs `range` (inclusive), or around all its content for
    /// `None`, followed by the comment's reference run.
    ///
    /// Record 0345, after OfficeCLI's `AddComment`: for a whole-paragraph
    /// anchor, `start_after` starts the range after that direct run (a run
    /// the paragraph lacks keeps the paragraph start); `reference_only`
    /// writes the reference run alone, with no range markers, at the end
    /// of the range (after `start_after` for a whole paragraph).
    pub(crate) fn anchor_comment_at(
        &mut self,
        paragraph: &ElementPath,
        range: Option<(Position, Position)>,
        id: u32,
        start_after: Option<Position>,
        reference_only: bool,
    ) -> Result<&mut Self> {
        self.ensure_unmanaged("add_comment_at")?;
        let (located, inherited) = self.element_with_scope(paragraph, b"p")?;
        let before = self.element_bytes(located)?;
        let base = match super::root_tag(&before)? {
            (_, true) => {
                let mut open = opened(&before)?;
                open.extend_from_slice(b"</");
                open.extend_from_slice(super::tag_name(&before)?);
                open.push(b'>');
                open
            },
            (_, false) => before.to_vec(),
        };
        let layout = element_children(&base, b"p", &inherited)?;
        let (start_child, end_child) = match range {
            Some((first, last)) => {
                if last < first {
                    return Err(refusal(Refusal::InvalidPropertyValue));
                }
                let start = layout
                    .index_of(b"r", first)
                    .ok_or_else(|| refusal(Refusal::RunNotFound))?;
                let end = layout
                    .index_of(b"r", last)
                    .ok_or_else(|| refusal(Refusal::RunNotFound))?;
                (start, end + 1)
            },
            None => {
                let after = start_after
                    .and_then(|run| layout.index_of(b"r", run))
                    .map(|child| child + 1);
                let start = after.unwrap_or_else(|| {
                    layout
                        .children
                        .iter()
                        .position(|child| child.2.as_deref() != Some(b"pPr"))
                        .unwrap_or(layout.children.len())
                });
                // A reference-only comment sits where the range would
                // start when a run names the place, else at the end.
                match (reference_only, after) {
                    (true, Some(after)) => (after, after),
                    _ => (start, layout.children.len()),
                }
            },
        };
        let start = layout
            .insertion_offset(start_child)
            .ok_or_else(|| refusal(Refusal::RunNotFound))?;
        let end = layout
            .insertion_offset(end_child)
            .ok_or_else(|| refusal(Refusal::RunNotFound))?;
        let prefix = layout.prefix.clone().unwrap_or_default();
        let (prefix, declare) = if prefix.is_empty() {
            (b"w".to_vec(), Some(self.projected.conformance.namespace()))
        } else {
            (prefix, None)
        };
        let marker = |local: &str| {
            let mut xml = Vec::new();
            open(&mut xml, &prefix, local);
            if let Some(namespace) = declare {
                xml.extend_from_slice(b" xmlns:");
                xml.extend_from_slice(&prefix);
                xml.extend_from_slice(b"=\"");
                xml.extend_from_slice(namespace.as_bytes());
                xml.push(b'"');
            }
            attribute(&mut xml, &prefix, "id", &id.to_string());
            xml.extend_from_slice(b"/>");
            xml
        };
        let mut reference = Vec::new();
        open(&mut reference, &prefix, "r");
        if let Some(namespace) = declare {
            reference.extend_from_slice(b" xmlns:");
            reference.extend_from_slice(&prefix);
            reference.extend_from_slice(b"=\"");
            reference.extend_from_slice(namespace.as_bytes());
            reference.push(b'"');
        }
        reference.push(b'>');
        open(&mut reference, &prefix, "rPr");
        reference.push(b'>');
        open(&mut reference, &prefix, "rStyle");
        attribute(&mut reference, &prefix, "val", "CommentReference");
        reference.extend_from_slice(b"/>");
        close(&mut reference, &prefix, "rPr");
        open(&mut reference, &prefix, "commentReference");
        attribute(&mut reference, &prefix, "id", &id.to_string());
        reference.extend_from_slice(b"/>");
        close(&mut reference, &prefix, "r");
        let (start_marker, end_marker) = if reference_only {
            (Vec::new(), Vec::new())
        } else {
            (marker("commentRangeStart"), marker("commentRangeEnd"))
        };
        let mut after = buffer(staging_add(
            base.len(),
            start_marker.len() + end_marker.len() + reference.len(),
        )?)?;
        after.extend_from_slice(&base[..start]);
        after.extend_from_slice(&start_marker);
        after.extend_from_slice(&base[start..end]);
        after.extend_from_slice(&end_marker);
        after.extend_from_slice(&reference);
        after.extend_from_slice(&base[end..]);
        self.replace_element(paragraph, &before, &Arc::new(after))
    }

    /// The exact bytes of the body-final section properties, with the
    /// namespace bindings in scope there, or `None` when the body has none
    /// (record 0319).
    pub(crate) fn body_section(
        &self,
    ) -> Result<Option<(Arc<Vec<u8>>, crate::part_scan::Bindings)>> {
        let path = [ElementStep::SectionProperties(Position::new(0))];
        let located = match locate(&self.projected, &path) {
            Ok(located) => located,
            Err(_) => return Ok(None),
        };
        let inherited = bindings_at(self.projected.xml_bytes(), located.start)?;
        Ok(Some((self.element_bytes(located)?, inherited)))
    }

    /// Replace the body-final section properties' exact bytes `before` with
    /// `after` (record 0319).
    pub(crate) fn replace_body_section(
        &mut self,
        before: &Arc<Vec<u8>>,
        after: &Arc<Vec<u8>>,
    ) -> Result<&mut Self> {
        self.replace_element(
            &ElementPath::new([ElementStep::SectionProperties(Position::new(0))]),
            before,
            after,
        )
    }

    /// Append body-final section properties to a body that has none (record
    /// 0319).
    pub(crate) fn append_body_section(&mut self, xml: &Arc<Vec<u8>>) -> Result<&mut Self> {
        let (children, _) = container_children(&self.projected, &ElementPath::body())?;
        if children.section.is_some() {
            return Err(Error::SemanticPrecondition);
        }
        let at = Position::new(children.children.len());
        self.insert_element(&ElementPath::body(), at, xml)
    }

    /// The projected main-document XML.
    pub(crate) fn main_xml(&self) -> &[u8] {
        self.projected.xml_bytes()
    }

    /// Insert one inline run's exact markup into the paragraph at
    /// `paragraph` before its direct run `at` (after its last child for
    /// `None`), and return its position among the paragraph's direct runs
    /// (record 0319).
    pub(crate) fn insert_run_markup_at(
        &mut self,
        paragraph: &ElementPath,
        at: Option<Position>,
        markup: impl FnOnce(&[u8]) -> Result<Vec<u8>>,
    ) -> Result<Position> {
        self.insert_inline(paragraph, at, b"r", markup)
    }

    /// Remove every direct run of every paragraph of the main story that
    /// holds a note reference (`footnoteReference` or `endnoteReference`,
    /// by `local`) to note `id`. Returns whether anything changed (record
    /// 0319).
    pub(crate) fn strip_note_references(&mut self, local: &[u8], id: i64) -> Result<bool> {
        self.ensure_unmanaged("remove_note")?;
        let mut changed = false;
        let families: [(usize, fn(Position) -> ElementStep); 3] = [
            (self.projected.paragraphs.len(), ElementStep::Paragraph),
            (self.projected.tables.len(), ElementStep::Table),
            (
                self.projected.block_controls.len(),
                ElementStep::ContentControl,
            ),
        ];
        for (count, step) in families {
            for index in 0..count {
                let path = ElementPath::new([step(Position::new(index))]);
                let located = locate(&self.projected, path.steps())?;
                let before = self.element_bytes(located)?;
                if let Some(after) = without_note_references(&before, local, id)? {
                    self.replace_element(&path, &before, &Arc::new(after))?;
                    changed = true;
                }
            }
        }
        Ok(changed)
    }

    /// The largest comment identifier the main document's range markers and
    /// references use.
    /// Every `paraId` value of the projected main document, for choosing
    /// a fresh paragraph identifier (record 0345).
    pub(crate) fn paragraph_ids(&self) -> Vec<u32> {
        crate::part_scan::paragraph_ids(self.projected.xml_bytes())
    }

    pub(crate) fn max_comment_marker_id(&self) -> Result<Option<u32>> {
        let mut maximum = None;
        scan_comment_markers(self.projected.xml_bytes(), |_, id, _| {
            maximum = maximum.max(Some(id));
        })?;
        Ok(maximum)
    }

    /// Remove every range marker of comment `id`, and every run holding its
    /// reference, from the main document. Returns whether anything changed.
    pub(crate) fn strip_comment_markers(&mut self, id: u32) -> Result<bool> {
        self.ensure_unmanaged("remove_comment")?;
        let mut changed = false;
        let families: [(usize, fn(Position) -> ElementStep); 3] = [
            (self.projected.paragraphs.len(), ElementStep::Paragraph),
            (self.projected.tables.len(), ElementStep::Table),
            (
                self.projected.block_controls.len(),
                ElementStep::ContentControl,
            ),
        ];
        for (count, step) in families {
            for index in 0..count {
                let path = ElementPath::new([step(Position::new(index))]);
                let located = locate(&self.projected, path.steps())?;
                let before = self.element_bytes(located)?;
                if let Some(after) = without_comment_markers(&before, id)? {
                    self.replace_element(&path, &before, &Arc::new(after))?;
                    changed = true;
                }
            }
        }
        Ok(changed)
    }

    /// Insert inline markup (a run or a hyperlink) into a paragraph or
    /// hyperlink before its direct run `at` (after its last child for
    /// `None`), and return its position among the owner's direct `counted`
    /// children.
    fn insert_inline(
        &mut self,
        owner: &ElementPath,
        at: Option<Position>,
        counted: &[u8],
        markup: impl FnOnce(&[u8]) -> Result<Vec<u8>>,
    ) -> Result<Position> {
        self.insert_inline_counted(owner, at, counted, false, markup)
            .map(|(position, _)| position)
    }

    /// As [`Self::insert_inline`]; with `images`, also return how many
    /// inline pictures of the owner precede the insertion point.
    fn insert_inline_counted(
        &mut self,
        owner: &ElementPath,
        at: Option<Position>,
        counted: &[u8],
        images: bool,
        markup: impl FnOnce(&[u8]) -> Result<Vec<u8>>,
    ) -> Result<(Position, usize)> {
        let local = owner
            .steps()
            .last()
            .map(|step| step.local())
            .filter(|local| matches!(*local, b"p" | b"hyperlink"))
            .ok_or_else(not_found)?;
        let located = locate(&self.projected, owner.steps())?;
        let inherited = bindings_at(self.projected.xml_bytes(), located.start)?;
        let before = self.element_bytes(located)?;
        // An empty owner (`<w:p/>`) is opened first.
        let base = match super::root_tag(&before)? {
            (_, true) => {
                let mut open = opened(&before)?;
                open.extend_from_slice(b"</");
                open.extend_from_slice(super::tag_name(&before)?);
                open.push(b'>');
                open
            },
            (_, false) => before.to_vec(),
        };
        let layout = element_children(&base, local, &inherited)?;
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
        if local == b"p" && inside_complex_field(&base, offset, &inherited).map_err(refusal)? {
            return Err(refusal(Refusal::FieldInterior));
        }
        let prefix = layout.prefix.clone().unwrap_or_default();
        let xml = markup(&prefix)?;
        let mut after = buffer(staging_add(base.len(), xml.len())?)?;
        after.extend_from_slice(&base[..offset]);
        after.extend_from_slice(&xml);
        after.extend_from_slice(&base[offset..]);
        let position = layout.count_before(counted, child);
        let preceding = if images {
            crate::image::parse_inline_images(&base[..offset])
                .map_err(Error::Document)?
                .len()
        } else {
            0
        };
        self.replace_element(owner, &before, &Arc::new(after))?;
        Ok((Position::new(position), preceding))
    }

    /// Insert a paragraph into the body, a table cell or a block content
    /// control at `anchor` (in the body, `End` is before the body-final
    /// section properties), and return its position among the container's
    /// direct paragraphs.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::document::{
    ///     raw, ElementAnchor, ElementPath, ElementStep, NewParagraph, NewRun,
    /// };
    ///
    /// let mut edit = raw::Tree::from_xml(
    ///     br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:body></w:document>"#
    ///         .to_vec(),
    ///     raw::Limits::default(),
    /// )?;
    /// let cell = ElementPath::new([
    ///     ElementStep::Table(Position::new(0)),
    ///     ElementStep::Row(Position::new(0)),
    ///     ElementStep::Cell(Position::new(0)),
    /// ]);
    /// let position = edit.insert_paragraph_at(
    ///     &cell,
    ///     ElementAnchor::End,
    ///     &NewParagraph::new().with_run(NewRun::new("second")),
    /// )?;
    /// assert_eq!(position, Position::new(1));
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for a missing container or sibling,
    /// a typed refusal for the paragraph's text or properties, or a
    /// resource-limit error.
    pub fn insert_paragraph_at(
        &mut self,
        container: &ElementPath,
        anchor: ElementAnchor,
        paragraph: &NewParagraph,
    ) -> Result<Position> {
        if container.steps().is_empty() {
            let (children, _) = container_children(&self.projected, container)?;
            return self.insert_formatted_paragraph(body_anchor(&children, anchor)?, paragraph);
        }
        self.ensure_unmanaged("insert_paragraph_at")?;
        let (children, local) = container_children(&self.projected, container)?;
        if !ElementStep::Paragraph(Position::new(0)).fits(local) {
            return Err(not_found());
        }
        let child = child_index(&children, local, anchor)?;
        let strict = matches!(self.projected.conformance, Conformance::Strict);
        let xml = paragraph_markup(
            children.prefix.as_deref(),
            self.projected.conformance,
            paragraph,
            strict,
        )
        .map_err(spliced(Position::new(0)))?;
        let position = children.count_before(b"p", child);
        self.insert_element(container, Position::new(child), &Arc::new(xml))?;
        Ok(Position::new(position))
    }

    /// Insert a table into the body, a table cell or a block content control
    /// at `anchor`, and return its position among the container's direct
    /// tables.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::document::{raw, ElementAnchor, ElementPath, NewTable};
    ///
    /// let mut edit = raw::Tree::from_xml(
    ///     br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/><w:sectPr/></w:body></w:document>"#
    ///         .to_vec(),
    ///     raw::Limits::default(),
    /// )?;
    /// let table = NewTable::new(2, 2)?.with_cell_text(1, 1, "d");
    /// assert_eq!(edit.insert_table_at(&ElementPath::body(), ElementAnchor::End, &table)?, Position::new(0));
    /// assert_eq!(edit.projected().table_count(), 1);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for a missing container or sibling,
    /// a typed refusal for cell text, or a resource-limit error.
    pub fn insert_table_at(
        &mut self,
        container: &ElementPath,
        anchor: ElementAnchor,
        table: &NewTable,
    ) -> Result<Position> {
        self.ensure_unmanaged("insert_table_at")?;
        let (children, local) = container_children(&self.projected, container)?;
        if !ElementStep::Table(Position::new(0)).fits(local) {
            return Err(not_found());
        }
        let child = child_index(&children, local, anchor)?;
        let prefix = children.prefix.clone();
        let xml = self.table_markup(prefix.as_deref(), table)?;
        let position = children.count_before(b"tbl", child);
        self.insert_element(container, Position::new(child), &Arc::new(xml))?;
        Ok(Position::new(position))
    }

    /// Insert a row of `cells` empty-paragraph cells (holding `texts`, in
    /// order, where given) into the table at `table`, before its row `at`
    /// (or after its last row), and return the new row's position.
    ///
    /// When `cells` is `None` the row has one cell per grid column of the
    /// table, with the grid's widths.
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::BlockNotFound`] for a missing table or row, a typed
    /// refusal for cell text, or a resource-limit error.
    pub fn insert_row_at(
        &mut self,
        table: &ElementPath,
        at: Option<Position>,
        cells: Option<usize>,
        texts: &[String],
    ) -> Result<Position> {
        self.ensure_unmanaged("insert_row_at")?;
        if !matches!(table.steps().last(), Some(ElementStep::Table(_))) {
            return Err(not_found());
        }
        let (children, _) = container_children(&self.projected, table)?;
        // Row positions count direct rows; a row-level wrapper would make
        // them differ from what readers count.
        if children
            .children
            .iter()
            .any(|child| matches!(child.2.as_deref(), Some(b"sdt" | b"customXml")))
        {
            return Err(refusal(Refusal::ComplexContent));
        }
        let rows = children.count(b"tr");
        let after_last = children
            .children
            .iter()
            .rposition(|child| child.2.as_deref() == Some(b"tr"))
            .map_or(children.children.len(), |index| index + 1);
        let child = match at {
            None => after_last,
            Some(position) if position.get() == rows => after_last,
            Some(position) => children.index_of(b"tr", position).ok_or_else(not_found)?,
        };
        let widths = grid_widths(self.projected.xml_bytes(), &children)?;
        let count = cells.unwrap_or(widths.len()).max(1);
        if count > MAX_NEW_TABLE_DIMENSION {
            return Err(Error::Limit {
                kind: crate::document::LimitKind::TableColumns,
                observed: count,
                maximum: MAX_NEW_TABLE_DIMENSION,
            });
        }
        let prefix = children.prefix.clone().filter(|prefix| !prefix.is_empty());
        let declare = prefix
            .is_none()
            .then(|| self.projected.conformance.namespace().as_bytes());
        let prefix = prefix.unwrap_or_else(|| b"w".to_vec());
        let mut row = Vec::new();
        self.row_markup(&mut row, &prefix, declare, count, |index| {
            (
                widths.get(index).copied().flatten(),
                texts.get(index).map_or("", String::as_str),
            )
        })?;
        let position = children.count_before(b"tr", child);
        self.insert_element(table, Position::new(child), &Arc::new(row))?;
        Ok(Position::new(position))
    }

    /// Remove the element at `path`: a paragraph, run, hyperlink, table,
    /// row, cell or content control.
    ///
    /// The removal is refused when the element holds one end of a range
    /// whose other end is outside it, a story reference, a field character
    /// or a section break, and when it is the last paragraph of a cell, the
    /// last cell of a row or the last row of a table (remove the container
    /// instead).
    ///
    /// # Errors
    ///
    /// Returns one of the refusals above, [`Refusal::BlockNotFound`] for a
    /// missing element, or a resource-limit error.
    pub fn remove_element_at(&mut self, path: &ElementPath) -> Result<&mut Self> {
        if let Some(position) = path.body_paragraph() {
            return self.remove_paragraph(position);
        }
        if let Some((paragraph, run)) = path.body_run() {
            return self.remove_run(paragraph, run);
        }
        self.ensure_unmanaged("remove_element_at")?;
        let (parent, step) = path.split_last().ok_or_else(not_found)?;
        let located = locate(&self.projected, path.steps())?;
        let inherited = bindings_at(self.projected.xml_bytes(), located.start)?;
        let xml = self.element_bytes(located)?;
        match step {
            ElementStep::Run(_) => run_dependencies(&xml, &inherited).map_err(refusal)?,
            _ => paragraph_dependencies(&xml, &inherited)
                .refuse(Removal::Delete)
                .map_err(refusal)?,
        }
        let (children, local) = container_children(&self.projected, &parent)?;
        let child = children
            .index_of(step.local(), step.position())
            .ok_or_else(not_found)?;
        let last_required = match (local, step) {
            (b"tc", ElementStep::Paragraph(_)) => children.count(b"p") == 1,
            (b"tr", ElementStep::Cell(_)) => children.count(b"tc") == 1,
            (b"tbl", ElementStep::Row(_)) => children.count(b"tr") == 1,
            _ => false,
        };
        if last_required {
            return Err(refusal(Refusal::ComplexContent));
        }
        self.remove_element(&parent, Position::new(child), &xml)
    }

    /// Move the element at `path` into `container` at `anchor` (evaluated
    /// before the move), and return its new path.
    ///
    /// The element's bytes move unchanged, so references and ranges travel
    /// with it; it is refused for an element holding one end of a range
    /// whose other end is outside it.
    ///
    /// # Errors
    ///
    /// As [`Self::remove_element_at`], and [`Refusal::BlockNotFound`] for a
    /// destination that cannot hold the element.
    pub fn move_element_at(
        &mut self,
        path: &ElementPath,
        container: &ElementPath,
        anchor: ElementAnchor,
    ) -> Result<ElementPath> {
        self.ensure_unmanaged("move_element_at")?;
        let (parent, step) = path.split_last().ok_or_else(not_found)?;
        if container.steps().starts_with(path.steps()) {
            return Err(refusal(Refusal::ComplexContent));
        }
        let located = locate(&self.projected, path.steps())?;
        let inherited = bindings_at(self.projected.xml_bytes(), located.start)?;
        let xml = self.element_bytes(located)?;
        match step {
            ElementStep::Run(_) => run_dependencies(&xml, &inherited).map_err(refusal)?,
            _ => paragraph_dependencies(&xml, &inherited)
                .refuse(Removal::Move)
                .map_err(refusal)?,
        }
        let (source_children, source_local) = container_children(&self.projected, &parent)?;
        let source_child = source_children
            .index_of(step.local(), step.position())
            .ok_or_else(not_found)?;
        let last_required = match (source_local, step) {
            (b"tc", ElementStep::Paragraph(_)) => source_children.count(b"p") == 1,
            (b"tr", ElementStep::Cell(_)) => source_children.count(b"tc") == 1,
            (b"tbl", ElementStep::Row(_)) => source_children.count(b"tr") == 1,
            _ => false,
        };
        let (target_children, target_local) = container_children(&self.projected, container)?;
        if !step.fits(target_local) {
            return Err(not_found());
        }
        let same = parent == *container;
        if last_required && !same {
            return Err(refusal(Refusal::ComplexContent));
        }
        let mut target_child = child_index(&target_children, target_local, anchor)?;
        if same && target_child > source_child {
            target_child -= 1;
        }
        if same && target_child == source_child {
            return Ok(path.clone());
        }
        let mut candidate = self.try_clone()?;
        candidate.remove_element(&parent, Position::new(source_child), &xml)?;
        // The destination path's own positions may shift when the source
        // preceded one of its steps in the same container; re-derive it.
        let destination = shifted_after_removal(container, &parent, step);
        let (moved_children, _) = container_children(&candidate.projected, &destination)?;
        let position = moved_children.count_before(step.local(), target_child);
        candidate.insert_element(&destination, Position::new(target_child), &xml)?;
        *self = candidate;
        Ok(destination.child(step.with_position(Position::new(position))))
    }

    // -----------------------------------------------------------------
    // Table markup
    // -----------------------------------------------------------------

    fn table_markup(&self, prefix: Option<&[u8]>, table: &NewTable) -> Result<Vec<u8>> {
        // An unprefixed (default-namespace) container still needs a bound
        // prefix for the WordprocessingML attributes.
        let prefix = prefix.filter(|prefix| !prefix.is_empty());
        let declare = prefix.is_none();
        let prefix = prefix.unwrap_or(b"w");
        let namespace = self.projected.conformance.namespace().as_bytes();
        let mut xml = Vec::new();
        open(&mut xml, prefix, "tbl");
        if declare {
            xml.extend_from_slice(b" xmlns:");
            xml.extend_from_slice(prefix);
            xml.extend_from_slice(b"=\"");
            xml.extend_from_slice(namespace);
            xml.push(b'"');
        }
        xml.push(b'>');
        open(&mut xml, prefix, "tblPr");
        xml.push(b'>');
        if let Some(style) = &table.style {
            if style.is_empty() || style.len() > 253 {
                return Err(refusal(Refusal::InvalidPropertyValue));
            }
            open(&mut xml, prefix, "tblStyle");
            xml.push(b' ');
            qualified_name(&mut xml, prefix, "val");
            xml.extend_from_slice(b"=\"");
            super::super::append_literal_escaped_xml(&mut xml, style);
            xml.extend_from_slice(b"\"/>");
        }
        open(&mut xml, prefix, "tblW");
        attribute(&mut xml, prefix, "w", "0");
        attribute(&mut xml, prefix, "type", "auto");
        xml.extend_from_slice(b"/>");
        if table.borders {
            open(&mut xml, prefix, "tblBorders");
            xml.push(b'>');
            for side in ["top", "left", "bottom", "right", "insideH", "insideV"] {
                open(&mut xml, prefix, side);
                attribute(&mut xml, prefix, "val", "single");
                attribute(&mut xml, prefix, "sz", "4");
                attribute(&mut xml, prefix, "space", "0");
                attribute(&mut xml, prefix, "color", "auto");
                xml.extend_from_slice(b"/>");
            }
            close(&mut xml, prefix, "tblBorders");
        }
        close(&mut xml, prefix, "tblPr");
        open(&mut xml, prefix, "tblGrid");
        xml.push(b'>');
        for width in &table.widths {
            open(&mut xml, prefix, "gridCol");
            attribute(&mut xml, prefix, "w", &width.to_string());
            xml.extend_from_slice(b"/>");
        }
        close(&mut xml, prefix, "tblGrid");
        for row in &table.cells {
            self.row_markup(&mut xml, prefix, None, table.columns, |index| {
                (
                    table.widths.get(index).copied(),
                    row.get(index).map_or("", String::as_str),
                )
            })?;
        }
        close(&mut xml, prefix, "tbl");
        staging_add(0, xml.len())?;
        ensure_utf8(&xml)?;
        Ok(xml)
    }

    fn row_markup<'t>(
        &self,
        xml: &mut Vec<u8>,
        prefix: &[u8],
        declare: Option<&[u8]>,
        cells: usize,
        cell: impl Fn(usize) -> (Option<u32>, &'t str),
    ) -> Result<()> {
        let strict = matches!(self.projected.conformance, Conformance::Strict);
        open(xml, prefix, "tr");
        if let Some(namespace) = declare {
            xml.extend_from_slice(b" xmlns:");
            xml.extend_from_slice(prefix);
            xml.extend_from_slice(b"=\"");
            xml.extend_from_slice(namespace);
            xml.push(b'"');
        }
        xml.push(b'>');
        for index in 0..cells {
            let (width, text) = cell(index);
            validate_authored_text(text).map_err(refusal)?;
            open(xml, prefix, "tc");
            xml.push(b'>');
            if let Some(width) = width {
                open(xml, prefix, "tcPr");
                xml.push(b'>');
                open(xml, prefix, "tcW");
                attribute(xml, prefix, "w", &width.to_string());
                attribute(xml, prefix, "type", "dxa");
                xml.extend_from_slice(b"/>");
                close(xml, prefix, "tcPr");
            }
            let paragraph = if text.is_empty() {
                NewParagraph::new()
            } else {
                NewParagraph::new().with_run(NewRun::new(text))
            };
            let markup =
                paragraph_markup(Some(prefix), self.projected.conformance, &paragraph, strict)
                    .map_err(spliced(Position::new(0)))?;
            xml.extend_from_slice(&markup);
            close(xml, prefix, "tc");
            staging_add(0, xml.len())?;
        }
        close(xml, prefix, "tr");
        Ok(())
    }
}

/// The local names a container's content children have, for
/// [`ElementAnchor::Start`].
fn content_children(container: &[u8]) -> &'static [&'static [u8]] {
    match container {
        b"body" | b"tc" | b"sdtContent" => &[b"p", b"tbl", b"sdt"],
        b"tbl" => &[b"tr"],
        b"tr" => &[b"tc"],
        b"p" => &[b"r", b"hyperlink", b"sdt"],
        b"hyperlink" => &[b"r"],
        _ => &[],
    }
}

/// The child index an anchor inserts at, before any change.
fn child_index(children: &Children, container: &[u8], anchor: ElementAnchor) -> Result<usize> {
    let limit = children.insertion_limit();
    let index = match anchor {
        ElementAnchor::End => Some(limit),
        ElementAnchor::Start => {
            let content = content_children(container);
            Some(
                children
                    .children
                    .iter()
                    .position(|child| {
                        child
                            .2
                            .as_deref()
                            .is_some_and(|local| content.contains(&local))
                    })
                    .map_or(limit, |index| index.min(limit)),
            )
        },
        ElementAnchor::Before(step) => children.index_of(step.local(), step.position()),
        ElementAnchor::After(step) => children
            .index_of(step.local(), step.position())
            .map(|index| index + 1),
    };
    index.filter(|index| *index <= limit).ok_or_else(not_found)
}

/// Map a body anchor onto the body-position anchor of the direct-body
/// paragraph route.
fn body_anchor(children: &Children, anchor: ElementAnchor) -> Result<BodyAnchor> {
    let block = |step: ElementStep| match step {
        ElementStep::Paragraph(position) => Ok(BodyBlock::Paragraph(position)),
        ElementStep::Table(position) => Ok(BodyBlock::Table(position)),
        ElementStep::ContentControl(position) => Ok(BodyBlock::ContentControl(position)),
        _ => Err(not_found()),
    };
    Ok(match anchor {
        ElementAnchor::End => BodyAnchor::End,
        ElementAnchor::Before(step) => BodyAnchor::Before(block(step)?),
        ElementAnchor::After(step) => BodyAnchor::After(block(step)?),
        ElementAnchor::Start => {
            let index = child_index(children, b"body", ElementAnchor::Start)?;
            match children
                .children
                .get(index)
                .and_then(|child| child.2.as_deref())
            {
                Some(local) if index < children.insertion_limit() => {
                    let position = Position::new(children.count_before(local, index));
                    BodyAnchor::Before(block(match local {
                        b"p" => ElementStep::Paragraph(position),
                        b"tbl" => ElementStep::Table(position),
                        _ => ElementStep::ContentControl(position),
                    })?)
                },
                _ => BodyAnchor::End,
            }
        },
    })
}

/// The path `container` has after the element `removed` (a direct child of
/// `parent`) is taken out: a step of `container` in `parent` that counts the
/// same element name after it moves back by one.
fn shifted_after_removal(
    container: &ElementPath,
    parent: &ElementPath,
    removed: ElementStep,
) -> ElementPath {
    let depth = parent.steps().len();
    if !container.steps().starts_with(parent.steps()) || container.steps().len() <= depth {
        return container.clone();
    }
    let mut steps = container.steps().to_vec();
    let step = steps[depth];
    if step.local() == removed.local() && step.position() > removed.position() {
        steps[depth] = step.with_position(Position::new(step.position().get() - 1));
    }
    ElementPath::new(steps)
}

fn qualified_name(xml: &mut Vec<u8>, prefix: &[u8], local: &str) {
    if !prefix.is_empty() {
        xml.extend_from_slice(prefix);
        xml.push(b':');
    }
    xml.extend_from_slice(local.as_bytes());
}

fn open(xml: &mut Vec<u8>, prefix: &[u8], local: &str) {
    xml.push(b'<');
    qualified_name(xml, prefix, local);
}

/// A WordprocessingML attribute, qualified with the element's prefix.
fn attribute(xml: &mut Vec<u8>, prefix: &[u8], name: &str, value: &str) {
    xml.push(b' ');
    qualified_name(xml, prefix, name);
    xml.extend_from_slice(b"=\"");
    xml.extend_from_slice(value.as_bytes());
    xml.push(b'"');
}

fn close(xml: &mut Vec<u8>, prefix: &[u8], local: &str) {
    xml.extend_from_slice(b"</");
    qualified_name(xml, prefix, local);
    xml.push(b'>');
}

/// The `w:gridCol` widths of a table, in order; `None` for a column whose
/// width is absent or unreadable.
fn grid_widths(xml: &[u8], table: &Children) -> Result<Vec<Option<u32>>> {
    let Some((start, end, _)) = table
        .children
        .iter()
        .find(|child| child.2.as_deref() == Some(b"tblGrid"))
    else {
        return Ok(Vec::new());
    };
    let grid = checked_slice(xml, *start, *end, "table grid")?;
    let mut reader = quick_xml::Reader::from_reader(grid);
    let mut widths = Vec::new();
    loop {
        match reader
            .read_event()
            .map_err(|_| Error::Document(crate::Error::Xml("invalid table grid".into())))?
        {
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == b"gridCol" =>
            {
                let width = element
                    .unchecked_attributes()
                    .filter_map(std::result::Result::ok)
                    .find(|attribute| attribute.key.local_name().as_ref() == b"w")
                    .and_then(|attribute| std::str::from_utf8(&attribute.value).ok()?.parse().ok());
                widths.push(width);
            },
            Event::Eof => return Ok(widths),
            _ => {},
        }
    }
}

const RELATIONSHIPS_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS_NAMESPACE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships";

/// The rewrite of an existing picture (record 0335): each field, when
/// present, replaces the drawing's extent (in EMUs, also the shape's own
/// extent), its name, its description or the relationship of its image.
#[derive(Default)]
pub(crate) struct PictureRewrite<'a> {
    pub(crate) extent: Option<(i64, i64)>,
    pub(crate) name: Option<&'a str>,
    pub(crate) description: Option<&'a str>,
    pub(crate) embed: Option<&'a str>,
}

/// The drawing of a new inline picture: its extent in EMUs, its drawing
/// object identifier, and its name and description.
pub(crate) struct PictureMarkup<'a> {
    pub(crate) width: i64,
    pub(crate) height: i64,
    pub(crate) id: u32,
    pub(crate) name: &'a str,
    pub(crate) description: &'a str,
}

/// The markup of a run holding an inline picture. Every drawing namespace,
/// and the relationships namespace of `r:embed`, is declared where it is
/// used, so the run resolves whatever the document root binds.
fn picture_markup(
    prefix: &[u8],
    conformance: Conformance,
    relationship: &str,
    picture: &PictureMarkup<'_>,
) -> Result<Vec<u8>> {
    let (prefix, declare) = if prefix.is_empty() {
        (b"w".as_slice(), true)
    } else {
        (prefix, false)
    };
    let (wordprocessing, main, pic) = match conformance {
        Conformance::Strict => (
            "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing",
            "http://purl.oclc.org/ooxml/drawingml/main",
            "http://purl.oclc.org/ooxml/drawingml/picture",
        ),
        Conformance::Transitional => (
            "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
            "http://schemas.openxmlformats.org/drawingml/2006/main",
            "http://schemas.openxmlformats.org/drawingml/2006/picture",
        ),
    };
    let mut name = Vec::new();
    super::super::append_literal_escaped_xml(&mut name, picture.name);
    let name = String::from_utf8(name).map_err(|_| refusal(Refusal::InvalidPropertyValue))?;
    let mut description = Vec::new();
    super::super::append_literal_escaped_xml(&mut description, picture.description);
    let description =
        String::from_utf8(description).map_err(|_| refusal(Refusal::InvalidPropertyValue))?;
    let mut relationship_id = Vec::new();
    super::super::append_literal_escaped_xml(&mut relationship_id, relationship);
    let relationship_id =
        String::from_utf8(relationship_id).map_err(|_| refusal(Refusal::InvalidPropertyValue))?;
    let (width, height, id) = (picture.width, picture.height, picture.id);
    let relationships = relationships_namespace(conformance);
    let drawing = format!(
        r#"<wp:inline xmlns:wp="{wordprocessing}" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{width}" cy="{height}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{id}" name="{name}" descr="{description}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="{main}" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="{main}"><a:graphicData uri="{pic}"><pic:pic xmlns:pic="{pic}"><pic:nvPicPr><pic:cNvPr id="0" name="{name}" descr="{description}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip xmlns:r="{relationships}" r:embed="{relationship_id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{width}" cy="{height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline>"#
    );
    let mut xml = buffer(staging_add(drawing.len(), 128)?)?;
    open(&mut xml, prefix, "r");
    if declare {
        xml.extend_from_slice(b" xmlns:");
        xml.extend_from_slice(prefix);
        xml.extend_from_slice(b"=\"");
        xml.extend_from_slice(conformance.namespace().as_bytes());
        xml.push(b'"');
    }
    xml.push(b'>');
    open(&mut xml, prefix, "drawing");
    xml.push(b'>');
    xml.extend_from_slice(drawing.as_bytes());
    close(&mut xml, prefix, "drawing");
    close(&mut xml, prefix, "r");
    Ok(xml)
}

/// The drawing of a new inline chart: its extent in EMUs, its drawing
/// object identifier, its name and description, and whether it is an
/// extended (`cx:chart`) one.
pub(crate) struct ChartMarkup<'a> {
    pub(crate) width: i64,
    pub(crate) height: i64,
    pub(crate) id: u32,
    pub(crate) name: &'a str,
    pub(crate) description: &'a str,
    pub(crate) extended: bool,
}

/// The markup of a run holding an inline chart, declaring every namespace
/// it uses where it is used.
fn chart_markup(
    prefix: &[u8],
    conformance: Conformance,
    relationship: &str,
    chart: &ChartMarkup<'_>,
) -> Result<Vec<u8>> {
    let (prefix, declare) = if prefix.is_empty() {
        (b"w".as_slice(), true)
    } else {
        (prefix, false)
    };
    let (wordprocessing, main, classic) = match conformance {
        Conformance::Strict => (
            "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing",
            "http://purl.oclc.org/ooxml/drawingml/main",
            crate::chart::inline::STRICT_CHART_URI,
        ),
        Conformance::Transitional => (
            "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
            "http://schemas.openxmlformats.org/drawingml/2006/main",
            crate::chart::inline::CHART_URI,
        ),
    };
    let (uri, element) = if chart.extended {
        (litchi_drawingml::chart::chartex::GRAPHIC_DATA_URI, "cx")
    } else {
        (classic, "c")
    };
    let escaped = |text: &str| -> Result<String> {
        let mut out = Vec::new();
        super::super::append_literal_escaped_xml(&mut out, text);
        String::from_utf8(out).map_err(|_| refusal(Refusal::InvalidPropertyValue))
    };
    let name = escaped(chart.name)?;
    let description = escaped(chart.description)?;
    let relationship_id = escaped(relationship)?;
    let relationships = relationships_namespace(conformance);
    let (width, height, id) = (chart.width, chart.height, chart.id);
    let drawing = format!(
        r#"<wp:inline xmlns:wp="{wordprocessing}" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{width}" cy="{height}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{id}" name="{name}" descr="{description}"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="{main}"><a:graphicData uri="{uri}"><{element}:chart xmlns:{element}="{uri}" xmlns:r="{relationships}" r:id="{relationship_id}"/></a:graphicData></a:graphic></wp:inline>"#
    );
    let mut xml = buffer(staging_add(drawing.len(), 128)?)?;
    open(&mut xml, prefix, "r");
    if declare {
        xml.extend_from_slice(b" xmlns:");
        xml.extend_from_slice(prefix);
        xml.extend_from_slice(b"=\"");
        xml.extend_from_slice(conformance.namespace().as_bytes());
        xml.push(b'"');
    }
    xml.push(b'>');
    open(&mut xml, prefix, "drawing");
    xml.push(b'>');
    xml.extend_from_slice(drawing.as_bytes());
    close(&mut xml, prefix, "drawing");
    close(&mut xml, prefix, "r");
    Ok(xml)
}

/// The drawing of a new inline SmartArt graphic (record 0295): its extent
/// in EMUs, its drawing object identifier, name and description.
pub(crate) struct SmartArtMarkup<'a> {
    pub(crate) width: i64,
    pub(crate) height: i64,
    pub(crate) id: u32,
    pub(crate) name: &'a str,
    pub(crate) description: &'a str,
}

/// The markup of a run holding an inline SmartArt graphic: a `dgm:relIds`
/// reference to its data, layout, quick-style and colors parts.
fn smartart_markup(
    prefix: &[u8],
    conformance: Conformance,
    [data, layout, style, colors]: [&str; 4],
    smartart: &SmartArtMarkup<'_>,
) -> Result<Vec<u8>> {
    let (prefix, declare) = if prefix.is_empty() {
        (b"w".as_slice(), true)
    } else {
        (prefix, false)
    };
    let (wordprocessing, main, diagram) = match conformance {
        Conformance::Strict => (
            "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing",
            "http://purl.oclc.org/ooxml/drawingml/main",
            litchi_drawingml::diagram::DGM_NAMESPACE_STRICT,
        ),
        Conformance::Transitional => (
            "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
            "http://schemas.openxmlformats.org/drawingml/2006/main",
            litchi_drawingml::diagram::DGM_NAMESPACE,
        ),
    };
    let escaped = |text: &str| -> Result<String> {
        let mut out = Vec::new();
        super::super::append_literal_escaped_xml(&mut out, text);
        String::from_utf8(out).map_err(|_| refusal(Refusal::InvalidPropertyValue))
    };
    let name = escaped(smartart.name)?;
    let description = escaped(smartart.description)?;
    let (data, layout, style, colors) = (
        escaped(data)?,
        escaped(layout)?,
        escaped(style)?,
        escaped(colors)?,
    );
    let relationships = relationships_namespace(conformance);
    let (width, height, id) = (smartart.width, smartart.height, smartart.id);
    let drawing = format!(
        r#"<wp:inline xmlns:wp="{wordprocessing}" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{width}" cy="{height}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{id}" name="{name}" descr="{description}"/><wp:cNvGraphicFramePr/><a:graphic xmlns:a="{main}"><a:graphicData uri="{diagram}"><dgm:relIds xmlns:dgm="{diagram}" xmlns:r="{relationships}" r:dm="{data}" r:lo="{layout}" r:qs="{style}" r:cs="{colors}"/></a:graphicData></a:graphic></wp:inline>"#
    );
    let mut xml = buffer(staging_add(drawing.len(), 128)?)?;
    open(&mut xml, prefix, "r");
    if declare {
        xml.extend_from_slice(b" xmlns:");
        xml.extend_from_slice(prefix);
        xml.extend_from_slice(b"=\"");
        xml.extend_from_slice(conformance.namespace().as_bytes());
        xml.push(b'"');
    }
    xml.push(b'>');
    open(&mut xml, prefix, "drawing");
    xml.push(b'>');
    xml.extend_from_slice(drawing.as_bytes());
    close(&mut xml, prefix, "drawing");
    close(&mut xml, prefix, "r");
    Ok(xml)
}

fn relationships_namespace(conformance: Conformance) -> &'static str {
    match conformance {
        Conformance::Strict => STRICT_RELATIONSHIPS_NAMESPACE,
        Conformance::Transitional => RELATIONSHIPS_NAMESPACE,
    }
}

/// The markup of a new hyperlink holding one run. The relationships
/// namespace is declared on the hyperlink itself, so its `r:id` resolves
/// whatever the document root binds.
fn hyperlink_markup(
    prefix: &[u8],
    conformance: Conformance,
    relationship: Option<&str>,
    anchor: Option<&str>,
    run: &NewRun,
) -> Result<Vec<u8>> {
    let (prefix, declare) = if prefix.is_empty() {
        (b"w".as_slice(), true)
    } else {
        (prefix, false)
    };
    let run = run_markup(prefix, conformance, run).map_err(spliced(Position::new(0)))?;
    let mut xml = buffer(staging_add(run.len(), 512)?)?;
    open(&mut xml, prefix, "hyperlink");
    if declare {
        xml.extend_from_slice(b" xmlns:");
        xml.extend_from_slice(prefix);
        xml.extend_from_slice(b"=\"");
        xml.extend_from_slice(conformance.namespace().as_bytes());
        xml.push(b'"');
    }
    match (relationship, anchor) {
        (Some(relationship), None) => {
            xml.extend_from_slice(b" xmlns:r=\"");
            xml.extend_from_slice(relationships_namespace(conformance).as_bytes());
            xml.extend_from_slice(b"\" r:id=\"");
            super::super::append_literal_escaped_xml(&mut xml, relationship);
            xml.push(b'"');
            attribute(&mut xml, prefix, "history", "1");
        },
        (None, Some(anchor)) => {
            xml.push(b' ');
            qualified_name(&mut xml, prefix, "anchor");
            xml.extend_from_slice(b"=\"");
            super::super::append_literal_escaped_xml(&mut xml, anchor);
            xml.push(b'"');
            attribute(&mut xml, prefix, "history", "1");
        },
        _ => return Err(refusal(Refusal::InvalidPropertyValue)),
    }
    xml.push(b'>');
    xml.extend_from_slice(&run);
    close(&mut xml, prefix, "hyperlink");
    ensure_utf8(&xml)?;
    Ok(xml)
}

/// A hyperlink with its relationship reference and bookmark replaced: the
/// start tag keeps every other attribute as written.
fn retarget_markup(
    xml: &[u8],
    inherited: &[(Option<Vec<u8>>, Vec<u8>)],
    conformance: Conformance,
    relationship: Option<&str>,
    anchor: Option<&str>,
) -> Result<Vec<u8>> {
    let (tag_end, empty) = super::root_tag(xml)?;
    let mut reader = NsReader::from_reader(&xml[..tag_end]);
    super::seed_namespaces(&mut reader, inherited)
        .map_err(|()| invalid("invalid inherited Word namespace"))?;
    let event = reader
        .read_event()
        .map_err(|_| Error::Document(crate::Error::Xml("invalid hyperlink tag".into())))?;
    let element = match event {
        Event::Start(element) | Event::Empty(element) => element,
        _ => return Err(invalid("hyperlink fragment has no start tag")),
    };
    let name = element.name();
    let element_prefix = name.prefix().map(|prefix| prefix.into_inner().to_vec());
    let mut output = buffer(staging_add(xml.len(), 512)?)?;
    output.push(b'<');
    output.extend_from_slice(name.as_ref());
    for attribute_result in element.unchecked_attributes() {
        let attribute = attribute_result.map_err(|_| {
            Error::Document(crate::Error::Xml("invalid hyperlink attribute".into()))
        })?;
        let key = attribute.key;
        let local = key.local_name();
        let (namespace, _) = reader.resolver().resolve_attribute(key);
        let relationship_id = local.as_ref() == b"id"
            && matches!(
                namespace,
                quick_xml::name::ResolveResult::Bound(quick_xml::name::Namespace(uri))
                    if uri == RELATIONSHIPS_NAMESPACE.as_bytes()
                        || uri == STRICT_RELATIONSHIPS_NAMESPACE.as_bytes()
            );
        let bookmark = local.as_ref() == b"anchor"
            && key.prefix().map(|prefix| prefix.into_inner().to_vec()) == element_prefix;
        let relationship_declaration = relationship.is_some() && key.as_ref() == b"xmlns:r";
        if relationship_id || bookmark || relationship_declaration {
            continue;
        }
        output.push(b' ');
        output.extend_from_slice(key.as_ref());
        output.extend_from_slice(b"=\"");
        for byte in attribute.value.iter() {
            if *byte == b'"' {
                output.extend_from_slice(b"&quot;");
            } else {
                output.push(*byte);
            }
        }
        output.push(b'"');
    }
    let prefix = element_prefix.unwrap_or_default();
    match (relationship, anchor) {
        (Some(relationship), None) => {
            output.extend_from_slice(b" xmlns:r=\"");
            output.extend_from_slice(relationships_namespace(conformance).as_bytes());
            output.extend_from_slice(b"\" r:id=\"");
            super::super::append_literal_escaped_xml(&mut output, relationship);
            output.push(b'"');
        },
        (None, Some(anchor)) => {
            output.push(b' ');
            qualified_name(&mut output, &prefix, "anchor");
            output.extend_from_slice(b"=\"");
            super::super::append_literal_escaped_xml(&mut output, anchor);
            output.push(b'"');
        },
        _ => return Err(refusal(Refusal::InvalidPropertyValue)),
    }
    output.extend_from_slice(if empty { b"/>" } else { b">" });
    output.extend_from_slice(&xml[tag_end..]);
    ensure_utf8(&output)?;
    Ok(output)
}

/// Visit every `commentRangeStart`, `commentRangeEnd` and
/// `commentReference` of a fragment: `(element range, id, local name)`.
fn scan_comment_markers(
    xml: &[u8],
    mut visit: impl FnMut((usize, usize), u32, &[u8]),
) -> Result<()> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("Word fragment offset overflow"))?;
        let event = reader
            .read_event()
            .map_err(|_| Error::Document(crate::Error::Xml("invalid Word XML".into())))?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("Word fragment offset overflow"))?;
        match event {
            Event::Start(element) | Event::Empty(element)
                if matches!(
                    element.local_name().as_ref(),
                    b"commentRangeStart" | b"commentRangeEnd" | b"commentReference"
                ) =>
            {
                let id = element
                    .unchecked_attributes()
                    .filter_map(std::result::Result::ok)
                    .find(|attribute| attribute.key.local_name().as_ref() == b"id")
                    .and_then(|attribute| std::str::from_utf8(&attribute.value).ok()?.parse().ok());
                if let Some(id) = id {
                    visit((start, end), id, element.local_name().as_ref());
                }
            },
            Event::Eof => return Ok(()),
            _ => {},
        }
    }
}

/// The fragment without comment `id`'s range markers and reference runs,
/// or `None` when it holds none.
/// `xml` without the runs that hold an element named `local` whose `id`
/// attribute is `id`, or `None` when it has none (record 0319).
fn without_note_references(xml: &[u8], local: &[u8], id: i64) -> Result<Option<Vec<u8>>> {
    let mut reader = quick_xml::Reader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let mut runs: Vec<(usize, bool)> = Vec::new();
    let mut removals: Vec<(usize, usize)> = Vec::new();
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("Word fragment offset overflow"))?;
        let event = reader
            .read_event()
            .map_err(|_| Error::Document(crate::Error::Xml("invalid Word XML".into())))?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("Word fragment offset overflow"))?;
        match event {
            Event::Start(element) if element.local_name().as_ref() == b"r" => {
                runs.push((start, false));
            },
            Event::Start(element) | Event::Empty(element)
                if element.local_name().as_ref() == local =>
            {
                let matches = element
                    .unchecked_attributes()
                    .filter_map(std::result::Result::ok)
                    .find(|attribute| attribute.key.local_name().as_ref() == b"id")
                    .and_then(|attribute| {
                        std::str::from_utf8(&attribute.value)
                            .ok()?
                            .trim()
                            .parse::<i64>()
                            .ok()
                    })
                    == Some(id);
                if matches && let Some(run) = runs.last_mut() {
                    run.1 = true;
                }
            },
            Event::End(element) if element.local_name().as_ref() == b"r" => {
                if let Some((run_start, true)) = runs.pop() {
                    removals.push((run_start, end));
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if removals.is_empty() {
        return Ok(None);
    }
    removals.sort_unstable();
    let mut output = buffer(xml.len())?;
    let mut cursor = 0usize;
    for (start, end) in removals {
        if start < cursor {
            continue;
        }
        output.extend_from_slice(&xml[cursor..start]);
        cursor = end;
    }
    output.extend_from_slice(&xml[cursor..]);
    Ok(Some(output))
}

fn without_comment_markers(xml: &[u8], id: u32) -> Result<Option<Vec<u8>>> {
    let mut removals: Vec<(usize, usize)> = Vec::new();
    let mut references = Vec::new();
    scan_comment_markers(xml, |range, found, local| {
        if found == id {
            if local == b"commentReference" {
                references.push(range.0);
            } else {
                removals.push(range);
            }
        }
    })?;
    if removals.is_empty() && references.is_empty() {
        return Ok(None);
    }
    // A reference is removed with the run that holds it.
    if !references.is_empty() {
        let mut reader = quick_xml::Reader::from_reader(xml);
        let origin = ReaderOrigin::of(xml);
        let mut runs: Vec<usize> = Vec::new();
        loop {
            let start = origin
                .offset(reader.buffer_position())
                .ok_or_else(|| invalid("Word fragment offset overflow"))?;
            let event = reader
                .read_event()
                .map_err(|_| Error::Document(crate::Error::Xml("invalid Word XML".into())))?;
            let end = origin
                .offset(reader.buffer_position())
                .ok_or_else(|| invalid("Word fragment offset overflow"))?;
            match event {
                Event::Start(element) if element.local_name().as_ref() == b"r" => runs.push(start),
                Event::End(element) if element.local_name().as_ref() == b"r" => {
                    if let Some(run_start) = runs.pop()
                        && references
                            .iter()
                            .any(|reference| *reference > run_start && *reference < end)
                    {
                        removals.push((run_start, end));
                    }
                },
                Event::Eof => break,
                _ => {},
            }
        }
    }
    removals.sort_unstable();
    let mut output = buffer(xml.len())?;
    let mut cursor = 0usize;
    for (start, end) in removals {
        if start < cursor {
            continue;
        }
        output.extend_from_slice(&xml[cursor..start]);
        cursor = end;
    }
    output.extend_from_slice(&xml[cursor..]);
    Ok(Some(output))
}
