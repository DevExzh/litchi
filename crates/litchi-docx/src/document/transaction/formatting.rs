//! Formatting and structure of existing body content (unified-ops L-DOCX-W5).
//!
//! These [`Edit`] operations change the direct properties of existing runs
//! and paragraphs, insert paragraphs and runs that carry properties, and
//! remove and move body paragraphs and runs. Like every other operation of
//! the main-document transaction they
//!
//! - address paragraphs, tables and block content controls by their
//!   transaction positions (separate zero-based positions for each block kind)
//!   and runs by their position among a paragraph's direct `w:r` children,
//!   resolved against the projected document;
//! - record one semantic [`Operation`] each, with an exact inverse, so a
//!   commit's [`Patch`](super::Patch) undoes them, composes under the
//!   existing read/write effect rules, and serializes durably;
//! - refuse with a typed [`Refusal`] instead of guessing.
//!
//! Property changes are typed: the caller edits the very
//! [`properties::Run`] or [`properties::Paragraph`] value the readers return
//! for the element, and the engine writes back only what changed (see
//! `crate::properties::edit` for the splice rules). A read, an unchanged
//! write and a read again therefore agree, and writing a property's current
//! value is an exact no-op.
//!
//! The operations record exact markup, so a commit that holds one publishes
//! the projected main document as staged, without compaction, exactly as a
//! revision disposition does.

use std::sync::Arc;

use litchi_core::Position;
use litchi_core::xml::ReaderOrigin;
use quick_xml::Reader;
use quick_xml::events::Event;
use quick_xml::name::{QName, ResolveResult};
use quick_xml::reader::NsReader;

use super::{
    Conformance, Edit, Error, FragmentPrefix, Operation, ParagraphSplice, Refusal, SplicePolicy,
    checked_end, checked_start, is_fragment_word_element, is_transaction_fragment_word_name,
    replace_range, scan_text_owner, try_run_content_fragment, validate_authored_run_content,
};
use crate::properties;
type Result<T, E = Error> = std::result::Result<T, E>;

mod element;
pub(crate) use element::{ChartMarkup, PictureMarkup, PictureRewrite, SmartArtMarkup};
pub use element::{
    ElementAnchor, ElementPath, ElementStep, MAX_NEW_TABLE_DIMENSION, NewCell, NewTable,
};
pub(in crate::document::transaction) use element::{body_children, locate_element};

/// A run to insert: its direct properties and its text.
///
/// Tabs, line feeds, carriage returns, non-breaking hyphens (U+2011) and soft
/// hyphens (U+00AD) in the text are written as `w:tab`, `w:br`, `w:cr`,
/// `w:noBreakHyphen` and `w:softHyphen`, as [`Edit::replace_run_text`]
/// writes them; every other character must be one XML can carry.
///
/// ```
/// use litchi_docx::document::NewRun;
/// use litchi_docx::properties::HalfPoints;
///
/// let mut run = NewRun::new("Bold 14 pt");
/// run.properties_mut().bold = Some(true);
/// run.properties_mut().size = Some(HalfPoints(28));
/// assert_eq!(run.text(), "Bold 14 pt");
/// ```
#[derive(Clone, Default, PartialEq, Eq)]
pub struct NewRun {
    properties: properties::Run,
    text: String,
}

impl NewRun {
    /// A run with `text` and no direct properties.
    #[must_use]
    pub fn new(text: impl Into<String>) -> Self {
        Self {
            properties: properties::Run::default(),
            text: text.into(),
        }
    }

    /// Replace the run's direct properties.
    #[must_use]
    pub fn with_properties(mut self, properties: properties::Run) -> Self {
        self.properties = properties;
        self
    }

    /// Borrow the run's direct properties for editing.
    pub fn properties_mut(&mut self) -> &mut properties::Run {
        &mut self.properties
    }

    /// The run's direct properties.
    #[must_use]
    pub const fn properties(&self) -> &properties::Run {
        &self.properties
    }

    /// The run's text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}

/// A paragraph to insert: its direct properties and its runs.
///
/// ```
/// use litchi_docx::document::{NewParagraph, NewRun};
/// use litchi_docx::properties::Justification;
///
/// let mut paragraph = NewParagraph::new().with_run(NewRun::new("Title"));
/// paragraph.properties_mut().style = Some("Heading1".to_owned());
/// paragraph.properties_mut().justification = Some(Justification::Center);
/// assert_eq!(paragraph.runs().len(), 1);
/// ```
#[derive(Clone, Default, PartialEq, Eq)]
pub struct NewParagraph {
    properties: properties::Paragraph,
    runs: Vec<NewRun>,
}

impl NewParagraph {
    /// An empty paragraph with no direct properties.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Replace the paragraph's direct properties.
    #[must_use]
    pub fn with_properties(mut self, properties: properties::Paragraph) -> Self {
        self.properties = properties;
        self
    }

    /// Append a run.
    #[must_use]
    pub fn with_run(mut self, run: NewRun) -> Self {
        self.runs.push(run);
        self
    }

    /// Append a run.
    pub fn push_run(&mut self, run: NewRun) -> &mut Self {
        self.runs.push(run);
        self
    }

    /// Borrow the paragraph's direct properties for editing.
    pub fn properties_mut(&mut self) -> &mut properties::Paragraph {
        &mut self.properties
    }

    /// The paragraph's direct properties.
    #[must_use]
    pub const fn properties(&self) -> &properties::Paragraph {
        &self.properties
    }

    /// The paragraph's runs, in order.
    #[must_use]
    pub fn runs(&self) -> &[NewRun] {
        &self.runs
    }
}

/// A direct child of the document body, addressed by its transaction
/// position: its position among the body's direct children of the same
/// element, which is what the transaction's per-kind position sequence
/// returns for a body block.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BodyBlock {
    /// A direct `w:p` (a paragraph, or a paragraph holding display math).
    Paragraph(Position),
    /// A direct `w:tbl`.
    Table(Position),
    /// A direct block-level `w:sdt`.
    ContentControl(Position),
}

/// Where a paragraph goes among the direct children of the document body.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BodyAnchor {
    /// Immediately before the block.
    Before(BodyBlock),
    /// Immediately after the block.
    After(BodyBlock),
    /// After every other child, before the body-final section properties.
    End,
}

/// Where a direct run goes inside a paragraph.
///
/// The position is the projected position of a direct `w:r`, and is checked
/// against the paragraph that is passed to [`Edit::move_run`]. `Before` may
/// name the run count and is then equivalent to the end of the paragraph.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum RunAnchor {
    /// Immediately before the direct run.
    Before(Position),
    /// Immediately after the direct run.
    After(Position),
    /// After every direct child of the paragraph.
    End,
}

/// What one direct child of `w:body` is.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ChildKind {
    Paragraph(usize),
    Table(usize),
    ContentControl(usize),
    Other,
}

/// One direct element child of `w:body`.
#[derive(Debug, Clone, Copy)]
struct BodyChild {
    start: usize,
    end: usize,
    kind: ChildKind,
}

/// The direct element children of the projected document's `w:body`.
struct Body {
    /// Offset of `</w:body>`.
    body_end: usize,
    children: Vec<BodyChild>,
    /// Index of the body-final section properties in `children`.
    section: Option<usize>,
    /// A prefix bound to the document's WordprocessingML namespace where the
    /// body's children are, if any is.
    prefix: Option<Vec<u8>>,
}

impl Body {
    /// The child index an anchor inserts at, before any change.
    fn resolve(&self, anchor: BodyAnchor) -> Result<usize> {
        let find = |block: BodyBlock| {
            let (wanted, position) = match block {
                BodyBlock::Paragraph(position) => (ChildKind::Paragraph(position.get()), position),
                BodyBlock::Table(position) => (ChildKind::Table(position.get()), position),
                BodyBlock::ContentControl(position) => {
                    (ChildKind::ContentControl(position.get()), position)
                },
            };
            self.children
                .iter()
                .position(|child| child.kind == wanted)
                .ok_or(Error::Refused {
                    position: position.get(),
                    reason: Refusal::BlockNotFound,
                })
        };
        match anchor {
            BodyAnchor::Before(block) => find(block),
            BodyAnchor::After(block) => Ok(find(block)? + 1),
            BodyAnchor::End => Ok(self.insertion_limit()),
        }
    }

    /// The largest child index an insertion may use: the body-final section
    /// properties stay last.
    fn insertion_limit(&self) -> usize {
        self.section.unwrap_or(self.children.len())
    }

    /// The byte offset of an insertion at child index `child`.
    fn insertion_offset(&self, child: usize) -> Option<usize> {
        if child > self.insertion_limit() {
            return None;
        }
        Some(
            self.children
                .get(child)
                .map_or(self.body_end, |found| found.start),
        )
    }

    /// The number of direct paragraphs before child index `child`.
    fn paragraphs_before(&self, child: usize) -> usize {
        self.children
            .iter()
            .take(child)
            .filter(|found| matches!(found.kind, ChildKind::Paragraph(_)))
            .count()
    }

    /// The child index of direct paragraph `position`.
    fn paragraph_child(&self, position: usize) -> Option<usize> {
        self.children
            .iter()
            .position(|child| child.kind == ChildKind::Paragraph(position))
    }

    /// The bytes child `child` owns: from its start to the next child's start
    /// (or the end of the body), so the whitespace and comments that follow
    /// an element leave and return with it.
    fn extent(&self, child: usize) -> Option<(usize, usize)> {
        let found = self.children.get(child)?;
        let next = self
            .children
            .get(child + 1)
            .map_or(self.body_end, |next| next.start);
        Some((found.start, next))
    }
}

/// The direct element children of one paragraph fragment.
struct ParagraphChildren {
    /// Offset just past the paragraph's start tag.
    tag_end: usize,
    /// Whether the paragraph is an empty element (`<w:p/>`).
    empty: bool,
    /// Offset of the paragraph's end tag, or its end for an empty element.
    close_start: usize,
    /// The paragraph's prefix, as written.
    prefix: Vec<u8>,
    /// `(start, end, is a direct WordprocessingML run)`, in order.
    children: Vec<(usize, usize, bool)>,
}

impl ParagraphChildren {
    fn run_count(&self) -> usize {
        self.children.iter().filter(|child| child.2).count()
    }

    /// The child index of direct run `run`, or the end of the paragraph when
    /// `run` is the run count.
    fn run_insertion_child(&self, run: usize) -> Option<usize> {
        let mut seen = 0usize;
        for (index, child) in self.children.iter().enumerate() {
            if child.2 {
                if seen == run {
                    return Some(index);
                }
                seen += 1;
            }
        }
        (seen == run).then_some(self.children.len())
    }

    /// The child index of existing direct run `run`.
    fn run_child(&self, run: usize) -> Option<usize> {
        self.run_insertion_child(run)
            .filter(|child| *child < self.children.len())
    }

    /// The number of direct runs before child index `child`.
    fn runs_before(&self, child: usize) -> usize {
        self.children
            .iter()
            .take(child)
            .filter(|found| found.2)
            .count()
    }

    fn insertion_offset(&self, child: usize) -> Option<usize> {
        if child > self.children.len() {
            return None;
        }
        Some(
            self.children
                .get(child)
                .map_or(self.close_start, |found| found.0),
        )
    }

    /// The bytes child `child` owns; see [`Body::extent`].
    fn extent(&self, child: usize) -> Option<(usize, usize)> {
        let found = self.children.get(child)?;
        let next = self
            .children
            .get(child + 1)
            .map_or(self.close_start, |next| next.0);
        Some((found.0, next))
    }
}

fn run_anchor_child(layout: &ParagraphChildren, anchor: RunAnchor) -> Result<usize, Refusal> {
    match anchor {
        RunAnchor::Before(position) => layout
            .run_insertion_child(position.get())
            .ok_or(Refusal::RunNotFound),
        RunAnchor::After(position) => layout
            .run_child(position.get())
            .map(|child| child + 1)
            .ok_or(Refusal::RunNotFound),
        RunAnchor::End => Ok(layout.children.len()),
    }
}

/// The length of the element a structure fragment starts with; whatever
/// follows it (the whitespace or comments a removed element owned) must hold
/// no element.
fn element_len(fragment: &[u8]) -> Result<usize> {
    let mut reader = Reader::from_reader(fragment);
    let origin = ReaderOrigin::of(fragment);
    let mut depth = 0usize;
    let mut element = None::<usize>;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("Word fragment offset overflow"))?;
        let event = reader.read_event().map_err(|_| {
            Error::Document(crate::Error::Xml("invalid Word formatting XML".into()))
        })?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("Word fragment offset overflow"))?;
        match event {
            Event::Start(_) | Event::Empty(_) if element.is_some() => {
                return Err(invalid("structure fragment holds more than one element"));
            },
            Event::Start(_) | Event::Empty(_) if depth == 0 && start != 0 => {
                return Err(invalid(
                    "structure fragment does not start with its element",
                ));
            },
            Event::Start(_) => depth += 1,
            Event::Empty(_) if depth == 0 => element = Some(end),
            Event::End(_) => {
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| invalid("structure fragment is unbalanced"))?;
                if depth == 0 {
                    element = Some(end);
                }
            },
            Event::Eof => {
                return element
                    .filter(|_| depth == 0)
                    .ok_or_else(|| invalid("structure fragment holds no element"));
            },
            _ => {},
        }
    }
}

fn refused(position: Position) -> impl Fn(Refusal) -> Error {
    move |reason| Error::Refused {
        position: position.get(),
        reason,
    }
}

fn spliced(position: Position) -> impl Fn(properties::SpliceError) -> Error {
    move |error| match error {
        properties::SpliceError::Refused(reason) => Error::Refused {
            position: position.get(),
            reason,
        },
        properties::SpliceError::Document(error) => Error::Document(error),
        properties::SpliceError::Transaction(error) => error,
    }
}

fn invalid(message: &str) -> Error {
    Error::Document(crate::Error::InvalidFormat(message.to_owned()))
}

/// A changed paragraph or inserted fragment must be UTF-8, as a commit
/// requires of the paragraphs an edit changes.
fn ensure_utf8(xml: &[u8]) -> Result<()> {
    std::str::from_utf8(xml).map(|_| ()).map_err(|error| {
        Error::Document(crate::Error::InvalidFormat(format!(
            "changed main-document XML is not UTF-8: {error}"
        )))
    })
}

impl Edit {
    /// Change the direct properties (`w:rPr`) of one direct run.
    ///
    /// `change` receives the run's direct properties exactly as
    /// [`crate::Run::direct_properties`] reads them and edits them in place:
    /// set a field to give the property a value, set it to `None` to remove
    /// the property. Only properties whose typed value changes are written:
    /// a new property element is inserted in schema order, a changed one
    /// keeps its unknown attributes and children, a removed one takes only
    /// its own element, and the container's unknown children and its
    /// `w:rPrChange` stay as they are. A property the reader could not type
    /// stays unless `change` gives it a value. When nothing changes the edit
    /// records nothing.
    ///
    /// `run` counts the paragraph's direct `w:r` children, as
    /// [`Self::replace_run_text`] does.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::document::raw;
    /// use litchi_docx::properties::HalfPoints;
    ///
    /// let mut edit = raw::Tree::from_xml(
    ///     br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:i/></w:rPr><w:t>x</w:t></w:r></w:p></w:body></w:document>"#
    ///         .to_vec(),
    ///     raw::Limits::default(),
    /// )?;
    /// edit.format_run(Position::new(0), Position::new(0), |run| {
    ///     run.bold = Some(true);
    ///     run.size = Some(HalfPoints(28));
    /// })?;
    /// let bytes = edit.to_bytes(raw::Limits::default())?;
    /// let xml = std::str::from_utf8(&bytes)?;
    /// assert!(xml.contains("<w:rPr><w:b/><w:i/><w:sz w:val=\"28\"/></w:rPr>"));
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::RunNotFound`] for a missing run, a typed refusal
    /// for a value outside its schema domain, a Transitional-only value in a
    /// Strict document, a duplicated property element that has to change, or
    /// a container removal that would drop change-tracking or unknown markup,
    /// a resource-limit error, or a malformed-document error, without
    /// changing the projected snapshot. Budget-managed snapshots are refused.
    pub fn format_run(
        &mut self,
        paragraph: Position,
        run: Position,
        change: impl FnOnce(&mut properties::Run),
    ) -> Result<&mut Self> {
        self.ensure_unmanaged("format_run")?;
        let (start, end) = self.paragraph_bounds(paragraph)?;
        let inherited = self.projected.paragraph_namespaces[paragraph.get()].as_ref();
        let paragraph_xml = &self.projected.xml_bytes()[start..end];
        let (run_start, run_end) =
            direct_run(paragraph_xml, run, inherited).map_err(refused(paragraph))?;
        let run_xml = &paragraph_xml[run_start..run_end];
        let current = properties::parse_run(run_xml, inherited)?;
        let mut target = current.clone();
        change(&mut target);
        let Some(spliced_run) = properties::splice_run(run_xml, &current, &target, inherited)
            .map_err(spliced(paragraph))?
        else {
            return Ok(self);
        };
        let before = head(run_xml, b"r", b"rPr", inherited)?.to_vec();
        let after = head(&spliced_run, b"r", b"rPr", inherited)?.to_vec();
        // The recorded operation replaces the run's head alone; it must
        // reproduce the splice, or replay and the inverse would diverge.
        if replace_head(run_xml, b"r", b"rPr", &before, &after, inherited)? != spliced_run {
            return Err(invalid(
                "run property splice is not confined to the run's head",
            ));
        }
        self.replace_run_properties(paragraph, run, &Arc::new(before), &Arc::new(after))?;
        Ok(self)
    }

    /// Change the direct properties (`w:pPr`) of one direct-body paragraph,
    /// including the paragraph mark's run properties
    /// ([`properties::Paragraph::mark`]).
    ///
    /// The rules of [`Self::format_run`] apply. Style application is
    /// [`properties::Paragraph::style`] (`w:pStyle`) and numbering
    /// application is [`properties::Paragraph::numbering`] (`w:numPr`); the
    /// edit owns the main document only, so it does not check that the style
    /// or numbering instance exists. Removing `w:numPr`, `w:pBdr`, `w:tabs`
    /// or the paragraph mark's `w:rPr` is refused while it holds
    /// change-tracking or unknown markup; `w:pPrChange` and `w:sectPr` are
    /// kept.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::document::raw;
    /// use litchi_docx::properties::{Justification, Numbering};
    ///
    /// let mut edit = raw::Tree::from_xml(
    ///     br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Item</w:t></w:r></w:p></w:body></w:document>"#
    ///         .to_vec(),
    ///     raw::Limits::default(),
    /// )?;
    /// edit.format_paragraph(Position::new(0), |paragraph| {
    ///     paragraph.style = Some("ListParagraph".to_owned());
    ///     let mut numbering = Numbering::default();
    ///     numbering.id = Some(3);
    ///     numbering.level = Some(0);
    ///     paragraph.numbering = Some(numbering);
    ///     paragraph.justification = Some(Justification::Both);
    /// })?;
    /// let paragraph = edit.projected().paragraph(Position::new(0)).unwrap();
    /// let properties = paragraph.direct_properties()?;
    /// assert_eq!(properties.style.as_deref(), Some("ListParagraph"));
    /// assert_eq!(properties.numbering.and_then(|numbering| numbering.id), Some(3));
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// As [`Self::format_run`].
    pub fn format_paragraph(
        &mut self,
        paragraph: Position,
        change: impl FnOnce(&mut properties::Paragraph),
    ) -> Result<&mut Self> {
        self.ensure_unmanaged("format_paragraph")?;
        let (start, end) = self.paragraph_bounds(paragraph)?;
        let inherited = self.projected.paragraph_namespaces[paragraph.get()].as_ref();
        let paragraph_xml = &self.projected.xml_bytes()[start..end];
        let current = properties::parse_paragraph(paragraph_xml, inherited)?;
        let mut target = current.clone();
        change(&mut target);
        let strict = matches!(self.projected.conformance, Conformance::Strict);
        let Some(spliced_paragraph) =
            properties::splice_paragraph(paragraph_xml, &current, &target, strict, inherited)
                .map_err(spliced(paragraph))?
        else {
            return Ok(self);
        };
        let before = head(paragraph_xml, b"p", b"pPr", inherited)?.to_vec();
        let after = head(&spliced_paragraph, b"p", b"pPr", inherited)?.to_vec();
        if replace_head(paragraph_xml, b"p", b"pPr", &before, &after, inherited)?
            != spliced_paragraph
        {
            return Err(invalid(
                "paragraph property splice is not confined to the paragraph's head",
            ));
        }
        self.replace_paragraph_properties(paragraph, &Arc::new(before), &Arc::new(after))?;
        Ok(self)
    }

    /// Insert a paragraph with direct properties and runs as a direct child
    /// of the document body, and return its paragraph position.
    ///
    /// The paragraph is written with the prefix the body binds to the
    /// document's WordprocessingML namespace (it declares one only when the
    /// body binds none), carries no identifier, and is read back before it
    /// is accepted.
    ///
    /// ```
    /// use litchi_core::Position;
    /// use litchi_docx::document::{raw, BodyAnchor, BodyBlock, NewParagraph, NewRun};
    ///
    /// let mut edit = raw::Tree::from_xml(
    ///     br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/><w:tbl/><w:sectPr/></w:body></w:document>"#
    ///         .to_vec(),
    ///     raw::Limits::default(),
    /// )?;
    /// let mut run = NewRun::new("after the table");
    /// run.properties_mut().italic = Some(true);
    /// let position = edit.insert_formatted_paragraph(
    ///     BodyAnchor::After(BodyBlock::Table(Position::new(0))),
    ///     &NewParagraph::new().with_run(run),
    /// )?;
    /// assert_eq!(position, Position::new(1));
    /// assert_eq!(edit.projected().paragraph_count(), 2);
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for a missing anchor block, invalid run text
    /// or property values, or a resource-limit or malformed-document error,
    /// without changing the projected snapshot.
    pub fn insert_formatted_paragraph(
        &mut self,
        at: BodyAnchor,
        paragraph: &NewParagraph,
    ) -> Result<Position> {
        self.ensure_unmanaged("insert_formatted_paragraph")?;
        let body = body(self)?;
        let child = body.resolve(at)?;
        let position = Position::new(body.paragraphs_before(child));
        let strict = matches!(self.projected.conformance, Conformance::Strict);
        let xml = paragraph_markup(
            body.prefix.as_deref(),
            self.projected.conformance,
            paragraph,
            strict,
        )
        .map_err(spliced(position))?;
        self.insert_paragraph_fragment(Position::new(child), position, &Arc::new(xml))?;
        Ok(position)
    }

    /// Remove one direct-body paragraph and everything it holds.
    ///
    /// The removal is refused, rather than leaving a dangling reference or
    /// changing the document's structure, when the paragraph carries section
    /// properties ([`Refusal::SectionBreak`]); holds one end of a bookmark,
    /// comment, permission, move, custom-XML or complex-field range whose
    /// other end is outside it ([`Refusal::OpenRange`]); holds a comment,
    /// footnote or endnote reference, or a side of a tracked move
    /// ([`Refusal::StoryReference`]); or its paragraph mark is a tracked
    /// insertion, deletion or move ([`Refusal::TrackedChange`]).
    ///
    /// # Errors
    ///
    /// Returns a checked-position error, one of the typed refusals above, or
    /// a resource-limit or malformed-document error, without changing the
    /// projected snapshot.
    pub fn remove_paragraph(&mut self, position: Position) -> Result<&mut Self> {
        self.ensure_unmanaged("remove_paragraph")?;
        let (start, end) = self.paragraph_bounds(position)?;
        let inherited = self.projected.paragraph_namespaces[position.get()].as_ref();
        paragraph_dependencies(&self.projected.xml_bytes()[start..end], inherited)
            .refuse(Removal::Delete)
            .map_err(refused(position))?;
        let body = body(self)?;
        let child = body
            .paragraph_child(position.get())
            .ok_or_else(|| invalid("direct paragraph is missing from the body"))?;
        let (from, to) = body
            .extent(child)
            .ok_or_else(|| invalid("direct paragraph is missing from the body"))?;
        let xml = Arc::new(self.projected.xml_bytes()[from..to].to_vec());
        self.remove_paragraph_fragment(Position::new(child), position, &xml)?;
        Ok(self)
    }

    /// Move one direct-body paragraph, unchanged, to `to`, and return its
    /// new paragraph position.
    ///
    /// `to` is resolved before the paragraph is taken out; anchoring the
    /// paragraph to itself or to a neighbour that leaves it where it is is
    /// an exact no-op. The move is recorded as the removal and the insertion
    /// of the same bytes, so its inverse restores the original position. It
    /// is refused for the reasons of [`Self::remove_paragraph`] except
    /// references, which move with the paragraph.
    ///
    /// # Errors
    ///
    /// As [`Self::remove_paragraph`], and a typed refusal for a missing
    /// anchor block.
    pub fn move_paragraph(&mut self, position: Position, to: BodyAnchor) -> Result<Position> {
        self.ensure_unmanaged("move_paragraph")?;
        let (start, end) = self.paragraph_bounds(position)?;
        let inherited = self.projected.paragraph_namespaces[position.get()].as_ref();
        paragraph_dependencies(&self.projected.xml_bytes()[start..end], inherited)
            .refuse(Removal::Move)
            .map_err(refused(position))?;
        let body = body(self)?;
        let from = body
            .paragraph_child(position.get())
            .ok_or_else(|| invalid("direct paragraph is missing from the body"))?;
        // The paragraph moves with the whitespace and comments it owns.
        let (owned_start, owned_end) = body
            .extent(from)
            .ok_or_else(|| invalid("direct paragraph is missing from the body"))?;
        let xml = Arc::new(self.projected.xml_bytes()[owned_start..owned_end].to_vec());
        let target = body.resolve(to)?;
        if target == from || target == from + 1 {
            return Ok(position);
        }
        let destination = if target > from { target - 1 } else { target };
        let mut candidate = self.try_clone()?;
        candidate.remove_paragraph_fragment(Position::new(from), position, &xml)?;
        let moved = Position::new(body_of(&candidate)?.paragraphs_before(destination));
        candidate.insert_paragraph_fragment(Position::new(destination), moved, &xml)?;
        *self = candidate;
        Ok(moved)
    }

    /// Swap two direct-body paragraphs, retaining their complete owned
    /// fragments and the bytes between each fragment and its next sibling.
    ///
    /// The operation is represented by the existing exact paragraph removal
    /// and insertion records. Consequently its inverse and durable replay
    /// use the same source fragments as ordinary structure operations.
    /// Paragraphs are addressed in the projected document; equal positions
    /// are an exact no-op.
    ///
    /// # Errors
    ///
    /// Returns the same checked-position, dependency, section, and resource
    /// errors as [`Self::move_paragraph`]. Managed snapshots are refused
    /// because they do not retain a source-preserving structural publication
    /// route.
    pub fn swap_paragraphs(&mut self, first: Position, second: Position) -> Result<&mut Self> {
        self.ensure_unmanaged("swap_paragraphs")?;
        if first == second {
            self.paragraph_bounds(first)?;
            return Ok(self);
        }
        let (left, right) = if first < second {
            (first, second)
        } else {
            (second, first)
        };
        let (left_start, left_end) = self.paragraph_bounds(left)?;
        let (right_start, right_end) = self.paragraph_bounds(right)?;
        let left_inherited = self.projected.paragraph_namespaces[left.get()].as_ref();
        let right_inherited = self.projected.paragraph_namespaces[right.get()].as_ref();
        let left_xml = &self.projected.xml_bytes()[left_start..left_end];
        let right_xml = &self.projected.xml_bytes()[right_start..right_end];
        paragraph_dependencies(left_xml, left_inherited)
            .refuse(Removal::Move)
            .map_err(refused(left))?;
        paragraph_dependencies(right_xml, right_inherited)
            .refuse(Removal::Move)
            .map_err(refused(right))?;
        let body = body(self)?;
        let left_child = body
            .paragraph_child(left.get())
            .ok_or_else(|| invalid("direct paragraph is missing from the body"))?;
        let right_child = body
            .paragraph_child(right.get())
            .ok_or_else(|| invalid("direct paragraph is missing from the body"))?;
        let (left_owned_start, left_owned_end) = body
            .extent(left_child)
            .ok_or_else(|| invalid("direct paragraph is missing from the body"))?;
        let (right_owned_start, right_owned_end) = body
            .extent(right_child)
            .ok_or_else(|| invalid("direct paragraph is missing from the body"))?;
        let left_fragment =
            Arc::new(self.projected.xml_bytes()[left_owned_start..left_owned_end].to_vec());
        let right_fragment =
            Arc::new(self.projected.xml_bytes()[right_owned_start..right_owned_end].to_vec());
        // Removing both fragments and inserting them at their original child
        // slots swaps the paragraphs while leaving tables, controls, unknown
        // elements, comments, and their physical gaps in those slots.  A
        // pair of ordinal paragraph moves cannot provide that guarantee.
        let mut candidate = self.try_clone()?;
        candidate.remove_paragraph_fragment(Position::new(right_child), right, &right_fragment)?;
        candidate.remove_paragraph_fragment(Position::new(left_child), left, &left_fragment)?;
        let body_after_removal = body_of(&candidate)?;
        let moved_right = Position::new(body_after_removal.paragraphs_before(left_child));
        candidate.insert_paragraph_fragment(
            Position::new(left_child),
            moved_right,
            &right_fragment,
        )?;
        let body_after_right = body_of(&candidate)?;
        let moved_left = Position::new(body_after_right.paragraphs_before(right_child));
        candidate.insert_paragraph_fragment(
            Position::new(right_child),
            moved_left,
            &left_fragment,
        )?;
        *self = candidate;
        Ok(self)
    }

    /// Insert a run with direct properties as a direct child of a paragraph,
    /// before direct run `run`, or after the paragraph's last child when
    /// `run` is the paragraph's direct run count; return `run`.
    ///
    /// The run is written with the paragraph's prefix. An insertion point
    /// inside a complex field (between its `begin` and `end` markers, as far
    /// as the paragraph shows them) is refused with
    /// [`Refusal::FieldInterior`].
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::RunNotFound`] past the end, a typed refusal for
    /// the insertion point or for invalid text or property values, or a
    /// resource-limit or malformed-document error, without changing the
    /// projected snapshot.
    pub fn insert_run(
        &mut self,
        paragraph: Position,
        run: Position,
        new_run: &NewRun,
    ) -> Result<Position> {
        self.ensure_unmanaged("insert_run")?;
        let (start, end) = self.paragraph_bounds(paragraph)?;
        let inherited = self.projected.paragraph_namespaces[paragraph.get()].as_ref();
        let paragraph_xml = &self.projected.xml_bytes()[start..end];
        let layout = paragraph_children(paragraph_xml, inherited).map_err(refused(paragraph))?;
        let child = layout
            .run_insertion_child(run.get())
            .ok_or(Refusal::RunNotFound)
            .map_err(refused(paragraph))?;
        let offset = layout
            .insertion_offset(child)
            .ok_or(Refusal::RunNotFound)
            .map_err(refused(paragraph))?;
        if inside_complex_field(paragraph_xml, offset, inherited).map_err(refused(paragraph))? {
            return Err(refused(paragraph)(Refusal::FieldInterior));
        }
        let xml = run_markup(&layout.prefix, self.projected.conformance, new_run)
            .map_err(spliced(paragraph))?;
        if layout.empty {
            // An empty paragraph (`<w:p/>`) is first opened by its own head
            // replacement, whose inverse closes it again once the run is
            // gone, so the inverse restores the exact element.
            let mut candidate = self.try_clone()?;
            let closed = Arc::new(paragraph_xml.to_vec());
            let open = Arc::new(opened(paragraph_xml)?);
            candidate.replace_paragraph_properties(paragraph, &closed, &open)?;
            candidate.insert_run_fragment(paragraph, Position::new(child), run, &Arc::new(xml))?;
            *self = candidate;
            return Ok(run);
        }
        self.insert_run_fragment(paragraph, Position::new(child), run, &Arc::new(xml))?;
        Ok(run)
    }

    /// Append a run after the paragraph's last child and return its direct
    /// run position; see [`Self::insert_run`].
    ///
    /// # Errors
    ///
    /// As [`Self::insert_run`].
    pub fn append_run(&mut self, paragraph: Position, new_run: &NewRun) -> Result<Position> {
        self.ensure_unmanaged("append_run")?;
        let (start, end) = self.paragraph_bounds(paragraph)?;
        let inherited = self.projected.paragraph_namespaces[paragraph.get()].as_ref();
        let runs = paragraph_children(&self.projected.xml_bytes()[start..end], inherited)
            .map_err(refused(paragraph))?
            .run_count();
        self.insert_run(paragraph, Position::new(runs), new_run)
    }

    /// Remove one direct run of a direct-body paragraph.
    ///
    /// A run holding a complex-field character or field code is refused
    /// with [`Refusal::OpenRange`], and one holding a comment, footnote or
    /// endnote reference with [`Refusal::StoryReference`].
    ///
    /// # Errors
    ///
    /// Returns [`Refusal::RunNotFound`] for a missing run, one of the typed
    /// refusals above, or a resource-limit or malformed-document error,
    /// without changing the projected snapshot.
    pub fn remove_run(&mut self, paragraph: Position, run: Position) -> Result<&mut Self> {
        self.ensure_unmanaged("remove_run")?;
        let (start, end) = self.paragraph_bounds(paragraph)?;
        let inherited = self.projected.paragraph_namespaces[paragraph.get()].as_ref();
        let paragraph_xml = &self.projected.xml_bytes()[start..end];
        let layout = paragraph_children(paragraph_xml, inherited).map_err(refused(paragraph))?;
        let child = layout
            .run_child(run.get())
            .ok_or(Refusal::RunNotFound)
            .map_err(refused(paragraph))?;
        let (run_start, run_end, _) = layout.children[child];
        run_dependencies(&paragraph_xml[run_start..run_end], inherited)
            .map_err(refused(paragraph))?;
        let (from, to) = layout
            .extent(child)
            .ok_or_else(|| invalid("direct run is missing from the paragraph"))?;
        let xml = Arc::new(paragraph_xml[from..to].to_vec());
        self.remove_run_fragment(paragraph, Position::new(child), run, &xml)?;
        Ok(self)
    }

    /// Move one direct run within its paragraph, unchanged, to `to`.
    ///
    /// The run's complete owned fragment is spliced, including whitespace and
    /// comments up to the next direct child. The source run must be outside a
    /// complex field and must not contain a field character or instruction;
    /// the destination must also be outside a complex field. This keeps field
    /// ownership intact while allowing inert run markup to travel byte for
    /// byte.
    ///
    /// # Errors
    ///
    /// Returns a checked run or field/dependency refusal, a malformed XML or
    /// resource error, or an unsafe-edit refusal for managed snapshots. The
    /// edit remains unchanged on every error.
    pub fn move_run(
        &mut self,
        paragraph: Position,
        run: Position,
        to: RunAnchor,
    ) -> Result<Position> {
        self.ensure_unmanaged("move_run")?;
        let (start, end) = self.paragraph_bounds(paragraph)?;
        let inherited = self.projected.paragraph_namespaces[paragraph.get()].as_ref();
        let paragraph_xml = &self.projected.xml_bytes()[start..end];
        let layout = paragraph_children(paragraph_xml, inherited).map_err(refused(paragraph))?;
        let source_child = layout
            .run_child(run.get())
            .ok_or(Refusal::RunNotFound)
            .map_err(refused(paragraph))?;
        let (run_start, _run_end, _) = layout.children[source_child];
        if inside_complex_field(paragraph_xml, run_start, inherited).map_err(refused(paragraph))? {
            return Err(refused(paragraph)(Refusal::FieldInterior));
        }
        let target_child = run_anchor_child(&layout, to).map_err(refused(paragraph))?;
        if target_child == source_child || target_child == source_child + 1 {
            return Ok(run);
        }
        let (owned_start, owned_end) = layout
            .extent(source_child)
            .ok_or_else(|| invalid("direct run is missing from the paragraph"))?;
        let xml = Arc::new(paragraph_xml[owned_start..owned_end].to_vec());
        let insertion_offset = layout
            .insertion_offset(target_child)
            .ok_or(Refusal::RunNotFound)
            .map_err(refused(paragraph))?;
        if inside_complex_field(paragraph_xml, insertion_offset, inherited)
            .map_err(refused(paragraph))?
        {
            return Err(refused(paragraph)(Refusal::FieldInterior));
        }
        run_dependencies(
            &paragraph_xml[run_start..layout.children[source_child].1],
            inherited,
        )
        .map_err(refused(paragraph))?;

        let destination = if target_child > source_child {
            target_child - 1
        } else {
            target_child
        };
        let mut candidate = self.try_clone()?;
        candidate.remove_run_fragment(paragraph, Position::new(source_child), run, &xml)?;
        let (candidate_start, candidate_end) = candidate.paragraph_bounds(paragraph)?;
        let candidate_inherited =
            candidate.projected.paragraph_namespaces[paragraph.get()].as_ref();
        let candidate_xml = &candidate.projected.xml_bytes()[candidate_start..candidate_end];
        let candidate_layout =
            paragraph_children(candidate_xml, candidate_inherited).map_err(refused(paragraph))?;
        let moved = Position::new(candidate_layout.runs_before(destination));
        candidate.insert_run_fragment(paragraph, Position::new(destination), moved, &xml)?;
        *self = candidate;
        Ok(moved)
    }

    /// Move one direct run to another direct-body paragraph.
    ///
    /// Cross-paragraph movement is admitted only when the source and target
    /// paragraphs have identical captured namespace scopes. That is a
    /// conservative proof that every prefix used by the retained fragment
    /// keeps the same meaning after the splice; a destination with a
    /// different scope is refused instead of rewriting unknown XML.
    ///
    /// `to` is resolved against the target paragraph before the move. When
    /// both paragraphs are equal this delegates to [`Self::move_run`].
    ///
    /// # Errors
    ///
    /// Returns the same checked run, field, namespace, dependency, malformed
    /// XML, resource, or managed-owner refusal as the focused move. The edit
    /// remains unchanged on every error.
    pub fn move_run_to(
        &mut self,
        from_paragraph: Position,
        run: Position,
        to_paragraph: Position,
        to: RunAnchor,
    ) -> Result<Position> {
        self.ensure_unmanaged("move_run_to")?;
        if from_paragraph == to_paragraph {
            return self.move_run(from_paragraph, run, to);
        }

        let (from_start, from_end) = self.paragraph_bounds(from_paragraph)?;
        let (to_start, to_end) = self.paragraph_bounds(to_paragraph)?;
        let source_inherited = self.projected.paragraph_namespaces[from_paragraph.get()].as_ref();
        let target_inherited = self.projected.paragraph_namespaces[to_paragraph.get()].as_ref();
        if source_inherited != target_inherited {
            return Err(refused(from_paragraph)(Refusal::ComplexContent));
        }
        let source_xml = &self.projected.xml_bytes()[from_start..from_end];
        let target_xml = &self.projected.xml_bytes()[to_start..to_end];
        let source_layout =
            paragraph_children(source_xml, source_inherited).map_err(refused(from_paragraph))?;
        let target_layout =
            paragraph_children(target_xml, target_inherited).map_err(refused(to_paragraph))?;
        let source_child = source_layout
            .run_child(run.get())
            .ok_or(Refusal::RunNotFound)
            .map_err(refused(from_paragraph))?;
        let (run_start, run_end, _) = source_layout.children[source_child];
        if inside_complex_field(source_xml, run_start, source_inherited)
            .map_err(refused(from_paragraph))?
        {
            return Err(refused(from_paragraph)(Refusal::FieldInterior));
        }
        run_dependencies(&source_xml[run_start..run_end], source_inherited)
            .map_err(refused(from_paragraph))?;
        let target_child = run_anchor_child(&target_layout, to).map_err(refused(to_paragraph))?;
        let target_offset = target_layout
            .insertion_offset(target_child)
            .ok_or(Refusal::RunNotFound)
            .map_err(refused(to_paragraph))?;
        if inside_complex_field(target_xml, target_offset, target_inherited)
            .map_err(refused(to_paragraph))?
        {
            return Err(refused(to_paragraph)(Refusal::FieldInterior));
        }
        let (owned_start, owned_end) = source_layout
            .extent(source_child)
            .ok_or_else(|| invalid("direct run is missing from the paragraph"))?;
        let xml = Arc::new(source_xml[owned_start..owned_end].to_vec());

        let mut candidate = self.try_clone()?;
        candidate.remove_run_fragment(from_paragraph, Position::new(source_child), run, &xml)?;
        // Removing a run does not remove its paragraph, so the target
        // paragraph's position remains stable even when it follows the
        // source paragraph.
        let target_after_removal = to_paragraph;
        let (candidate_start, candidate_end) = candidate.paragraph_bounds(target_after_removal)?;
        let candidate_inherited =
            candidate.projected.paragraph_namespaces[target_after_removal.get()].as_ref();
        let candidate_xml = &candidate.projected.xml_bytes()[candidate_start..candidate_end];
        let mut candidate_layout = paragraph_children(candidate_xml, candidate_inherited)
            .map_err(refused(target_after_removal))?;
        if candidate_layout.empty {
            // Match the empty-paragraph route of `insert_run`: opening the
            // retained root is itself an exact, reversible head operation.
            let closed = Arc::new(candidate_xml.to_vec());
            let open = Arc::new(opened(candidate_xml)?);
            candidate.replace_paragraph_properties(target_after_removal, &closed, &open)?;
            let (opened_start, opened_end) = candidate.paragraph_bounds(target_after_removal)?;
            let opened_inherited =
                candidate.projected.paragraph_namespaces[target_after_removal.get()].as_ref();
            candidate_layout = paragraph_children(
                &candidate.projected.xml_bytes()[opened_start..opened_end],
                opened_inherited,
            )
            .map_err(refused(target_after_removal))?;
        }
        let moved = Position::new(candidate_layout.runs_before(target_child));
        candidate.insert_run_fragment(
            target_after_removal,
            Position::new(target_child),
            moved,
            &xml,
        )?;
        *self = candidate;
        Ok(moved)
    }

    /// Swap two direct runs in one paragraph, preserving each run's exact
    /// fragment. Equal positions are an exact no-op.
    ///
    /// # Errors
    ///
    /// Returns the same checked-position, field/dependency, malformed XML,
    /// resource, or managed-owner refusal as [`Self::move_run`].
    pub fn swap_runs(
        &mut self,
        paragraph: Position,
        first: Position,
        second: Position,
    ) -> Result<&mut Self> {
        self.ensure_unmanaged("swap_runs")?;
        if first == second {
            let (start, end) = self.paragraph_bounds(paragraph)?;
            let inherited = self.projected.paragraph_namespaces[paragraph.get()].as_ref();
            let layout = paragraph_children(&self.projected.xml_bytes()[start..end], inherited)
                .map_err(refused(paragraph))?;
            if layout.run_child(first.get()).is_none() {
                return Err(refused(paragraph)(Refusal::RunNotFound));
            }
            return Ok(self);
        }
        let (left, right) = if first < second {
            (first, second)
        } else {
            (second, first)
        };
        let (start, end) = self.paragraph_bounds(paragraph)?;
        let inherited = self.projected.paragraph_namespaces[paragraph.get()].as_ref();
        let paragraph_xml = &self.projected.xml_bytes()[start..end];
        let layout = paragraph_children(paragraph_xml, inherited).map_err(refused(paragraph))?;
        let left_child = layout
            .run_child(left.get())
            .ok_or(Refusal::RunNotFound)
            .map_err(refused(paragraph))?;
        let right_child = layout
            .run_child(right.get())
            .ok_or(Refusal::RunNotFound)
            .map_err(refused(paragraph))?;
        let (left_start, left_end) = layout
            .extent(left_child)
            .ok_or_else(|| invalid("direct run is missing from the paragraph"))?;
        let (right_start, right_end) = layout
            .extent(right_child)
            .ok_or_else(|| invalid("direct run is missing from the paragraph"))?;
        if inside_complex_field(paragraph_xml, left_start, inherited).map_err(refused(paragraph))?
            || inside_complex_field(paragraph_xml, right_start, inherited)
                .map_err(refused(paragraph))?
        {
            return Err(refused(paragraph)(Refusal::FieldInterior));
        }
        let left_element_end = left_start
            .checked_add(element_len(&paragraph_xml[left_start..left_end])?)
            .ok_or_else(|| invalid("direct run offset overflow"))?;
        let right_element_end = right_start
            .checked_add(element_len(&paragraph_xml[right_start..right_end])?)
            .ok_or_else(|| invalid("direct run offset overflow"))?;
        run_dependencies(&paragraph_xml[left_start..left_element_end], inherited)
            .map_err(refused(paragraph))?;
        run_dependencies(&paragraph_xml[right_start..right_element_end], inherited)
            .map_err(refused(paragraph))?;
        let left_offset = layout
            .insertion_offset(left_child)
            .ok_or(Refusal::RunNotFound)
            .map_err(refused(paragraph))?;
        let right_offset = layout
            .insertion_offset(right_child)
            .ok_or(Refusal::RunNotFound)
            .map_err(refused(paragraph))?;
        if inside_complex_field(paragraph_xml, left_offset, inherited)
            .map_err(refused(paragraph))?
            || inside_complex_field(paragraph_xml, right_offset, inherited)
                .map_err(refused(paragraph))?
        {
            return Err(refused(paragraph)(Refusal::FieldInterior));
        }
        let left_fragment = Arc::new(paragraph_xml[left_start..left_end].to_vec());
        let right_fragment = Arc::new(paragraph_xml[right_start..right_end].to_vec());
        // As with paragraphs, operate on one candidate and use the original
        // direct-child slots. This exchanges runs around opaque direct
        // children without rotating those children with either run, and an
        // error in either half leaves the editor unchanged.
        let mut candidate = self.try_clone()?;
        candidate.remove_run_fragment(
            paragraph,
            Position::new(right_child),
            right,
            &right_fragment,
        )?;
        candidate.remove_run_fragment(
            paragraph,
            Position::new(left_child),
            left,
            &left_fragment,
        )?;
        let after_removal = {
            let (start, end) = candidate.paragraph_bounds(paragraph)?;
            let inherited = candidate.projected.paragraph_namespaces[paragraph.get()].as_ref();
            paragraph_children(&candidate.projected.xml_bytes()[start..end], inherited)
                .map_err(refused(paragraph))?
        };
        let moved_right = Position::new(after_removal.runs_before(left_child));
        candidate.insert_run_fragment(
            paragraph,
            Position::new(left_child),
            moved_right,
            &right_fragment,
        )?;
        let after_right = {
            let (start, end) = candidate.paragraph_bounds(paragraph)?;
            let inherited = candidate.projected.paragraph_namespaces[paragraph.get()].as_ref();
            paragraph_children(&candidate.projected.xml_bytes()[start..end], inherited)
                .map_err(refused(paragraph))?
        };
        let moved_left = Position::new(after_right.runs_before(right_child));
        candidate.insert_run_fragment(
            paragraph,
            Position::new(right_child),
            moved_left,
            &left_fragment,
        )?;
        *self = candidate;
        Ok(self)
    }

    /// Swap direct runs in two different direct-body paragraphs while
    /// retaining each run's complete fragment and each paragraph's physical
    /// child slots.
    ///
    /// The paragraphs must have identical captured namespace scopes. This
    /// conservative proof permits unknown run markup to cross the paragraph
    /// boundary without rewriting its prefixes. Opaque direct children,
    /// comments, and whitespace remain in their original paragraph slots.
    ///
    /// # Errors
    ///
    /// Returns a checked-position, namespace, field/dependency, malformed
    /// XML, resource, or managed-owner refusal. All validation and both
    /// splices happen on one candidate, so an error leaves this edit intact.
    pub fn swap_runs_to(
        &mut self,
        first_paragraph: Position,
        first: Position,
        second_paragraph: Position,
        second: Position,
    ) -> Result<&mut Self> {
        self.ensure_unmanaged("swap_runs_to")?;
        if first_paragraph == second_paragraph {
            return self.swap_runs(first_paragraph, first, second);
        }

        let (first_start, first_end) = self.paragraph_bounds(first_paragraph)?;
        let (second_start, second_end) = self.paragraph_bounds(second_paragraph)?;
        let first_inherited = self.projected.paragraph_namespaces[first_paragraph.get()].as_ref();
        let second_inherited = self.projected.paragraph_namespaces[second_paragraph.get()].as_ref();
        if first_inherited != second_inherited {
            return Err(refused(first_paragraph)(Refusal::ComplexContent));
        }
        let first_xml = &self.projected.xml_bytes()[first_start..first_end];
        let second_xml = &self.projected.xml_bytes()[second_start..second_end];
        let first_layout =
            paragraph_children(first_xml, first_inherited).map_err(refused(first_paragraph))?;
        let second_layout =
            paragraph_children(second_xml, second_inherited).map_err(refused(second_paragraph))?;
        let first_child = first_layout
            .run_child(first.get())
            .ok_or(Refusal::RunNotFound)
            .map_err(refused(first_paragraph))?;
        let second_child = second_layout
            .run_child(second.get())
            .ok_or(Refusal::RunNotFound)
            .map_err(refused(second_paragraph))?;
        let (first_fragment_start, first_fragment_end) = first_layout
            .extent(first_child)
            .ok_or_else(|| invalid("direct run is missing from the paragraph"))?;
        let (second_fragment_start, second_fragment_end) = second_layout
            .extent(second_child)
            .ok_or_else(|| invalid("direct run is missing from the paragraph"))?;
        let first_element_end = first_fragment_start
            .checked_add(element_len(
                &first_xml[first_fragment_start..first_fragment_end],
            )?)
            .ok_or_else(|| invalid("direct run offset overflow"))?;
        let second_element_end = second_fragment_start
            .checked_add(element_len(
                &second_xml[second_fragment_start..second_fragment_end],
            )?)
            .ok_or_else(|| invalid("direct run offset overflow"))?;
        if inside_complex_field(first_xml, first_fragment_start, first_inherited)
            .map_err(refused(first_paragraph))?
            || inside_complex_field(second_xml, second_fragment_start, second_inherited)
                .map_err(refused(second_paragraph))?
        {
            return Err(refused(first_paragraph)(Refusal::FieldInterior));
        }
        run_dependencies(
            &first_xml[first_fragment_start..first_element_end],
            first_inherited,
        )
        .map_err(refused(first_paragraph))?;
        run_dependencies(
            &second_xml[second_fragment_start..second_element_end],
            second_inherited,
        )
        .map_err(refused(second_paragraph))?;
        let first_offset = first_layout
            .insertion_offset(first_child)
            .ok_or(Refusal::RunNotFound)
            .map_err(refused(first_paragraph))?;
        let second_offset = second_layout
            .insertion_offset(second_child)
            .ok_or(Refusal::RunNotFound)
            .map_err(refused(second_paragraph))?;
        if inside_complex_field(first_xml, first_offset, first_inherited)
            .map_err(refused(first_paragraph))?
            || inside_complex_field(second_xml, second_offset, second_inherited)
                .map_err(refused(second_paragraph))?
        {
            return Err(refused(first_paragraph)(Refusal::FieldInterior));
        }
        let first_fragment = Arc::new(first_xml[first_fragment_start..first_fragment_end].to_vec());
        let second_fragment =
            Arc::new(second_xml[second_fragment_start..second_fragment_end].to_vec());

        let mut candidate = self.try_clone()?;
        candidate.remove_run_fragment(
            second_paragraph,
            Position::new(second_child),
            second,
            &second_fragment,
        )?;
        candidate.remove_run_fragment(
            first_paragraph,
            Position::new(first_child),
            first,
            &first_fragment,
        )?;
        let first_after_removal = {
            let (start, end) = candidate.paragraph_bounds(first_paragraph)?;
            let inherited =
                candidate.projected.paragraph_namespaces[first_paragraph.get()].as_ref();
            paragraph_children(&candidate.projected.xml_bytes()[start..end], inherited)
                .map_err(refused(first_paragraph))?
        };
        let moved_second = Position::new(first_after_removal.runs_before(first_child));
        candidate.insert_run_fragment(
            first_paragraph,
            Position::new(first_child),
            moved_second,
            &second_fragment,
        )?;
        let second_after_removal = {
            let (start, end) = candidate.paragraph_bounds(second_paragraph)?;
            let inherited =
                candidate.projected.paragraph_namespaces[second_paragraph.get()].as_ref();
            paragraph_children(&candidate.projected.xml_bytes()[start..end], inherited)
                .map_err(refused(second_paragraph))?
        };
        let moved_first = Position::new(second_after_removal.runs_before(second_child));
        candidate.insert_run_fragment(
            second_paragraph,
            Position::new(second_child),
            moved_first,
            &first_fragment,
        )?;
        *self = candidate;
        Ok(self)
    }

    fn paragraph_bounds(&self, position: Position) -> Result<(usize, usize)> {
        let range = self.range(position)?;
        Ok((
            checked_start(range, "paragraph")?,
            checked_end(range, "paragraph")?,
        ))
    }

    /// Replace one whole direct-body paragraph with `replacement` under the
    /// bounded-growth layout policy, and record `operation`.
    fn replace_paragraph_bytes(
        &mut self,
        position: Position,
        replacement: &[u8],
        operation: Operation,
        recorded_bytes: usize,
    ) -> Result<()> {
        let (start, end) = self.paragraph_bounds(position)?;
        let replacement_text_bytes = self.checked_text_total(recorded_bytes)?;
        ensure_utf8(replacement)?;
        let xml = replace_range(self.projected.xml_bytes(), start, end, replacement)?;
        let candidate = self.projected_with_splices_under(
            xml,
            &[ParagraphSplice {
                start,
                end,
                replacement,
            }],
            SplicePolicy::BoundedGrowth,
        )?;
        let readback = candidate.paragraph(position);
        if candidate.paragraph_count() != self.projected.paragraph_count()
            || readback.as_ref().map(crate::Paragraph::xml_bytes) != Some(replacement)
        {
            return Err(invalid("paragraph rewrite failed exact readback"));
        }
        self.operations.push(operation);
        self.replacement_text_bytes = replacement_text_bytes;
        self.projected = candidate;
        Ok(())
    }

    /// Replace a run's property container; the replay of
    /// [`Operation::ReplaceRunProperties`].
    pub(super) fn replace_run_properties(
        &mut self,
        paragraph: Position,
        run: Position,
        before: &Arc<Vec<u8>>,
        after: &Arc<Vec<u8>>,
    ) -> Result<&mut Self> {
        self.ensure_unmanaged("format_run")?;
        self.reserve_operation()?;
        let (start, end) = self.paragraph_bounds(paragraph)?;
        let inherited = self.projected.paragraph_namespaces[paragraph.get()].as_ref();
        let paragraph_xml = &self.projected.xml_bytes()[start..end];
        let (run_start, run_end) =
            direct_run(paragraph_xml, run, inherited).map_err(refused(paragraph))?;
        let run_xml = &paragraph_xml[run_start..run_end];
        let new_run = replace_head(run_xml, b"r", b"rPr", before, after, inherited)?;
        properties::parse_run(&new_run, inherited)?;
        let mut replacement = buffer(staging_add(paragraph_xml.len(), new_run.len())?)?;
        replacement.extend_from_slice(&paragraph_xml[..run_start]);
        replacement.extend_from_slice(&new_run);
        replacement.extend_from_slice(&paragraph_xml[run_end..]);
        if direct_run(&replacement, run, inherited).ok()
            != Some((run_start, run_start + new_run.len()))
        {
            return Err(invalid("run property replacement moved the run"));
        }
        self.replace_paragraph_bytes(
            paragraph,
            &replacement,
            Operation::ReplaceRunProperties {
                paragraph,
                run,
                before: Arc::clone(before),
                after: Arc::clone(after),
            },
            after.len(),
        )?;
        Ok(self)
    }

    /// Replace a paragraph's property container; the replay of
    /// [`Operation::ReplaceParagraphProperties`].
    pub(super) fn replace_paragraph_properties(
        &mut self,
        paragraph: Position,
        before: &Arc<Vec<u8>>,
        after: &Arc<Vec<u8>>,
    ) -> Result<&mut Self> {
        self.ensure_unmanaged("format_paragraph")?;
        self.reserve_operation()?;
        let (start, end) = self.paragraph_bounds(paragraph)?;
        let inherited = self.projected.paragraph_namespaces[paragraph.get()].as_ref();
        let paragraph_xml = &self.projected.xml_bytes()[start..end];
        let replacement = replace_head(paragraph_xml, b"p", b"pPr", before, after, inherited)?;
        properties::parse_paragraph(&replacement, inherited)?;
        self.replace_paragraph_bytes(
            paragraph,
            &replacement,
            Operation::ReplaceParagraphProperties {
                paragraph,
                before: Arc::clone(before),
                after: Arc::clone(after),
            },
            after.len(),
        )?;
        Ok(self)
    }

    /// Insert exact paragraph bytes at a body child index; the replay of
    /// [`Operation::InsertParagraphFragment`].
    pub(super) fn insert_paragraph_fragment(
        &mut self,
        child: Position,
        position: Position,
        xml: &Arc<Vec<u8>>,
    ) -> Result<&mut Self> {
        self.ensure_unmanaged("insert_formatted_paragraph")?;
        self.reserve_operation()?;
        let replacement_text_bytes = self.checked_text_total(xml.len())?;
        ensure_utf8(xml)?;
        let body = body(self)?;
        let offset = body
            .insertion_offset(child.get())
            .ok_or(Error::SemanticPrecondition)?;
        if body.paragraphs_before(child.get()) != position.get() {
            return Err(Error::SemanticPrecondition);
        }
        let element = element_len(xml)?;
        let count = self.projected.paragraph_count();
        let candidate = self.projected.with_inserted_paragraph(
            replace_range(self.projected.xml_bytes(), offset, offset, xml)?,
            offset,
            position.get(),
            xml,
            element,
        )?;
        let inserted = candidate.paragraph(position);
        if let Some(paragraph) = &inserted {
            paragraph.direct_properties()?;
        }
        if candidate.paragraph_count() != count + 1
            || inserted.as_ref().map(crate::Paragraph::xml_bytes) != Some(&xml[..element])
        {
            return Err(invalid("paragraph insertion failed exact readback"));
        }
        self.operations.push(Operation::InsertParagraphFragment {
            child,
            position,
            xml: Arc::clone(xml),
        });
        self.replacement_text_bytes = replacement_text_bytes;
        self.projected = candidate;
        Ok(self)
    }

    /// Remove exact paragraph bytes at a body child index; the replay of
    /// [`Operation::RemoveParagraphFragment`].
    pub(super) fn remove_paragraph_fragment(
        &mut self,
        child: Position,
        position: Position,
        xml: &Arc<Vec<u8>>,
    ) -> Result<&mut Self> {
        self.ensure_unmanaged("remove_paragraph")?;
        self.reserve_operation()?;
        let body = body(self)?;
        let (from, to) = body
            .children
            .get(child.get())
            .filter(|found| found.kind == ChildKind::Paragraph(position.get()))
            .and_then(|_| body.extent(child.get()))
            .ok_or(Error::SemanticPrecondition)?;
        if self.projected.xml_bytes().get(from..to) != Some(xml.as_slice()) {
            return Err(Error::SemanticPrecondition);
        }
        let count = self.projected.paragraph_count();
        let candidate = self.projected.with_removed_paragraph(
            replace_range(self.projected.xml_bytes(), from, to, &[])?,
            position.get(),
            to - from,
        )?;
        if candidate.paragraph_count().checked_add(1) != Some(count) {
            return Err(invalid("paragraph removal failed readback"));
        }
        self.operations.push(Operation::RemoveParagraphFragment {
            child,
            position,
            xml: Arc::clone(xml),
        });
        self.projected = candidate;
        Ok(self)
    }

    /// Insert exact run bytes at a paragraph child index; the replay of
    /// [`Operation::InsertRun`].
    pub(super) fn insert_run_fragment(
        &mut self,
        paragraph: Position,
        child: Position,
        run: Position,
        xml: &Arc<Vec<u8>>,
    ) -> Result<&mut Self> {
        self.ensure_unmanaged("insert_run")?;
        self.reserve_operation()?;
        let (start, end) = self.paragraph_bounds(paragraph)?;
        let inherited = self.projected.paragraph_namespaces[paragraph.get()].as_ref();
        let paragraph_xml = &self.projected.xml_bytes()[start..end];
        let layout = paragraph_children(paragraph_xml, inherited).map_err(refused(paragraph))?;
        let offset = layout
            .insertion_offset(child.get())
            .ok_or(Error::SemanticPrecondition)?;
        // An empty paragraph is opened by its own recorded head replacement
        // first (see `insert_run`).
        if layout.empty || layout.runs_before(child.get()) != run.get() {
            return Err(Error::SemanticPrecondition);
        }
        let mut replacement = buffer(staging_add(paragraph_xml.len(), xml.len())?)?;
        replacement.extend_from_slice(&paragraph_xml[..offset]);
        replacement.extend_from_slice(xml);
        replacement.extend_from_slice(&paragraph_xml[offset..]);
        let element = element_len(xml)?;
        properties::parse_run(&xml[..element], inherited)?;
        let inserted = direct_run(&replacement, run, inherited)
            .map_err(refused(paragraph))
            .map(|(run_start, run_end)| &replacement[run_start..run_end])?;
        if inserted != &xml[..element] {
            return Err(invalid("run insertion failed exact readback"));
        }
        self.replace_paragraph_bytes(
            paragraph,
            &replacement,
            Operation::InsertRun {
                paragraph,
                child,
                run,
                xml: Arc::clone(xml),
            },
            xml.len(),
        )?;
        Ok(self)
    }

    /// Remove exact run bytes at a paragraph child index; the replay of
    /// [`Operation::RemoveRun`].
    pub(super) fn remove_run_fragment(
        &mut self,
        paragraph: Position,
        child: Position,
        run: Position,
        xml: &Arc<Vec<u8>>,
    ) -> Result<&mut Self> {
        self.ensure_unmanaged("remove_run")?;
        self.reserve_operation()?;
        let (start, end) = self.paragraph_bounds(paragraph)?;
        let inherited = self.projected.paragraph_namespaces[paragraph.get()].as_ref();
        let paragraph_xml = &self.projected.xml_bytes()[start..end];
        let layout = paragraph_children(paragraph_xml, inherited).map_err(refused(paragraph))?;
        let is_run = layout
            .children
            .get(child.get())
            .is_some_and(|found| found.2);
        let (run_start, run_end) = layout
            .extent(child.get())
            .ok_or(Error::SemanticPrecondition)?;
        let element = element_len(xml)?;
        if !is_run
            || layout.runs_before(child.get()) != run.get()
            || paragraph_xml.get(run_start..run_end.min(run_start + element))
                != Some(&xml[..element])
        {
            return Err(Error::SemanticPrecondition);
        }
        let mut replacement = buffer(paragraph_xml.len())?;
        replacement.extend_from_slice(&paragraph_xml[..run_start]);
        replacement.extend_from_slice(&paragraph_xml[run_end..]);
        self.replace_paragraph_bytes(
            paragraph,
            &replacement,
            Operation::RemoveRun {
                paragraph,
                child,
                run,
                xml: Arc::clone(xml),
            },
            0,
        )?;
        Ok(self)
    }
}

/// The byte range of direct run `run` within a paragraph fragment.
fn direct_run(
    paragraph_xml: &[u8],
    run: Position,
    inherited: &[(Option<Vec<u8>>, Vec<u8>)],
) -> Result<(usize, usize), Refusal> {
    let layout = paragraph_children(paragraph_xml, inherited)?;
    let child = layout.run_child(run.get()).ok_or(Refusal::RunNotFound)?;
    let (start, end, _) = layout.children[child];
    Ok((start, end))
}

/// A fragment's head: its root start tag through its first direct
/// WordprocessingML `container` child, or just the start tag when it has
/// none, or the whole element when the root is an empty element. Property
/// changes rewrite the head alone.
fn head<'x>(
    xml: &'x [u8],
    root: &[u8],
    container: &[u8],
    inherited: &[(Option<Vec<u8>>, Vec<u8>)],
) -> Result<&'x [u8]> {
    let (tag_end, empty) = root_tag(xml)?;
    if empty {
        return Ok(xml);
    }
    let end = locate_first_child(xml, root, container, inherited)?.map_or(tag_end, |(_, end)| end);
    Ok(&xml[..end])
}

/// Whether a head is a whole empty element (`<w:p .../>`) rather than an
/// open start tag with what follows it.
fn is_empty_head(head: &[u8]) -> Result<bool> {
    Ok(root_tag(head)?.1)
}

/// Replace a fragment's head, which must be `before`, with `after`; see
/// [`head`]. Turning an empty element into an open one appends its end tag,
/// and the reverse requires that nothing but the end tag follows the head.
fn replace_head(
    xml: &[u8],
    root: &[u8],
    container: &[u8],
    before: &[u8],
    after: &[u8],
    inherited: &[(Option<Vec<u8>>, Vec<u8>)],
) -> Result<Vec<u8>> {
    let current = head(xml, root, container, inherited)?;
    if current != before {
        return Err(Error::SemanticPrecondition);
    }
    let rest = &xml[current.len()..];
    let mut close = b"</".to_vec();
    close.extend_from_slice(tag_name(xml)?);
    close.push(b'>');
    let mut output = buffer(staging_add(
        staging_add(after.len(), rest.len())?,
        close.len(),
    )?)?;
    output.extend_from_slice(after);
    match (is_empty_head(before)?, is_empty_head(after)?) {
        (true, false) => output.extend_from_slice(&close),
        (false, true) => {
            if rest != close.as_slice() {
                return Err(Error::SemanticPrecondition);
            }
        },
        (true, true) | (false, false) => output.extend_from_slice(rest),
    }
    Ok(output)
}

/// The byte range of the first direct WordprocessingML `child` of a fragment
/// whose root is `root`.
fn locate_first_child(
    xml: &[u8],
    root: &[u8],
    child: &[u8],
    inherited: &[(Option<Vec<u8>>, Vec<u8>)],
) -> Result<Option<(usize, usize)>> {
    let mut reader = NsReader::from_reader(xml);
    seed_namespaces(&mut reader, inherited)
        .map_err(|_| invalid("invalid inherited Word namespace"))?;
    let origin = ReaderOrigin::of(xml);
    let mut fragment_prefix = FragmentPrefix::Unseen;
    let mut depth = 0usize;
    let mut capture = None::<usize>;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("Word fragment offset overflow"))?;
        let raw = reader.read_event().map_err(|_| {
            Error::Document(crate::Error::Xml("invalid Word formatting XML".into()))
        })?;
        let (namespace, event) = reader.resolver().resolve_event(raw);
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("Word fragment offset overflow"))?;
        match event {
            Event::Start(element) => {
                depth += 1;
                if depth == 1 {
                    fragment_prefix = FragmentPrefix::from_name(element.name());
                    if !is_transaction_fragment_word_name(
                        &namespace,
                        element.name(),
                        root,
                        &fragment_prefix,
                    ) {
                        return Err(invalid("Word fragment has an unexpected root"));
                    }
                } else if depth == 2
                    && is_transaction_fragment_word_name(
                        &namespace,
                        element.name(),
                        child,
                        &fragment_prefix,
                    )
                {
                    capture = Some(start);
                }
            },
            Event::Empty(element) => {
                if depth == 0 {
                    return Ok(None);
                }
                if depth == 1
                    && is_transaction_fragment_word_name(
                        &namespace,
                        element.name(),
                        child,
                        &fragment_prefix,
                    )
                {
                    return Ok(Some((start, end)));
                }
            },
            Event::End(_) => {
                if depth == 2
                    && let Some(capture) = capture
                {
                    return Ok(Some((capture, end)));
                }
                depth = depth.saturating_sub(1);
                if depth == 0 {
                    return Ok(None);
                }
            },
            Event::Eof => return Ok(None),
            _ => {},
        }
    }
}

/// The end of a fragment's root start tag and whether it is an empty
/// element.
fn root_tag(xml: &[u8]) -> Result<(usize, bool)> {
    let mut reader = Reader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    loop {
        let event = reader.read_event().map_err(|_| {
            Error::Document(crate::Error::Xml("invalid Word formatting XML".into()))
        })?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("Word fragment offset overflow"))?;
        match event {
            Event::Start(_) => return Ok((end, false)),
            Event::Empty(_) => return Ok((end, true)),
            Event::Eof => return Err(invalid("Word fragment has no root")),
            _ => {},
        }
    }
}

/// The qualified name of a fragment's root, as written.
fn tag_name(xml: &[u8]) -> Result<&[u8]> {
    let start = xml
        .iter()
        .position(|byte| *byte == b'<')
        .ok_or_else(|| invalid("Word fragment has no root"))?
        + 1;
    let length = xml[start..]
        .iter()
        .position(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n' | b'/' | b'>'))
        .ok_or_else(|| invalid("Word fragment has a malformed root"))?;
    Ok(&xml[start..start + length])
}

/// Turn an empty-element tag into the matching start tag.
fn opened(tag: &[u8]) -> Result<Vec<u8>> {
    let slash = tag
        .len()
        .checked_sub(2)
        .filter(|slash| tag.get(*slash..) == Some(b"/>".as_slice()))
        .ok_or_else(|| invalid("Word fragment has a malformed empty tag"))?;
    let mut output = tag[..slash].to_vec();
    output.push(b'>');
    Ok(output)
}

/// The direct element children of a paragraph fragment.
fn paragraph_children(
    xml: &[u8],
    inherited: &[(Option<Vec<u8>>, Vec<u8>)],
) -> Result<ParagraphChildren, Refusal> {
    let mut reader = NsReader::from_reader(xml);
    seed_namespaces(&mut reader, inherited).map_err(|_| Refusal::ComplexContent)?;
    let origin = ReaderOrigin::of(xml);
    let mut fragment_prefix = FragmentPrefix::Unseen;
    let mut depth = 0usize;
    let mut layout = ParagraphChildren {
        tag_end: 0,
        empty: false,
        close_start: 0,
        prefix: Vec::new(),
        children: Vec::new(),
    };
    let mut open = None::<(usize, bool)>;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or(Refusal::ComplexContent)?;
        let raw = reader
            .read_event()
            .map_err(|_error| Refusal::ComplexContent)?;
        let (namespace, event) = reader.resolver().resolve_event(raw);
        let end = origin
            .offset(reader.buffer_position())
            .ok_or(Refusal::ComplexContent)?;
        match event {
            Event::Start(element) | Event::Empty(element) if depth == 0 => {
                fragment_prefix = FragmentPrefix::from_name(element.name());
                if !is_transaction_fragment_word_name(
                    &namespace,
                    element.name(),
                    b"p",
                    &fragment_prefix,
                ) {
                    return Err(Refusal::ComplexContent);
                }
                layout.prefix = element
                    .name()
                    .prefix()
                    .map_or_else(Vec::new, |prefix| prefix.into_inner().to_vec());
                layout.tag_end = end;
                if end == xml.len() && xml.get(end.wrapping_sub(2)..end) == Some(b"/>") {
                    layout.empty = true;
                    layout.close_start = end;
                    return Ok(layout);
                }
                depth = 1;
            },
            Event::Start(element) => {
                if depth == 1 {
                    let run = is_transaction_fragment_word_name(
                        &namespace,
                        element.name(),
                        b"r",
                        &fragment_prefix,
                    );
                    open = Some((start, run));
                }
                depth += 1;
            },
            Event::Empty(element) => {
                if depth == 1 {
                    let run = is_transaction_fragment_word_name(
                        &namespace,
                        element.name(),
                        b"r",
                        &fragment_prefix,
                    );
                    layout.children.push((start, end, run));
                }
            },
            Event::End(_) => {
                depth = depth.checked_sub(1).ok_or(Refusal::ComplexContent)?;
                if depth == 1
                    && let Some((child_start, run)) = open.take()
                {
                    layout.children.push((child_start, end, run));
                } else if depth == 0 {
                    layout.close_start = start;
                    return Ok(layout);
                }
            },
            Event::Eof => return Err(Refusal::ComplexContent),
            _ => {},
        }
    }
}

/// Whether `offset` of a paragraph fragment lies inside a complex field, as
/// far as the paragraph's own field characters show: a field begun before
/// it is still open there, or a field character after it ends a field it
/// did not begin.
fn inside_complex_field(
    xml: &[u8],
    offset: usize,
    inherited: &[(Option<Vec<u8>>, Vec<u8>)],
) -> Result<bool, Refusal> {
    let mut reader = NsReader::from_reader(xml);
    seed_namespaces(&mut reader, inherited).map_err(|_| Refusal::ComplexContent)?;
    let origin = ReaderOrigin::of(xml);
    let mut fragment_prefix = FragmentPrefix::Unseen;
    let mut depth = 0i64;
    let mut at_offset = None::<i64>;
    let mut lowest_after = i64::MAX;
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or(Refusal::ComplexContent)?;
        if start >= offset && at_offset.is_none() {
            at_offset = Some(depth);
        }
        let raw = reader
            .read_event()
            .map_err(|_error| Refusal::ComplexContent)?;
        let (namespace, event) = reader.resolver().resolve_event(raw);
        let element = match event {
            Event::Start(element) | Event::Empty(element) => element,
            Event::Eof => break,
            _ => continue,
        };
        if matches!(fragment_prefix, FragmentPrefix::Unseen) {
            fragment_prefix = FragmentPrefix::from_name(element.name());
        }
        if !is_transaction_fragment_word_name(
            &namespace,
            element.name(),
            b"fldChar",
            &fragment_prefix,
        ) {
            continue;
        }
        let kind = word_attribute(&element, b"fldCharType", &reader)
            .map_err(|_| Refusal::ComplexContent)?;
        match kind.as_deref() {
            Some("begin") => depth += 1,
            Some("end") => depth -= 1,
            Some("separate") => {},
            _ => return Err(Refusal::ComplexContent),
        }
        if at_offset.is_some() {
            lowest_after = lowest_after.min(depth);
        }
    }
    let at_offset = at_offset.unwrap_or(depth);
    Ok(at_offset > 0 || lowest_after < at_offset)
}

/// Why a paragraph is being taken out.
#[derive(Clone, Copy, PartialEq, Eq)]
enum Removal {
    /// Removed for good: references would be orphaned.
    Delete,
    /// Moved: references travel with it.
    Move,
}

/// What a paragraph holds that its removal or move must not break.
#[derive(Default)]
struct Dependencies {
    malformed: bool,
    section_break: bool,
    open_range: bool,
    story_reference: bool,
    tracked_move: bool,
    mark_revision: bool,
}

impl Dependencies {
    fn refuse(&self, removal: Removal) -> Result<(), Refusal> {
        if self.malformed {
            return Err(Refusal::ComplexContent);
        }
        if self.section_break {
            return Err(Refusal::SectionBreak);
        }
        if self.open_range {
            return Err(Refusal::OpenRange);
        }
        if self.mark_revision {
            return Err(Refusal::TrackedChange);
        }
        if removal == Removal::Delete && (self.story_reference || self.tracked_move) {
            return Err(Refusal::StoryReference);
        }
        Ok(())
    }
}

/// Range markers and the family that pairs them, by local name.
fn range_marker(local: &[u8]) -> Option<(&'static str, bool)> {
    Some(match local {
        b"bookmarkStart" => ("bookmark", true),
        b"bookmarkEnd" => ("bookmark", false),
        b"commentRangeStart" => ("comment", true),
        b"commentRangeEnd" => ("comment", false),
        b"moveFromRangeStart" => ("moveFrom", true),
        b"moveFromRangeEnd" => ("moveFrom", false),
        b"moveToRangeStart" => ("moveTo", true),
        b"moveToRangeEnd" => ("moveTo", false),
        b"permStart" => ("perm", true),
        b"permEnd" => ("perm", false),
        b"customXmlInsRangeStart" => ("customXmlIns", true),
        b"customXmlInsRangeEnd" => ("customXmlIns", false),
        b"customXmlDelRangeStart" => ("customXmlDel", true),
        b"customXmlDelRangeEnd" => ("customXmlDel", false),
        b"customXmlMoveFromRangeStart" => ("customXmlMoveFrom", true),
        b"customXmlMoveFromRangeEnd" => ("customXmlMoveFrom", false),
        b"customXmlMoveToRangeStart" => ("customXmlMoveTo", true),
        b"customXmlMoveToRangeEnd" => ("customXmlMoveTo", false),
        _ => return None,
    })
}

/// Scan a paragraph for what its removal or move must not break.
fn paragraph_dependencies(xml: &[u8], inherited: &[(Option<Vec<u8>>, Vec<u8>)]) -> Dependencies {
    let mut dependencies = Dependencies::default();
    if scan_dependencies(xml, &mut dependencies, inherited).is_err() {
        dependencies.malformed = true;
    }
    dependencies
}

fn scan_dependencies(
    xml: &[u8],
    dependencies: &mut Dependencies,
    inherited: &[(Option<Vec<u8>>, Vec<u8>)],
) -> Result<(), ()> {
    let mut reader = NsReader::from_reader(xml);
    seed_namespaces(&mut reader, inherited).map_err(|_| ())?;
    let mut fragment_prefix = FragmentPrefix::Unseen;
    // Each open element: whether it is a Word element and its local name.
    let mut path: Vec<(bool, Vec<u8>)> = Vec::new();
    let mut ranges: Vec<(&'static str, String, bool)> = Vec::new();
    let mut field_depth = 0i64;
    loop {
        let raw = reader.read_event().map_err(|_error| ())?;
        let (namespace, event) = reader.resolver().resolve_event(raw);
        let (element, empty) = match event {
            Event::Start(element) => (element, false),
            Event::Empty(element) => (element, true),
            Event::End(_) => {
                path.pop().ok_or(())?;
                continue;
            },
            Event::Eof => break,
            _ => continue,
        };
        if matches!(fragment_prefix, FragmentPrefix::Unseen) {
            fragment_prefix = FragmentPrefix::from_name(element.name());
        }
        let word = is_fragment_word_element(&namespace, &fragment_prefix);
        let local = element.local_name().as_ref().to_vec();
        if word {
            let parent = |depth: usize, name: &[u8]| {
                path.len() >= depth
                    && path[path.len() - depth].0
                    && path[path.len() - depth].1 == name
            };
            match local.as_slice() {
                b"sectPr" if parent(1, b"pPr") => dependencies.section_break = true,
                b"ins" | b"del" | b"moveFrom" | b"moveTo"
                    if parent(1, b"rPr") && parent(2, b"pPr") =>
                {
                    dependencies.mark_revision = true;
                },
                b"moveFrom" | b"moveTo" => dependencies.tracked_move = true,
                b"commentReference" | b"footnoteReference" | b"endnoteReference" => {
                    dependencies.story_reference = true;
                },
                b"fldChar" => {
                    let kind = word_attribute(&element, b"fldCharType", &reader)?;
                    match kind.as_deref() {
                        Some("begin") => field_depth += 1,
                        Some("end") => field_depth -= 1,
                        Some("separate") => {},
                        _ => return Err(()),
                    }
                    if field_depth < 0 {
                        dependencies.open_range = true;
                    }
                },
                name => {
                    if let Some((family, start)) = range_marker(name) {
                        let id = word_attribute(&element, b"id", &reader)?.ok_or(())?;
                        ranges.push((family, id, start));
                    }
                },
            }
        }
        if !empty {
            path.push((word, local));
        }
    }
    if field_depth != 0 {
        dependencies.open_range = true;
    }
    for (family, id, start) in &ranges {
        if !ranges.iter().any(|(other, other_id, other_start)| {
            other == family && other_id == id && other_start != start
        }) {
            dependencies.open_range = true;
        }
    }
    Ok(())
}

/// The decoded value of an attribute by local name.
fn word_attribute(
    element: &quick_xml::events::BytesStart<'_>,
    local: &[u8],
    reader: &NsReader<&[u8]>,
) -> Result<Option<String>, ()> {
    properties::audit_attributes(element, reader.resolver(), reader.decoder()).map_err(|_| ())?;
    crate::namespace::word_attribute_value(element, local, reader.decoder(), reader.resolver())
        .map_err(|_| ())
}

/// Refuse to remove a run that holds a field character or field code, or a
/// story reference.
fn run_dependencies(xml: &[u8], inherited: &[(Option<Vec<u8>>, Vec<u8>)]) -> Result<(), Refusal> {
    let mut reader = NsReader::from_reader(xml);
    seed_namespaces(&mut reader, inherited).map_err(|_| Refusal::ComplexContent)?;
    let mut fragment_prefix = FragmentPrefix::Unseen;
    loop {
        let raw = reader
            .read_event()
            .map_err(|_error| Refusal::ComplexContent)?;
        let (namespace, event) = reader.resolver().resolve_event(raw);
        let element = match event {
            Event::Start(element) | Event::Empty(element) => element,
            Event::Eof => return Ok(()),
            _ => continue,
        };
        if matches!(fragment_prefix, FragmentPrefix::Unseen) {
            fragment_prefix = FragmentPrefix::from_name(element.name());
        }
        if !is_fragment_word_element(&namespace, &fragment_prefix) {
            continue;
        }
        match element.local_name().as_ref() {
            b"fldChar" | b"instrText" | b"delInstrText" => return Err(Refusal::OpenRange),
            b"commentReference" | b"footnoteReference" | b"endnoteReference" => {
                return Err(Refusal::StoryReference);
            },
            _ => {},
        }
    }
}

fn body(edit: &Edit) -> Result<Body> {
    body_of(edit)
}

/// The direct element children of the projected document's body.
fn body_of(edit: &Edit) -> Result<Body> {
    body_of_snapshot(&edit.projected)
}

/// The direct element children of a snapshot's body.
fn body_of_snapshot(snapshot: &super::Snapshot) -> Result<Body> {
    let xml = snapshot.xml_bytes();
    let (content_start, prefix) = body_start(xml, snapshot.conformance)?;
    let to_usize =
        |value: u32| usize::try_from(value).map_err(|_| invalid("body offset does not fit usize"));
    let mut known = Vec::with_capacity(
        snapshot.paragraphs.len() + snapshot.tables.len() + snapshot.block_controls.len(),
    );
    let families: [(&[super::Range], fn(usize) -> ChildKind); 3] = [
        (&snapshot.paragraphs, ChildKind::Paragraph),
        (&snapshot.tables, ChildKind::Table),
        (&snapshot.block_controls, ChildKind::ContentControl),
    ];
    for (ranges, kind) in families {
        for (index, range) in ranges.iter().enumerate() {
            let start = to_usize(range.start)?;
            known.push(BodyChild {
                start,
                end: start + to_usize(range.length)?,
                kind: kind(index),
            });
        }
    }
    known.sort_unstable_by_key(|child| child.start);
    let mut children = Vec::with_capacity(known.len() + 1);
    let mut cursor = content_start;
    for child in known {
        gap_children(xml, cursor, child.start, &mut children)?;
        cursor = child.end;
        children.push(child);
    }
    let content_end = to_usize(snapshot.content_end)?;
    gap_children(xml, cursor, content_end, &mut children)?;
    let (section, body_end) = if xml.get(content_end..content_end + 2) == Some(b"</") {
        (None, content_end)
    } else {
        let mut tail = Vec::new();
        let end = gap_children(xml, content_end, xml.len(), &mut tail)?;
        let section = tail
            .first()
            .copied()
            .ok_or_else(|| invalid("body-final section properties are missing"))?;
        children.push(section);
        (Some(children.len() - 1), end)
    };
    Ok(Body {
        body_end,
        children,
        section,
        prefix,
    })
}

/// The offset just past the body's start tag, and a prefix bound to the
/// document's WordprocessingML namespace there.
fn body_start(xml: &[u8], conformance: Conformance) -> Result<(usize, Option<Vec<u8>>)> {
    let mut reader = NsReader::from_reader(xml);
    let origin = ReaderOrigin::of(xml);
    let mut depth = 0usize;
    let namespace = conformance.namespace().as_bytes();
    loop {
        let event = reader.read_event().map_err(|_| {
            Error::Document(crate::Error::Xml("invalid Word formatting XML".into()))
        })?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("document offset overflow"))?;
        match event {
            Event::Start(element) => {
                depth += 1;
                if depth == 2
                    && element.local_name().as_ref() == b"body"
                    && super::is_wordprocessing_namespace(
                        &reader.resolver().resolve_element(element.name()).0,
                    )
                {
                    let resolver = reader.resolver();
                    let bound = |prefix: &[u8]| {
                        let mut name = prefix.to_vec();
                        name.extend_from_slice(b":p");
                        matches!(
                            resolver.resolve_element(QName(&name)).0,
                            ResolveResult::Bound(found) if found.into_inner() == namespace
                        )
                    };
                    let prefix = if bound(b"w") {
                        Some(b"w".to_vec())
                    } else {
                        resolver
                            .bindings()
                            .filter_map(|(declaration, found)| match declaration {
                                quick_xml::name::PrefixDeclaration::Named(prefix)
                                    if found.into_inner() == namespace =>
                                {
                                    Some(prefix.to_vec())
                                },
                                _ => None,
                            })
                            .find(|prefix| bound(prefix))
                    };
                    return Ok((end, prefix));
                }
            },
            Event::End(_) => depth = depth.saturating_sub(1),
            Event::Eof => return Err(invalid("main document has no body")),
            _ => {},
        }
    }
}

/// Record the complete elements of `xml[start..end]` as other body children;
/// return the offset where the slice's content ends (its first unmatched end
/// tag, or `end`).
fn gap_children(
    xml: &[u8],
    start: usize,
    end: usize,
    children: &mut Vec<BodyChild>,
) -> Result<usize> {
    let slice = xml
        .get(start..end)
        .ok_or_else(|| invalid("body gap is outside the document"))?;
    let mut reader = Reader::from_reader(slice);
    // The slice ends inside `w:body`: its closing tag is the first end tag
    // the slice does not open, which this scan reports instead of refusing.
    reader.config_mut().allow_unmatched_ends = true;
    let origin = ReaderOrigin::of(slice);
    let mut depth = 0usize;
    let mut open = 0usize;
    loop {
        let event_start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("body offset overflow"))?;
        let event = reader.read_event().map_err(|_| {
            Error::Document(crate::Error::Xml("invalid Word formatting XML".into()))
        })?;
        let event_end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| invalid("body offset overflow"))?;
        match event {
            Event::Start(_) => {
                if depth == 0 {
                    open = event_start;
                }
                depth += 1;
            },
            Event::Empty(_) if depth == 0 => children.push(BodyChild {
                start: start + event_start,
                end: start + event_end,
                kind: ChildKind::Other,
            }),
            Event::End(_) => {
                if depth == 0 {
                    return Ok(start + event_start);
                }
                depth -= 1;
                if depth == 0 {
                    children.push(BodyChild {
                        start: start + open,
                        end: start + event_end,
                        kind: ChildKind::Other,
                    });
                }
            },
            Event::Eof => return Ok(end),
            _ => {},
        }
    }
}

fn qualified(
    output: &mut Vec<u8>,
    prefix: &[u8],
    local: &[u8],
) -> Result<(), properties::SpliceError> {
    let bytes = staging_add(
        staging_add(prefix.len(), local.len()).map_err(properties::SpliceError::Transaction)?,
        128,
    )
    .map_err(properties::SpliceError::Transaction)?;
    staging_add(output.len(), bytes).map_err(properties::SpliceError::Transaction)?;
    output
        .try_reserve_exact(bytes)
        .map_err(|source| crate::Error::Allocation {
            resource: "Word formatting tag",
            source,
        })?;
    if !prefix.is_empty() {
        output.extend_from_slice(prefix);
        output.push(b':');
    }
    output.extend_from_slice(local);
    Ok(())
}

/// Open a generated root element, declaring `prefix` when the document does
/// not bind it.
fn open_root(
    output: &mut Vec<u8>,
    prefix: &[u8],
    local: &[u8],
    declare: Option<Conformance>,
) -> Result<(), properties::SpliceError> {
    let bytes = staging_add(
        staging_add(prefix.len(), local.len()).map_err(properties::SpliceError::Transaction)?,
        128,
    )
    .map_err(properties::SpliceError::Transaction)?;
    staging_add(output.len(), bytes).map_err(properties::SpliceError::Transaction)?;
    output
        .try_reserve_exact(bytes)
        .map_err(|source| crate::Error::Allocation {
            resource: "Word formatting tag",
            source,
        })?;
    output.push(b'<');
    qualified(output, prefix, local)?;
    if let Some(conformance) = declare {
        output.extend_from_slice(b" xmlns:");
        output.extend_from_slice(prefix);
        output.extend_from_slice(b"=\"");
        output.extend_from_slice(conformance.namespace().as_bytes());
        output.push(b'"');
    }
    Ok(())
}

/// The markup of a new run, written with `prefix` (declared when it is
/// empty), read back before it is returned.
fn run_markup(
    prefix: &[u8],
    conformance: Conformance,
    run: &NewRun,
) -> Result<Vec<u8>, properties::SpliceError> {
    let (prefix, declare) = if prefix.is_empty() {
        (b"w".as_slice(), Some(conformance))
    } else {
        (prefix, None)
    };
    let capacity = admit_new_run(prefix, run)?;
    let mut output = buffer(capacity).map_err(properties::SpliceError::Transaction)?;
    open_root(&mut output, prefix, b"r", declare)?;
    output.push(b'>');
    write_run_body(&mut output, prefix, run)?;
    output.extend_from_slice(b"</");
    qualified(&mut output, prefix, b"r")?;
    output.push(b'>');
    verify_run(&output, run)?;
    Ok(output)
}

fn write_run_body(
    output: &mut Vec<u8>,
    prefix: &[u8],
    run: &NewRun,
) -> Result<(), properties::SpliceError> {
    validate_authored_run_content(&run.text)?;
    if let Some(properties) = properties::run_properties_element(prefix, &run.properties)? {
        append_output(output, &properties).map_err(properties::SpliceError::Transaction)?;
    }
    let content = try_run_content_fragment(prefix, b"t", &run.text)
        .map_err(properties::SpliceError::Transaction)?;
    append_output(output, &content).map_err(properties::SpliceError::Transaction)?;
    Ok(())
}

/// Read a generated run back: its properties and its text must be exactly
/// what was asked.
fn verify_run(xml: &[u8], run: &NewRun) -> Result<(), properties::SpliceError> {
    let readback = properties::parse_run(xml, &[])?;
    let text = scan_text_owner(xml, b"r").map(|owner| owner.text);
    if !properties::same_run(&readback, &run.properties) || text.as_deref() != Ok(run.text.as_str())
    {
        return Err(crate::Error::InvalidFormat("new run failed typed readback".into()).into());
    }
    Ok(())
}

/// The markup of a new paragraph, read back before it is returned.
fn paragraph_markup(
    body_prefix: Option<&[u8]>,
    conformance: Conformance,
    paragraph: &NewParagraph,
    strict: bool,
) -> Result<Vec<u8>, properties::SpliceError> {
    let (prefix, declare) = match body_prefix {
        Some(prefix) => (prefix, None),
        None => (b"w".as_slice(), Some(conformance)),
    };
    let mut capacity = properties::admit_paragraph(&paragraph.properties)?;
    for run in &paragraph.runs {
        capacity = staging_add(capacity, admit_new_run(prefix, run)?)
            .map_err(properties::SpliceError::Transaction)?;
    }
    let properties =
        properties::paragraph_properties_element(prefix, &paragraph.properties, strict)?;
    let mut output = buffer(capacity).map_err(properties::SpliceError::Transaction)?;
    open_root(&mut output, prefix, b"p", declare)?;
    if properties.is_none() && paragraph.runs.is_empty() {
        output.extend_from_slice(b"/>");
    } else {
        output.push(b'>');
        if let Some(properties) = properties {
            output.extend_from_slice(&properties);
        }
        for run in &paragraph.runs {
            let start = output.len();
            open_root(&mut output, prefix, b"r", None)?;
            output.push(b'>');
            write_run_body(&mut output, prefix, run)?;
            output.extend_from_slice(b"</");
            qualified(&mut output, prefix, b"r")?;
            output.push(b'>');
            if declare.is_none() {
                verify_run(&output[start..], run)?;
            }
        }
        output.extend_from_slice(b"</");
        qualified(&mut output, prefix, b"p")?;
        output.push(b'>');
    }
    let readback = properties::parse_paragraph(&output, &[])?;
    if !properties::same_paragraph(&readback, &paragraph.properties) {
        return Err(
            crate::Error::InvalidFormat("new paragraph failed typed readback".into()).into(),
        );
    }
    Ok(output)
}

/// [`direct_run`] for the tests of this module's parent.
#[cfg(test)]
pub(super) fn direct_run_for_tests(paragraph_xml: &[u8], run: usize) -> (usize, usize) {
    direct_run(paragraph_xml, Position::new(run), &[]).unwrap()
}

/// The number of direct runs of a paragraph, for tests.
#[cfg(test)]
pub(super) fn direct_run_count_for_tests(paragraph_xml: &[u8]) -> usize {
    paragraph_children(paragraph_xml, &[]).unwrap().run_count()
}

/// The length of a fragment's [`head`], for tests.
#[cfg(test)]
pub(super) fn head_len_for_tests(xml: &[u8], root: &[u8], container: &[u8]) -> usize {
    head(xml, root, container, &[]).unwrap().len()
}

/// The start offsets of the projected body's direct element children, for
/// tests.
#[cfg(test)]
pub(super) fn body_for_tests(edit: &Edit) -> Vec<usize> {
    body_of(edit)
        .unwrap()
        .children
        .iter()
        .map(|child| child.start)
        .collect()
}

fn seed_namespaces(
    reader: &mut NsReader<&[u8]>,
    inherited: &[(Option<Vec<u8>>, Vec<u8>)],
) -> std::result::Result<(), ()> {
    for (prefix, namespace) in inherited {
        let prefix = prefix.as_deref().map_or(
            quick_xml::name::PrefixDeclaration::Default,
            quick_xml::name::PrefixDeclaration::Named,
        );
        reader
            .resolver_mut()
            .add(prefix, quick_xml::name::Namespace(namespace))
            .map_err(|_| ())?;
    }
    Ok(())
}

impl std::fmt::Debug for NewRun {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("NewRun").finish_non_exhaustive()
    }
}
impl std::fmt::Debug for NewParagraph {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("NewParagraph").finish_non_exhaustive()
    }
}

const MAX_FORMATTING_STAGING_BYTES: usize = 64 * 1024 * 1024;
fn staging_add(left: usize, right: usize) -> Result<usize> {
    let observed = left.checked_add(right).unwrap_or(usize::MAX);
    if observed > MAX_FORMATTING_STAGING_BYTES {
        return Err(Error::Limit {
            kind: super::super::LimitKind::FormattingStagingBytes,
            observed,
            maximum: MAX_FORMATTING_STAGING_BYTES,
        });
    }
    Ok(observed)
}
fn buffer(capacity: usize) -> Result<Vec<u8>> {
    staging_add(0, capacity)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(capacity)
        .map_err(|source| crate::Error::Allocation {
            resource: "Word formatting output",
            source,
        })?;
    Ok(output)
}
fn admit_new_run(prefix: &[u8], run: &NewRun) -> Result<usize, properties::SpliceError> {
    let properties = properties::admit_run(&run.properties)?;
    let text = super::run_content_fragment_len(prefix, b"t", &run.text)
        .map_err(properties::SpliceError::Transaction)?;
    staging_add(properties, text).map_err(properties::SpliceError::Transaction)
}

fn append_output(output: &mut Vec<u8>, bytes: &[u8]) -> Result<()> {
    staging_add(output.len(), bytes.len())?;
    output
        .try_reserve_exact(bytes.len())
        .map_err(|source| crate::Error::Allocation {
            resource: "Word formatting output",
            source,
        })?;
    output.extend_from_slice(bytes);
    Ok(())
}
