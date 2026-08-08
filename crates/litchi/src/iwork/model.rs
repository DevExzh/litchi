use std::{fmt, sync::Arc};

use litchi_iwa_structured::StructuredData;

use super::Format;

struct State {
    format: Format,
    data: StructuredData,
}

/// A successfully decoded, immutable Apple iWork document.
///
/// Construction is eager and fallible; every operation on a successful value
/// is an infallible view over one archive-free semantic snapshot. Cloning the
/// value shares that snapshot without cloning tables, slides, sections, cells,
/// or text.
#[derive(Clone)]
pub struct Document {
    state: Arc<State>,
}

impl Document {
    pub(super) fn from_data(format: Format, data: StructuredData) -> Self {
        Self {
            state: Arc::new(State { format, data }),
        }
    }

    /// Return the detected application family.
    #[must_use]
    pub fn format(&self) -> Format {
        self.state.format
    }

    /// Capture a cheap, lifetime-independent semantic snapshot.
    #[must_use]
    pub fn snapshot(&self) -> Snapshot {
        Snapshot {
            state: Arc::clone(&self.state),
        }
    }
}

impl fmt::Debug for Document {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Document")
            .field("format", &self.format())
            .finish_non_exhaustive()
    }
}

/// A cheap handle to one immutable, archive-free iWork semantic snapshot.
#[derive(Clone)]
pub struct Snapshot {
    state: Arc<State>,
}

impl Snapshot {
    /// Return the detected application family.
    #[must_use]
    pub fn format(&self) -> Format {
        self.state.format
    }

    /// Return the number of semantic Numbers tables.
    #[must_use]
    pub fn table_count(&self) -> usize {
        self.state.data.table_count()
    }

    /// Return the number of semantic Keynote slides.
    #[must_use]
    pub fn slide_count(&self) -> usize {
        self.state.data.slide_count()
    }

    /// Return the number of semantic Pages sections.
    #[must_use]
    pub fn section_count(&self) -> usize {
        self.state.data.section_count()
    }

    /// Return whether the snapshot has no tables, slides, or sections.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.state.data.is_empty()
    }

    /// Return deterministic semantic collection counts without allocating.
    #[must_use]
    pub fn summary(&self) -> Summary {
        Summary {
            tables: self.table_count(),
            slides: self.slide_count(),
            sections: self.section_count(),
        }
    }

    /// Select a table by checked zero-based semantic position.
    #[must_use]
    pub fn table(&self, position: usize) -> Option<Table> {
        self.state.data.table(position).map(|_table| Table {
            state: Arc::clone(&self.state),
            position,
        })
    }

    /// Iterate over lifetime-independent table handles in source order.
    #[must_use = "iterators are lazy and do nothing unless consumed"]
    pub fn tables(&self) -> impl ExactSizeIterator<Item = Table> + '_ {
        (0..self.table_count()).map(|position| Table {
            state: Arc::clone(&self.state),
            position,
        })
    }

    /// Select a slide by checked zero-based semantic position.
    #[must_use]
    pub fn slide(&self, position: usize) -> Option<Slide> {
        self.state.data.slide(position).map(|_slide| Slide {
            state: Arc::clone(&self.state),
            position,
        })
    }

    /// Iterate over lifetime-independent slide handles in presentation order.
    #[must_use = "iterators are lazy and do nothing unless consumed"]
    pub fn slides(&self) -> impl ExactSizeIterator<Item = Slide> + '_ {
        (0..self.slide_count()).map(|position| Slide {
            state: Arc::clone(&self.state),
            position,
        })
    }

    /// Select a section by checked zero-based semantic position.
    #[must_use]
    pub fn section(&self, position: usize) -> Option<Section> {
        self.state.data.section(position).map(|_section| Section {
            state: Arc::clone(&self.state),
            position,
        })
    }

    /// Iterate over lifetime-independent section handles in document order.
    #[must_use = "iterators are lazy and do nothing unless consumed"]
    pub fn sections(&self) -> impl ExactSizeIterator<Item = Section> + '_ {
        (0..self.section_count()).map(|position| Section {
            state: Arc::clone(&self.state),
            position,
        })
    }

    /// Iterate over borrowed semantic text without allocating.
    ///
    /// Values are ordered by table, slide, then section. Within a slide the
    /// order is title, ordinary content, additional rich text, then notes.
    #[must_use = "iterators are lazy and do nothing unless consumed"]
    pub fn iter_text(&self) -> impl Iterator<Item = Text<'_>> + '_ {
        let tables = self.state.data.tables().iter().flat_map(table_text);
        let slides = self.state.data.slides().iter().flat_map(slide_text);
        let sections = self.state.data.sections().iter().flat_map(section_text);
        tables.chain(slides).chain(sections)
    }

    /// Collect semantic text as explicitly owned strings.
    #[must_use]
    pub fn all_text(&self) -> Vec<String> {
        self.iter_text()
            .map(|item| item.value().to_owned())
            .collect()
    }
}

impl fmt::Debug for Snapshot {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Snapshot")
            .field("format", &self.format())
            .field("summary", &self.summary())
            .finish_non_exhaustive()
    }
}

/// Collection counts for one archive-free semantic snapshot.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Summary {
    tables: usize,
    slides: usize,
    sections: usize,
}

impl Summary {
    /// Return the Numbers table count.
    #[must_use]
    pub const fn tables(self) -> usize {
        self.tables
    }

    /// Return the Keynote slide count.
    #[must_use]
    pub const fn slides(self) -> usize {
        self.slides
    }

    /// Return the Pages section count.
    #[must_use]
    pub const fn sections(self) -> usize {
        self.sections
    }
}

impl fmt::Display for Summary {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            formatter,
            "Tables: {}, Slides: {}, Sections: {}",
            self.tables, self.slides, self.sections
        )
    }
}

/// A cheap, lifetime-independent handle to one Numbers table.
#[derive(Clone)]
pub struct Table {
    state: Arc<State>,
    position: usize,
}

impl Table {
    fn inner(&self) -> &litchi_numbers::Table {
        self.state
            .data
            .table(self.position)
            .unwrap_or_else(|| unreachable!("validated iWork table handle"))
    }

    /// Return the table's zero-based position in the snapshot.
    #[must_use]
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Borrow the semantic table name.
    #[must_use]
    pub fn name(&self) -> &str {
        self.inner().name()
    }

    /// Return the declared row count.
    #[must_use]
    pub fn row_count(&self) -> u32 {
        self.inner().row_count()
    }

    /// Return the declared column count.
    #[must_use]
    pub fn column_count(&self) -> u32 {
        self.inner().column_count()
    }

    /// Return the number of materialized sparse cells.
    #[must_use]
    pub fn cell_count(&self) -> usize {
        self.inner().cell_count()
    }

    /// Return the number of materialized non-empty cells.
    #[must_use]
    pub fn non_empty_cell_count(&self) -> usize {
        self.inner().non_empty_cell_count()
    }

    /// Look up a zero-based coordinate without conflating a missing cell with
    /// an explicitly stored empty value.
    ///
    /// `None` means the coordinate lies outside the declared table extent.
    #[must_use]
    pub fn cell(&self, row: u32, column: u32) -> Option<CellView<'_>> {
        if row >= self.row_count() || column >= self.column_count() {
            return None;
        }
        let position = litchi_numbers::CellPosition::new(row, column);
        Some(match self.inner().view(position) {
            litchi_numbers::View::Missing => CellView::Missing,
            litchi_numbers::View::Covered => CellView::Covered,
            litchi_numbers::View::Stored(value) => CellView::Stored(map_value(value)),
        })
    }

    /// Iterate over materialized sparse cells in row-major order.
    #[must_use = "iterators are lazy and do nothing unless consumed"]
    pub fn cells(&self) -> impl ExactSizeIterator<Item = Cell<'_>> + '_ {
        self.inner().iter_cells().map(|cell| {
            let position = cell.position();
            Cell {
                row: position.row(),
                column: position.column(),
                value: map_value(cell.value()),
            }
        })
    }

    /// Iterate over the table name, headers, and textual cells without
    /// allocating.
    #[must_use = "iterators are lazy and do nothing unless consumed"]
    pub fn iter_text(&self) -> impl Iterator<Item = Text<'_>> + '_ {
        table_text(self.inner())
    }
}

impl fmt::Debug for Table {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Table")
            .field("position", &self.position)
            .field("rows", &self.row_count())
            .field("columns", &self.column_count())
            .field("cells", &self.cell_count())
            .finish_non_exhaustive()
    }
}

/// The semantic result of looking up a Numbers coordinate.
#[derive(Debug, Clone, Copy, PartialEq)]
#[non_exhaustive]
pub enum CellView<'a> {
    /// No semantic cell is stored at the coordinate.
    Missing,
    /// The coordinate is covered by a merged region.
    Covered,
    /// A semantic value is explicitly stored at the coordinate.
    Stored(Value<'a>),
}

/// One materialized sparse Numbers cell.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Cell<'a> {
    row: u32,
    column: u32,
    value: Value<'a>,
}

impl<'a> Cell<'a> {
    /// Return the zero-based row.
    #[must_use]
    pub const fn row(self) -> u32 {
        self.row
    }

    /// Return the zero-based column.
    #[must_use]
    pub const fn column(self) -> u32 {
        self.column
    }

    /// Return the borrowed semantic value.
    #[must_use]
    pub const fn value(self) -> Value<'a> {
        self.value
    }
}

/// A borrowed, facade-owned Numbers cell value.
#[derive(Debug, Clone, Copy, PartialEq)]
#[non_exhaustive]
pub enum Value<'a> {
    /// An explicitly stored empty value.
    Empty,
    /// User-entered text.
    Text(&'a str),
    /// A finite number.
    Number(f64),
    /// A Boolean value.
    Boolean(bool),
    /// Finite seconds since Apple's 2001-01-01 UTC epoch.
    Date(f64),
    /// A finite duration in seconds.
    Duration(f64),
    /// Formula source or rendered formula text.
    Formula(&'a str),
    /// Producer-reported cell error text.
    Error(&'a str),
}

impl Value<'_> {
    /// Return the stable semantic value category.
    #[must_use]
    pub const fn kind(self) -> ValueKind {
        match self {
            Self::Empty => ValueKind::Empty,
            Self::Text(_) => ValueKind::Text,
            Self::Number(_) => ValueKind::Number,
            Self::Boolean(_) => ValueKind::Boolean,
            Self::Date(_) => ValueKind::Date,
            Self::Duration(_) => ValueKind::Duration,
            Self::Formula(_) => ValueKind::Formula,
            Self::Error(_) => ValueKind::Error,
        }
    }
}

/// Stable category of a Numbers cell value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ValueKind {
    /// Explicit empty.
    Empty,
    /// User text.
    Text,
    /// Finite number.
    Number,
    /// Boolean.
    Boolean,
    /// Apple-epoch date.
    Date,
    /// Duration.
    Duration,
    /// Formula text.
    Formula,
    /// Producer error text.
    Error,
}

/// A cheap, lifetime-independent handle to one Keynote slide.
#[derive(Clone)]
pub struct Slide {
    state: Arc<State>,
    position: usize,
}

impl Slide {
    fn inner(&self) -> &litchi_keynote::Slide {
        self.state
            .data
            .slide(self.position)
            .unwrap_or_else(|| unreachable!("validated iWork slide handle"))
    }

    /// Return the zero-based presentation position.
    #[must_use]
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Return whether presentation playback skips this slide.
    #[must_use]
    pub fn is_skipped(&self) -> bool {
        self.inner().is_skipped()
    }

    /// Borrow the optional navigator name, distinct from visible title text.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.inner().name()
    }

    /// Borrow the optional visible slide title.
    #[must_use]
    pub fn title(&self) -> Option<&str> {
        self.inner().title()
    }

    /// Iterate over ordinary slide text blocks in source order.
    #[must_use = "iterators are lazy and do nothing unless consumed"]
    pub fn text_blocks(&self) -> impl ExactSizeIterator<Item = &str> + '_ {
        self.inner().text_content().iter().map(String::as_str)
    }

    /// Iterate over additional rich text in source order.
    #[must_use = "iterators are lazy and do nothing unless consumed"]
    pub fn additional_text(&self) -> impl Iterator<Item = &str> + '_ {
        self.inner()
            .text_storages()
            .iter()
            .filter(|item| !item.is_empty())
            .map(|item| item.text())
    }

    /// Borrow optional speaker notes.
    #[must_use]
    pub fn notes(&self) -> Option<&str> {
        self.inner().notes()
    }

    /// Return the number of build animations.
    #[must_use]
    pub fn build_count(&self) -> usize {
        self.inner().builds().len()
    }

    /// Return whether the slide carries a transition.
    #[must_use]
    pub fn has_transition(&self) -> bool {
        self.inner().transition().is_some()
    }

    /// Iterate over visible and authored slide text without allocating.
    #[must_use = "iterators are lazy and do nothing unless consumed"]
    pub fn iter_text(&self) -> impl Iterator<Item = Text<'_>> + '_ {
        slide_text(self.inner())
    }
}

impl fmt::Debug for Slide {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Slide")
            .field("position", &self.position)
            .field("skipped", &self.is_skipped())
            .field("builds", &self.build_count())
            .field("transition", &self.has_transition())
            .finish_non_exhaustive()
    }
}

/// A cheap, lifetime-independent handle to one Pages section.
#[derive(Clone)]
pub struct Section {
    state: Arc<State>,
    position: usize,
}

impl Section {
    fn inner(&self) -> &litchi_pages::Section {
        self.state
            .data
            .section(self.position)
            .unwrap_or_else(|| unreachable!("validated iWork section handle"))
    }

    /// Return the zero-based document position.
    #[must_use]
    pub const fn position(&self) -> usize {
        self.position
    }

    /// Return the semantic section kind.
    #[must_use]
    pub fn kind(&self) -> SectionKind {
        match self.inner().section_type() {
            litchi_pages::SectionType::Body => SectionKind::Body,
            litchi_pages::SectionType::Header => SectionKind::Header,
            litchi_pages::SectionType::Footer => SectionKind::Footer,
            litchi_pages::SectionType::Floating => SectionKind::Floating,
        }
    }

    /// Borrow the optional producer-visible section name.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.inner().name()
    }

    /// Borrow the optional section heading.
    #[must_use]
    pub fn heading(&self) -> Option<&str> {
        self.inner().heading()
    }

    /// Iterate over paragraphs in source order.
    #[must_use = "iterators are lazy and do nothing unless consumed"]
    pub fn paragraphs(&self) -> impl ExactSizeIterator<Item = &str> + '_ {
        self.inner().paragraphs().iter().map(String::as_str)
    }

    /// Iterate over additional rich text in source order.
    #[must_use = "iterators are lazy and do nothing unless consumed"]
    pub fn additional_text(&self) -> impl Iterator<Item = &str> + '_ {
        self.inner()
            .text_storages()
            .iter()
            .filter(|item| !item.is_empty())
            .map(|item| item.text())
    }

    /// Return the known page count, when present.
    #[must_use]
    pub fn page_count(&self) -> Option<usize> {
        self.inner().page_count()
    }

    /// Iterate over section text without allocating.
    #[must_use = "iterators are lazy and do nothing unless consumed"]
    pub fn iter_text(&self) -> impl Iterator<Item = Text<'_>> + '_ {
        section_text(self.inner())
    }
}

impl fmt::Debug for Section {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Section")
            .field("position", &self.position)
            .field("kind", &self.kind())
            .field("page_count", &self.page_count())
            .finish_non_exhaustive()
    }
}

/// Semantic kind of a Pages section.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SectionKind {
    /// Main document body.
    Body,
    /// Header content.
    Header,
    /// Footer content.
    Footer,
    /// Floating or anchored content.
    Floating,
}

/// Semantic role of one borrowed text value.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum TextRole {
    /// Numbers table name.
    TableName,
    /// Numbers column header.
    TableColumnHeader,
    /// Numbers row header.
    TableRowHeader,
    /// Textual, formula, or error Numbers cell.
    TableCell,
    /// Visible Keynote slide title.
    SlideTitle,
    /// Ordinary Keynote slide content.
    SlideContent,
    /// Additional Keynote rich text.
    SlideAdditional,
    /// Keynote speaker notes.
    SlideNotes,
    /// Pages section heading.
    SectionHeading,
    /// Pages paragraph.
    SectionParagraph,
    /// Additional Pages rich text.
    SectionAdditional,
}

/// One borrowed, allocation-free text item from an iWork snapshot.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Text<'a> {
    role: TextRole,
    value: &'a str,
}

impl<'a> Text<'a> {
    const fn new(role: TextRole, value: &'a str) -> Self {
        Self { role, value }
    }

    /// Return the semantic role.
    #[must_use]
    pub const fn role(self) -> TextRole {
        self.role
    }

    /// Borrow the exact semantic text.
    #[must_use]
    pub const fn value(self) -> &'a str {
        self.value
    }
}

impl fmt::Display for Text<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.value)
    }
}

fn map_value(value: &litchi_numbers::cell::Value) -> Value<'_> {
    match value {
        litchi_numbers::cell::Value::Empty => Value::Empty,
        litchi_numbers::cell::Value::Text(value) => Value::Text(value),
        litchi_numbers::cell::Value::Number(value) => Value::Number(value.get()),
        litchi_numbers::cell::Value::Boolean(value) => Value::Boolean(*value),
        litchi_numbers::cell::Value::Date(value) => Value::Date(value.get()),
        litchi_numbers::cell::Value::Duration(value) => Value::Duration(value.get()),
        litchi_numbers::cell::Value::Formula(value) => Value::Formula(value),
        litchi_numbers::cell::Value::Error(value) => Value::Error(value),
    }
}

fn table_text(table: &litchi_numbers::Table) -> impl Iterator<Item = Text<'_>> + '_ {
    let name = std::iter::once(Text::new(TextRole::TableName, table.name()));
    let column_headers = table
        .column_headers()
        .map(|value| Text::new(TextRole::TableColumnHeader, value));
    let row_headers = table
        .row_headers()
        .map(|value| Text::new(TextRole::TableRowHeader, value));
    let cells = table.iter_cells().filter_map(|cell| match cell.value() {
        litchi_numbers::cell::Value::Text(value)
        | litchi_numbers::cell::Value::Formula(value)
        | litchi_numbers::cell::Value::Error(value) => Some(Text::new(TextRole::TableCell, value)),
        litchi_numbers::cell::Value::Empty
        | litchi_numbers::cell::Value::Number(_)
        | litchi_numbers::cell::Value::Boolean(_)
        | litchi_numbers::cell::Value::Date(_)
        | litchi_numbers::cell::Value::Duration(_) => None,
    });
    name.chain(column_headers).chain(row_headers).chain(cells)
}

fn slide_text(slide: &litchi_keynote::Slide) -> impl Iterator<Item = Text<'_>> + '_ {
    let title = slide
        .title()
        .into_iter()
        .map(|value| Text::new(TextRole::SlideTitle, value));
    let content = slide
        .text_content()
        .iter()
        .map(|value| Text::new(TextRole::SlideContent, value));
    let additional = slide
        .text_storages()
        .iter()
        .filter(|value| !value.is_empty())
        .map(|value| Text::new(TextRole::SlideAdditional, value.text()));
    let notes = slide
        .notes()
        .into_iter()
        .map(|value| Text::new(TextRole::SlideNotes, value));
    title.chain(content).chain(additional).chain(notes)
}

fn section_text(section: &litchi_pages::Section) -> impl Iterator<Item = Text<'_>> + '_ {
    let heading = section
        .heading()
        .into_iter()
        .map(|value| Text::new(TextRole::SectionHeading, value));
    let paragraphs = section
        .paragraphs()
        .iter()
        .map(|value| Text::new(TextRole::SectionParagraph, value));
    let additional = section
        .text_storages()
        .iter()
        .filter(|value| !value.is_empty())
        .map(|value| Text::new(TextRole::SectionAdditional, value.text()));
    heading.chain(paragraphs).chain(additional)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn assert_send_sync<T: Send + Sync>() {}

    #[test]
    fn major_handles_are_send_sync() {
        assert_send_sync::<Document>();
        assert_send_sync::<Snapshot>();
        assert_send_sync::<Table>();
        assert_send_sync::<Slide>();
        assert_send_sync::<Section>();
    }

    #[test]
    fn snapshot_and_handles_share_one_state() {
        let document = Document::from_data(Format::Pages, StructuredData::empty());
        let first = document.snapshot();
        let second = first.clone();
        assert!(Arc::ptr_eq(&first.state, &second.state));
    }
}
