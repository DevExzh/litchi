//! Archive-free structured iWork snapshots.
//!
//! The physical IWA adapter owns archive traversal and native decoding. This
//! crate owns only the bounded aggregation of semantic values from the three
//! concrete iWork format leaves, so it does not depend on protobufs, package
//! IDs, ZIP state, or the `litchi-iwa` facade.

#![forbid(unsafe_code)]

use std::fmt;
use std::sync::Arc;

use litchi_keynote::Slide;
use litchi_numbers::Table;
use litchi_numbers::cell::Value;
use litchi_pages::Section;

/// Maximum number of tables retained by one structured snapshot.
pub const MAX_TABLES: usize = litchi_numbers::MAX_TABLES;
/// Maximum number of slides retained by one structured snapshot.
pub const MAX_SLIDES: usize = 65_536;
/// Maximum number of sections retained by one structured snapshot.
pub const MAX_SECTIONS: usize = litchi_pages::MAX_SECTIONS;
/// Maximum UTF-8 bytes retained by one structured snapshot.
pub const DEFAULT_MAX_TEXT_BYTES: usize = litchi_numbers::DEFAULT_MAX_TEXT_BYTES;

/// Finite construction limits for an immutable structured snapshot.
#[allow(
    clippy::struct_field_names,
    reason = "The public budget accessors intentionally share one max_* vocabulary"
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_tables: usize,
    max_slides: usize,
    max_sections: usize,
    max_text_bytes: usize,
}

impl Limits {
    /// Create a bounded profile. Values above the hard semantic ceilings are
    /// clamped immediately, so the accessors always report the effective
    /// limits used by snapshot construction.
    #[must_use]
    pub const fn new(
        max_tables: usize,
        max_slides: usize,
        max_sections: usize,
        max_text_bytes: usize,
    ) -> Self {
        Self {
            max_tables: clamp_limit(max_tables, MAX_TABLES),
            max_slides: clamp_limit(max_slides, MAX_SLIDES),
            max_sections: clamp_limit(max_sections, MAX_SECTIONS),
            max_text_bytes: clamp_limit(max_text_bytes, DEFAULT_MAX_TEXT_BYTES),
        }
    }

    /// Return the configured table ceiling.
    #[must_use]
    pub const fn max_tables(self) -> usize {
        self.max_tables
    }

    /// Return the configured slide ceiling.
    #[must_use]
    pub const fn max_slides(self) -> usize {
        self.max_slides
    }

    /// Return the configured section ceiling.
    #[must_use]
    pub const fn max_sections(self) -> usize {
        self.max_sections
    }

    /// Return the configured semantic text-byte ceiling.
    #[must_use]
    pub const fn max_text_bytes(self) -> usize {
        self.max_text_bytes
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::new(MAX_TABLES, MAX_SLIDES, MAX_SECTIONS, DEFAULT_MAX_TEXT_BYTES)
    }
}

/// Errors raised while constructing a bounded structured snapshot.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// The supplied table sequence exceeds the selected bound.
    TooManyTables {
        /// Number of supplied tables.
        actual: usize,
        /// Maximum accepted tables.
        limit: usize,
    },
    /// The supplied slide sequence exceeds the selected bound.
    TooManySlides {
        /// Number of supplied slides.
        actual: usize,
        /// Maximum accepted slides.
        limit: usize,
    },
    /// The supplied section sequence exceeds the selected bound.
    TooManySections {
        /// Number of supplied sections.
        actual: usize,
        /// Maximum accepted sections.
        limit: usize,
    },
    /// A slide does not carry its canonical position in the ordered sequence.
    InvalidSlideIndex {
        /// Position occupied by the slide in the supplied sequence.
        expected: usize,
        /// Index stored by the slide.
        actual: usize,
    },
    /// A section does not carry its canonical position in the ordered sequence.
    InvalidSectionIndex {
        /// Position occupied by the section in the supplied sequence.
        expected: usize,
        /// Index stored by the section.
        actual: usize,
    },
    /// Semantic text values exceed the selected byte budget.
    TextTooLarge {
        /// Maximum accepted UTF-8 bytes.
        limit: usize,
    },
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::TooManyTables { actual, limit } => write!(
                formatter,
                "structured snapshot contains {actual} tables; maximum is {limit}"
            ),
            Self::TooManySlides { actual, limit } => write!(
                formatter,
                "structured snapshot contains {actual} slides; maximum is {limit}"
            ),
            Self::TooManySections { actual, limit } => write!(
                formatter,
                "structured snapshot contains {actual} sections; maximum is {limit}"
            ),
            Self::InvalidSlideIndex { expected, actual } => write!(
                formatter,
                "structured slide index {actual} is not the expected index {expected}"
            ),
            Self::InvalidSectionIndex { expected, actual } => write!(
                formatter,
                "structured section index {actual} is not the expected index {expected}"
            ),
            Self::TextTooLarge { limit } => write!(
                formatter,
                "structured snapshot semantic text exceeds {limit} bytes"
            ),
        }
    }
}

impl std::error::Error for Error {}

/// Result type for bounded structured snapshot construction.
pub type Result<T> = std::result::Result<T, Error>;

/// The semantic role of one borrowed text value in a structured snapshot.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum TextKind {
    /// A Numbers table name. [`Text::to_string`] adds the `Table: ` label.
    TableName,
    /// A Numbers column header.
    TableColumnHeader,
    /// A Numbers row header.
    TableRowHeader,
    /// A textual, formula, or error value in a materialized Numbers cell.
    TableCell,
    /// A Keynote slide title.
    SlideTitle,
    /// A Keynote plain-text content block.
    SlideContent,
    /// Keynote speaker notes.
    SlideNotes,
    /// Text from a Keynote rich-text storage.
    SlideStorage,
    /// A Pages section heading.
    SectionHeading,
    /// A Pages paragraph.
    SectionParagraph,
    /// Text from a Pages rich-text storage.
    SectionStorage,
}

/// One borrowed, allocation-free text item from a structured snapshot.
///
/// The item keeps its source role so callers can filter or format it without
/// reparsing synthetic strings. The value points directly into the immutable
/// leaf-owned value; no text is cloned while iterating.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Text<'a> {
    kind: TextKind,
    value: &'a str,
}

impl<'a> Text<'a> {
    const fn new(kind: TextKind, value: &'a str) -> Self {
        Self { kind, value }
    }

    /// Return the semantic role of this text item.
    #[must_use]
    pub const fn kind(self) -> TextKind {
        self.kind
    }

    /// Borrow the original semantic text without allocating.
    #[must_use]
    pub const fn value(self) -> &'a str {
        self.value
    }
}

impl fmt::Display for Text<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        if self.kind == TextKind::TableName {
            write!(formatter, "Table: {}", self.value)
        } else {
            formatter.write_str(self.value)
        }
    }
}

#[derive(Debug, Default)]
struct Snapshot {
    tables: Vec<Table>,
    slides: Vec<Slide>,
    sections: Vec<Section>,
}

/// An immutable, archive-free structured iWork snapshot.
///
/// The snapshot has one reference-counted owner containing the three original
/// semantic vectors. A clone or [`Self::snapshot`] shares that owner and never
/// clones a table, slide, section, sparse cell, or text allocation.
/// Native archive objects, protobuf messages, and physical identifiers are
/// intentionally absent from this API.
#[derive(Debug, Clone)]
pub struct StructuredData {
    snapshot: Arc<Snapshot>,
}

impl Default for StructuredData {
    fn default() -> Self {
        Self {
            snapshot: Arc::new(Snapshot::default()),
        }
    }
}

impl StructuredData {
    /// Build a snapshot from ordered semantic values using the default bounds.
    ///
    /// The input vectors are consumed without cloning their elements.
    ///
    /// # Errors
    ///
    /// Returns a typed error when a collection exceeds its hard bound, a
    /// slide or section has a non-canonical position, or semantic text
    /// exceeds [`DEFAULT_MAX_TEXT_BYTES`].
    pub fn from_parts(
        tables: Vec<Table>,
        slides: Vec<Slide>,
        sections: Vec<Section>,
    ) -> Result<Self> {
        Self::from_parts_with_limits(tables, slides, sections, Limits::default())
    }

    /// Build a snapshot under explicit finite semantic budgets.
    ///
    /// Caller-selected bounds can only tighten the hard semantic ceilings.
    /// Validation runs before the input vectors are moved into one shared
    /// immutable owner, so malformed input cannot publish a partial snapshot.
    ///
    /// # Errors
    ///
    /// Returns a typed error when a collection exceeds its selected bound, a
    /// slide or section has a non-canonical position, or semantic text
    /// exceeds the selected byte budget.
    pub fn from_parts_with_limits(
        tables: Vec<Table>,
        slides: Vec<Slide>,
        sections: Vec<Section>,
        limits: Limits,
    ) -> Result<Self> {
        let max_tables = limits.max_tables;
        let max_slides = limits.max_slides;
        let max_sections = limits.max_sections;
        let max_text_bytes = limits.max_text_bytes;

        if tables.len() > max_tables {
            return Err(Error::TooManyTables {
                actual: tables.len(),
                limit: max_tables,
            });
        }
        if slides.len() > max_slides {
            return Err(Error::TooManySlides {
                actual: slides.len(),
                limit: max_slides,
            });
        }
        if sections.len() > max_sections {
            return Err(Error::TooManySections {
                actual: sections.len(),
                limit: max_sections,
            });
        }

        for (expected, slide) in slides.iter().enumerate() {
            if slide.index() != expected {
                return Err(Error::InvalidSlideIndex {
                    expected,
                    actual: slide.index(),
                });
            }
        }
        for (expected, section) in sections.iter().enumerate() {
            if section.index() != expected {
                return Err(Error::InvalidSectionIndex {
                    expected,
                    actual: section.index(),
                });
            }
        }

        let mut text_bytes = 0;
        for table in &tables {
            text_bytes = checked_text_add(text_bytes, table.name().len(), max_text_bytes)?;
            for header in table.column_headers().chain(table.row_headers()) {
                text_bytes = checked_text_add(text_bytes, header.len(), max_text_bytes)?;
            }
            for cell in table.iter_cells() {
                let value_bytes = match cell.value() {
                    Value::Text(value) | Value::Formula(value) | Value::Error(value) => value.len(),
                    Value::Empty
                    | Value::Number(_)
                    | Value::Boolean(_)
                    | Value::Date(_)
                    | Value::Duration(_) => 0,
                };
                text_bytes = checked_text_add(text_bytes, value_bytes, max_text_bytes)?;
            }
        }
        for slide in &slides {
            if let Some(title) = slide.title() {
                text_bytes = checked_text_add(text_bytes, title.len(), max_text_bytes)?;
            }
            for text in slide.text_content() {
                text_bytes = checked_text_add(text_bytes, text.len(), max_text_bytes)?;
            }
            if let Some(notes) = slide.notes() {
                text_bytes = checked_text_add(text_bytes, notes.len(), max_text_bytes)?;
            }
            for storage in slide.text_storages() {
                text_bytes = checked_text_add(text_bytes, storage.len(), max_text_bytes)?;
            }
        }
        for section in &sections {
            if let Some(heading) = section.heading() {
                text_bytes = checked_text_add(text_bytes, heading.len(), max_text_bytes)?;
            }
            for paragraph in section.paragraphs() {
                text_bytes = checked_text_add(text_bytes, paragraph.len(), max_text_bytes)?;
            }
            for storage in section.text_storages() {
                text_bytes = checked_text_add(text_bytes, storage.len(), max_text_bytes)?;
            }
        }

        Ok(Self {
            snapshot: Arc::new(Snapshot {
                tables,
                slides,
                sections,
            }),
        })
    }

    /// Return an empty immutable structured snapshot.
    #[must_use]
    pub fn empty() -> Self {
        Self::default()
    }

    /// Capture another cheap handle to the same immutable snapshot.
    #[must_use]
    pub fn snapshot(&self) -> Self {
        self.clone()
    }

    /// Borrow tables in stable source order without cloning them.
    #[must_use]
    pub fn tables(&self) -> &[Table] {
        &self.snapshot.tables
    }

    /// Iterate over tables in stable source order without cloning them.
    #[must_use]
    pub fn iter_tables(&self) -> impl ExactSizeIterator<Item = &Table> + '_ {
        self.snapshot.tables.iter()
    }

    /// Select a table by checked zero-based position.
    #[must_use]
    pub fn table(&self, index: usize) -> Option<&Table> {
        self.snapshot.tables.get(index)
    }

    /// Return the number of tables in the snapshot.
    #[must_use]
    pub fn table_count(&self) -> usize {
        self.snapshot.tables.len()
    }

    /// Borrow slides in stable source order without cloning them.
    #[must_use]
    pub fn slides(&self) -> &[Slide] {
        &self.snapshot.slides
    }

    /// Iterate over slides in stable source order without cloning them.
    #[must_use]
    pub fn iter_slides(&self) -> impl ExactSizeIterator<Item = &Slide> + '_ {
        self.snapshot.slides.iter()
    }

    /// Select a slide by checked zero-based position.
    #[must_use]
    pub fn slide(&self, index: usize) -> Option<&Slide> {
        self.snapshot.slides.get(index)
    }

    /// Return the number of slides in the snapshot.
    #[must_use]
    pub fn slide_count(&self) -> usize {
        self.snapshot.slides.len()
    }

    /// Borrow sections in stable source order without cloning them.
    #[must_use]
    pub fn sections(&self) -> &[Section] {
        &self.snapshot.sections
    }

    /// Iterate over sections in stable source order without cloning them.
    #[must_use]
    pub fn iter_sections(&self) -> impl ExactSizeIterator<Item = &Section> + '_ {
        self.snapshot.sections.iter()
    }

    /// Select a section by checked zero-based position.
    #[must_use]
    pub fn section(&self, index: usize) -> Option<&Section> {
        self.snapshot.sections.get(index)
    }

    /// Return the number of sections in the snapshot.
    #[must_use]
    pub fn section_count(&self) -> usize {
        self.snapshot.sections.len()
    }

    /// Return whether no semantic values were extracted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.snapshot.tables.is_empty()
            && self.snapshot.slides.is_empty()
            && self.snapshot.sections.is_empty()
    }

    /// Return deterministic summary counts for the contained values.
    #[must_use]
    pub fn summary(&self) -> String {
        format!(
            "Tables: {}, Slides: {}, Sections: {}",
            self.snapshot.tables.len(),
            self.snapshot.slides.len(),
            self.snapshot.sections.len()
        )
    }

    /// Iterate over modeled text in stable table, slide, and section order.
    ///
    /// This is the zero-allocation text API. Each item borrows the original
    /// leaf-owned string, and empty rich-text storages are omitted just as they
    /// are from [`Self::all_text`]. The table-name label is represented by the
    /// [`TextKind::TableName`] role and is only materialized by formatting the
    /// returned item.
    #[must_use = "iterating the borrowed text requires consuming the iterator"]
    pub fn iter_text(&self) -> impl Iterator<Item = Text<'_>> + '_ {
        let tables = self.snapshot.tables.iter().flat_map(table_text);
        let slides = self.snapshot.slides.iter().flat_map(slide_text);
        let sections = self.snapshot.sections.iter().flat_map(section_text);
        tables.chain(slides).chain(sections)
    }

    /// Collect all human-readable text as an explicit allocating projection.
    ///
    /// Use [`Self::iter_text`] for borrowed traversal when the caller does not
    /// need owned strings.
    #[must_use]
    pub fn all_text(&self) -> Vec<String> {
        self.iter_text().map(|text| text.to_string()).collect()
    }
}

const fn clamp_limit(value: usize, maximum: usize) -> usize {
    if value > maximum { maximum } else { value }
}

fn checked_text_add(current: usize, added: usize, limit: usize) -> Result<usize> {
    let total = current
        .checked_add(added)
        .ok_or(Error::TextTooLarge { limit })?;
    if total > limit {
        return Err(Error::TextTooLarge { limit });
    }
    Ok(total)
}

fn table_text(table: &Table) -> impl Iterator<Item = Text<'_>> + '_ {
    let name = std::iter::once(Text::new(TextKind::TableName, table.name()));
    let column_headers = table
        .column_headers()
        .map(|text| Text::new(TextKind::TableColumnHeader, text));
    let row_headers = table
        .row_headers()
        .map(|text| Text::new(TextKind::TableRowHeader, text));
    let cells = table.iter_cells().filter_map(|cell| match cell.value() {
        Value::Text(text) | Value::Formula(text) | Value::Error(text) => {
            Some(Text::new(TextKind::TableCell, text))
        },
        Value::Empty
        | Value::Number(_)
        | Value::Boolean(_)
        | Value::Date(_)
        | Value::Duration(_) => None,
    });
    name.chain(column_headers).chain(row_headers).chain(cells)
}

fn slide_text(slide: &Slide) -> impl Iterator<Item = Text<'_>> + '_ {
    let title = slide
        .title()
        .map(|text| Text::new(TextKind::SlideTitle, text))
        .into_iter();
    let content = slide
        .text_content()
        .iter()
        .map(|text| Text::new(TextKind::SlideContent, text));
    let notes = slide
        .notes()
        .map(|text| Text::new(TextKind::SlideNotes, text))
        .into_iter();
    let storages = slide
        .text_storages()
        .iter()
        .filter(|storage| !storage.is_empty())
        .map(|storage| Text::new(TextKind::SlideStorage, storage.text()));
    title.chain(content).chain(notes).chain(storages)
}

fn section_text(section: &Section) -> impl Iterator<Item = Text<'_>> + '_ {
    let heading = section
        .heading()
        .map(|text| Text::new(TextKind::SectionHeading, text))
        .into_iter();
    let paragraphs = section
        .paragraphs()
        .iter()
        .map(|text| Text::new(TextKind::SectionParagraph, text));
    let storages = section
        .text_storages()
        .iter()
        .filter(|storage| !storage.is_empty())
        .map(|storage| Text::new(TextKind::SectionStorage, storage.text()));
    heading.chain(paragraphs).chain(storages)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn empty_parts() -> (Vec<Table>, Vec<Slide>, Vec<Section>) {
        (Vec::new(), Vec::new(), Vec::new())
    }

    #[test]
    fn empty_data_is_a_valid_empty_snapshot() {
        let (tables, slides, sections) = empty_parts();
        let data = StructuredData::from_parts(tables, slides, sections)
            .unwrap_or_else(|error| panic!("empty snapshot should be valid: {error}"));

        assert!(data.is_empty());
        assert_eq!(data.summary(), "Tables: 0, Slides: 0, Sections: 0");
        assert!(data.all_text().is_empty());
        assert!(data.tables().is_empty());
        assert!(data.slides().is_empty());
        assert!(data.sections().is_empty());
        assert!(data.table(0).is_none());
        assert!(data.slide(0).is_none());
        assert!(data.section(0).is_none());
    }

    #[test]
    fn default_and_empty_are_cheap_valid_snapshots() {
        let default = StructuredData::default();
        let empty = StructuredData::empty();

        assert!(default.is_empty());
        assert!(empty.is_empty());
        assert_eq!(default.summary(), "Tables: 0, Slides: 0, Sections: 0");
    }

    #[test]
    fn accessors_and_iterators_borrow_ordered_semantic_values() {
        let table = Table::new("Data", litchi_numbers::Dimensions::new(1, 1));
        let mut slide_builder = Slide::builder(0);
        slide_builder.set_title(Some("Title".to_owned()));
        slide_builder.push_text("Body".to_owned());
        let slide = slide_builder.build();
        let mut section_builder = Section::builder(0, litchi_pages::SectionType::Body);
        section_builder.set_heading(Some("Heading".to_owned()));
        let section = section_builder.build();

        let data = StructuredData::from_parts(vec![table], vec![slide], vec![section])
            .unwrap_or_else(|error| panic!("semantic values should be valid: {error}"));

        assert_eq!(data.table_count(), 1);
        assert_eq!(data.slide_count(), 1);
        assert_eq!(data.section_count(), 1);
        assert_eq!(data.table(0).map(Table::name), Some("Data"));
        assert_eq!(data.slide(0).map(Slide::title), Some(Some("Title")));
        assert_eq!(data.section(0).and_then(Section::heading), Some("Heading"));
        assert!(data.table(1).is_none());
        assert!(data.slide(1).is_none());
        assert!(data.section(1).is_none());
        assert_eq!(
            data.iter_tables().map(Table::name).collect::<Vec<_>>(),
            ["Data"]
        );
        assert_eq!(
            data.iter_slides().map(Slide::index).collect::<Vec<_>>(),
            [0]
        );
        assert_eq!(
            data.iter_sections().map(Section::index).collect::<Vec<_>>(),
            [0]
        );
        assert_eq!(
            data.iter_text()
                .map(|text| (text.kind(), text.value()))
                .collect::<Vec<_>>(),
            [
                (TextKind::TableName, "Data"),
                (TextKind::SlideTitle, "Title"),
                (TextKind::SlideContent, "Body"),
                (TextKind::SectionHeading, "Heading"),
            ]
        );
        assert_eq!(data.all_text(), ["Table: Data", "Title", "Body", "Heading"]);
    }

    #[test]
    fn text_iteration_keeps_table_headers_and_textual_cells() {
        let mut builder = Table::builder("Data", litchi_numbers::Dimensions::new(1, 1));
        builder
            .set_column_headers(["Column"])
            .unwrap_or_else(|error| panic!("column header should be valid: {error:?}"));
        builder
            .set_row_headers(["Row"])
            .unwrap_or_else(|error| panic!("row header should be valid: {error:?}"));
        builder
            .push(litchi_numbers::Cell::new(
                litchi_numbers::CellPosition::new(0, 0),
                Value::Text("Cell".to_owned()),
            ))
            .unwrap_or_else(|error| panic!("cell should be valid: {error:?}"));
        let table = builder
            .finish()
            .unwrap_or_else(|error| panic!("table should be valid: {error:?}"));

        let data = StructuredData::from_parts(vec![table], Vec::new(), Vec::new())
            .unwrap_or_else(|error| panic!("semantic values should be valid: {error}"));

        assert_eq!(
            data.iter_text()
                .map(|text| (text.kind(), text.value()))
                .collect::<Vec<_>>(),
            [
                (TextKind::TableName, "Data"),
                (TextKind::TableColumnHeader, "Column"),
                (TextKind::TableRowHeader, "Row"),
                (TextKind::TableCell, "Cell"),
            ]
        );
        assert_eq!(data.all_text(), ["Table: Data", "Column", "Row", "Cell"]);
    }

    #[test]
    fn construction_rejects_malformed_order_and_tightened_budgets() {
        let invalid_slide = Slide::builder(1).build();
        assert!(matches!(
            StructuredData::from_parts(Vec::new(), vec![invalid_slide], Vec::new()),
            Err(Error::InvalidSlideIndex {
                expected: 0,
                actual: 1,
            })
        ));

        let invalid_section = Section::builder(1, litchi_pages::SectionType::Body).build();
        assert!(matches!(
            StructuredData::from_parts(Vec::new(), Vec::new(), vec![invalid_section]),
            Err(Error::InvalidSectionIndex {
                expected: 0,
                actual: 1,
            })
        ));

        let too_many_tables = StructuredData::from_parts_with_limits(
            vec![Table::new("Data", litchi_numbers::Dimensions::new(1, 1))],
            Vec::new(),
            Vec::new(),
            Limits::new(0, MAX_SLIDES, MAX_SECTIONS, DEFAULT_MAX_TEXT_BYTES),
        );
        assert!(matches!(
            too_many_tables,
            Err(Error::TooManyTables {
                actual: 1,
                limit: 0,
            })
        ));

        let too_much_text = StructuredData::from_parts_with_limits(
            vec![Table::new("Data", litchi_numbers::Dimensions::new(1, 1))],
            Vec::new(),
            Vec::new(),
            Limits::new(MAX_TABLES, MAX_SLIDES, MAX_SECTIONS, 3),
        );
        assert!(matches!(
            too_much_text,
            Err(Error::TextTooLarge { limit: 3 })
        ));

        let limits = Limits::new(usize::MAX, usize::MAX, usize::MAX, usize::MAX);
        assert_eq!(limits.max_tables(), MAX_TABLES);
        assert_eq!(limits.max_slides(), MAX_SLIDES);
        assert_eq!(limits.max_sections(), MAX_SECTIONS);
        assert_eq!(limits.max_text_bytes(), DEFAULT_MAX_TEXT_BYTES);
    }

    #[test]
    fn snapshot_clones_share_semantic_storage_and_transfer_input_allocation() {
        let mut tables = Vec::new();
        tables.push(Table::new("Data", litchi_numbers::Dimensions::new(1, 1)));
        let table_ptr = tables.as_ptr();
        let data = StructuredData::from_parts(tables, Vec::new(), Vec::new())
            .unwrap_or_else(|error| panic!("semantic values should be valid: {error}"));
        let snapshot = data.snapshot();

        assert!(Arc::ptr_eq(&data.snapshot, &snapshot.snapshot));
        assert_eq!(data.tables().as_ptr(), table_ptr);
        assert_eq!(snapshot.table(0).map(Table::name), Some("Data"));
    }

    #[test]
    fn structured_snapshots_are_send_and_sync() {
        fn assert_send_sync<T: Send + Sync>() {}

        assert_send_sync::<StructuredData>();
    }
}
