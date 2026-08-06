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
    /// clamped during snapshot construction.
    #[must_use]
    pub const fn new(
        max_tables: usize,
        max_slides: usize,
        max_sections: usize,
        max_text_bytes: usize,
    ) -> Self {
        Self {
            max_tables,
            max_slides,
            max_sections,
            max_text_bytes,
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

/// An immutable, archive-free structured iWork snapshot.
///
/// The three semantic collections are private, reference-counted slices. A
/// clone or [`Self::snapshot`] shares those allocations and never clones a
/// table, slide, or section. Native archive objects, protobuf messages, and
/// physical identifiers are intentionally absent from this API.
#[derive(Debug, Clone)]
pub struct StructuredData {
    tables: Arc<[Table]>,
    slides: Arc<[Slide]>,
    sections: Arc<[Section]>,
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
    /// Validation runs before the input vectors are moved into shared
    /// immutable slices, so malformed input cannot publish a partial snapshot.
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
        let max_tables = limits.max_tables.min(MAX_TABLES);
        let max_slides = limits.max_slides.min(MAX_SLIDES);
        let max_sections = limits.max_sections.min(MAX_SECTIONS);
        let max_text_bytes = limits.max_text_bytes.min(DEFAULT_MAX_TEXT_BYTES);

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
                    litchi_numbers::cell::Value::Text(value)
                    | litchi_numbers::cell::Value::Formula(value)
                    | litchi_numbers::cell::Value::Error(value) => value.len(),
                    litchi_numbers::cell::Value::Empty
                    | litchi_numbers::cell::Value::Number(_)
                    | litchi_numbers::cell::Value::Boolean(_)
                    | litchi_numbers::cell::Value::Date(_)
                    | litchi_numbers::cell::Value::Duration(_) => 0,
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
            tables: Arc::from(tables.into_boxed_slice()),
            slides: Arc::from(slides.into_boxed_slice()),
            sections: Arc::from(sections.into_boxed_slice()),
        })
    }

    /// Capture another cheap handle to the same immutable snapshot.
    #[must_use]
    pub fn snapshot(&self) -> Self {
        self.clone()
    }

    /// Borrow tables in stable source order without cloning them.
    #[must_use]
    pub fn tables(&self) -> &[Table] {
        &self.tables
    }

    /// Iterate over tables in stable source order without cloning them.
    #[must_use]
    pub fn iter_tables(&self) -> impl ExactSizeIterator<Item = &Table> + '_ {
        self.tables.iter()
    }

    /// Select a table by checked zero-based position.
    #[must_use]
    pub fn table(&self, index: usize) -> Option<&Table> {
        self.tables.get(index)
    }

    /// Return the number of tables in the snapshot.
    #[must_use]
    pub fn table_count(&self) -> usize {
        self.tables.len()
    }

    /// Borrow slides in stable source order without cloning them.
    #[must_use]
    pub fn slides(&self) -> &[Slide] {
        &self.slides
    }

    /// Iterate over slides in stable source order without cloning them.
    #[must_use]
    pub fn iter_slides(&self) -> impl ExactSizeIterator<Item = &Slide> + '_ {
        self.slides.iter()
    }

    /// Select a slide by checked zero-based position.
    #[must_use]
    pub fn slide(&self, index: usize) -> Option<&Slide> {
        self.slides.get(index)
    }

    /// Return the number of slides in the snapshot.
    #[must_use]
    pub fn slide_count(&self) -> usize {
        self.slides.len()
    }

    /// Borrow sections in stable source order without cloning them.
    #[must_use]
    pub fn sections(&self) -> &[Section] {
        &self.sections
    }

    /// Iterate over sections in stable source order without cloning them.
    #[must_use]
    pub fn iter_sections(&self) -> impl ExactSizeIterator<Item = &Section> + '_ {
        self.sections.iter()
    }

    /// Select a section by checked zero-based position.
    #[must_use]
    pub fn section(&self, index: usize) -> Option<&Section> {
        self.sections.get(index)
    }

    /// Return the number of sections in the snapshot.
    #[must_use]
    pub fn section_count(&self) -> usize {
        self.sections.len()
    }

    /// Return whether no semantic values were extracted.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.tables.is_empty() && self.slides.is_empty() && self.sections.is_empty()
    }

    /// Return deterministic summary counts for the contained values.
    #[must_use]
    pub fn summary(&self) -> String {
        format!(
            "Tables: {}, Slides: {}, Sections: {}",
            self.tables.len(),
            self.slides.len(),
            self.sections.len()
        )
    }

    /// Collect all human-readable text without materializing intermediate
    /// per-format collections.
    #[must_use]
    pub fn all_text(&self) -> Vec<String> {
        let mut all_text = Vec::new();

        for table in self.tables.iter() {
            all_text.push(format!("Table: {}", table.name()));
        }
        for slide in self.slides.iter() {
            append_slide_text(&mut all_text, slide);
        }
        for section in self.sections.iter() {
            append_section_text(&mut all_text, section);
        }

        all_text
    }
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

fn append_slide_text(output: &mut Vec<String>, slide: &Slide) {
    if let Some(title) = slide.title() {
        output.push(title.to_owned());
    }
    output.extend(slide.text_content().iter().cloned());
    if let Some(notes) = slide.notes() {
        output.push(notes.to_owned());
    }
    output.extend(
        slide
            .text_storages()
            .iter()
            .filter(|storage| !storage.is_empty())
            .map(|storage| storage.text().to_owned()),
    );
}

fn append_section_text(output: &mut Vec<String>, section: &Section) {
    if let Some(heading) = section.heading() {
        output.push(heading.to_owned());
    }
    output.extend(section.paragraphs().iter().cloned());
    output.extend(
        section
            .text_storages()
            .iter()
            .filter(|storage| !storage.is_empty())
            .map(|storage| storage.text().to_owned()),
    );
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
        assert_eq!(data.iter_sections().map(Section::index).collect::<Vec<_>>(), [0]);
        assert_eq!(data.all_text(), ["Table: Data", "Title", "Body", "Heading"]);
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
    }

    #[test]
    fn snapshot_clones_share_semantic_storage() {
        let data = StructuredData::from_parts(
            vec![Table::new("Data", litchi_numbers::Dimensions::new(1, 1))],
            Vec::new(),
            Vec::new(),
        )
        .unwrap_or_else(|error| panic!("semantic values should be valid: {error}"));
        let snapshot = data.snapshot();

        assert!(Arc::ptr_eq(&data.tables, &snapshot.tables));
        assert!(Arc::ptr_eq(&data.slides, &snapshot.slides));
        assert!(Arc::ptr_eq(&data.sections, &snapshot.sections));
        assert_eq!(snapshot.table(0).map(Table::name), Some("Data"));
    }
}
