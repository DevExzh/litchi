pub mod pagination;

use crate::selector::SectionSelector;
use litchi_iwa_text::storage::Storage;
use thiserror::Error;

pub use pagination::{PageNumber, PageNumbering, Pagination, Start};

const INHERIT_PREVIOUS_HEADER_FOOTER: u8 = 1;
const FIRST_PAGE_DIFFERENT: u8 = 2;
const EVEN_ODD_PAGES_DIFFERENT: u8 = 4;
const FIRST_PAGE_HIDES_HEADER_FOOTER: u8 = 8;

/// Maximum bytes retained by one opaque section-background payload.
pub const MAX_BACKGROUND_PAYLOAD_BYTES: usize = 4 * 1024 * 1024;

/// Validation failures for section semantic values.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum Error {
    /// A section name contains a native string terminator.
    #[error("Pages section names cannot contain NUL")]
    NameContainsNul,
    /// A known native section-start value was represented by an unknown value.
    #[error("Pages section start must use its canonical variant for a known value")]
    NonCanonicalStart,
    /// A known native page-numbering value was represented by an unknown value.
    #[error("Pages page numbering must use its canonical variant for a known value")]
    NonCanonicalNumbering,
    /// An opaque fill payload was empty.
    #[error("Pages section background payload cannot be empty")]
    EmptyBackgroundPayload,
    /// An opaque fill payload exceeded the semantic storage budget.
    #[error("Pages section background payload exceeds the semantic byte budget")]
    BackgroundPayloadTooLarge,
}

/// Result type for section semantic value construction and validation.
pub type Result<T> = std::result::Result<T, Error>;

/// Lossless settings stored directly on a Pages section.
///
/// Native boolean presence and values are packed into two bytes. The section
/// name owns exactly one UTF-8 allocation, while pagination values retain
/// unknown native discriminants without carrying archive or protobuf state.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Settings {
    name: Option<Box<str>>,
    present: u8,
    values: u8,
    start: Option<Start>,
    page_numbering: Option<PageNumbering>,
    starting_page_number: Option<PageNumber>,
}

impl Settings {
    /// Create settings with every optional native field absent.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            name: None,
            present: 0,
            values: 0,
            start: None,
            page_numbering: None,
            starting_page_number: None,
        }
    }

    /// Return the optional display name.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Set or clear the display name.
    ///
    /// The name is validated before it is moved into its exact-size owned
    /// representation.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NameContainsNul`] when the supplied name contains a
    /// native string terminator.
    pub fn set_name(&mut self, input: Option<impl Into<Box<str>>>) -> Result<()> {
        let boxed = input.map(Into::into);
        if boxed
            .as_deref()
            .is_some_and(|candidate| candidate.contains('\0'))
        {
            return Err(Error::NameContainsNul);
        }
        self.name = boxed;
        Ok(())
    }

    /// Remove the display name.
    pub fn clear_name(&mut self) {
        self.name = None;
    }

    /// Return whether the section inherits the previous header and footer.
    #[must_use]
    pub const fn inherit_previous_header_footer(&self) -> Option<bool> {
        self.get_flag(INHERIT_PREVIOUS_HEADER_FOOTER)
    }

    /// Set or clear header/footer inheritance.
    pub const fn set_inherit_previous_header_footer(&mut self, value: Option<bool>) {
        self.set_flag(INHERIT_PREVIOUS_HEADER_FOOTER, value);
    }

    /// Return whether the first page has a distinct template.
    #[must_use]
    pub const fn first_page_different(&self) -> Option<bool> {
        self.get_flag(FIRST_PAGE_DIFFERENT)
    }

    /// Set or clear first-page template distinction.
    pub const fn set_first_page_different(&mut self, value: Option<bool>) {
        self.set_flag(FIRST_PAGE_DIFFERENT, value);
    }

    /// Return whether even and odd pages use distinct templates.
    #[must_use]
    pub const fn even_odd_pages_different(&self) -> Option<bool> {
        self.get_flag(EVEN_ODD_PAGES_DIFFERENT)
    }

    /// Set or clear even/odd template distinction.
    pub const fn set_even_odd_pages_different(&mut self, value: Option<bool>) {
        self.set_flag(EVEN_ODD_PAGES_DIFFERENT, value);
    }

    /// Return whether the first page hides its header and footer.
    #[must_use]
    pub const fn first_page_hides_header_footer(&self) -> Option<bool> {
        self.get_flag(FIRST_PAGE_HIDES_HEADER_FOOTER)
    }

    /// Set or clear first-page header/footer hiding.
    pub const fn set_first_page_hides_header_footer(&mut self, value: Option<bool>) {
        self.set_flag(FIRST_PAGE_HIDES_HEADER_FOOTER, value);
    }

    /// Return the optional section-start behavior.
    #[must_use]
    pub const fn start(&self) -> Option<Start> {
        self.start
    }

    /// Set or clear section-start behavior.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonCanonicalStart`] when a known native value is
    /// represented by an `Unknown` variant.
    pub const fn set_start(&mut self, value: Option<Start>) -> Result<()> {
        if let Some(candidate) = value
            && !candidate.is_canonical()
        {
            return Err(Error::NonCanonicalStart);
        }
        self.start = value;
        Ok(())
    }

    /// Return the optional page-numbering behavior.
    #[must_use]
    pub const fn page_numbering(&self) -> Option<PageNumbering> {
        self.page_numbering
    }

    /// Set or clear page-numbering behavior.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonCanonicalNumbering`] when a known native value is
    /// represented by an `Unknown` variant.
    pub const fn set_page_numbering(&mut self, value: Option<PageNumbering>) -> Result<()> {
        if let Some(candidate) = value
            && !candidate.is_canonical()
        {
            return Err(Error::NonCanonicalNumbering);
        }
        self.page_numbering = value;
        Ok(())
    }

    /// Return the optional starting page number.
    #[must_use]
    pub const fn starting_page_number(&self) -> Option<PageNumber> {
        self.starting_page_number
    }

    /// Set or clear the starting page number.
    pub const fn set_starting_page_number(&mut self, value: Option<PageNumber>) {
        self.starting_page_number = value;
    }

    /// Validate all semantic invariants before an archive adapter publishes
    /// this value.
    ///
    /// # Errors
    ///
    /// Returns a typed error when a name or pagination value is invalid.
    pub fn validate(&self) -> Result<()> {
        if self.name.as_deref().is_some_and(|name| name.contains('\0')) {
            return Err(Error::NameContainsNul);
        }
        if self.start.is_some_and(|value| !value.is_canonical()) {
            return Err(Error::NonCanonicalStart);
        }
        if self
            .page_numbering
            .is_some_and(|value| !value.is_canonical())
        {
            return Err(Error::NonCanonicalNumbering);
        }
        Ok(())
    }

    const fn get_flag(&self, bit: u8) -> Option<bool> {
        if self.present & bit != 0 {
            Some(self.values & bit != 0)
        } else {
            None
        }
    }

    const fn set_flag(&mut self, bit: u8, option: Option<bool>) {
        if let Some(value) = option {
            self.present |= bit;
            if value {
                self.values |= bit;
            } else {
                self.values &= !bit;
            }
        } else {
            self.present &= !bit;
            self.values &= !bit;
        }
    }
}

/// The semantic fill of a Pages section.
#[non_exhaustive]
#[derive(Debug, Clone, PartialEq)]
pub enum Background {
    /// No section background fill is present.
    None,
    /// A single validated color fills the section.
    Solid(litchi_iwa_common::color::Rgba),
    /// A native fill payload not modeled by this version of the crate.
    Opaque(Opaque),
}

/// A lossless, non-empty native section-background payload.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Opaque(Box<[u8]>);

impl Opaque {
    /// Retain a non-empty native payload in exact-size bounded storage.
    ///
    /// # Errors
    ///
    /// Returns [`Error::EmptyBackgroundPayload`] for an empty payload or
    /// [`Error::BackgroundPayloadTooLarge`] when it exceeds the semantic byte
    /// budget.
    pub fn new(input: impl Into<Box<[u8]>>) -> Result<Self> {
        let payload = input.into();
        validate_background_payload(&payload)?;
        Ok(Self(payload))
    }

    /// Copy a borrowed native payload into an opaque value.
    ///
    /// # Errors
    ///
    /// Returns [`Error::EmptyBackgroundPayload`] for an empty payload or
    /// [`Error::BackgroundPayloadTooLarge`] when it exceeds the semantic byte
    /// budget. The length is checked before copying the borrowed bytes.
    pub fn from_slice(payload: &[u8]) -> Result<Self> {
        validate_background_payload(payload)?;
        Self::new(payload.to_vec().into_boxed_slice())
    }

    /// Borrow the exact retained native bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.0
    }

    /// Consume the value and return its exact native bytes.
    #[must_use]
    pub fn into_bytes(self) -> Box<[u8]> {
        self.0
    }
}

/// A logical Pages document section.
#[allow(
    clippy::struct_field_names,
    reason = "The focused Section API names its semantic kind section_type"
)]
#[derive(Debug, Clone)]
pub struct Section {
    /// Zero-based section index.
    index: usize,
    /// Semantic kind of the section.
    section_type: SectionType,
    /// Optional producer-visible section name.
    name: Option<Box<str>>,
    /// Optional section heading.
    heading: Option<Box<str>>,
    /// Paragraph values extracted from the section.
    paragraphs: Box<[String]>,
    /// Rich-text storages belonging to the section.
    text_storages: Box<[Storage]>,
    /// Number of pages represented by the section, when known.
    page_count: Option<usize>,
}

impl Section {
    /// Creates an empty section with `index` and `section_type`.
    #[must_use]
    pub fn new(index: usize, section_type: SectionType) -> Self {
        Self {
            index,
            section_type,
            name: None,
            heading: None,
            paragraphs: Box::new([]),
            text_storages: Box::new([]),
            page_count: None,
        }
    }

    /// Starts a detached builder for a section at `index`.
    #[must_use]
    pub fn builder(index: usize, section_type: SectionType) -> Builder {
        Builder::new(index, section_type)
    }

    /// Returns the zero-based position in the semantic document.
    #[must_use]
    pub const fn index(&self) -> usize {
        self.index
    }

    /// Returns the semantic kind of the section.
    #[must_use]
    pub const fn section_type(&self) -> SectionType {
        self.section_type
    }

    /// Returns the optional producer-visible section name.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Return a semantic selector, preferring the producer-visible name.
    ///
    /// Exact-name resolution can report ambiguity when malformed or authored
    /// input repeats a name. Use [`Self::position_selector`] when a stable,
    /// snapshot-local selector is required.
    #[must_use]
    pub fn selector(&self) -> SectionSelector<'_> {
        self.name
            .as_deref()
            .map_or_else(|| self.position_selector(), SectionSelector::name)
    }

    /// Return a typed selector for this section's zero-based source position.
    #[must_use]
    pub const fn position_selector(&self) -> SectionSelector<'static> {
        SectionSelector::index(self.index)
    }

    /// Returns the optional section heading.
    #[must_use]
    pub fn heading(&self) -> Option<&str> {
        self.heading.as_deref()
    }

    /// Borrows paragraph values in source order.
    #[must_use]
    pub fn paragraphs(&self) -> &[String] {
        &self.paragraphs
    }

    /// Borrows rich-text storages in source order without copying them.
    #[must_use]
    pub fn text_storages(&self) -> &[Storage] {
        &self.text_storages
    }

    /// Returns the known page count, when present.
    #[must_use]
    pub const fn page_count(&self) -> Option<usize> {
        self.page_count
    }

    /// Returns all non-empty text values in document order.
    #[must_use]
    pub fn all_text(&self) -> Vec<String> {
        let mut all = Vec::with_capacity(
            usize::from(self.heading.is_some())
                .saturating_add(self.paragraphs.len())
                .saturating_add(self.text_storages.len()),
        );
        if let Some(heading) = &self.heading {
            all.push(heading.to_string());
        }
        all.extend(self.paragraphs.iter().cloned());

        all.extend(
            self.text_storages
                .iter()
                .filter(|storage| !storage.is_empty())
                .map(|storage| storage.text().to_owned()),
        );

        all
    }

    /// Returns all section text joined with newlines.
    #[must_use]
    pub fn plain_text(&self) -> String {
        let mut text = String::new();
        if let Some(length) = self.checked_text_len() {
            text.reserve(length);
        }
        self.append_plain_text(&mut text);
        text
    }

    /// Returns whether the section has no modeled content.
    ///
    /// Empty rich-text storages are native container artifacts rather than
    /// semantic content. Heading and paragraph presence remains modeled
    /// content, preserving their structural meaning even when their text is
    /// empty.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        if self.heading.is_some() || !self.paragraphs.is_empty() {
            return false;
        }

        let mut index = 0;
        while index < self.text_storages.len() {
            if !self.text_storages[index].is_empty() {
                return false;
            }
            index += 1;
        }
        true
    }

    /// Return the rendered section length when all additions are representable.
    pub(crate) fn checked_text_len(&self) -> Option<usize> {
        let mut length = 0usize;
        let mut values = 0usize;

        if let Some(heading) = &self.heading {
            length = length.checked_add(heading.len())?;
            values = values.checked_add(1)?;
        }
        for paragraph in &self.paragraphs {
            length = length.checked_add(paragraph.len())?;
            values = values.checked_add(1)?;
        }
        for storage in &self.text_storages {
            if !storage.is_empty() {
                length = length.checked_add(storage.len())?;
                values = values.checked_add(1)?;
            }
        }

        length.checked_add(values.saturating_sub(1))
    }

    pub(crate) fn append_plain_text(&self, output: &mut String) {
        let mut first = true;
        if let Some(heading) = &self.heading {
            append_value(output, &mut first, heading);
        }
        for paragraph in &self.paragraphs {
            append_value(output, &mut first, paragraph);
        }
        for storage in &self.text_storages {
            if !storage.is_empty() {
                append_value(output, &mut first, storage.text());
            }
        }
    }
}

/// A detached, mutable builder for an immutable section.
#[derive(Debug)]
pub struct Builder {
    index: usize,
    section_type: SectionType,
    name: Option<Box<str>>,
    heading: Option<Box<str>>,
    paragraphs: Vec<String>,
    text_storages: Vec<Storage>,
    page_count: Option<usize>,
}

impl Builder {
    /// Creates an empty section builder at `index`.
    #[must_use]
    pub fn new(index: usize, section_type: SectionType) -> Self {
        Self {
            index,
            section_type,
            name: None,
            heading: None,
            paragraphs: Vec::new(),
            text_storages: Vec::new(),
            page_count: None,
        }
    }

    /// Sets or clears the producer-visible section name.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NameContainsNul`] when the name contains a native
    /// string terminator. The current value is preserved when validation
    /// fails.
    pub fn set_name(&mut self, input: Option<impl Into<Box<str>>>) -> Result<()> {
        let name = input.map(Into::into);
        if name.as_deref().is_some_and(|value| value.contains('\0')) {
            return Err(Error::NameContainsNul);
        }
        self.name = name;
        Ok(())
    }

    /// Sets or clears the section heading.
    pub fn set_heading(&mut self, heading: Option<String>) {
        self.heading = heading.map(String::into_boxed_str);
    }

    /// Appends one paragraph in source order.
    pub fn push_paragraph(&mut self, paragraph: String) {
        self.paragraphs.push(paragraph);
    }

    /// Appends one rich-text storage in source order.
    pub fn push_text_storage(&mut self, storage: Storage) {
        self.text_storages.push(storage);
    }

    /// Sets or clears the known page count.
    pub fn set_page_count(&mut self, page_count: Option<usize>) {
        self.page_count = page_count;
    }

    /// Finishes the detached builder as an immutable section snapshot.
    #[must_use]
    pub fn build(self) -> Section {
        Section {
            index: self.index,
            section_type: self.section_type,
            name: self.name,
            heading: self.heading,
            paragraphs: self.paragraphs.into_boxed_slice(),
            text_storages: self.text_storages.into_boxed_slice(),
            page_count: self.page_count,
        }
    }
}

/// Semantic section kinds used by Pages.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[allow(
    clippy::module_name_repetitions,
    reason = "SectionType is the established public semantic name for this module."
)]
pub enum SectionType {
    /// Main body content.
    Body,
    /// Header content.
    Header,
    /// Footer content.
    Footer,
    /// Floating or anchored section content.
    Floating,
}

impl SectionType {
    /// Returns a stable human-readable name.
    #[must_use]
    pub const fn name(self) -> &'static str {
        match self {
            Self::Body => "Body",
            Self::Header => "Header",
            Self::Footer => "Footer",
            Self::Floating => "Floating",
        }
    }
}

fn validate_background_payload(payload: &[u8]) -> Result<()> {
    if payload.is_empty() {
        return Err(Error::EmptyBackgroundPayload);
    }
    if payload.len() > MAX_BACKGROUND_PAYLOAD_BYTES {
        return Err(Error::BackgroundPayloadTooLarge);
    }
    Ok(())
}

fn append_value(output: &mut String, first: &mut bool, value: &str) {
    if !*first {
        output.push('\n');
    }
    output.push_str(value);
    *first = false;
}

#[cfg(test)]
mod tests {
    use std::mem::size_of;

    use super::*;

    #[test]
    fn settings_pack_presence_and_retain_lossless_values() {
        let mut settings = Settings::new();
        settings
            .set_name(Some("Overview"))
            .unwrap_or_else(|error| panic!("valid section name: {error}"));
        settings.set_inherit_previous_header_footer(Some(false));
        settings.set_first_page_different(None);
        settings.set_even_odd_pages_different(Some(true));
        settings.set_first_page_hides_header_footer(Some(true));
        settings
            .set_start(Some(Start::Unknown(7)))
            .unwrap_or_else(|error| panic!("valid section start: {error}"));
        settings
            .set_page_numbering(Some(PageNumbering::Restart))
            .unwrap_or_else(|error| panic!("valid page numbering: {error}"));
        settings.set_starting_page_number(Some(
            PageNumber::new(4)
                .unwrap_or_else(|error| panic!("valid starting page number: {error}")),
        ));

        assert!(size_of::<Settings>() <= 48);
        assert_eq!(settings.name(), Some("Overview"));
        assert_eq!(settings.inherit_previous_header_footer(), Some(false));
        assert_eq!(settings.first_page_different(), None);
        assert_eq!(settings.even_odd_pages_different(), Some(true));
        assert_eq!(settings.first_page_hides_header_footer(), Some(true));
        assert_eq!(settings.start(), Some(Start::Unknown(7)));
        assert_eq!(settings.page_numbering(), Some(PageNumbering::Restart));
        assert_eq!(
            settings.starting_page_number().map(PageNumber::get),
            Some(4)
        );
        assert!(settings.validate().is_ok());
    }

    #[test]
    fn settings_reject_noncanonical_values_and_nul_names() {
        let mut settings = Settings::new();
        assert_eq!(
            settings.set_name(Some("bad\0name")),
            Err(Error::NameContainsNul)
        );
        assert_eq!(
            settings.set_start(Some(Start::Unknown(0))),
            Err(Error::NonCanonicalStart)
        );
        assert_eq!(
            settings.set_page_numbering(Some(PageNumbering::Unknown(1))),
            Err(Error::NonCanonicalNumbering)
        );
    }

    #[test]
    fn background_opaque_owns_exact_storage() {
        let opaque = Opaque::from_slice(&[0x0a, 0xff])
            .unwrap_or_else(|error| panic!("valid opaque background: {error}"));
        assert_eq!(opaque.as_bytes(), [0x0a, 0xff]);
        assert_eq!(opaque.into_bytes().as_ref(), [0x0a, 0xff]);
        assert_eq!(
            Opaque::from_slice(&[]).err(),
            Some(Error::EmptyBackgroundPayload)
        );
        let oversized = vec![0_u8; MAX_BACKGROUND_PAYLOAD_BYTES + 1];
        assert_eq!(
            Opaque::from_slice(&oversized).err(),
            Some(Error::BackgroundPayloadTooLarge)
        );
        assert_eq!(
            Background::Opaque(
                Opaque::from_slice(&[0x01])
                    .unwrap_or_else(|error| panic!("valid opaque background: {error}")),
            ),
            Background::Opaque(
                Opaque::from_slice(&[0x01])
                    .unwrap_or_else(|error| panic!("valid opaque background: {error}")),
            )
        );
    }

    #[test]
    fn section_creation_and_text() {
        let empty = Section::new(0, SectionType::Body);
        assert_eq!(empty.index(), 0);
        assert_eq!(empty.section_type(), SectionType::Body);
        assert_eq!(empty.page_count(), None);
        assert!(empty.is_empty());

        let mut builder = Section::builder(0, SectionType::Body);
        builder
            .set_name(Some("Chapter One"))
            .unwrap_or_else(|error| panic!("valid section name: {error}"));
        builder.set_heading(Some("Introduction".to_owned()));
        builder.push_paragraph("First paragraph".to_owned());
        builder.push_text_storage(Storage::from_text("Storage text".to_owned()));
        builder.set_page_count(Some(3));
        let section = builder.build();

        assert!(!section.is_empty());
        assert_eq!(section.section_type(), SectionType::Body);
        assert_eq!(section.name(), Some("Chapter One"));
        assert_eq!(section.selector(), SectionSelector::name("Chapter One"));
        assert_eq!(
            section
                .position_selector()
                .as_position()
                .map(litchi_core::Position::get),
            Some(0)
        );
        assert_eq!(section.page_count(), Some(3));
        assert_eq!(section.heading(), Some("Introduction"));
        assert_eq!(section.paragraphs(), ["First paragraph"]);
        assert_eq!(
            section
                .text_storages()
                .iter()
                .map(Storage::text)
                .collect::<Vec<_>>(),
            ["Storage text"]
        );
        assert_eq!(
            section.all_text(),
            ["Introduction", "First paragraph", "Storage text"]
        );
        assert_eq!(
            section.plain_text(),
            "Introduction\nFirst paragraph\nStorage text"
        );
        assert_eq!(section.checked_text_len(), Some(section.plain_text().len()));
    }

    #[test]
    fn section_without_text_storages_is_semantically_empty() {
        let section = Section::new(0, SectionType::Body);

        assert!(section.text_storages().is_empty());
        assert_eq!(section.plain_text(), "");
        assert!(section.is_empty());
    }

    #[test]
    fn one_empty_text_storage_is_semantically_empty() {
        let mut builder = Section::builder(0, SectionType::Body);
        builder.push_text_storage(Storage::new());
        let section = builder.build();

        assert_eq!(section.text_storages().len(), 1);
        assert_eq!(section.plain_text(), "");
        assert!(section.is_empty());
    }

    #[test]
    fn multiple_empty_text_storages_are_semantically_empty() {
        let mut builder = Section::builder(0, SectionType::Body);
        builder.push_text_storage(Storage::new());
        builder.push_text_storage(Storage::from_text(String::new()));
        builder.push_text_storage(Storage::new());
        let section = builder.build();

        assert_eq!(section.text_storages().len(), 3);
        assert_eq!(section.plain_text(), "");
        assert!(section.is_empty());
    }

    #[test]
    fn any_nonempty_text_storage_makes_section_nonempty() {
        let mut builder = Section::builder(0, SectionType::Body);
        builder.push_text_storage(Storage::new());
        builder.push_text_storage(Storage::from_text("content".to_owned()));
        builder.push_text_storage(Storage::new());
        let section = builder.build();

        assert_eq!(section.plain_text(), "content");
        assert!(!section.is_empty());
    }

    #[test]
    fn unnamed_sections_select_by_typed_position() {
        let section = Section::new(7, SectionType::Body);
        assert_eq!(section.selector(), section.position_selector());
        assert_eq!(
            section
                .selector()
                .as_position()
                .map(litchi_core::Position::get),
            Some(7)
        );
    }

    #[test]
    fn section_builder_rejects_invalid_names_without_mutation() {
        let mut builder = Section::builder(0, SectionType::Body);
        builder
            .set_name(Some("Valid"))
            .unwrap_or_else(|error| panic!("valid section name: {error}"));

        assert_eq!(
            builder.set_name(Some("invalid\0name")),
            Err(Error::NameContainsNul)
        );
        assert_eq!(builder.build().name(), Some("Valid"));
    }

    #[test]
    fn section_type_names_are_stable() {
        assert_eq!(SectionType::Body.name(), "Body");
        assert_eq!(SectionType::Header.name(), "Header");
        assert_eq!(SectionType::Footer.name(), "Footer");
        assert_eq!(SectionType::Floating.name(), "Floating");
    }
}
