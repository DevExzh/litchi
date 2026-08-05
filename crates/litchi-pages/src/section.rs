//! Archive-free Pages section content, pagination, settings, and backgrounds.

pub mod pagination;

use litchi_iwa_common::color::Rgba;
use litchi_iwa_text::TextStorage;
use thiserror::Error;

pub use pagination::{PageNumber, PageNumbering, Start};

const INHERIT_HEADER_FOOTER_PRESENT: u16 = 1 << 0;
const INHERIT_HEADER_FOOTER_VALUE: u16 = 1 << 1;
const FIRST_PAGE_DIFFERENT_PRESENT: u16 = 1 << 2;
const FIRST_PAGE_DIFFERENT_VALUE: u16 = 1 << 3;
const EVEN_ODD_PAGES_DIFFERENT_PRESENT: u16 = 1 << 4;
const EVEN_ODD_PAGES_DIFFERENT_VALUE: u16 = 1 << 5;
const FIRST_PAGE_HIDES_HEADER_FOOTER_PRESENT: u16 = 1 << 6;
const FIRST_PAGE_HIDES_HEADER_FOOTER_VALUE: u16 = 1 << 7;
const START_PRESENT: u16 = 1 << 8;
const PAGE_NUMBERING_PRESENT: u16 = 1 << 9;

/// Validation failures for Pages section semantic values.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum Error {
    /// A section name contains an embedded NUL character.
    #[error("Pages section names cannot contain NUL")]
    NameContainsNul,
    /// A known native section-start value was represented by `Unknown`.
    #[error("Pages section start must use its canonical variant for a known value")]
    NonCanonicalStart,
    /// A known native page-numbering value was represented by `Unknown`.
    #[error("Pages page numbering must use its canonical variant for a known value")]
    NonCanonicalNumbering,
}

/// Result type for Pages section semantic construction and mutation.
pub type Result<T> = std::result::Result<T, Error>;

/// Lossless, archive-free settings attached to one Pages section.
///
/// The value keeps native field presence in a compact bitset. Pagination
/// values retain their native discriminants through the focused types in
/// [`pagination`], while section names use one owned boxed string. Protobuf
/// messages, object identifiers, background-fill bytes, and package state are
/// deliberately absent; those belong to the IWA adapter.
#[repr(C)]
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Settings {
    name: Option<Box<str>>,
    start: u32,
    page_numbering: u32,
    starting_page_number: Option<PageNumber>,
    flags: u16,
}

impl Settings {
    /// Construct settings while retaining the presence of every native field.
    ///
    /// # Errors
    ///
    /// Returns an error when `name` contains NUL or a pagination value uses an
    /// `Unknown` variant for a known native discriminant.
    #[allow(
        clippy::too_many_arguments,
        reason = "The constructor mirrors the complete semantic section record."
    )]
    pub fn new(
        name: Option<String>,
        inherit_previous_header_footer: Option<bool>,
        first_page_different: Option<bool>,
        even_odd_pages_different: Option<bool>,
        start: Option<Start>,
        page_numbering: Option<PageNumbering>,
        starting_page_number: Option<PageNumber>,
        first_page_hides_header_footer: Option<bool>,
    ) -> Result<Self> {
        let mut settings = Self::empty();
        settings.set_name(name)?;
        settings.set_inherit_previous_header_footer(inherit_previous_header_footer);
        settings.set_first_page_different(first_page_different);
        settings.set_even_odd_pages_different(even_odd_pages_different);
        settings.set_start(start)?;
        settings.set_page_numbering(page_numbering)?;
        settings.set_starting_page_number(starting_page_number);
        settings.set_first_page_hides_header_footer(first_page_hides_header_footer);
        Ok(settings)
    }

    /// Return empty settings with every native field absent.
    #[must_use]
    pub const fn empty() -> Self {
        Self {
            name: None,
            start: 0,
            page_numbering: 0,
            starting_page_number: None,
            flags: 0,
        }
    }

    /// Return the optional semantic section name.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Replace or clear the section name.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NameContainsNul`] when `name` contains an embedded
    /// NUL character.
    pub fn set_name(&mut self, name: Option<String>) -> Result<()> {
        if name.as_deref().is_some_and(|value| value.contains('\0')) {
            return Err(Error::NameContainsNul);
        }
        self.name = name.map(String::into_boxed_str);
        Ok(())
    }

    /// Return the optional inherit-previous-header/footer flag.
    #[must_use]
    pub const fn inherit_previous_header_footer(&self) -> Option<bool> {
        flag_value(
            self.flags,
            INHERIT_HEADER_FOOTER_PRESENT,
            INHERIT_HEADER_FOOTER_VALUE,
        )
    }

    /// Set or clear the inherit-previous-header/footer flag.
    pub const fn set_inherit_previous_header_footer(&mut self, value: Option<bool>) {
        set_flag(
            &mut self.flags,
            INHERIT_HEADER_FOOTER_PRESENT,
            INHERIT_HEADER_FOOTER_VALUE,
            value,
        );
    }

    /// Return the optional first-page-different flag.
    #[must_use]
    pub const fn first_page_different(&self) -> Option<bool> {
        flag_value(
            self.flags,
            FIRST_PAGE_DIFFERENT_PRESENT,
            FIRST_PAGE_DIFFERENT_VALUE,
        )
    }

    /// Set or clear the first-page-different flag.
    pub const fn set_first_page_different(&mut self, value: Option<bool>) {
        set_flag(
            &mut self.flags,
            FIRST_PAGE_DIFFERENT_PRESENT,
            FIRST_PAGE_DIFFERENT_VALUE,
            value,
        );
    }

    /// Return the optional even/odd-pages-different flag.
    #[must_use]
    pub const fn even_odd_pages_different(&self) -> Option<bool> {
        flag_value(
            self.flags,
            EVEN_ODD_PAGES_DIFFERENT_PRESENT,
            EVEN_ODD_PAGES_DIFFERENT_VALUE,
        )
    }

    /// Set or clear the even/odd-pages-different flag.
    pub const fn set_even_odd_pages_different(&mut self, value: Option<bool>) {
        set_flag(
            &mut self.flags,
            EVEN_ODD_PAGES_DIFFERENT_PRESENT,
            EVEN_ODD_PAGES_DIFFERENT_VALUE,
            value,
        );
    }

    /// Return the optional page on which the section starts.
    #[must_use]
    pub const fn start(&self) -> Option<Start> {
        if self.flags & START_PRESENT != 0 {
            Some(Start::from_raw(self.start))
        } else {
            None
        }
    }

    /// Set or clear the section start.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonCanonicalStart`] when a known native value is
    /// passed through [`Start::Unknown`].
    pub const fn set_start(&mut self, value: Option<Start>) -> Result<()> {
        if let Some(candidate) = value
            && !candidate.is_canonical()
        {
            return Err(Error::NonCanonicalStart);
        }
        if let Some(candidate) = value {
            self.start = candidate.as_raw();
            self.flags |= START_PRESENT;
        } else {
            self.start = 0;
            self.flags &= !START_PRESENT;
        }
        Ok(())
    }

    /// Return the optional page-numbering behavior.
    #[must_use]
    pub const fn page_numbering(&self) -> Option<PageNumbering> {
        if self.flags & PAGE_NUMBERING_PRESENT != 0 {
            Some(PageNumbering::from_raw(self.page_numbering))
        } else {
            None
        }
    }

    /// Set or clear page-numbering behavior.
    ///
    /// # Errors
    ///
    /// Returns [`Error::NonCanonicalNumbering`] when a known native value is
    /// passed through [`PageNumbering::Unknown`].
    pub const fn set_page_numbering(&mut self, value: Option<PageNumbering>) -> Result<()> {
        if let Some(candidate) = value
            && !candidate.is_canonical()
        {
            return Err(Error::NonCanonicalNumbering);
        }
        if let Some(candidate) = value {
            self.page_numbering = candidate.as_raw();
            self.flags |= PAGE_NUMBERING_PRESENT;
        } else {
            self.page_numbering = 0;
            self.flags &= !PAGE_NUMBERING_PRESENT;
        }
        Ok(())
    }

    /// Return the optional first page number for this section.
    #[must_use]
    pub const fn starting_page_number(&self) -> Option<PageNumber> {
        self.starting_page_number
    }

    /// Set or clear the first page number for this section.
    pub const fn set_starting_page_number(&mut self, value: Option<PageNumber>) {
        self.starting_page_number = value;
    }

    /// Return the optional first-page header/footer hiding flag.
    #[must_use]
    pub const fn first_page_hides_header_footer(&self) -> Option<bool> {
        flag_value(
            self.flags,
            FIRST_PAGE_HIDES_HEADER_FOOTER_PRESENT,
            FIRST_PAGE_HIDES_HEADER_FOOTER_VALUE,
        )
    }

    /// Set or clear the first-page header/footer hiding flag.
    pub const fn set_first_page_hides_header_footer(&mut self, value: Option<bool>) {
        set_flag(
            &mut self.flags,
            FIRST_PAGE_HIDES_HEADER_FOOTER_PRESENT,
            FIRST_PAGE_HIDES_HEADER_FOOTER_VALUE,
            value,
        );
    }

    /// Return whether all optional section settings are absent.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.name.is_none() && self.starting_page_number.is_none() && self.flags == 0
    }

    /// Validate all semantic invariants.
    ///
    /// # Errors
    ///
    /// Returns an error if an invalid name or pagination discriminant is
    /// present. Values created through this module's setters are already
    /// checked; this method is useful at format boundaries before publication.
    pub fn validate(&self) -> Result<()> {
        if self
            .name
            .as_deref()
            .is_some_and(|value| value.contains('\0'))
        {
            return Err(Error::NameContainsNul);
        }
        if self.start().is_some_and(|value| !value.is_canonical()) {
            return Err(Error::NonCanonicalStart);
        }
        if self
            .page_numbering()
            .is_some_and(|value| !value.is_canonical())
        {
            return Err(Error::NonCanonicalNumbering);
        }
        Ok(())
    }
}

impl Default for Settings {
    fn default() -> Self {
        Self::empty()
    }
}

/// A Pages section background with an archive-free solid-color fast path.
///
/// Unsupported, future, or format-specific fills are retained as exact owned
/// bytes in [`Self::Opaque`]. The semantic crate does not decode or validate
/// those bytes; the IWA adapter owns that format-specific responsibility.
#[derive(Debug, Clone, PartialEq, Default)]
pub enum Background {
    /// No section background is present.
    #[default]
    None,
    /// A validated solid RGBA background.
    Solid(Rgba),
    /// An unsupported or future background retained byte-for-byte.
    Opaque(Box<[u8]>),
}

impl Background {
    /// Create an opaque background without retaining spare vector capacity.
    #[must_use]
    pub fn opaque(payload: impl Into<Box<[u8]>>) -> Self {
        Self::Opaque(payload.into())
    }

    /// Borrow an opaque background's exact bytes, when present.
    #[must_use]
    pub fn opaque_bytes(&self) -> Option<&[u8]> {
        match self {
            Self::Opaque(payload) => Some(payload),
            Self::None | Self::Solid(_) => None,
        }
    }

    /// Return the solid color, when this is a solid background.
    #[must_use]
    pub const fn solid_color(&self) -> Option<Rgba> {
        match self {
            Self::Solid(color) => Some(*color),
            Self::None | Self::Opaque(_) => None,
        }
    }
}

/// A logical Pages document section.
#[derive(Debug, Clone)]
pub struct Section {
    /// Zero-based section index.
    pub index: usize,
    /// Semantic kind of the section.
    pub section_type: SectionType,
    /// Optional section heading.
    pub heading: Option<String>,
    /// Paragraph values extracted from the section.
    pub paragraphs: Vec<String>,
    /// Rich-text storages belonging to the section.
    pub text_storages: Vec<TextStorage>,
    /// Number of pages represented by the section, when known.
    pub page_count: Option<usize>,
}

impl Section {
    /// Creates an empty section with `index` and `section_type`.
    #[must_use]
    pub fn new(index: usize, section_type: SectionType) -> Self {
        Self {
            index,
            section_type,
            heading: None,
            paragraphs: Vec::new(),
            text_storages: Vec::new(),
            page_count: None,
        }
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
            all.push(heading.clone());
        }
        all.extend(self.paragraphs.iter().cloned());

        all.extend(
            self.text_storages
                .iter()
                .filter(|storage| !storage.is_empty())
                .map(|storage| storage.plain_text().to_owned()),
        );

        all
    }

    /// Returns all section text joined with newlines.
    #[must_use]
    pub fn plain_text(&self) -> String {
        let mut text = String::with_capacity(self.text_len());
        self.append_plain_text(&mut text);
        text
    }

    /// Returns whether the section has no modeled content.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.heading.is_none() && self.paragraphs.is_empty() && self.text_storages.is_empty()
    }

    /// Returns the UTF-8 byte length of the rendered section text.
    pub(crate) fn text_len(&self) -> usize {
        let mut length = 0usize;
        let mut values = 0usize;

        if let Some(heading) = &self.heading {
            length = length.saturating_add(heading.len());
            values = values.saturating_add(1);
        }
        for paragraph in &self.paragraphs {
            length = length.saturating_add(paragraph.len());
            values = values.saturating_add(1);
        }
        for storage in &self.text_storages {
            if !storage.is_empty() {
                length = length.saturating_add(storage.len());
                values = values.saturating_add(1);
            }
        }

        length.saturating_add(values.saturating_sub(1))
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
                append_value(output, &mut first, storage.plain_text());
            }
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

/// Return a packed optional boolean.
const fn flag_value(flags: u16, present: u16, value: u16) -> Option<bool> {
    if flags & present != 0 {
        Some(flags & value != 0)
    } else {
        None
    }
}

/// Update a packed optional boolean.
const fn set_flag(flags: &mut u16, present: u16, value: u16, option: Option<bool>) {
    if let Some(enabled) = option {
        *flags |= present;
        if enabled {
            *flags |= value;
        } else {
            *flags &= !value;
        }
    } else {
        *flags &= !(present | value);
    }
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

    use litchi_iwa_common::color::{RgbColorSpace, Rgba};

    use super::*;

    #[test]
    fn section_creation_and_text() {
        let mut section = Section::new(0, SectionType::Body);
        assert_eq!(section.index, 0);
        assert_eq!(section.section_type, SectionType::Body);
        assert!(section.is_empty());

        section.heading = Some("Introduction".to_owned());
        section.paragraphs.push("First paragraph".to_owned());
        section
            .text_storages
            .push(TextStorage::from_text("Storage text".to_owned()));

        assert!(!section.is_empty());
        assert_eq!(
            section.all_text(),
            ["Introduction", "First paragraph", "Storage text"]
        );
        assert_eq!(
            section.plain_text(),
            "Introduction\nFirst paragraph\nStorage text"
        );
        assert_eq!(section.text_len(), section.plain_text().len());
    }

    #[test]
    fn section_type_names_are_stable() {
        assert_eq!(SectionType::Body.name(), "Body");
        assert_eq!(SectionType::Header.name(), "Header");
        assert_eq!(SectionType::Footer.name(), "Footer");
        assert_eq!(SectionType::Floating.name(), "Floating");
    }

    #[test]
    fn settings_pack_presence_and_retain_semantics() {
        assert_eq!(size_of::<Settings>(), 32);
        let page_number =
            PageNumber::new(7).unwrap_or_else(|error| panic!("valid Pages page number: {error}"));
        let settings = Settings::new(
            Some("Body".to_owned()),
            Some(true),
            Some(false),
            None,
            Some(Start::RightPage),
            Some(PageNumbering::Restart),
            Some(page_number),
            Some(true),
        )
        .unwrap_or_else(|error| panic!("valid Pages section settings: {error}"));

        assert_eq!(settings.name(), Some("Body"));
        assert_eq!(settings.inherit_previous_header_footer(), Some(true));
        assert_eq!(settings.first_page_different(), Some(false));
        assert_eq!(settings.even_odd_pages_different(), None);
        assert_eq!(settings.start(), Some(Start::RightPage));
        assert_eq!(settings.page_numbering(), Some(PageNumbering::Restart));
        assert_eq!(settings.starting_page_number(), Some(page_number));
        assert_eq!(settings.first_page_hides_header_footer(), Some(true));
        assert!(!settings.is_empty());
        assert!(settings.validate().is_ok());

        let mut empty = Settings::default();
        assert!(empty.is_empty());
        empty.set_first_page_different(Some(false));
        assert_eq!(empty.first_page_different(), Some(false));
        empty.set_first_page_different(None);
        assert!(empty.is_empty());
    }

    #[test]
    fn settings_reject_noncanonical_values_before_mutation() {
        let mut settings = Settings::default();
        settings
            .set_start(Some(Start::RightPage))
            .unwrap_or_else(|error| panic!("valid section start: {error}"));
        settings
            .set_page_numbering(Some(PageNumbering::Restart))
            .unwrap_or_else(|error| panic!("valid page numbering: {error}"));
        assert_eq!(
            settings.set_start(Some(Start::Unknown(1))),
            Err(Error::NonCanonicalStart)
        );
        assert_eq!(settings.start(), Some(Start::RightPage));
        assert_eq!(
            settings.set_page_numbering(Some(PageNumbering::Unknown(0))),
            Err(Error::NonCanonicalNumbering)
        );
        assert_eq!(settings.page_numbering(), Some(PageNumbering::Restart));
        assert_eq!(
            settings.set_name(Some("bad\0name".to_owned())),
            Err(Error::NameContainsNul)
        );
        assert_eq!(settings.name(), None);
    }

    #[test]
    fn background_has_solid_fast_path_and_exact_opaque_storage() {
        assert_eq!(size_of::<Background>(), 24);
        let color = Rgba::new(0.1, 0.2, 0.3, 0.4, RgbColorSpace::DisplayP3)
            .unwrap_or_else(|error| panic!("valid background color: {error}"));
        assert_eq!(Background::Solid(color).solid_color(), Some(color));
        assert_eq!(Background::None.solid_color(), None);

        let background = Background::opaque(vec![0, 1, 2, 3]);
        assert_eq!(background.opaque_bytes(), Some(&[0, 1, 2, 3][..]));
        assert_eq!(background.solid_color(), None);
        assert_eq!(Background::default(), Background::None);
    }
}
