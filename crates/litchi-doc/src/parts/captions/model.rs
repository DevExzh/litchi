//! Semantic values for the Word caption tables.

use super::Format;
use super::validation::{
    validate_auto_entries, validate_auto_entry, validate_definition, validate_definitions,
};
use crate::package::Result;

/// Where Word places an automatically generated caption relative to its host.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Location {
    /// Place the caption below the selected item.
    Below = 0x0,
    /// Place the caption above the selected item.
    Above = 0x1,
}

impl Location {
    pub(crate) fn from_raw(value: u8) -> Result<Self> {
        super::validation::location_from_raw(value)
    }
}

/// Heading level that starts a new chapter for caption numbering.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u8)]
pub enum Heading {
    Level1 = 0x1,
    Level2 = 0x2,
    Level3 = 0x3,
    Level4 = 0x4,
    Level5 = 0x5,
    Level6 = 0x6,
    Level7 = 0x7,
    Level8 = 0x8,
    Level9 = 0x9,
}

impl Heading {
    pub(crate) fn from_raw(value: u8) -> Result<Self> {
        super::validation::heading_from_raw(value)
    }
}

/// Character separating a chapter number from a caption number.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u16)]
pub enum Separator {
    Hyphen = 0x001E,
    Period = 0x002E,
    Colon = 0x003A,
    EnDash = 0x2013,
    EmDash = 0x2014,
}

impl Separator {
    pub(crate) fn from_raw(value: u16) -> Result<Self> {
        super::validation::separator_from_raw(value)
    }
}

/// Chapter-numbering options carried by a caption definition.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Numbering {
    heading: Heading,
    separator: Separator,
}

impl Numbering {
    pub const fn new(heading: Heading, separator: Separator) -> Self {
        Self { heading, separator }
    }

    /// Heading level that marks the beginning of a new chapter.
    pub const fn heading(self) -> Heading {
        self.heading
    }

    /// Character between the chapter and caption numbers.
    pub const fn separator(self) -> Separator {
        self.separator
    }
}

/// Caption insertion and numbering metadata (`CAPI`, MS-DOC 2.9.24).
#[derive(Debug, Clone, Copy)]
pub struct Info {
    location: Location,
    numbering: Option<Numbering>,
    omit_label: bool,
    number_format: Format,
    /// Undefined CAPI flag bits retained for lossless round-trips.
    pub(crate) raw_flags: u16,
    /// The separator is ignored when chapter numbering is disabled, but its
    /// source value is retained so editing a neighboring field is lossless.
    pub(crate) raw_separator: u16,
}

impl Info {
    /// Create valid semantic caption metadata using the shared MSONFC value
    /// defined by MS-OSHARED 2.2.1.3.
    pub const fn new(
        location: Location,
        numbering: Option<Numbering>,
        omit_label: bool,
        number_format: Format,
    ) -> Self {
        Self {
            location,
            numbering,
            omit_label,
            number_format,
            raw_flags: 0,
            raw_separator: 0,
        }
    }

    /// Where the caption is inserted relative to its host.
    pub const fn location(self) -> Location {
        self.location
    }

    /// Chapter-numbering options, if chapter numbers are included.
    pub const fn numbering(self) -> Option<Numbering> {
        self.numbering
    }

    /// Chapter-numbering options, phrased using the protocol field name.
    pub const fn chapter_numbering(self) -> Option<Numbering> {
        self.numbering()
    }

    /// Whether the label is omitted from the generated caption.
    pub const fn omit_label(self) -> bool {
        self.omit_label
    }

    /// Opaque MSONFC number format for the caption number.
    pub const fn number_format(self) -> Format {
        self.number_format
    }
}

impl PartialEq for Info {
    fn eq(&self, other: &Self) -> bool {
        self.location == other.location
            && self.numbering == other.numbering
            && self.omit_label == other.omit_label
            && self.number_format == other.number_format
    }
}

impl Eq for Info {}

/// One caption label and its insertion metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Definition {
    label: String,
    info: Info,
}

impl Definition {
    /// Construct a definition while enforcing the 40 UTF-16-unit label cap.
    pub fn try_new(label: String, info: Info) -> Result<Self> {
        let value = Self { label, info };
        validate_definition(&value)?;
        Ok(value)
    }

    /// Caption label text.
    pub fn label(&self) -> &str {
        &self.label
    }

    /// Caption insertion metadata.
    pub const fn info(&self) -> Info {
        self.info
    }
}

/// One OLE ProgID mapping from `SttbfAutoCaption` to `LabelTable`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct AutoEntry {
    prog_id: String,
    caption_index: u16,
}

impl AutoEntry {
    /// Construct a validated ProgID-to-label mapping.
    pub fn try_new(prog_id: String, caption_index: u16) -> Result<Self> {
        let value = Self {
            prog_id,
            caption_index,
        };
        validate_auto_entry(&value)?;
        Ok(value)
    }

    /// OLE ProgID associated with the automatic caption rule.
    pub fn prog_id(&self) -> &str {
        &self.prog_id
    }

    /// Zero-based label-table index selected by the rule.
    pub const fn caption_index(&self) -> u16 {
        self.caption_index
    }
}

/// The ordered `SttbfCaption` label definitions.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct LabelTable {
    definitions: Vec<Definition>,
}

impl LabelTable {
    /// Construct a validated label table.
    pub fn try_new(definitions: Vec<Definition>) -> Result<Self> {
        validate_definitions(&definitions)?;
        Ok(Self { definitions })
    }

    /// Definitions in their on-disk order.
    pub fn definitions(&self) -> &[Definition] {
        &self.definitions
    }

    pub const fn len(&self) -> usize {
        self.definitions.len()
    }

    pub const fn is_empty(&self) -> bool {
        self.definitions.is_empty()
    }
}

/// The ordered `SttbfAutoCaption` ProgID mappings.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct AutoTable {
    entries: Vec<AutoEntry>,
}

impl AutoTable {
    /// Construct a validated automatic-caption table.
    pub fn try_new(entries: Vec<AutoEntry>) -> Result<Self> {
        validate_auto_entries(&entries)?;
        Ok(Self { entries })
    }

    /// Mappings in their on-disk order.
    pub fn entries(&self) -> &[AutoEntry] {
        &self.entries
    }

    pub const fn len(&self) -> usize {
        self.entries.len()
    }

    pub const fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }
}
