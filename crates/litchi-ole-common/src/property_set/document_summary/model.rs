//! Semantic Document Summary Information snapshots and PIDDSI identifiers.

use super::super::Binding;
use super::super::codec::PropertySetReader;
use super::super::model::{
    CodePage, DOCUMENT_SUMMARY_INFORMATION_FMTID, Section, Stream, Value, try_clone_property_set,
};
use super::transaction::Transaction;
use super::validation::validate_section;
use litchi_cfb::{OleError, OleFile};
use std::io::{Read, Seek};

/// The Document Summary Information CodePage property identifier.
pub const CODEPAGE: u32 = super::super::model::PID_CODEPAGE;
/// The Category property identifier.
pub const CATEGORY: u32 = 0x0000_0002;
/// The Presentation Format property identifier.
pub const PRESENTATION_FORMAT: u32 = 0x0000_0003;
/// The estimated document byte-count property identifier.
pub const BYTE_COUNT: u32 = 0x0000_0004;
/// The estimated text-line count property identifier.
pub const LINE_COUNT: u32 = 0x0000_0005;
/// The paragraph-count property identifier.
pub const PARAGRAPH_COUNT: u32 = 0x0000_0006;
/// The slide-count property identifier.
pub const SLIDE_COUNT: u32 = 0x0000_0007;
/// The note-count property identifier.
pub const NOTE_COUNT: u32 = 0x0000_0008;
/// The hidden-slide count property identifier.
pub const HIDDEN_COUNT: u32 = 0x0000_0009;
/// The multimedia-clip count property identifier.
pub const MULTIMEDIA_CLIP_COUNT: u32 = 0x0000_000A;
/// The scale property identifier.
pub const SCALE: u32 = 0x0000_000B;
/// The HeadingPairs property identifier.
pub const HEADING_PAIRS: u32 = super::super::model::PID_HEADING_PAIRS;
/// The DocParts property identifier.
pub const DOCUMENT_PARTS: u32 = super::super::model::PID_DOC_PARTS;
/// The Manager property identifier.
pub const MANAGER: u32 = 0x0000_000E;
/// The Company property identifier.
pub const COMPANY: u32 = 0x0000_000F;
/// The LinksDirty property identifier.
pub const LINKS_DIRTY: u32 = 0x0000_0010;
/// The character-count-with-spaces property identifier.
pub const CHARACTER_COUNT_WITH_SPACES: u32 = 0x0000_0011;
/// The SharedDocument property identifier.
pub const SHARED_DOCUMENT: u32 = 0x0000_0013;
/// The LinkBase property identifier, which MS-OSHARED reserves from writing.
pub const LINK_BASE: u32 = 0x0000_0014;
/// The Hyperlinks property identifier, which MS-OSHARED reserves from writing.
pub const HYPERLINKS: u32 = 0x0000_0015;
/// The HyperlinksChanged property identifier.
pub const HYPERLINKS_CHANGED: u32 = 0x0000_0016;
/// The application-version property identifier.
pub const VERSION: u32 = 0x0000_0017;
/// The Digital Signature property identifier.
pub const DIGITAL_SIGNATURE: u32 = 0x0000_0018;
/// The ContentType property identifier.
pub const CONTENT_TYPE: u32 = 0x0000_001A;
/// The ContentStatus property identifier.
pub const CONTENT_STATUS: u32 = 0x0000_001B;
/// The Language property identifier.
pub const LANGUAGE: u32 = 0x0000_001C;
/// The document-version property identifier.
pub const DOCUMENT_VERSION: u32 = 0x0000_001D;

/// Maximum UTF-8 payload accepted by a typed string edit.
///
/// The limit keeps a single semantic field below the conservative simple
/// Property Set stream budget while allowing ordinary Office metadata values.
pub const MAX_TEXT_BYTES: usize = super::super::model::MAX_TYPED_TEXT_BYTES;

/// The two unsigned 16-bit components encoded by the PIDDSI `Version` I4.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Version {
    major: u16,
    minor: u16,
}

impl Version {
    /// Constructs a version, rejecting a zero major component required by
    /// [MS-OSHARED].
    pub const fn new(major: u16, minor: u16) -> Option<Self> {
        if major == 0 {
            None
        } else {
            Some(Self { major, minor })
        }
    }

    /// Decodes the signed `VT_I4` wire representation into its unsigned
    /// major/minor components.
    pub fn from_raw(raw: i32) -> Result<Self, OleError> {
        let raw = raw as u32;
        Self::new((raw >> 16) as u16, raw as u16)
            .ok_or_else(|| super::super::model::invalid("Document version major is zero"))
    }

    /// Returns the major application version.
    pub const fn major(self) -> u16 {
        self.major
    }

    /// Returns the minor application version.
    pub const fn minor(self) -> u16 {
        self.minor
    }

    /// Returns the exact signed `VT_I4` representation.
    pub const fn raw(self) -> i32 {
        (((self.major as u32) << 16) | self.minor as u32) as i32
    }
}

/// An immutable, validated Document Summary Information property-set view.
///
/// The complete generic section is retained, including properties that this
/// typed owner does not understand.
#[derive(Debug, Clone, PartialEq)]
pub struct Snapshot {
    pub(crate) section: Section,
}

impl Snapshot {
    /// Creates a new empty document-summary section with a required code page.
    pub fn new(page: CodePage) -> Result<Self, OleError> {
        let mut section = Section::new(DOCUMENT_SUMMARY_INFORMATION_FMTID);
        section.set_page(page);
        Self::from_section(&section)
    }

    /// Validates and clones a generic Document Summary Information section.
    pub fn from_section(section: &Section) -> Result<Self, OleError> {
        validate_section(section)?;
        Ok(Self {
            section: try_clone_property_set(section)?,
        })
    }

    /// Projects the Document Summary Information section from a version-zero
    /// Property Set stream.
    pub fn from_stream(stream: &Stream) -> Result<Self, OleError> {
        if stream.version != Stream::VERSION_0 {
            return Err(super::super::model::invalid(
                "Document Summary Information requires Property Set version 0",
            ));
        }
        let section = stream
            .section(DOCUMENT_SUMMARY_INFORMATION_FMTID)
            .ok_or_else(|| {
                super::super::model::invalid(
                    "Property Set stream has no Document Summary Information section",
                )
            })?;
        Self::from_section(section)
    }

    /// Reads and projects the standard Document Summary Information stream.
    pub fn from_ole<R: Read + Seek>(ole: &mut OleFile<R>) -> Result<Self, OleError> {
        let stream = ole.property_set(Binding::DocumentSummaryInformation)?;
        Self::from_stream(&stream)
    }

    /// Borrows the complete generic section, including opaque properties.
    pub const fn section(&self) -> &Section {
        &self.section
    }

    /// Returns a raw property by PID for extensions not modeled here.
    pub fn property(&self, identifier: u32) -> Option<&Value> {
        self.section.property(identifier)
    }

    /// Starts a source-bound transactional edit.
    pub fn transaction(&self) -> Result<Transaction<'_>, OleError> {
        Transaction::from_snapshot(self)
    }

    /// Consumes the view and returns its complete generic section.
    pub fn into_section(self) -> Section {
        self.section
    }
}
