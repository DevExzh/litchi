//! Typed `DocumentAtom` reader (MS-PPT 2.4.2).
//!
//! The atom carries document-wide slide geometry, OLE server zoom, master
//! persist references, and display flags. Parsing is limited to bytes
//! already present in caller-supplied PPT records; no persist object is
//! resolved and nothing is rendered.

use super::package::{Error, Result};
use super::records::Record;
use super::view_info::Ratio;
use crate::consts::RecordType;

/// Byte length of a `DocumentAtom` payload.
const DOCUMENT_ATOM_PAYLOAD_LEN: usize = 40;
/// Record version of a `DocumentAtom`.
const DOCUMENT_ATOM_VERSION: u16 = 0x1;
/// Smallest slide or notes dimension permitted by MS-PPT, in master units.
const MIN_MASTER_DIMENSION: i32 = 0x240;
/// Largest slide or notes dimension permitted by MS-PPT, in master units.
const MAX_MASTER_DIMENSION: i32 = 0x7E00;
/// Largest permitted starting slide number (`firstSlideNumber` MUST be less
/// than 10000).
const MAX_FIRST_SLIDE_NUMBER: u16 = 9999;

/// Presentation slide size types (MS-PPT 2.13.26).
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum SlideSizeType {
    /// Slide size ratio is consistent with a computer screen.
    #[default]
    Screen,
    /// Slide size ratio is consistent with letter paper.
    LetterPaper,
    /// Slide size ratio is consistent with A4 paper.
    A4Paper,
    /// Slide size ratio is consistent with 35mm photo slides.
    Photo35mm,
    /// Slide size ratio is consistent with overhead projector slides.
    Overhead,
    /// Slide size ratio is consistent with a banner.
    Banner,
    /// Slide size ratio not covered by any other variant.
    Custom,
}

impl SlideSizeType {
    fn from_u16(value: u16) -> Result<Self> {
        match value {
            0x0000 => Ok(Self::Screen),
            0x0001 => Ok(Self::LetterPaper),
            0x0002 => Ok(Self::A4Paper),
            0x0003 => Ok(Self::Photo35mm),
            0x0004 => Ok(Self::Overhead),
            0x0005 => Ok(Self::Banner),
            0x0006 => Ok(Self::Custom),
            _ => corrupted("DocumentAtom slideSizeType is not a SlideSizeEnum value"),
        }
    }

    /// The `SlideSizeEnum` value of this slide size type.
    #[must_use]
    pub fn as_u16(self) -> u16 {
        match self {
            Self::Screen => 0x0000,
            Self::LetterPaper => 0x0001,
            Self::A4Paper => 0x0002,
            Self::Photo35mm => 0x0003,
            Self::Overhead => 0x0004,
            Self::Banner => 0x0005,
            Self::Custom => 0x0006,
        }
    }
}

/// Width and height in master units, used for `slideSize` and `notesSize`.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct DocumentDimensions {
    pub width: i32,
    pub height: i32,
}

impl DocumentDimensions {
    /// Construct validated slide or notes dimensions in master units.
    ///
    /// # Errors
    ///
    /// Returns an error if either dimension is outside the MS-PPT master-unit
    /// range.
    pub fn new(width: i32, height: i32) -> Result<Self> {
        let dimensions = Self { width, height };
        dimensions.validate()?;
        Ok(dimensions)
    }

    fn validate(self) -> Result<()> {
        for dimension in [self.width, self.height] {
            if !(MIN_MASTER_DIMENSION..=MAX_MASTER_DIMENSION).contains(&dimension) {
                return corrupted("DocumentAtom dimension is outside the permitted range");
            }
        }
        Ok(())
    }
}

/// A validated `DocumentAtom` (MS-PPT 2.4.2).
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[allow(
    clippy::struct_excessive_bools,
    reason = "fields mirror the independent `DocumentAtom` wire-format flags one-to-one"
)]
pub struct DocumentAtom {
    /// Dimensions of the presentation slides in master units.
    pub slide_size: DocumentDimensions,
    /// Dimensions of the notes and handout slides in master units.
    pub notes_size: DocumentDimensions,
    /// Zoom level for OLE server representations of the document.
    pub server_zoom: Ratio,
    /// Persist identifier reference of the notes master slide.
    pub notes_master_persist_id_ref: u32,
    /// Persist identifier reference of the handout master slide.
    pub handout_master_persist_id_ref: u32,
    /// Starting number for slide numbering.
    pub first_slide_number: u16,
    /// Type of the presentation slide size.
    pub slide_size_type: SlideSizeType,
    /// Whether fonts are embedded in the document.
    pub save_with_fonts: bool,
    /// Whether placeholder shapes on the title slide are not displayed.
    pub omit_title_place: bool,
    /// Whether the user interface is optimized for right-to-left languages.
    pub right_to_left: bool,
    /// Whether presentation comments are displayed.
    pub show_comments: bool,
}

impl DocumentAtom {
    /// Construct and validate a `DocumentAtom` value.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[allow(
        clippy::too_many_arguments,
        clippy::fn_params_excessive_bools,
        reason = "constructor parameters mirror the atom wire fields one-to-one"
    )]
    pub fn new(
        slide_size: DocumentDimensions,
        notes_size: DocumentDimensions,
        server_zoom: Ratio,
        notes_master_persist_id_ref: u32,
        handout_master_persist_id_ref: u32,
        first_slide_number: u16,
        slide_size_type: SlideSizeType,
        save_with_fonts: bool,
        omit_title_place: bool,
        right_to_left: bool,
        show_comments: bool,
    ) -> Result<Self> {
        let atom = Self {
            slide_size,
            notes_size,
            server_zoom,
            notes_master_persist_id_ref,
            handout_master_persist_id_ref,
            first_slide_number,
            slide_size_type,
            save_with_fonts,
            omit_title_place,
            right_to_left,
            show_comments,
        };
        atom.validate()?;
        Ok(atom)
    }

    /// Strictly parse one already-materialized `RT_DocumentAtom` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the record header is malformed, the payload has the
    /// wrong size, or any field violates its MS-PPT constraints.
    #[allow(
        clippy::cast_possible_truncation,
        reason = "the document-atom payload length is the fixed constant 40, always representable as u32"
    )]
    pub fn parse(record: &Record) -> Result<Self> {
        if record.version != DOCUMENT_ATOM_VERSION
            || record.instance != 0
            || record.record_type_raw != RecordType::DocumentAtom.as_u16()
            || record.data.len() != DOCUMENT_ATOM_PAYLOAD_LEN
            || record.data_length != DOCUMENT_ATOM_PAYLOAD_LEN as u32
        {
            return corrupted("DocumentAtom has an invalid record header or size");
        }
        let data = &record.data;
        let atom = Self {
            slide_size: DocumentDimensions {
                width: read_i32(data, 0),
                height: read_i32(data, 4),
            },
            notes_size: DocumentDimensions {
                width: read_i32(data, 8),
                height: read_i32(data, 12),
            },
            server_zoom: Ratio::new(read_i32(data, 16), read_i32(data, 20))?,
            notes_master_persist_id_ref: read_u32(data, 24),
            handout_master_persist_id_ref: read_u32(data, 28),
            first_slide_number: read_u16(data, 32),
            slide_size_type: SlideSizeType::from_u16(read_u16(data, 34))?,
            save_with_fonts: strict_bool(data[36], "fSaveWithFonts")?,
            omit_title_place: strict_bool(data[37], "fOmitTitlePlace")?,
            right_to_left: strict_bool(data[38], "fRightToLeft")?,
            show_comments: strict_bool(data[39], "fShowComments")?,
        };
        atom.validate()?;
        Ok(atom)
    }

    /// Serialize this value as one canonical `RT_DocumentAtom` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn to_record(&self) -> Result<Record> {
        let bytes = self.to_record_bytes()?;
        let (record, end) = Record::parse(&bytes, 0)?;
        if end != bytes.len() {
            return corrupted("canonical DocumentAtom did not consume its bytes");
        }
        Ok(record)
    }

    /// Serialize this value as canonical `RT_DocumentAtom` record bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[allow(
        clippy::cast_possible_truncation,
        reason = "the document-atom payload length is the fixed constant 40, always representable as u32"
    )]
    pub fn to_record_bytes(&self) -> Result<[u8; 8 + DOCUMENT_ATOM_PAYLOAD_LEN]> {
        self.validate()?;
        let mut bytes = [0u8; 8 + DOCUMENT_ATOM_PAYLOAD_LEN];
        bytes[0..2].copy_from_slice(&DOCUMENT_ATOM_VERSION.to_le_bytes());
        bytes[2..4].copy_from_slice(&RecordType::DocumentAtom.as_u16().to_le_bytes());
        bytes[4..8].copy_from_slice(&(DOCUMENT_ATOM_PAYLOAD_LEN as u32).to_le_bytes());
        bytes[8..12].copy_from_slice(&self.slide_size.width.to_le_bytes());
        bytes[12..16].copy_from_slice(&self.slide_size.height.to_le_bytes());
        bytes[16..20].copy_from_slice(&self.notes_size.width.to_le_bytes());
        bytes[20..24].copy_from_slice(&self.notes_size.height.to_le_bytes());
        bytes[24..28].copy_from_slice(&self.server_zoom.numerator().to_le_bytes());
        bytes[28..32].copy_from_slice(&self.server_zoom.denominator().to_le_bytes());
        bytes[32..36].copy_from_slice(&self.notes_master_persist_id_ref.to_le_bytes());
        bytes[36..40].copy_from_slice(&self.handout_master_persist_id_ref.to_le_bytes());
        bytes[40..42].copy_from_slice(&self.first_slide_number.to_le_bytes());
        bytes[42..44].copy_from_slice(&self.slide_size_type.as_u16().to_le_bytes());
        bytes[44] = u8::from(self.save_with_fonts);
        bytes[45] = u8::from(self.omit_title_place);
        bytes[46] = u8::from(self.right_to_left);
        bytes[47] = u8::from(self.show_comments);
        Ok(bytes)
    }

    fn validate(&self) -> Result<()> {
        self.slide_size.validate()?;
        self.notes_size.validate()?;
        if self.first_slide_number > MAX_FIRST_SLIDE_NUMBER {
            return corrupted("DocumentAtom firstSlideNumber must be less than 10000");
        }
        let numerator = self.server_zoom.numerator();
        let denominator = self.server_zoom.denominator();
        if denominator == 0 || (numerator < 0) != (denominator < 0) || numerator == 0 {
            return corrupted("DocumentAtom serverZoom ratio must be greater than zero");
        }
        Ok(())
    }
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Corrupted(message.into()))
}

fn strict_bool(value: u8, field: &str) -> Result<bool> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => corrupted(format!("DocumentAtom {field} is not a bool1")),
    }
}

fn read_i32(data: &[u8], offset: usize) -> i32 {
    i32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
mod tests {
    use super::*;

    fn sample_atom() -> DocumentAtom {
        DocumentAtom::new(
            DocumentDimensions::new(5760, 4320).unwrap(),
            DocumentDimensions::new(4320, 5760).unwrap(),
            Ratio::new(1, 2).unwrap(),
            2,
            0,
            1,
            SlideSizeType::Screen,
            false,
            false,
            false,
            true,
        )
        .unwrap()
    }

    #[test]
    fn document_atom_roundtrips() {
        let atom = sample_atom();
        let record = atom.to_record().unwrap();
        assert_eq!(record.record_type, RecordType::DocumentAtom);
        assert_eq!(record.version, 1);
        assert_eq!(record.instance, 0);
        let parsed = DocumentAtom::parse(&record).unwrap();
        assert_eq!(parsed, atom);
        assert_eq!(parsed.slide_size_type.as_u16(), 0);
    }

    #[test]
    fn rejects_out_of_range_dimensions_slide_numbers_and_zoom() {
        assert!(DocumentDimensions::new(0x23f, 4320).is_err());
        assert!(DocumentDimensions::new(5760, 0x7e01).is_err());

        let mut atom = sample_atom();
        atom.first_slide_number = 10000;
        assert!(atom.to_record().is_err());

        let mut zero_zoom = sample_atom();
        zero_zoom.server_zoom = Ratio::new(0, 2).unwrap();
        assert!(zero_zoom.to_record().is_err());

        let mut negative_zoom = sample_atom();
        negative_zoom.server_zoom = Ratio::new(1, -2).unwrap();
        assert!(negative_zoom.to_record().is_err());
    }

    #[test]
    fn rejects_malformed_records() {
        let mut bad_flag = sample_atom().to_record_bytes().unwrap();
        bad_flag[8 + 36] = 2; // fSaveWithFonts is not a bool1
        let bad_flag_record = Record::parse(&bad_flag, 0).unwrap().0;
        assert!(DocumentAtom::parse(&bad_flag_record).is_err());

        let mut bad_size_type = sample_atom().to_record_bytes().unwrap();
        bad_size_type[8 + 34] = 7; // not a SlideSizeEnum value
        let bad_size_type_record = Record::parse(&bad_size_type, 0).unwrap().0;
        assert!(DocumentAtom::parse(&bad_size_type_record).is_err());

        let mut bad_version = sample_atom().to_record_bytes().unwrap();
        bad_version[0] = 0;
        let bad_version_record = Record::parse(&bad_version, 0).unwrap().0;
        assert!(DocumentAtom::parse(&bad_version_record).is_err());
    }

    #[test]
    fn presentation_exposes_real_document_atom() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/ppt/SampleShow.ppt");
        let mut package = crate::Package::open(path).unwrap();
        let presentation = package.presentation().unwrap();
        let atom = presentation.document_atom().unwrap().unwrap();

        assert_eq!(atom.slide_size.width, 5760);
        assert_eq!(atom.slide_size.height, 4320);
        assert_eq!(atom.notes_size.width, 4320);
        assert_eq!(atom.notes_size.height, 5760);
        assert_eq!(atom.server_zoom.numerator(), 1);
        assert_eq!(atom.server_zoom.denominator(), 2);
        assert_eq!(atom.notes_master_persist_id_ref, 2);
        assert_eq!(atom.handout_master_persist_id_ref, 0);
        assert_eq!(atom.first_slide_number, 1);
        assert_eq!(atom.slide_size_type, SlideSizeType::Screen);
        assert!(!atom.save_with_fonts);
        assert!(!atom.omit_title_place);
        assert!(!atom.right_to_left);
        assert!(atom.show_comments);
    }
}
