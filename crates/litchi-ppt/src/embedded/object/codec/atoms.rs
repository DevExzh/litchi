//! Fixed-width OLE atoms and their exact binary codecs.

use super::super::model::{
    ColorFollow, DimensionPolicy, DrawAspect, EmbedPreferences, LinkInfo, Metadata, ObjectSubtype,
    ObjectType, UpdateMode,
};
use super::wire::{corrupted, parse_bool, record_bytes, require_atom, u32_at};
use crate::consts::RecordType;
use crate::package::Result;
use crate::records::Record;

impl DrawAspect {
    pub(crate) fn parse(value: u32) -> Result<Self> {
        match value {
            1 => Ok(Self::Content),
            2 => Ok(Self::Thumbnail),
            4 => Ok(Self::Icon),
            8 => Ok(Self::DocumentPrint),
            _ => corrupted("ExOleObjAtom contains an invalid draw aspect"),
        }
    }
}

impl ObjectType {
    pub(crate) fn parse(value: u32) -> Result<Self> {
        match value {
            0 => Ok(Self::Embedded),
            1 => Ok(Self::Linked),
            2 => Ok(Self::ActiveXControl),
            _ => corrupted("ExOleObjAtom contains an invalid object type"),
        }
    }
}

impl ObjectSubtype {
    pub(crate) fn parse(value: u32) -> Result<Self> {
        match value {
            0 => Ok(Self::Default),
            1 => Ok(Self::ClipArtGallery),
            2 => Ok(Self::WordTable),
            3 => Ok(Self::Excel),
            4 => Ok(Self::Graph),
            5 => Ok(Self::OrganizationChart),
            6 => Ok(Self::Equation),
            7 => Ok(Self::WordArt),
            8 => Ok(Self::Sound),
            9 => Ok(Self::Image),
            10 => Ok(Self::Presentation),
            11 => Ok(Self::Slide),
            12 => Ok(Self::Project),
            13 => Ok(Self::NoteIt),
            14 => Ok(Self::ExcelChart),
            15 => Ok(Self::MediaPlayer),
            _ => corrupted("ExOleObjAtom contains an invalid subtype"),
        }
    }
}

impl Metadata {
    /// Parse the 24-byte `ExOleObjAtom` payload carried by `record`.
    ///
    /// # Errors
    ///
    /// Returns an error if the record header or size is invalid, if a
    /// fixed-width field contains an out-of-range value, or if the object or
    /// persist ID is zero.
    pub fn parse(record: &Record) -> Result<Self> {
        require_atom(
            record,
            1,
            0,
            RecordType::ExternalOleObjectAtom,
            24,
            "ExOleObjAtom",
        )?;
        let value = Self {
            draw_aspect: DrawAspect::parse(u32_at(&record.data, 0))?,
            object_type: ObjectType::parse(u32_at(&record.data, 4))?,
            id: u32_at(&record.data, 8),
            subtype: ObjectSubtype::parse(u32_at(&record.data, 12))?,
            persist_id: u32_at(&record.data, 16),
            unused: [
                record.data[20],
                record.data[21],
                record.data[22],
                record.data[23],
            ],
        };
        value.validate()?;
        Ok(value)
    }

    pub(crate) fn validate(&self) -> Result<()> {
        if self.id == 0 {
            return corrupted("ExOleObjAtom object ID must be positive");
        }
        if self.persist_id == 0 {
            return corrupted("ExOleObjAtom persist ID must be positive");
        }
        Ok(())
    }

    /// Serialize as a fully parsed `ExOleObjAtom` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the object or persist ID is zero, if the record
    /// payload exceeds the encodable size, or if the encoded record cannot be
    /// re-parsed.
    pub fn to_record(&self) -> Result<Record> {
        Ok(Record::parse(&self.to_record_bytes()?, 0)?.0)
    }

    /// Serialize to the raw bytes of a complete `ExOleObjAtom` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the object or persist ID is zero or if the record
    /// payload exceeds the encodable size.
    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        self.validate()?;
        let mut data = [0; 24];
        data[0..4].copy_from_slice(&(self.draw_aspect as u32).to_le_bytes());
        data[4..8].copy_from_slice(&(self.object_type as u32).to_le_bytes());
        data[8..12].copy_from_slice(&self.id.to_le_bytes());
        data[12..16].copy_from_slice(&(self.subtype as u32).to_le_bytes());
        data[16..20].copy_from_slice(&self.persist_id.to_le_bytes());
        data[20..24].copy_from_slice(&self.unused);
        record_bytes(1, 0, RecordType::ExternalOleObjectAtom, &data)
    }
}

impl ColorFollow {
    pub(crate) fn parse(value: u32) -> Result<Self> {
        match value {
            0 => Ok(Self::None),
            1 => Ok(Self::EntireScheme),
            2 => Ok(Self::TextAndBackground),
            _ => corrupted("ExOleEmbedAtom contains an invalid color-follow value"),
        }
    }
}

impl DimensionPolicy {
    pub(crate) fn parse(value: u8) -> Self {
        match value {
            0 => Self::Send,
            1 => Self::Omit,
            other => Self::ProducerDefined(other),
        }
    }

    pub(crate) fn value(self) -> u8 {
        match self {
            Self::Send => 0,
            Self::Omit => 1,
            Self::ProducerDefined(value) => value,
        }
    }
}

impl EmbedPreferences {
    /// Parse the 8-byte `ExOleEmbedAtom` payload carried by `record`.
    ///
    /// # Errors
    ///
    /// Returns an error if the record header or size is invalid or if a
    /// fixed-width field contains an out-of-range value.
    pub fn parse(record: &Record) -> Result<Self> {
        require_atom(
            record,
            0,
            0,
            RecordType::ExternalOleEmbedAtom,
            8,
            "ExOleEmbedAtom",
        )?;
        Ok(Self {
            color_follow: ColorFollow::parse(u32_at(&record.data, 0))?,
            cannot_lock_server: parse_bool(record.data[4], "fCantLockServer")?,
            dimension_policy: DimensionPolicy::parse(record.data[5]),
            is_word_table: parse_bool(record.data[6], "fIsTable")?,
            unused: record.data[7],
        })
    }

    /// Serialize as a fully parsed `ExOleEmbedAtom` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the record payload exceeds the encodable size or
    /// if the encoded record cannot be re-parsed.
    pub fn to_record(&self) -> Result<Record> {
        Ok(Record::parse(&self.to_record_bytes()?, 0)?.0)
    }

    /// Serialize to the raw bytes of a complete `ExOleEmbedAtom` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the record payload exceeds the encodable size.
    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        let mut data = [0; 8];
        data[0..4].copy_from_slice(&(self.color_follow as u32).to_le_bytes());
        data[4] = u8::from(self.cannot_lock_server);
        data[5] = self.dimension_policy.value();
        data[6] = u8::from(self.is_word_table);
        data[7] = self.unused;
        record_bytes(0, 0, RecordType::ExternalOleEmbedAtom, &data)
    }
}

impl UpdateMode {
    pub(crate) fn parse(value: u32) -> Result<Self> {
        match value {
            0 => Ok(Self::Always),
            1 => Ok(Self::OnCall),
            _ => corrupted("ExOleLinkAtom contains an invalid update mode"),
        }
    }
}

impl LinkInfo {
    /// Parse the 12-byte `ExOleLinkAtom` payload carried by `record`.
    ///
    /// # Errors
    ///
    /// Returns an error if the record header or size is invalid or if the
    /// update-mode field contains an out-of-range value.
    pub fn parse(record: &Record) -> Result<Self> {
        require_atom(
            record,
            0,
            0,
            RecordType::ExternalOleLinkAtom,
            12,
            "ExOleLinkAtom",
        )?;
        let slide_id = u32_at(&record.data, 0);
        Ok(Self {
            slide_id: (slide_id != 0).then_some(slide_id),
            update_mode: UpdateMode::parse(u32_at(&record.data, 4))?,
            unused: [
                record.data[8],
                record.data[9],
                record.data[10],
                record.data[11],
            ],
        })
    }

    /// Serialize as a fully parsed `ExOleLinkAtom` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the record payload exceeds the encodable size or
    /// if the encoded record cannot be re-parsed.
    pub fn to_record(&self) -> Result<Record> {
        Ok(Record::parse(&self.to_record_bytes()?, 0)?.0)
    }

    /// Serialize to the raw bytes of a complete `ExOleLinkAtom` record.
    ///
    /// # Errors
    ///
    /// Returns an error if the record payload exceeds the encodable size.
    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        let mut data = [0; 12];
        data[0..4].copy_from_slice(&self.slide_id.unwrap_or(0).to_le_bytes());
        data[4..8].copy_from_slice(&(self.update_mode as u32).to_le_bytes());
        data[8..12].copy_from_slice(&self.unused);
        record_bytes(0, 0, RecordType::ExternalOleLinkAtom, &data)
    }
}
