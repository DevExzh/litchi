//! Typed, lossless values for a worksheet's BIFF8 comments.

use crate::error::{Error, Result};

/// BIFF8 maximum payload for a single record.
pub(crate) const MAX_RECORD_BYTES: usize = 8_224;
/// The largest number of comment objects addressable by a BIFF8 ObjId.
pub(crate) const MAX_COMMENTS: usize = 65_535;
/// A defensive ceiling for records retained by one worksheet comment collector.
pub(crate) const MAX_RETAINED_RECORDS: usize = 262_140;
/// A defensive ceiling for raw comment/object bytes retained by one collector.
pub(crate) const MAX_RETAINED_BYTES: usize = 64 * 1024 * 1024;

pub const RECORD_TYPE: u16 = 0x001C;
pub const OBJ_TYPE: u16 = 0x005D;
pub const TXO_TYPE: u16 = 0x01B6;
pub const CONTINUE_TYPE: u16 = 0x003C;
pub const MSODRAWING_TYPE: u16 = 0x00EC;
pub(crate) const COMMENT_OBJECT_TYPE: u16 = 0x0019;

/// Whether a NOTE is shown without user interaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Visibility {
    Hidden,
    Visible,
}

/// BIFF8 TxO horizontal alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HorizontalAlignment {
    Left,
    Centered,
    Right,
    Justified,
    Distributed,
    /// A value not assigned by the checked MS-XLS enumeration.
    Unknown(u8),
}

/// BIFF8 TxO vertical alignment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VerticalAlignment {
    Top,
    Centered,
    Bottom,
    Justified,
    Distributed,
    /// A value not assigned by the checked MS-XLS enumeration.
    Unknown(u8),
}

/// BIFF8 TxO text orientation.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextOrientation {
    None,
    Stacked,
    CounterClockwise,
    Clockwise,
    /// A value not assigned by the checked MS-XLS enumeration.
    Unknown(u16),
}

/// The BIFF record kinds making up one linked comment.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RecordKind {
    Note,
    Object,
    Drawing,
    TextObject,
    Continue,
}

impl RecordKind {
    pub const fn record_type(self) -> u16 {
        match self {
            Self::Note => RECORD_TYPE,
            Self::Object => OBJ_TYPE,
            Self::Drawing => MSODRAWING_TYPE,
            Self::TextObject => TXO_TYPE,
            Self::Continue => CONTINUE_TYPE,
        }
    }
}

/// An exact BIFF payload retained in the source order of a comment.
///
/// Keeping the complete payload makes reserved and future fields replayable
/// without interpreting or normalizing them.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct CommentRecord {
    kind: RecordKind,
    payload: Box<[u8]>,
}

impl CommentRecord {
    pub fn kind(&self) -> RecordKind {
        self.kind
    }

    pub fn record_type(&self) -> u16 {
        self.kind.record_type()
    }

    /// The payload excluding the four-byte BIFF record header.
    pub fn payload(&self) -> &[u8] {
        &self.payload
    }

    pub(crate) fn new(kind: RecordKind, data: &[u8]) -> Result<Self> {
        if data.len() > MAX_RECORD_BYTES {
            return Err(Error::InvalidData(
                "comment record payload exceeds the BIFF8 record bound".to_string(),
            ));
        }
        let mut payload = Vec::new();
        payload
            .try_reserve_exact(data.len())
            .map_err(|_| Error::Allocation("retaining comment record payload"))?;
        payload.extend_from_slice(data);
        Ok(Self {
            kind,
            payload: payload.into_boxed_slice(),
        })
    }
}

/// The common FtCmo object properties retained for a comment.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ObjectProperties {
    pub(crate) object_type: u16,
    pub(crate) object_id: u16,
    pub(crate) flags: u16,
    pub(crate) reserved_flags: u16,
    pub(crate) reserved_header: [u8; 4],
    pub(crate) unused: [u8; 12],
}

impl ObjectProperties {
    pub fn object_type(&self) -> u16 {
        self.object_type
    }

    pub fn object_id(&self) -> u16 {
        self.object_id
    }

    pub fn flags(&self) -> u16 {
        self.flags
    }

    /// The FtCmo flag bits not assigned by MS-XLS.
    pub fn reserved_flags(&self) -> u16 {
        self.reserved_flags
    }

    /// FtCmo's four reserved header bytes (ft and cb).
    pub fn reserved_header(&self) -> &[u8; 4] {
        &self.reserved_header
    }

    /// FtCmo's three four-byte undefined fields, in source order.
    pub fn unused_bytes(&self) -> &[u8; 12] {
        &self.unused
    }
}

/// An OBJ subrecord retained without interpreting its payload.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ObjectSubrecord {
    record_type: u16,
    payload: Box<[u8]>,
    known: bool,
}

impl ObjectSubrecord {
    pub fn record_type(&self) -> u16 {
        self.record_type
    }

    /// The subrecord body excluding its four-byte Ft header.
    pub fn payload(&self) -> &[u8] {
        &self.payload
    }

    /// Whether this is a subrecord understood by the comment codec.
    pub fn is_known(&self) -> bool {
        self.known
    }

    pub(crate) fn new(record_type: u16, payload: &[u8], known: bool) -> Result<Self> {
        if payload.len() > MAX_RECORD_BYTES {
            return Err(Error::InvalidData(
                "OBJ subrecord payload exceeds the BIFF8 record bound".to_string(),
            ));
        }
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(payload.len())
            .map_err(|_| Error::Allocation("retaining OBJ subrecord payload"))?;
        owned.extend_from_slice(payload);
        Ok(Self {
            record_type,
            payload: owned.into_boxed_slice(),
            known,
        })
    }
}

/// The reserved/padding bytes after an OBJ's FtEnd subrecord.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ObjectPadding(Box<[u8]>);

impl ObjectPadding {
    pub fn as_bytes(&self) -> &[u8] {
        &self.0
    }

    pub(crate) fn new(data: &[u8]) -> Result<Self> {
        if data.len() > MAX_RECORD_BYTES {
            return Err(Error::InvalidData(
                "OBJ padding exceeds the BIFF8 record bound".to_string(),
            ));
        }
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(data.len())
            .map_err(|_| Error::Allocation("retaining OBJ padding"))?;
        owned.extend_from_slice(data);
        Ok(Self(owned.into_boxed_slice()))
    }
}

/// Stable identity supplied by an OBJ's FtNts structure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ObjectIdentity {
    pub(crate) object_id: u16,
    pub(crate) guid: [u8; 16],
    pub(crate) shared: bool,
    pub(crate) shared_value: u16,
    pub(crate) unused: [u8; 4],
}

impl ObjectIdentity {
    pub fn object_id(&self) -> u16 {
        self.object_id
    }

    pub fn guid(&self) -> &[u8; 16] {
        &self.guid
    }

    pub fn shared(&self) -> bool {
        self.shared
    }

    /// The source FtNts Boolean value, including future nonzero values.
    pub fn shared_value(&self) -> u16 {
        self.shared_value
    }

    /// FtNts's four undefined bytes.
    pub fn unused_bytes(&self) -> &[u8; 4] {
        &self.unused
    }
}

/// One formatting run from the comment's TxORuns structure.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextRun {
    pub(crate) character_index: u16,
    pub(crate) font_index: u16,
}

impl TextRun {
    pub fn character_index(&self) -> u16 {
        self.character_index
    }

    pub fn font_index(&self) -> u16 {
        self.font_index
    }
}

/// The ignored/reserved and inert fields retained from a comment's TxO.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct TextProperties {
    pub(crate) horizontal_alignment: HorizontalAlignment,
    pub(crate) vertical_alignment: VerticalAlignment,
    pub(crate) orientation: TextOrientation,
    pub(crate) locked: bool,
    pub(crate) justify_last_line: bool,
    pub(crate) secret_edit: bool,
    pub(crate) font_when_empty: u16,
    pub(crate) reserved_options: u16,
    pub(crate) reserved_fields: [u8; 6],
    pub(crate) formula_bytes: Box<[u8]>,
}

impl TextProperties {
    pub fn horizontal_alignment(&self) -> HorizontalAlignment {
        self.horizontal_alignment
    }

    pub fn vertical_alignment(&self) -> VerticalAlignment {
        self.vertical_alignment
    }

    pub fn orientation(&self) -> TextOrientation {
        self.orientation
    }

    pub fn locked(&self) -> bool {
        self.locked
    }

    pub fn justify_last_line(&self) -> bool {
        self.justify_last_line
    }

    pub fn secret_edit(&self) -> bool {
        self.secret_edit
    }

    pub fn font_when_empty(&self) -> u16 {
        self.font_when_empty
    }

    /// The TxO option bits marked reserved by MS-XLS.
    pub fn reserved_options(&self) -> u16 {
        self.reserved_options
    }

    /// TxO's reserved4/reserved5 bytes, in source order.
    pub fn reserved_fields(&self) -> &[u8; 6] {
        &self.reserved_fields
    }

    /// Raw, unevaluated ObjFmla payload bytes, excluding its length field.
    pub fn formula_bytes(&self) -> &[u8] {
        &self.formula_bytes
    }
}

/// The ignored/reserved bytes in a BIFF8 NOTE record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct NoteMetadata {
    pub(crate) reserved_flags: u16,
    pub(crate) reserved_string_flags: u8,
    pub(crate) unused: u8,
}

impl NoteMetadata {
    pub fn reserved_flags(&self) -> u16 {
        self.reserved_flags
    }

    pub fn reserved_string_flags(&self) -> u8 {
        self.reserved_string_flags
    }

    pub fn unused_byte(&self) -> u8 {
        self.unused
    }
}

/// A fully linked, immutable BIFF8 cell comment.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Comment {
    pub(crate) row: u16,
    pub(crate) column: u8,
    pub(crate) visibility: Visibility,
    pub(crate) row_hidden: bool,
    pub(crate) column_hidden: bool,
    pub(crate) identity: ObjectIdentity,
    pub(crate) object_properties: ObjectProperties,
    pub(crate) object_subrecords: Vec<ObjectSubrecord>,
    pub(crate) object_padding: ObjectPadding,
    pub(crate) note_metadata: NoteMetadata,
    pub(crate) author: String,
    pub(crate) text: String,
    pub(crate) text_properties: TextProperties,
    pub(crate) text_runs: Vec<TextRun>,
    pub(crate) records: Vec<CommentRecord>,
}

impl Comment {
    pub fn row(&self) -> u16 {
        self.row
    }

    pub fn column(&self) -> u8 {
        self.column
    }

    pub fn visibility(&self) -> Visibility {
        self.visibility
    }

    pub fn row_hidden(&self) -> bool {
        self.row_hidden
    }

    pub fn column_hidden(&self) -> bool {
        self.column_hidden
    }

    pub fn identity(&self) -> &ObjectIdentity {
        &self.identity
    }

    pub fn object_properties(&self) -> &ObjectProperties {
        &self.object_properties
    }

    /// All OBJ subrecords after FtCmo, including known and unknown records.
    pub fn object_subrecords(&self) -> &[ObjectSubrecord] {
        &self.object_subrecords
    }

    pub fn object_padding(&self) -> &[u8] {
        self.object_padding.as_bytes()
    }

    pub fn note_metadata(&self) -> NoteMetadata {
        self.note_metadata
    }

    pub fn author(&self) -> &str {
        &self.author
    }

    pub fn text(&self) -> &str {
        &self.text
    }

    pub fn text_properties(&self) -> &TextProperties {
        &self.text_properties
    }

    pub fn text_runs(&self) -> &[TextRun] {
        &self.text_runs
    }

    /// The NOTE/OBJ/MSODRAWING/TXO/CONTINUE records in source order.
    pub fn records(&self) -> &[CommentRecord] {
        &self.records
    }
}

pub(crate) fn boxed_bytes(data: &[u8], context: &'static str) -> Result<Box<[u8]>> {
    if data.len() > MAX_RECORD_BYTES {
        return Err(Error::InvalidData(
            "retained comment bytes exceed the BIFF8 record bound".to_string(),
        ));
    }
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(data.len())
        .map_err(|_| Error::Allocation(context))?;
    owned.extend_from_slice(data);
    Ok(owned.into_boxed_slice())
}
