use litchi_core::error::{Error, Result};

/// EMF record type used to carry EMF+ records.
pub const EMR_COMMENT: u32 = 0x0000_0046;

/// `"EMF+"`, interpreted as a little-endian integer.
pub const EMFPLUS_COMMENT_IDENTIFIER: u32 = 0x2B46_4D45;

/// Size of the invariant EMF+ record header.
pub const EMFPLUS_RECORD_HEADER_SIZE: usize = 12;

/// Maximum table size permitted by [MS-EMFPLUS].
pub const MAX_EMFPLUS_OBJECT_SLOTS: usize = 64;

/// Resource ceilings used while framing untrusted EMF+ data.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ParserLimits {
    /// Maximum number of EMF+ bytes accepted in one logical stream.
    pub max_bytes: usize,
    /// Maximum number of EMF+ records accepted in one logical stream.
    pub max_records: usize,
    /// Number of usable slots at the start of the 64-entry EMF+ object table.
    pub max_object_slots: usize,
}

impl ParserLimits {
    /// Validate that the configured ceilings are meaningful and specification-safe.
    pub fn validate(self) -> Result<Self> {
        if self.max_bytes < EMFPLUS_RECORD_HEADER_SIZE {
            return Err(parse_error("max_bytes must be at least 12"));
        }
        if self.max_records == 0 {
            return Err(parse_error("max_records must be greater than zero"));
        }
        if !(1..=MAX_EMFPLUS_OBJECT_SLOTS).contains(&self.max_object_slots) {
            return Err(parse_error(
                "max_object_slots must be in the specification range 1..=64",
            ));
        }
        Ok(self)
    }
}

impl Default for ParserLimits {
    fn default() -> Self {
        Self {
            max_bytes: 64 * 1024 * 1024,
            max_records: 1_000_000,
            max_object_slots: MAX_EMFPLUS_OBJECT_SLOTS,
        }
    }
}

/// Every record identifier defined by the `RecordType` enumeration in
/// [MS-EMFPLUS] section 2.1.1.1.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
#[repr(u16)]
pub enum RecordType {
    Header = 0x4001,
    EndOfFile = 0x4002,
    Comment = 0x4003,
    GetDc = 0x4004,
    MultiFormatStart = 0x4005,
    MultiFormatSection = 0x4006,
    MultiFormatEnd = 0x4007,
    Object = 0x4008,
    Clear = 0x4009,
    FillRects = 0x400A,
    DrawRects = 0x400B,
    FillPolygon = 0x400C,
    DrawLines = 0x400D,
    FillEllipse = 0x400E,
    DrawEllipse = 0x400F,
    FillPie = 0x4010,
    DrawPie = 0x4011,
    DrawArc = 0x4012,
    FillRegion = 0x4013,
    FillPath = 0x4014,
    DrawPath = 0x4015,
    FillClosedCurve = 0x4016,
    DrawClosedCurve = 0x4017,
    DrawCurve = 0x4018,
    DrawBeziers = 0x4019,
    DrawImage = 0x401A,
    DrawImagePoints = 0x401B,
    DrawString = 0x401C,
    SetRenderingOrigin = 0x401D,
    SetAntiAliasMode = 0x401E,
    SetTextRenderingHint = 0x401F,
    SetTextContrast = 0x4020,
    SetInterpolationMode = 0x4021,
    SetPixelOffsetMode = 0x4022,
    SetCompositingMode = 0x4023,
    SetCompositingQuality = 0x4024,
    Save = 0x4025,
    Restore = 0x4026,
    BeginContainer = 0x4027,
    BeginContainerNoParams = 0x4028,
    EndContainer = 0x4029,
    SetWorldTransform = 0x402A,
    ResetWorldTransform = 0x402B,
    MultiplyWorldTransform = 0x402C,
    TranslateWorldTransform = 0x402D,
    ScaleWorldTransform = 0x402E,
    RotateWorldTransform = 0x402F,
    SetPageTransform = 0x4030,
    ResetClip = 0x4031,
    SetClipRect = 0x4032,
    SetClipPath = 0x4033,
    SetClipRegion = 0x4034,
    OffsetClip = 0x4035,
    DrawDriverString = 0x4036,
    StrokeFillPath = 0x4037,
    SerializableObject = 0x4038,
    SetTsGraphics = 0x4039,
    SetTsClip = 0x403A,
}

impl RecordType {
    /// All recognized record identifiers, in their wire order.
    pub const ALL: [Self; 58] = [
        Self::Header,
        Self::EndOfFile,
        Self::Comment,
        Self::GetDc,
        Self::MultiFormatStart,
        Self::MultiFormatSection,
        Self::MultiFormatEnd,
        Self::Object,
        Self::Clear,
        Self::FillRects,
        Self::DrawRects,
        Self::FillPolygon,
        Self::DrawLines,
        Self::FillEllipse,
        Self::DrawEllipse,
        Self::FillPie,
        Self::DrawPie,
        Self::DrawArc,
        Self::FillRegion,
        Self::FillPath,
        Self::DrawPath,
        Self::FillClosedCurve,
        Self::DrawClosedCurve,
        Self::DrawCurve,
        Self::DrawBeziers,
        Self::DrawImage,
        Self::DrawImagePoints,
        Self::DrawString,
        Self::SetRenderingOrigin,
        Self::SetAntiAliasMode,
        Self::SetTextRenderingHint,
        Self::SetTextContrast,
        Self::SetInterpolationMode,
        Self::SetPixelOffsetMode,
        Self::SetCompositingMode,
        Self::SetCompositingQuality,
        Self::Save,
        Self::Restore,
        Self::BeginContainer,
        Self::BeginContainerNoParams,
        Self::EndContainer,
        Self::SetWorldTransform,
        Self::ResetWorldTransform,
        Self::MultiplyWorldTransform,
        Self::TranslateWorldTransform,
        Self::ScaleWorldTransform,
        Self::RotateWorldTransform,
        Self::SetPageTransform,
        Self::ResetClip,
        Self::SetClipRect,
        Self::SetClipPath,
        Self::SetClipRegion,
        Self::OffsetClip,
        Self::DrawDriverString,
        Self::StrokeFillPath,
        Self::SerializableObject,
        Self::SetTsGraphics,
        Self::SetTsClip,
    ];

    /// Return the exact 16-bit value serialized on the wire.
    #[must_use]
    pub const fn raw(self) -> u16 {
        self as u16
    }

    /// The three multi-format identifiers are reserved and MUST NOT occur.
    #[must_use]
    pub const fn is_reserved(self) -> bool {
        matches!(
            self,
            Self::MultiFormatStart | Self::MultiFormatSection | Self::MultiFormatEnd
        )
    }

    /// Convert a wire value without accepting vendor-private values.
    #[must_use]
    pub const fn from_raw(value: u16) -> Option<Self> {
        if value < Self::Header.raw() || value > Self::SetTsClip.raw() {
            return None;
        }
        Some(Self::ALL[(value - Self::Header.raw()) as usize])
    }
}

impl TryFrom<u16> for RecordType {
    type Error = Error;

    fn try_from(value: u16) -> Result<Self> {
        Self::from_raw(value)
            .ok_or_else(|| parse_error(format!("unknown EMF+ record type 0x{value:04X}")))
    }
}

/// Typed view of the 16-bit, record-specific flags word.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
#[repr(transparent)]
pub struct RecordFlags(u16);

impl RecordFlags {
    pub const CONTINUED_OBJECT: u16 = 0x8000;
    pub const OBJECT_ID_MASK: u16 = 0x00FF;
    pub const OBJECT_TYPE_MASK: u16 = 0x7F00;

    #[must_use]
    pub const fn new(raw: u16) -> Self {
        Self(raw)
    }

    #[must_use]
    pub const fn raw(self) -> u16 {
        self.0
    }

    #[must_use]
    pub const fn contains(self, mask: u16) -> bool {
        self.0 & mask == mask
    }

    #[must_use]
    pub const fn low_byte(self) -> u8 {
        (self.0 & Self::OBJECT_ID_MASK) as u8
    }

    #[must_use]
    pub const fn high_byte(self) -> u8 {
        (self.0 >> 8) as u8
    }

    /// Decode the low-byte object table index used by many EMF+ records.
    pub fn object_id(self, limits: ParserLimits) -> Result<ObjectId> {
        ObjectId::new(self.low_byte(), limits.max_object_slots)
    }
}

/// Checked index into an EMF+ object table.
#[derive(Clone, Copy, Debug, Eq, Hash, Ord, PartialEq, PartialOrd)]
#[repr(transparent)]
pub struct ObjectId(u8);

impl ObjectId {
    pub fn new(raw: u8, slot_limit: usize) -> Result<Self> {
        if !(1..=MAX_EMFPLUS_OBJECT_SLOTS).contains(&slot_limit) {
            return Err(parse_error("object table slot limit must be in 1..=64"));
        }
        if usize::from(raw) >= slot_limit {
            return Err(parse_error(format!(
                "EMF+ object ID {raw} exceeds configured slot limit {slot_limit}"
            )));
        }
        Ok(Self(raw))
    }

    #[must_use]
    pub const fn get(self) -> u8 {
        self.0
    }
}

/// Graphics object kinds encoded in an `EmfPlusObject` flags word.
#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
#[repr(u8)]
pub enum ObjectType {
    Invalid = 0,
    Brush = 1,
    Pen = 2,
    Path = 3,
    Region = 4,
    Image = 5,
    Font = 6,
    StringFormat = 7,
    ImageAttributes = 8,
    CustomLineCap = 9,
}

impl ObjectType {
    #[must_use]
    pub const fn from_raw(value: u8) -> Option<Self> {
        match value {
            0 => Some(Self::Invalid),
            1 => Some(Self::Brush),
            2 => Some(Self::Pen),
            3 => Some(Self::Path),
            4 => Some(Self::Region),
            5 => Some(Self::Image),
            6 => Some(Self::Font),
            7 => Some(Self::StringFormat),
            8 => Some(Self::ImageAttributes),
            9 => Some(Self::CustomLineCap),
            _ => None,
        }
    }
}

/// Decoded flags specific to an `EmfPlusObject` record.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct ObjectRecordFlags {
    pub continued: bool,
    pub object_type: ObjectType,
    pub object_id: ObjectId,
}

impl ObjectRecordFlags {
    pub fn parse(flags: RecordFlags, limits: ParserLimits) -> Result<Self> {
        limits.validate()?;
        let raw_type = ((flags.raw() & RecordFlags::OBJECT_TYPE_MASK) >> 8) as u8;
        let object_type = ObjectType::from_raw(raw_type)
            .ok_or_else(|| parse_error(format!("unknown EMF+ object type {raw_type}")))?;
        if object_type == ObjectType::Invalid {
            return Err(parse_error("EmfPlusObject cannot use ObjectTypeInvalid"));
        }
        Ok(Self {
            continued: flags.contains(RecordFlags::CONTINUED_OBJECT),
            object_type,
            object_id: flags.object_id(limits)?,
        })
    }
}

/// Parsed invariant header of an EMF+ record.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct RecordHeader {
    pub record_type: RecordType,
    pub flags: RecordFlags,
    pub size: u32,
    pub data_size: u32,
}

pub(crate) fn parse_error(message: impl Into<String>) -> Error {
    Error::ParseError(message.into())
}
