/// CLSID selecting the MS-OSHARED `MsoEnvelope` payload.
pub const MSO_ENVELOPE_CLSID: [u8; 16] = [
    0x1a, 0xf0, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46,
];

/// A complete `PowerPoint` 9 envelope atom.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct EnvelopeData {
    pub clsid: [u8; 16],
    pub payload: EnvelopePayload,
}

/// Payload selected by the envelope CLSID.
#[derive(Debug, Clone, PartialEq, Eq)]
#[allow(
    clippy::large_enum_variant,
    reason = "public payload enum; boxing would break the API"
)]
pub enum EnvelopePayload {
    Mso(MsoEnvelope),
    /// A payload whose CLSID-defined syntax is outside MS-OSHARED.
    Opaque(Vec<u8>),
}

/// The two layouts defined for version-dependent envelope strings.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum MsoEnvelopeVersion {
    Office6 = 6,
    Office8 = 8,
}

/// A version-dependent MS-OSHARED string, retained without a lossy conversion.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum MsoEnvelopeText {
    Ansi(Vec<u8>),
    Unicode(Vec<u16>),
}

impl MsoEnvelopeText {
    /// Decode for display. Invalid ANSI bytes are mapped one-to-one as Latin-1.
    #[must_use]
    pub fn to_string_lossy(&self) -> String {
        match self {
            Self::Ansi(bytes) => bytes.iter().map(|byte| char::from(*byte)).collect(),
            Self::Unicode(units) => String::from_utf16_lossy(units),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum MsoFollowUpStatus {
    None = 0,
    Complete = 1,
    Flagged = 2,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum MsoSensitivity {
    Normal = 0,
    Personal = 1,
    Private = 2,
    Confidential = 3,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum MsoImportance {
    Low = 0,
    Normal = 1,
    High = 2,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct MsoSecurityFlags {
    pub signed: bool,
    pub encrypted: bool,
}

/// Fully decoded MS-OSHARED mail-envelope state.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MsoEnvelope {
    pub version: MsoEnvelopeVersion,
    pub last_sent_time: u32,
    pub flag_status: MsoFollowUpStatus,
    pub reply_time: u32,
    pub request: MsoEnvelopeText,
    pub sent_representing_entry_id: Vec<u8>,
    pub sent_representing_name: MsoEnvelopeText,
    pub internet_account_stamp: MsoEnvelopeText,
    pub internet_account_name: MsoEnvelopeText,
    pub expiry_time: u32,
    pub deferred_delivery_time: u32,
    pub delete_after_submit: bool,
    pub security: MsoSecurityFlags,
    pub delivery_report: bool,
    pub read_receipt: bool,
    pub categories: MsoEnvelopeText,
    pub sensitivity: MsoSensitivity,
    pub importance: MsoImportance,
    pub subject: MsoEnvelopeText,
    pub voting_options: Vec<u8>,
    pub reply_recipients: MsoRecipientCollection,
    /// Present exactly for version 8.
    pub contact_link_recipients: Option<MsoRecipientCollection>,
    pub recipients: MsoRecipientCollection,
    pub attachments: Vec<MsoAttachment>,
    /// Present exactly for version 8.
    pub intro_text: Option<Vec<u16>>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct MsoRecipientCollection {
    pub recipients: Vec<MsoRecipientProperties>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct MsoRecipientProperties {
    pub properties: Vec<MsoRecipientProperty>,
}

/// One tagged MAPI property from an envelope recipient.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MsoRecipientProperty {
    pub property_id: u16,
    pub value: MsoPropertyValue,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum MsoPropertyValue {
    Long(u32),
    Null(u32),
    Boolean(bool),
    SystemTime { high: u32, low: u32 },
    Error(u32),
    String8(Vec<u8>),
    Unicode(Vec<u16>),
    Binary(Vec<u8>),
    MultiString8(Vec<Vec<u8>>),
    MultiBinary(Vec<Vec<u8>>),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MsoAttachment {
    pub method: u32,
    pub name: Vec<u16>,
    pub data: Vec<u8>,
}
