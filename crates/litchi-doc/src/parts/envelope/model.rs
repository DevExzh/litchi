//! Typed, inert Word envelope metadata.

use crate::package::Result;

/// The CLSID selecting the [`Message`] payload in `[MS-OSHARED]` 2.3.8.
///
/// The bytes are in the serialized Windows GUID order used by `[MS-DOC]`.
pub const MSO_ENVELOPE_CLSID: [u8; 16] = [
    0x1A, 0xF0, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46,
];

/// One bounded `MsoEnvelopeCLSID` structure (`[MS-DOC]` 2.5.7 and
/// `[MS-OSHARED]` 2.3.8.1).
///
/// A known CLSID is decoded into [`Payload::Message`]. A producer-defined
/// CLSID remains an opaque payload, so reading an envelope never activates a
/// mail client, follows recipients, or interprets an out-of-scope format.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Envelope {
    class_id: [u8; 16],
    payload: Payload,
}

impl Envelope {
    pub(super) fn from_parts(class_id: [u8; 16], payload: Payload) -> Self {
        Self { class_id, payload }
    }

    /// Construct a typed, validated known envelope.
    pub fn from_message(message: Message) -> Result<Self> {
        let value = Self {
            class_id: MSO_ENVELOPE_CLSID,
            payload: Payload::Message(Box::new(message)),
        };
        super::validation::validate(&value)?;
        Ok(value)
    }

    /// Construct a bounded producer-defined envelope payload.
    ///
    /// The known Office envelope CLSID must use [`Self::from_message`] so the
    /// class identifier and typed payload cannot disagree.
    pub fn opaque(class_id: [u8; 16], payload: Vec<u8>) -> Result<Self> {
        let value = Self {
            class_id,
            payload: Payload::Opaque(payload.into_boxed_slice()),
        };
        super::validation::validate(&value)?;
        Ok(value)
    }

    /// The exact serialized GUID bytes from `MsoEnvelopeCLSID.CLSID`.
    pub const fn class_id(&self) -> &[u8; 16] {
        &self.class_id
    }

    /// The class-selected envelope payload.
    pub fn payload(&self) -> &Payload {
        &self.payload
    }

    /// The typed Office envelope body, when the documented CLSID is present.
    pub fn message(&self) -> Option<&Message> {
        match &self.payload {
            Payload::Message(message) => Some(message),
            Payload::Opaque(_) => None,
        }
    }

    /// Parse one complete `MsoEnvelopeCLSID` payload from a table-stream slice.
    pub fn parse_bytes(data: &[u8]) -> Result<Self> {
        super::codec::parse(data)
    }

    pub(crate) fn parse_fib(
        fib: &crate::parts::fib::FileInformationBlock,
        table_stream: &[u8],
    ) -> Result<Option<Self>> {
        super::codec::parse_fib(fib, table_stream)
    }

    /// Serialize one complete `MsoEnvelopeCLSID` payload.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        super::codec::write(self)
    }
}

/// Payload selected by [`Envelope::class_id`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Payload {
    /// The documented `[MS-OSHARED]` `MsoEnvelope` body.
    Message(Box<Message>),
    /// A bounded payload whose CLSID is outside this implementation's scope.
    Opaque(Box<[u8]>),
}

/// The two versioned layouts defined by `[MS-OSHARED]` 2.3.8.2.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u32)]
pub enum Version {
    /// ANSI versioned strings (`Ver == 6`).
    Office6 = 6,
    /// Unicode versioned strings (`Ver == 8`).
    Office8 = 8,
}

/// A versioned envelope string. ANSI bytes are retained without guessing a
/// code page; Unicode values are validated as scalar UTF-16 when encoded.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Text {
    Ansi(Box<[u8]>),
    Unicode(Box<[u16]>),
}

impl Text {
    pub fn ansi(value: Vec<u8>) -> Self {
        Self::Ansi(value.into_boxed_slice())
    }

    pub fn unicode(value: Vec<u16>) -> Self {
        Self::Unicode(value.into_boxed_slice())
    }

    pub fn as_ansi(&self) -> Option<&[u8]> {
        match self {
            Self::Ansi(value) => Some(value),
            Self::Unicode(_) => None,
        }
    }

    pub fn as_unicode(&self) -> Option<&[u16]> {
        match self {
            Self::Ansi(_) => None,
            Self::Unicode(value) => Some(value),
        }
    }

    /// Render for diagnostics without treating ANSI bytes as a claimed code
    /// page. Unicode replacement is limited to display and never used for
    /// serialization.
    pub fn to_string_lossy(&self) -> String {
        match self {
            Self::Ansi(value) => value.iter().map(|byte| char::from(*byte)).collect(),
            Self::Unicode(value) => String::from_utf16_lossy(value),
        }
    }
}

/// Follow-up state stored in `FlagStatus`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u32)]
pub enum FollowUpStatus {
    None = 0,
    Flagged = 1,
    Complete = 2,
}

/// Message sensitivity stored in `Sensitivity`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u32)]
pub enum Sensitivity {
    Normal = 0,
    Personal = 1,
    Private = 2,
    Confidential = 3,
}

/// Message importance stored in `Importance`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[repr(u32)]
pub enum Importance {
    Low = 0,
    Normal = 1,
    High = 2,
}

/// The two defined `SecurityFlags` bits.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub struct SecurityFlags {
    pub signed: bool,
    pub encrypted: bool,
}

/// A complete, inert `[MS-OSHARED]` `MsoEnvelope` body.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Message {
    pub version: Version,
    pub last_sent_time: u32,
    pub flag_status: FollowUpStatus,
    pub reply_time: u32,
    pub request: Text,
    pub sent_representing_entry_id: Box<[u8]>,
    pub sent_representing_name: Text,
    pub internet_account_stamp: Text,
    pub internet_account_name: Text,
    pub expiry_time: u32,
    pub deferred_delivery_time: u32,
    pub delete_after_submit: bool,
    pub security: SecurityFlags,
    pub delivery_report: bool,
    pub read_receipt: bool,
    pub categories: Text,
    pub sensitivity: Sensitivity,
    pub importance: Importance,
    pub subject: Text,
    pub voting_options: Box<[u8]>,
    pub reply_recipients: RecipientCollection,
    /// Present exactly when `version` is [`Version::Office8`].
    pub contact_link_recipients: Option<RecipientCollection>,
    pub recipients: RecipientCollection,
    pub attachments: Vec<Attachment>,
    /// Present exactly when `version` is [`Version::Office8`].
    pub intro_text: Option<Box<[u16]>>,
    /// Bytes following the documented body, retained for forward-compatible
    /// producers and emitted verbatim after a typed edit.
    pub tail: Box<[u8]>,
}

impl Default for Message {
    fn default() -> Self {
        let empty = || Text::Unicode(Box::default());
        Self {
            version: Version::Office8,
            last_sent_time: super::validation::MAX_MINUTE_TIME,
            flag_status: FollowUpStatus::None,
            reply_time: super::validation::MAX_MINUTE_TIME,
            request: empty(),
            sent_representing_entry_id: Box::default(),
            sent_representing_name: empty(),
            internet_account_stamp: empty(),
            internet_account_name: empty(),
            expiry_time: super::validation::MAX_MINUTE_TIME,
            deferred_delivery_time: super::validation::MAX_MINUTE_TIME,
            delete_after_submit: false,
            security: SecurityFlags::default(),
            delivery_report: false,
            read_receipt: false,
            categories: empty(),
            sensitivity: Sensitivity::Normal,
            importance: Importance::Normal,
            subject: empty(),
            voting_options: Box::default(),
            reply_recipients: RecipientCollection::default(),
            contact_link_recipients: Some(RecipientCollection::default()),
            recipients: RecipientCollection::default(),
            attachments: Vec::new(),
            intro_text: Some(Box::default()),
            tail: Box::default(),
        }
    }
}

/// Recipients represented by one tagged collection.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct RecipientCollection {
    pub recipients: Vec<RecipientProperties>,
}

/// One recipient's property bag.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct RecipientProperties {
    pub properties: Vec<RecipientProperty>,
}

/// One MAPI property in a recipient property bag.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RecipientProperty {
    pub property_id: u16,
    pub value: PropertyValue,
}

/// Property payloads defined by `[MS-OSHARED]` 2.3.8.7–2.3.8.16.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PropertyValue {
    Long(u32),
    Null(u32),
    Boolean(bool),
    SystemTime { high: u32, low: u32 },
    Error(u32),
    String8(Box<[u8]>),
    Unicode(Box<[u16]>),
    Binary(Box<[u8]>),
    MultiString8(Vec<Box<[u8]>>),
    MultiBinary(Vec<Box<[u8]>>),
}

/// One inert envelope attachment.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Attachment {
    pub method: u32,
    pub name: Box<[u16]>,
    pub data: Box<[u8]>,
}
