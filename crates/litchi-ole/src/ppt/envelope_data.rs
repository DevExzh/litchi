//! Strict, inert parsing of PowerPoint 9 `EnvelopeData9Atom` records.
//!
//! The known Office mail-envelope CLSID is decoded according to MS-OSHARED.
//! Other CLSIDs remain bounded opaque payloads. Nothing in this module sends
//! mail, opens attachments, invokes a mail client, or evaluates embedded data.

use crate::consts::PptRecordType;
use crate::ppt::package::{PptError, Result};
use crate::ppt::records::PptRecord;

const ENVELOPE_DATA_RECORD_TYPE: u16 = 0x1785;
const RECIPIENT_COLLECTION_TAG: u32 = 0xdcca_0123;
const MAX_ENVELOPE_BYTES: usize = 64 * 1024 * 1024;
const MAX_COLLECTION_ITEMS: usize = 65_536;
const MAX_ATTACHMENTS: usize = 4_096;
const MAX_MINUTE_TIME: u32 = 0x5ae9_80e0;

/// CLSID selecting the MS-OSHARED `MsoEnvelope` payload.
pub const MSO_ENVELOPE_CLSID: [u8; 16] = [
    0x1a, 0xf0, 0x06, 0x00, 0x00, 0x00, 0x00, 0x00, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46,
];

/// A complete PowerPoint 9 envelope atom.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointEnvelopeData {
    pub clsid: [u8; 16],
    pub payload: PowerPointEnvelopePayload,
}

/// Payload selected by the envelope CLSID.
#[derive(Debug, Clone, PartialEq, Eq)]
#[allow(clippy::large_enum_variant)] // public payload enum; boxing would break the API
pub enum PowerPointEnvelopePayload {
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

impl PowerPointEnvelopeData {
    /// Parse one exact `EnvelopeData9Atom`.
    pub fn parse(record: &PptRecord) -> Result<Self> {
        if record.record_type_raw != ENVELOPE_DATA_RECORD_TYPE
            || record.version != 0
            || record.instance != 0
            || record.data.len() < 16
            || record.data.len() > MAX_ENVELOPE_BYTES
        {
            return corrupted("EnvelopeData9Atom has an invalid header or size");
        }
        let mut clsid = [0u8; 16];
        clsid.copy_from_slice(&record.data[..16]);
        let body = &record.data[16..];
        let payload = if clsid == MSO_ENVELOPE_CLSID {
            PowerPointEnvelopePayload::Mso(MsoEnvelope::parse(body)?)
        } else {
            PowerPointEnvelopePayload::Opaque(body.to_vec())
        };
        Ok(Self { clsid, payload })
    }

    pub(crate) fn parse_document(document: &PptRecord) -> Result<Option<Self>> {
        let records = document.versioned_binary_tag_records(9)?;
        let mut matches = records
            .iter()
            .filter(|record| record.record_type_raw == ENVELOPE_DATA_RECORD_TYPE);
        let Some(record) = matches.next() else {
            return Ok(None);
        };
        if matches.next().is_some() {
            return corrupted("PPT9 document tag contains multiple EnvelopeData9Atom records");
        }
        Self::parse(record).map(Some)
    }

    /// Encode a canonical inert atom.
    pub fn to_record(&self) -> Result<PptRecord> {
        let mut data = Vec::from(self.clsid);
        match (&self.clsid, &self.payload) {
            (clsid, PowerPointEnvelopePayload::Mso(value)) if *clsid == MSO_ENVELOPE_CLSID => {
                value.write(&mut data)?;
            },
            (clsid, PowerPointEnvelopePayload::Opaque(value)) if *clsid != MSO_ENVELOPE_CLSID => {
                data.extend_from_slice(value);
            },
            _ => return corrupted("envelope CLSID and payload kind disagree"),
        }
        if data.len() > MAX_ENVELOPE_BYTES {
            return corrupted("EnvelopeData9Atom exceeds the resource cap");
        }
        let data_length = u32::try_from(data.len())
            .map_err(|_| PptError::Corrupted("envelope length overflow".to_string()))?;
        Ok(PptRecord {
            record_type: PptRecordType::from(ENVELOPE_DATA_RECORD_TYPE),
            record_type_raw: ENVELOPE_DATA_RECORD_TYPE,
            version: 0,
            instance: 0,
            data_length,
            data,
            children: Vec::new(),
        })
    }
}

impl MsoEnvelope {
    fn parse(data: &[u8]) -> Result<Self> {
        let mut input = Cursor::new(data);
        let version = match input.u32()? {
            6 => MsoEnvelopeVersion::Office6,
            8 => MsoEnvelopeVersion::Office8,
            _ => return corrupted("MsoEnvelope has an undefined version"),
        };
        let last_sent_time = minute_time(input.u32()?, "last-sent time")?;
        let flag_status = match input.u32()? {
            0 => MsoFollowUpStatus::None,
            1 => MsoFollowUpStatus::Complete,
            2 => MsoFollowUpStatus::Flagged,
            _ => return corrupted("MsoEnvelope has an invalid follow-up status"),
        };
        let reply_time = minute_time(input.u32()?, "reply time")?;
        let request = input.versioned_text(version)?;
        let sent_representing_entry_id = input.u32_blob()?;
        let sent_representing_name = input.versioned_text(version)?;
        let internet_account_stamp = input.versioned_text(version)?;
        let internet_account_name = input.versioned_text(version)?;
        let expiry_time = minute_time(input.u32()?, "expiry time")?;
        let deferred_delivery_time = minute_time(input.u32()?, "deferred-delivery time")?;
        let delete_after_submit = input.boolean32()?;
        let security_bits = input.u32()?;
        if security_bits & !3 != 0 {
            return corrupted("MsoEnvelope security flags contain reserved bits");
        }
        let security = MsoSecurityFlags {
            signed: security_bits & 1 != 0,
            encrypted: security_bits & 2 != 0,
        };
        let delivery_report = input.boolean32()?;
        let read_receipt = input.boolean32()?;
        let categories = input.versioned_text(version)?;
        let sensitivity = match input.u32()? {
            0 => MsoSensitivity::Normal,
            1 => MsoSensitivity::Personal,
            2 => MsoSensitivity::Private,
            3 => MsoSensitivity::Confidential,
            _ => return corrupted("MsoEnvelope has an invalid sensitivity"),
        };
        let importance = match input.u32()? {
            0 => MsoImportance::Low,
            1 => MsoImportance::Normal,
            2 => MsoImportance::High,
            _ => return corrupted("MsoEnvelope has an invalid importance"),
        };
        let subject = input.versioned_text(version)?;
        let voting_options = input.u16_blob()?;
        let reply_recipients = input.recipient_collection()?;
        let contact_link_recipients = if version == MsoEnvelopeVersion::Office8 {
            Some(input.recipient_collection()?)
        } else {
            None
        };
        let recipients = input.recipient_collection()?;
        let attachments = input.attachments()?;
        let intro_text = if version == MsoEnvelopeVersion::Office8 {
            Some(input.intro_text()?)
        } else {
            None
        };
        input.finish("MsoEnvelope")?;
        Ok(Self {
            version,
            last_sent_time,
            flag_status,
            reply_time,
            request,
            sent_representing_entry_id,
            sent_representing_name,
            internet_account_stamp,
            internet_account_name,
            expiry_time,
            deferred_delivery_time,
            delete_after_submit,
            security,
            delivery_report,
            read_receipt,
            categories,
            sensitivity,
            importance,
            subject,
            voting_options,
            reply_recipients,
            contact_link_recipients,
            recipients,
            attachments,
            intro_text,
        })
    }

    fn write(&self, out: &mut Vec<u8>) -> Result<()> {
        minute_time(self.last_sent_time, "last-sent time")?;
        minute_time(self.reply_time, "reply time")?;
        minute_time(self.expiry_time, "expiry time")?;
        minute_time(self.deferred_delivery_time, "deferred-delivery time")?;
        put_u32(out, self.version as u32);
        put_u32(out, self.last_sent_time);
        put_u32(out, self.flag_status as u32);
        put_u32(out, self.reply_time);
        write_versioned_text(out, self.version, &self.request)?;
        write_u32_blob(out, &self.sent_representing_entry_id)?;
        write_versioned_text(out, self.version, &self.sent_representing_name)?;
        write_versioned_text(out, self.version, &self.internet_account_stamp)?;
        write_versioned_text(out, self.version, &self.internet_account_name)?;
        put_u32(out, self.expiry_time);
        put_u32(out, self.deferred_delivery_time);
        put_u32(out, u32::from(self.delete_after_submit));
        put_u32(
            out,
            u32::from(self.security.signed) | (u32::from(self.security.encrypted) << 1),
        );
        put_u32(out, u32::from(self.delivery_report));
        put_u32(out, u32::from(self.read_receipt));
        write_versioned_text(out, self.version, &self.categories)?;
        put_u32(out, self.sensitivity as u32);
        put_u32(out, self.importance as u32);
        write_versioned_text(out, self.version, &self.subject)?;
        write_u16_blob(out, &self.voting_options)?;
        write_recipient_collection(out, &self.reply_recipients)?;
        match (self.version, &self.contact_link_recipients) {
            (MsoEnvelopeVersion::Office8, Some(value)) => {
                write_recipient_collection(out, value)?;
            },
            (MsoEnvelopeVersion::Office6, None) => {},
            _ => return corrupted("contact-link recipients do not match envelope version"),
        }
        write_recipient_collection(out, &self.recipients)?;
        write_attachments(out, &self.attachments)?;
        match (self.version, &self.intro_text) {
            (MsoEnvelopeVersion::Office8, Some(value)) => write_intro_text(out, value)?,
            (MsoEnvelopeVersion::Office6, None) => {},
            _ => return corrupted("intro text does not match envelope version"),
        }
        Ok(())
    }
}

struct Cursor<'a> {
    data: &'a [u8],
    position: usize,
}

impl<'a> Cursor<'a> {
    fn new(data: &'a [u8]) -> Self {
        Self { data, position: 0 }
    }

    fn take(&mut self, length: usize) -> Result<&'a [u8]> {
        let end = self
            .position
            .checked_add(length)
            .ok_or_else(|| PptError::Corrupted("envelope offset overflow".to_string()))?;
        if end > self.data.len() {
            return corrupted("MsoEnvelope is truncated");
        }
        let value = &self.data[self.position..end];
        self.position = end;
        Ok(value)
    }

    fn u8(&mut self) -> Result<u8> {
        Ok(self.take(1)?[0])
    }

    fn u16(&mut self) -> Result<u16> {
        let value = self.take(2)?;
        Ok(u16::from_le_bytes([value[0], value[1]]))
    }

    fn u32(&mut self) -> Result<u32> {
        let value = self.take(4)?;
        Ok(u32::from_le_bytes([value[0], value[1], value[2], value[3]]))
    }

    fn boolean32(&mut self) -> Result<bool> {
        match self.u32()? {
            0 => Ok(false),
            1 => Ok(true),
            _ => corrupted("MsoEnvelope contains a non-Boolean flag"),
        }
    }

    fn u16_blob(&mut self) -> Result<Vec<u8>> {
        let length = usize::from(self.u16()?);
        Ok(self.take(length)?.to_vec())
    }

    fn u32_blob(&mut self) -> Result<Vec<u8>> {
        let length = usize::try_from(self.u32()?)
            .map_err(|_| PptError::Corrupted("envelope blob length overflow".to_string()))?;
        Ok(self.take(length)?.to_vec())
    }

    fn utf16_bytes(&mut self, byte_length: usize) -> Result<Vec<u16>> {
        if !byte_length.is_multiple_of(2) {
            return corrupted("UTF-16 envelope string has an odd byte size");
        }
        let units: Vec<u16> = self
            .take(byte_length)?
            .chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
            .collect();
        validate_utf16(&units, "UTF-16 envelope string")?;
        Ok(units)
    }

    fn versioned_text(&mut self, version: MsoEnvelopeVersion) -> Result<MsoEnvelopeText> {
        let characters = usize::from(self.u16()?);
        match version {
            MsoEnvelopeVersion::Office6 => {
                Ok(MsoEnvelopeText::Ansi(self.take(characters)?.to_vec()))
            },
            MsoEnvelopeVersion::Office8 => Ok(MsoEnvelopeText::Unicode(
                self.utf16_bytes(checked_double(characters)?)?,
            )),
        }
    }

    fn recipient_collection(&mut self) -> Result<MsoRecipientCollection> {
        if self.u32()? != RECIPIENT_COLLECTION_TAG || self.u32()? != 1 {
            return corrupted("recipient collection has an invalid tag or version");
        }
        let count = bounded_count(self.u32()?, MAX_COLLECTION_ITEMS, "recipient")?;
        let mut recipients = Vec::with_capacity(count);
        for _ in 0..count {
            let property_count = bounded_count(self.u32()?, MAX_COLLECTION_ITEMS, "property")?;
            let _ignored = self.u32()?;
            let mut properties = Vec::with_capacity(property_count);
            for _ in 0..property_count {
                properties.push(self.property()?);
            }
            recipients.push(MsoRecipientProperties { properties });
        }
        Ok(MsoRecipientCollection { recipients })
    }

    fn property(&mut self) -> Result<MsoRecipientProperty> {
        let tag = self.u32()?;
        let property_id = (tag >> 16) as u16;
        let value = match tag as u16 {
            0x0003 => MsoPropertyValue::Long(self.u32()?),
            0x0001 => MsoPropertyValue::Null(self.u32()?),
            0x000b => match self.u16()? {
                0 => MsoPropertyValue::Boolean(false),
                1 => MsoPropertyValue::Boolean(true),
                _ => return corrupted("recipient property has a non-Boolean value"),
            },
            0x0040 => MsoPropertyValue::SystemTime {
                high: self.u32()?,
                low: self.u32()?,
            },
            0x000a => MsoPropertyValue::Error(self.u32()?),
            0x001e => MsoPropertyValue::String8(self.u16_blob()?),
            0x001f => {
                let size = usize::from(self.u16()?);
                MsoPropertyValue::Unicode(self.utf16_bytes(size)?)
            },
            0x0102 => MsoPropertyValue::Binary(self.u16_blob()?),
            0x101e => {
                let count = bounded_count(self.u32()?, MAX_COLLECTION_ITEMS, "multi-string")?;
                let mut values = Vec::with_capacity(count);
                for _ in 0..count {
                    values.push(self.u16_blob()?);
                }
                MsoPropertyValue::MultiString8(values)
            },
            0x1102 => {
                let count = bounded_count(self.u32()?, MAX_COLLECTION_ITEMS, "multi-binary")?;
                let mut values = Vec::with_capacity(count);
                for _ in 0..count {
                    values.push(self.u16_blob()?);
                }
                MsoPropertyValue::MultiBinary(values)
            },
            _ => return corrupted("recipient property has an unsupported property type"),
        };
        Ok(MsoRecipientProperty { property_id, value })
    }

    fn attachments(&mut self) -> Result<Vec<MsoAttachment>> {
        let count = bounded_count(self.u32()?, MAX_ATTACHMENTS, "attachment")?;
        let mut attachments = Vec::with_capacity(count);
        for _ in 0..count {
            let method = self.u32()?;
            let name_characters = usize::from(self.u8()?);
            let name = self.utf16_bytes(checked_double(name_characters)?)?;
            let low = u64::from(self.u32()?);
            let high = u64::from(self.u32()?);
            let size = (high << 32) | low;
            let size = usize::try_from(size)
                .map_err(|_| PptError::Corrupted("attachment length overflow".to_string()))?;
            let data = self.take(size)?.to_vec();
            attachments.push(MsoAttachment { method, name, data });
        }
        Ok(attachments)
    }

    fn intro_text(&mut self) -> Result<Vec<u16>> {
        let bytes = usize::try_from(self.u32()?)
            .map_err(|_| PptError::Corrupted("intro-text length overflow".to_string()))?;
        self.utf16_bytes(bytes)
    }

    fn finish(&self, structure: &str) -> Result<()> {
        if self.position == self.data.len() {
            Ok(())
        } else {
            corrupted(&format!("{structure} has trailing bytes"))
        }
    }
}

fn write_versioned_text(
    out: &mut Vec<u8>,
    version: MsoEnvelopeVersion,
    value: &MsoEnvelopeText,
) -> Result<()> {
    match (version, value) {
        (MsoEnvelopeVersion::Office6, MsoEnvelopeText::Ansi(bytes)) => {
            put_u16_len(out, bytes.len(), "ANSI envelope string")?;
            out.extend_from_slice(bytes);
        },
        (MsoEnvelopeVersion::Office8, MsoEnvelopeText::Unicode(units)) => {
            validate_utf16(units, "Unicode envelope string")?;
            put_u16_len(out, units.len(), "Unicode envelope string")?;
            put_utf16(out, units);
        },
        _ => return corrupted("envelope string encoding does not match version"),
    }
    Ok(())
}

fn write_recipient_collection(out: &mut Vec<u8>, value: &MsoRecipientCollection) -> Result<()> {
    if value.recipients.len() > MAX_COLLECTION_ITEMS {
        return corrupted("recipient count exceeds the resource cap");
    }
    put_u32(out, RECIPIENT_COLLECTION_TAG);
    put_u32(out, 1);
    put_u32_len(out, value.recipients.len(), "recipient")?;
    for recipient in &value.recipients {
        if recipient.properties.len() > MAX_COLLECTION_ITEMS {
            return corrupted("property count exceeds the resource cap");
        }
        put_u32_len(out, recipient.properties.len(), "property")?;
        put_u32(out, 0);
        for property in &recipient.properties {
            write_property(out, property)?;
        }
    }
    Ok(())
}

fn write_property(out: &mut Vec<u8>, property: &MsoRecipientProperty) -> Result<()> {
    let property_type = match &property.value {
        MsoPropertyValue::Long(_) => 0x0003,
        MsoPropertyValue::Null(_) => 0x0001,
        MsoPropertyValue::Boolean(_) => 0x000b,
        MsoPropertyValue::SystemTime { .. } => 0x0040,
        MsoPropertyValue::Error(_) => 0x000a,
        MsoPropertyValue::String8(_) => 0x001e,
        MsoPropertyValue::Unicode(_) => 0x001f,
        MsoPropertyValue::Binary(_) => 0x0102,
        MsoPropertyValue::MultiString8(_) => 0x101e,
        MsoPropertyValue::MultiBinary(_) => 0x1102,
    };
    put_u32(out, (u32::from(property.property_id) << 16) | property_type);
    match &property.value {
        MsoPropertyValue::Long(value)
        | MsoPropertyValue::Null(value)
        | MsoPropertyValue::Error(value) => put_u32(out, *value),
        MsoPropertyValue::Boolean(value) => put_u16(out, u16::from(*value)),
        MsoPropertyValue::SystemTime { high, low } => {
            put_u32(out, *high);
            put_u32(out, *low);
        },
        MsoPropertyValue::String8(value) | MsoPropertyValue::Binary(value) => {
            write_u16_blob(out, value)?;
        },
        MsoPropertyValue::Unicode(value) => {
            validate_utf16(value, "Unicode recipient property")?;
            put_u16_len(out, checked_double(value.len())?, "Unicode property")?;
            put_utf16(out, value);
        },
        MsoPropertyValue::MultiString8(values) | MsoPropertyValue::MultiBinary(values) => {
            if values.len() > MAX_COLLECTION_ITEMS {
                return corrupted("multi-value property count exceeds the resource cap");
            }
            put_u32_len(out, values.len(), "multi-value property")?;
            for value in values {
                write_u16_blob(out, value)?;
            }
        },
    }
    Ok(())
}

fn write_attachments(out: &mut Vec<u8>, attachments: &[MsoAttachment]) -> Result<()> {
    if attachments.len() > MAX_ATTACHMENTS {
        return corrupted("attachment count exceeds the resource cap");
    }
    put_u32_len(out, attachments.len(), "attachment")?;
    for attachment in attachments {
        validate_utf16(&attachment.name, "attachment name")?;
        let name_length = u8::try_from(attachment.name.len())
            .map_err(|_| PptError::Corrupted("attachment name is too long".to_string()))?;
        put_u32(out, attachment.method);
        out.push(name_length);
        put_utf16(out, &attachment.name);
        let length = u64::try_from(attachment.data.len())
            .map_err(|_| PptError::Corrupted("attachment length overflow".to_string()))?;
        put_u32(out, length as u32);
        put_u32(out, (length >> 32) as u32);
        out.extend_from_slice(&attachment.data);
    }
    Ok(())
}

fn write_intro_text(out: &mut Vec<u8>, value: &[u16]) -> Result<()> {
    validate_utf16(value, "intro text")?;
    put_u32_len(out, checked_double(value.len())?, "intro text")?;
    put_utf16(out, value);
    Ok(())
}

fn write_u16_blob(out: &mut Vec<u8>, value: &[u8]) -> Result<()> {
    put_u16_len(out, value.len(), "envelope blob")?;
    out.extend_from_slice(value);
    Ok(())
}

fn write_u32_blob(out: &mut Vec<u8>, value: &[u8]) -> Result<()> {
    put_u32_len(out, value.len(), "envelope blob")?;
    out.extend_from_slice(value);
    Ok(())
}

fn put_u16_len(out: &mut Vec<u8>, length: usize, name: &str) -> Result<()> {
    let length = u16::try_from(length)
        .map_err(|_| PptError::Corrupted(format!("{name} length overflow")))?;
    put_u16(out, length);
    Ok(())
}

fn put_u32_len(out: &mut Vec<u8>, length: usize, name: &str) -> Result<()> {
    let length = u32::try_from(length)
        .map_err(|_| PptError::Corrupted(format!("{name} length overflow")))?;
    put_u32(out, length);
    Ok(())
}

fn put_u16(out: &mut Vec<u8>, value: u16) {
    out.extend_from_slice(&value.to_le_bytes());
}

fn put_u32(out: &mut Vec<u8>, value: u32) {
    out.extend_from_slice(&value.to_le_bytes());
}

fn put_utf16(out: &mut Vec<u8>, value: &[u16]) {
    for unit in value {
        put_u16(out, *unit);
    }
}

fn checked_double(value: usize) -> Result<usize> {
    value
        .checked_mul(2)
        .ok_or_else(|| PptError::Corrupted("UTF-16 length overflow".to_string()))
}

fn validate_utf16(value: &[u16], name: &str) -> Result<()> {
    if char::decode_utf16(value.iter().copied()).any(|unit| unit.is_err()) {
        corrupted(&format!("{name} contains an unpaired surrogate"))
    } else {
        Ok(())
    }
}

fn bounded_count(value: u32, cap: usize, name: &str) -> Result<usize> {
    let value = usize::try_from(value)
        .map_err(|_| PptError::Corrupted(format!("{name} count overflow")))?;
    if value > cap {
        return corrupted(&format!("{name} count exceeds the resource cap"));
    }
    Ok(value)
}

fn minute_time(value: u32, name: &str) -> Result<u32> {
    if value > MAX_MINUTE_TIME {
        corrupted(&format!("{name} is outside the defined range"))
    } else {
        Ok(value)
    }
}

fn corrupted<T>(message: &str) -> Result<T> {
    Err(PptError::Corrupted(message.to_string()))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn empty_collection() -> MsoRecipientCollection {
        MsoRecipientCollection::default()
    }

    fn sample() -> PowerPointEnvelopeData {
        let unicode = |value: &str| MsoEnvelopeText::Unicode(value.encode_utf16().collect());
        PowerPointEnvelopeData {
            clsid: MSO_ENVELOPE_CLSID,
            payload: PowerPointEnvelopePayload::Mso(MsoEnvelope {
                version: MsoEnvelopeVersion::Office8,
                last_sent_time: 0,
                flag_status: MsoFollowUpStatus::Flagged,
                reply_time: MAX_MINUTE_TIME,
                request: unicode("reply"),
                sent_representing_entry_id: vec![1, 2, 3],
                sent_representing_name: unicode("sender"),
                internet_account_stamp: unicode("stamp"),
                internet_account_name: unicode("account"),
                expiry_time: MAX_MINUTE_TIME,
                deferred_delivery_time: 0,
                delete_after_submit: false,
                security: MsoSecurityFlags {
                    signed: true,
                    encrypted: false,
                },
                delivery_report: true,
                read_receipt: false,
                categories: unicode("category"),
                sensitivity: MsoSensitivity::Private,
                importance: MsoImportance::High,
                subject: unicode("subject"),
                voting_options: b"yes;no".to_vec(),
                reply_recipients: MsoRecipientCollection {
                    recipients: vec![MsoRecipientProperties {
                        properties: vec![
                            MsoRecipientProperty {
                                property_id: 0x3001,
                                value: MsoPropertyValue::Unicode(
                                    "Recipient".encode_utf16().collect(),
                                ),
                            },
                            MsoRecipientProperty {
                                property_id: 0x0c15,
                                value: MsoPropertyValue::Boolean(true),
                            },
                        ],
                    }],
                },
                contact_link_recipients: Some(empty_collection()),
                recipients: empty_collection(),
                attachments: vec![MsoAttachment {
                    method: 1,
                    name: "a.txt".encode_utf16().collect(),
                    data: vec![0xde, 0xad],
                }],
                intro_text: Some("intro".encode_utf16().collect()),
            }),
        }
    }

    #[test]
    fn known_envelope_round_trips() {
        let expected = sample();
        let record = expected.to_record().unwrap();
        assert_eq!(PowerPointEnvelopeData::parse(&record).unwrap(), expected);
    }

    #[test]
    fn unknown_clsid_is_bounded_opaque_data() {
        let expected = PowerPointEnvelopeData {
            clsid: [7; 16],
            payload: PowerPointEnvelopePayload::Opaque(vec![1, 2, 3]),
        };
        assert_eq!(
            PowerPointEnvelopeData::parse(&expected.to_record().unwrap()).unwrap(),
            expected
        );
    }

    #[test]
    fn rejects_reserved_flags_and_version_mismatches() {
        let mut record = sample().to_record().unwrap();
        let security_offset =
            16 + 4 * 4 + 2 + 5 * 2 + 4 + 3 + 2 + 6 * 2 + 2 + 5 * 2 + 2 + 7 * 2 + 4 * 2 + 4;
        record.data[security_offset..security_offset + 4].copy_from_slice(&4u32.to_le_bytes());
        assert!(PowerPointEnvelopeData::parse(&record).is_err());

        let mut value = sample();
        let PowerPointEnvelopePayload::Mso(envelope) = &mut value.payload else {
            unreachable!();
        };
        envelope.contact_link_recipients = None;
        assert!(value.to_record().is_err());
    }

    #[test]
    fn rejects_unpaired_utf16_on_parse_and_write() {
        let mut value = sample();
        let PowerPointEnvelopePayload::Mso(envelope) = &mut value.payload else {
            unreachable!();
        };
        envelope.subject = MsoEnvelopeText::Unicode(vec![0xd800]);
        assert!(value.to_record().is_err());

        let mut record = sample().to_record().unwrap();
        let subject = "subject".encode_utf16().collect::<Vec<_>>();
        let bytes = subject
            .iter()
            .flat_map(|unit| unit.to_le_bytes())
            .collect::<Vec<_>>();
        let offset = record
            .data
            .windows(bytes.len())
            .position(|window| window == bytes)
            .unwrap();
        record.data[offset..offset + 2].copy_from_slice(&0xd800u16.to_le_bytes());
        assert!(PowerPointEnvelopeData::parse(&record).is_err());
    }
}
