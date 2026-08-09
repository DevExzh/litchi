//! Binary codec for `[MS-DOC]` `MsoEnvelopeCLSID` metadata.

use super::model::{
    Attachment, Envelope, FollowUpStatus, Importance, MSO_ENVELOPE_CLSID, Message, Payload,
    PropertyValue, RecipientCollection, RecipientProperties, RecipientProperty, SecurityFlags,
    Sensitivity, Text, Version,
};
use super::validation::{self, MAX_ENVELOPE_BYTES};
use crate::package::{Error as PackageError, Result};
use crate::parts::fib::FileInformationBlock;

/// Table-pointer index of `fcMsoEnvelope`/`lcbMsoEnvelope` in
/// `FibRgFcLcb2000` (`[MS-DOC]` 2.5.7).
pub(super) const FIB_INDEX: usize = validation::FIB_INDEX;
const RECIPIENT_COLLECTION_TAG: u32 = 0xDCCA_0123;
const RECIPIENT_COLLECTION_VERSION: u32 = 1;

/// Parse the optional envelope range addressed by the FIB.
pub(super) fn parse_fib(
    fib: &FileInformationBlock,
    table_stream: &[u8],
) -> Result<Option<Envelope>> {
    let Some((offset, length)) = fib.get_table_pointer(FIB_INDEX) else {
        return Ok(None);
    };
    if length == 0 {
        return Ok(None);
    }
    let data = validation::table_range(table_stream, offset, length)?;
    parse(data).map(Some)
}

pub(super) fn parse(data: &[u8]) -> Result<Envelope> {
    if data.len() < 16 || data.len() > MAX_ENVELOPE_BYTES {
        return Err(corrupted("MsoEnvelopeCLSID has an invalid size"));
    }
    let mut class_id = [0u8; 16];
    class_id.copy_from_slice(&data[..16]);
    let body = &data[16..];
    let payload = if class_id == MSO_ENVELOPE_CLSID {
        Payload::Message(Box::new(parse_message(body)?))
    } else {
        Payload::Opaque(body.to_vec().into_boxed_slice())
    };
    let value = Envelope::from_parts(class_id, payload);
    validation::validate(&value)?;
    Ok(value)
}

pub(super) fn write(value: &Envelope) -> Result<Vec<u8>> {
    validation::validate(value)?;
    let mut output = Vec::new();
    output.extend_from_slice(value.class_id());
    match value.payload() {
        Payload::Message(message) => write_message(&mut output, message)?,
        Payload::Opaque(payload) => output.extend_from_slice(payload),
    }
    if output.len() > MAX_ENVELOPE_BYTES {
        return Err(corrupted("MsoEnvelopeCLSID exceeds the resource cap"));
    }
    Ok(output)
}

fn parse_message(data: &[u8]) -> Result<Message> {
    let mut input = Cursor::new(data);
    let version = match input.u32()? {
        6 => Version::Office6,
        8 => Version::Office8,
        _ => return Err(corrupted("MsoEnvelope has an undefined version")),
    };
    let last_sent_time = input.u32()?;
    let flag_status = match input.u32()? {
        0 => FollowUpStatus::None,
        1 => FollowUpStatus::Flagged,
        2 => FollowUpStatus::Complete,
        _ => return Err(corrupted("MsoEnvelope has an invalid follow-up status")),
    };
    let reply_time = input.u32()?;
    let request = input.versioned_text(version)?;
    let sent_representing_entry_id = input.blob32()?;
    let sent_representing_name = input.versioned_text(version)?;
    let internet_account_stamp = input.versioned_text(version)?;
    let internet_account_name = input.versioned_text(version)?;
    let expiry_time = input.u32()?;
    let deferred_delivery_time = input.u32()?;
    let delete_after_submit = input.boolean32()?;
    let security_bits = input.u32()?;
    if security_bits & !0x3 != 0 {
        return Err(corrupted(
            "MsoEnvelope security flags contain reserved bits",
        ));
    }
    let security = SecurityFlags {
        signed: security_bits & 1 != 0,
        encrypted: security_bits & 2 != 0,
    };
    let delivery_report = input.boolean32()?;
    let read_receipt = input.boolean32()?;
    let categories = input.versioned_text(version)?;
    let sensitivity = match input.u32()? {
        0 => Sensitivity::Normal,
        1 => Sensitivity::Personal,
        2 => Sensitivity::Private,
        3 => Sensitivity::Confidential,
        _ => return Err(corrupted("MsoEnvelope has an invalid sensitivity")),
    };
    let importance = match input.u32()? {
        0 => Importance::Low,
        1 => Importance::Normal,
        2 => Importance::High,
        _ => return Err(corrupted("MsoEnvelope has an invalid importance")),
    };
    let subject = input.versioned_text(version)?;
    let voting_options = input.blob16()?;
    let reply_recipients = input.recipient_collection()?;
    let contact_link_recipients = if version == Version::Office8 {
        Some(input.recipient_collection()?)
    } else {
        None
    };
    let recipients = input.recipient_collection()?;
    let attachments = input.attachments()?;
    let intro_text = if version == Version::Office8 {
        Some(input.intro_text()?)
    } else {
        None
    };
    let tail = input.take_remaining();
    let value = Message {
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
        tail,
    };
    validation::validate(&Envelope::from_parts(
        MSO_ENVELOPE_CLSID,
        Payload::Message(Box::new(value.clone())),
    ))?;
    Ok(value)
}

fn write_message(output: &mut Vec<u8>, message: &Message) -> Result<()> {
    put_u32(output, message.version as u32);
    put_u32(output, message.last_sent_time);
    put_u32(output, message.flag_status as u32);
    put_u32(output, message.reply_time);
    write_versioned_text(output, message.version, &message.request)?;
    write_blob32(output, &message.sent_representing_entry_id);
    write_versioned_text(output, message.version, &message.sent_representing_name)?;
    write_versioned_text(output, message.version, &message.internet_account_stamp)?;
    write_versioned_text(output, message.version, &message.internet_account_name)?;
    put_u32(output, message.expiry_time);
    put_u32(output, message.deferred_delivery_time);
    put_u32(output, u32::from(message.delete_after_submit));
    put_u32(
        output,
        u32::from(message.security.signed) | (u32::from(message.security.encrypted) << 1),
    );
    put_u32(output, u32::from(message.delivery_report));
    put_u32(output, u32::from(message.read_receipt));
    write_versioned_text(output, message.version, &message.categories)?;
    put_u32(output, message.sensitivity as u32);
    put_u32(output, message.importance as u32);
    write_versioned_text(output, message.version, &message.subject)?;
    write_blob16(output, &message.voting_options)?;
    write_recipient_collection(output, &message.reply_recipients)?;
    match (message.version, &message.contact_link_recipients) {
        (Version::Office8, Some(collection)) => write_recipient_collection(output, collection)?,
        (Version::Office6, None) => {},
        _ => {
            return Err(corrupted(
                "contact-link recipients do not match envelope version",
            ));
        },
    }
    write_recipient_collection(output, &message.recipients)?;
    write_attachments(output, &message.attachments)?;
    match (message.version, &message.intro_text) {
        (Version::Office8, Some(value)) => {
            put_u32(
                output,
                u32::try_from(
                    value
                        .len()
                        .checked_mul(2)
                        .ok_or_else(|| corrupted("intro text size overflows"))?,
                )
                .map_err(|_| corrupted("intro text exceeds the wire length"))?,
            );
            put_utf16(output, value);
        },
        (Version::Office6, None) => {},
        _ => return Err(corrupted("intro text does not match envelope version")),
    }
    output.extend_from_slice(&message.tail);
    Ok(())
}

fn write_versioned_text(output: &mut Vec<u8>, version: Version, value: &Text) -> Result<()> {
    match (version, value) {
        (Version::Office6, Text::Ansi(bytes)) => {
            put_u16(
                output,
                u16::try_from(bytes.len())
                    .map_err(|_| corrupted("ANSI envelope string is too long"))?,
            );
            output.extend_from_slice(bytes);
        },
        (Version::Office8, Text::Unicode(units)) => {
            put_u16(
                output,
                u16::try_from(units.len())
                    .map_err(|_| corrupted("Unicode envelope string is too long"))?,
            );
            put_utf16(output, units);
        },
        _ => return Err(corrupted("envelope string encoding does not match version")),
    }
    Ok(())
}

fn write_recipient_collection(
    output: &mut Vec<u8>,
    collection: &RecipientCollection,
) -> Result<()> {
    put_u32(output, RECIPIENT_COLLECTION_TAG);
    put_u32(output, RECIPIENT_COLLECTION_VERSION);
    put_u32(
        output,
        u32::try_from(collection.recipients.len())
            .map_err(|_| corrupted("recipient count exceeds the wire length"))?,
    );
    for recipient in &collection.recipients {
        put_u32(
            output,
            u32::try_from(recipient.properties.len())
                .map_err(|_| corrupted("property count exceeds the wire length"))?,
        );
        put_u32(output, 0);
        for property in &recipient.properties {
            write_property(output, property)?;
        }
    }
    Ok(())
}

fn write_property(output: &mut Vec<u8>, property: &RecipientProperty) -> Result<()> {
    let property_type = match &property.value {
        PropertyValue::Long(_) => 0x0003,
        PropertyValue::Null(_) => 0x0001,
        PropertyValue::Boolean(_) => 0x000B,
        PropertyValue::SystemTime { .. } => 0x0040,
        PropertyValue::Error(_) => 0x000A,
        PropertyValue::String8(_) => 0x001E,
        PropertyValue::Unicode(_) => 0x001F,
        PropertyValue::Binary(_) => 0x0102,
        PropertyValue::MultiString8(_) => 0x101E,
        PropertyValue::MultiBinary(_) => 0x1102,
    };
    put_u32(
        output,
        (u32::from(property.property_id) << 16) | property_type,
    );
    match &property.value {
        PropertyValue::Long(value) | PropertyValue::Null(value) | PropertyValue::Error(value) => {
            put_u32(output, *value);
        },
        PropertyValue::Boolean(value) => put_u16(output, u16::from(*value)),
        PropertyValue::SystemTime { high, low } => {
            put_u32(output, *high);
            put_u32(output, *low);
        },
        PropertyValue::String8(value) | PropertyValue::Binary(value) => {
            write_blob16(output, value)?;
        },
        PropertyValue::Unicode(value) => {
            put_u16(
                output,
                u16::try_from(
                    value
                        .len()
                        .checked_mul(2)
                        .ok_or_else(|| corrupted("Unicode recipient property size overflows"))?,
                )
                .map_err(|_| corrupted("Unicode recipient property is too long"))?,
            );
            put_utf16(output, value);
        },
        PropertyValue::MultiString8(values) | PropertyValue::MultiBinary(values) => {
            put_u32(
                output,
                u32::try_from(values.len())
                    .map_err(|_| corrupted("multi-value property count exceeds the wire length"))?,
            );
            for value in values {
                write_blob16(output, value)?;
            }
        },
    }
    Ok(())
}

fn write_attachments(output: &mut Vec<u8>, attachments: &[Attachment]) -> Result<()> {
    put_u32(
        output,
        u32::try_from(attachments.len())
            .map_err(|_| corrupted("attachment count exceeds the wire length"))?,
    );
    for attachment in attachments {
        put_u32(output, attachment.method);
        output.push(
            u8::try_from(attachment.name.len())
                .map_err(|_| corrupted("attachment name is too long"))?,
        );
        put_utf16(output, &attachment.name);
        let length = u64::try_from(attachment.data.len())
            .map_err(|_| corrupted("attachment length overflows"))?;
        put_u32(output, length as u32);
        put_u32(output, (length >> 32) as u32);
        output.extend_from_slice(&attachment.data);
    }
    Ok(())
}

fn write_blob16(output: &mut Vec<u8>, value: &[u8]) -> Result<()> {
    put_u16(
        output,
        u16::try_from(value.len()).map_err(|_| corrupted("16-bit envelope blob is too long"))?,
    );
    output.extend_from_slice(value);
    Ok(())
}

fn write_blob32(output: &mut Vec<u8>, value: &[u8]) {
    put_u32(output, value.len() as u32);
    output.extend_from_slice(value);
}

struct Cursor<'a> {
    data: &'a [u8],
    position: usize,
}

impl<'a> Cursor<'a> {
    fn new(data: &'a [u8]) -> Self {
        Self { data, position: 0 }
    }

    fn remaining(&self) -> usize {
        self.data.len().saturating_sub(self.position)
    }

    fn take(&mut self, length: usize) -> Result<&'a [u8]> {
        let end = self
            .position
            .checked_add(length)
            .ok_or_else(|| corrupted("envelope offset overflows"))?;
        if end > self.data.len() {
            return Err(corrupted("MsoEnvelope is truncated"));
        }
        let value = &self.data[self.position..end];
        self.position = end;
        Ok(value)
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
            _ => Err(corrupted("MsoEnvelope contains a non-Boolean flag")),
        }
    }

    fn blob16(&mut self) -> Result<Box<[u8]>> {
        let length = usize::from(self.u16()?);
        Ok(self.take(length)?.to_vec().into_boxed_slice())
    }

    fn blob32(&mut self) -> Result<Box<[u8]>> {
        let length = usize::try_from(self.u32()?)
            .map_err(|_| corrupted("32-bit envelope blob length overflows"))?;
        if length > MAX_ENVELOPE_BYTES {
            return Err(corrupted("32-bit envelope blob exceeds the resource cap"));
        }
        Ok(self.take(length)?.to_vec().into_boxed_slice())
    }

    fn utf16_bytes(&mut self, byte_length: usize, name: &str) -> Result<Box<[u16]>> {
        if !byte_length.is_multiple_of(2) {
            return Err(corrupted(format!("{name} has an odd UTF-16 byte size")));
        }
        let units = self
            .take(byte_length)?
            .chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
            .collect::<Vec<_>>();
        if char::decode_utf16(units.iter().copied()).any(|unit| unit.is_err()) {
            return Err(corrupted(format!("{name} contains an unpaired surrogate")));
        }
        Ok(units.into_boxed_slice())
    }

    fn versioned_text(&mut self, version: Version) -> Result<Text> {
        let characters = usize::from(self.u16()?);
        match version {
            Version::Office6 => Ok(Text::Ansi(
                self.take(characters)?.to_vec().into_boxed_slice(),
            )),
            Version::Office8 => Ok(Text::Unicode(
                self.utf16_bytes(
                    characters
                        .checked_mul(2)
                        .ok_or_else(|| corrupted("Unicode envelope string size overflows"))?,
                    "Unicode envelope string",
                )?,
            )),
        }
    }

    fn recipient_collection(&mut self) -> Result<RecipientCollection> {
        if self.u32()? != RECIPIENT_COLLECTION_TAG || self.u32()? != RECIPIENT_COLLECTION_VERSION {
            return Err(corrupted(
                "recipient collection has an invalid tag or version",
            ));
        }
        let count = bounded_count(self.u32()?, 65_536, 8, self.remaining(), "recipient")?;
        let mut recipients = Vec::with_capacity(count);
        for _ in 0..count {
            let property_count =
                bounded_count(self.u32()?, 65_536, 8, self.remaining(), "property")?;
            let _ignored = self.u32()?;
            let mut properties = Vec::with_capacity(property_count);
            for _ in 0..property_count {
                properties.push(self.property()?);
            }
            recipients.push(RecipientProperties { properties });
        }
        Ok(RecipientCollection { recipients })
    }

    fn property(&mut self) -> Result<RecipientProperty> {
        let tag = self.u32()?;
        let property_id = (tag >> 16) as u16;
        let value = match tag as u16 {
            0x0003 => PropertyValue::Long(self.u32()?),
            0x0001 => PropertyValue::Null(self.u32()?),
            0x000B => match self.u16()? {
                0 => PropertyValue::Boolean(false),
                1 => PropertyValue::Boolean(true),
                _ => return Err(corrupted("recipient property has a non-Boolean value")),
            },
            0x0040 => PropertyValue::SystemTime {
                high: self.u32()?,
                low: self.u32()?,
            },
            0x000A => PropertyValue::Error(self.u32()?),
            0x001E => PropertyValue::String8(self.blob16()?),
            0x001F => {
                let size = usize::from(self.u16()?);
                PropertyValue::Unicode(self.utf16_bytes(size, "Unicode recipient property")?)
            },
            0x0102 => PropertyValue::Binary(self.blob16()?),
            0x101E => PropertyValue::MultiString8(self.multi_blob("multi-string")?),
            0x1102 => PropertyValue::MultiBinary(self.multi_blob("multi-binary")?),
            _ => return Err(corrupted("recipient property has an unsupported type")),
        };
        Ok(RecipientProperty { property_id, value })
    }

    fn multi_blob(&mut self, name: &str) -> Result<Vec<Box<[u8]>>> {
        let count = bounded_count(self.u32()?, 65_536, 2, self.remaining(), name)?;
        let mut values = Vec::with_capacity(count);
        for _ in 0..count {
            values.push(self.blob16()?);
        }
        Ok(values)
    }

    fn attachments(&mut self) -> Result<Vec<Attachment>> {
        let count = bounded_count(self.u32()?, 4_096, 13, self.remaining(), "attachment")?;
        let mut attachments = Vec::with_capacity(count);
        for _ in 0..count {
            let method = self.u32()?;
            let name_length = usize::from(self.take(1)?[0]);
            let name = self.utf16_bytes(
                name_length
                    .checked_mul(2)
                    .ok_or_else(|| corrupted("attachment name size overflows"))?,
                "attachment name",
            )?;
            let low = u64::from(self.u32()?);
            let high = u64::from(self.u32()?);
            let size = usize::try_from((high << 32) | low)
                .map_err(|_| corrupted("attachment size exceeds the platform size"))?;
            if size > MAX_ENVELOPE_BYTES {
                return Err(corrupted("attachment exceeds the resource cap"));
            }
            let data = self.take(size)?.to_vec().into_boxed_slice();
            attachments.push(Attachment { method, name, data });
        }
        Ok(attachments)
    }

    fn intro_text(&mut self) -> Result<Box<[u16]>> {
        let bytes = usize::try_from(self.u32()?)
            .map_err(|_| corrupted("intro text length exceeds the platform size"))?;
        self.utf16_bytes(bytes, "intro text")
    }

    fn take_remaining(&mut self) -> Box<[u8]> {
        let tail = self.data[self.position..].to_vec().into_boxed_slice();
        self.position = self.data.len();
        tail
    }
}

fn bounded_count(
    value: u32,
    cap: usize,
    minimum_size: usize,
    remaining: usize,
    name: &str,
) -> Result<usize> {
    let value = usize::try_from(value).map_err(|_| corrupted(format!("{name} count overflows")))?;
    if value > cap || value > remaining / minimum_size {
        return Err(corrupted(format!("{name} count exceeds the resource cap")));
    }
    Ok(value)
}

fn put_u16(output: &mut Vec<u8>, value: u16) {
    output.extend_from_slice(&value.to_le_bytes());
}

fn put_u32(output: &mut Vec<u8>, value: u32) {
    output.extend_from_slice(&value.to_le_bytes());
}

fn put_utf16(output: &mut Vec<u8>, value: &[u16]) {
    for unit in value {
        put_u16(output, *unit);
    }
}

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}
