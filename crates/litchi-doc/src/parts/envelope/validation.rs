//! Resource and semantic validation for the envelope owner.

use super::model::MSO_ENVELOPE_CLSID;
use super::model::{
    Attachment, Envelope, Message, Payload, PropertyValue, RecipientCollection, RecipientProperty,
    Text, Version,
};
use crate::package::{Error as PackageError, Result};

/// Maximum serialized `MsoEnvelopeCLSID` size accepted from a DOC table
/// stream or produced by the detached codec.
pub(super) const MAX_ENVELOPE_BYTES: usize = 16 * 1024 * 1024;
/// Maximum value of the minute-based envelope timestamps from `[MS-OSHARED]`.
pub(super) const MAX_MINUTE_TIME: u32 = 0x5AE9_80E0;
const MAX_COLLECTION_ITEMS: usize = 65_536;
const MAX_TOTAL_PROPERTIES: usize = 1_000_000;
const MAX_ATTACHMENTS: usize = 4_096;
const MAX_BLOB_BYTES: usize = MAX_ENVELOPE_BYTES;

pub(super) fn validate(value: &Envelope) -> Result<()> {
    let body_length = match (value.class_id(), value.payload()) {
        (&MSO_ENVELOPE_CLSID, Payload::Message(message)) => message_length(message)?,
        (&MSO_ENVELOPE_CLSID, Payload::Opaque(_)) => {
            return Err(corrupted(
                "the Office envelope CLSID requires a typed Message payload",
            ));
        },
        (_, Payload::Message(_)) => {
            return Err(corrupted(
                "a producer-defined envelope CLSID cannot carry a typed Message payload",
            ));
        },
        (_, Payload::Opaque(payload)) => {
            if payload.len() > MAX_ENVELOPE_BYTES.saturating_sub(16) {
                return Err(corrupted(
                    "opaque envelope payload exceeds the resource cap",
                ));
            }
            payload.len()
        },
    };
    if body_length > MAX_ENVELOPE_BYTES.saturating_sub(16) {
        return Err(corrupted("envelope body exceeds the resource cap"));
    }
    Ok(())
}

fn message_length(message: &Message) -> Result<usize> {
    let mut length = 16usize; // Ver, LastSentTime, FlagStatus, ReplyTime.
    validate_time(message.last_sent_time, "last-sent time")?;
    validate_time(message.reply_time, "reply time")?;
    validate_time(message.expiry_time, "expiry time")?;
    validate_time(message.deferred_delivery_time, "deferred-delivery time")?;
    validate_text(&message.request, message.version, "request")?;
    length = add(length, text_length(&message.request)?, "message size")?;
    length = add(
        length,
        blob32_length(
            &message.sent_representing_entry_id,
            "sent-representing entry id",
        )?,
        "message size",
    )?;
    for (text, name) in [
        (&message.sent_representing_name, "sent-representing name"),
        (&message.internet_account_stamp, "internet account stamp"),
        (&message.internet_account_name, "internet account name"),
    ] {
        validate_text(text, message.version, name)?;
        length = add(length, text_length(text)?, "message size")?;
    }
    length = add(length, 24, "message size")?; // Four times and four u32 flags.
    validate_text(&message.categories, message.version, "categories")?;
    length = add(length, text_length(&message.categories)?, "message size")?;
    length = add(length, 8, "message size")?; // Sensitivity and importance.
    validate_text(&message.subject, message.version, "subject")?;
    length = add(length, text_length(&message.subject)?, "message size")?;
    length = add(
        length,
        blob16_length(&message.voting_options, "voting options")?,
        "message size",
    )?;

    let mut property_count = 0usize;
    length = add(
        length,
        collection_length(&message.reply_recipients, &mut property_count)?,
        "message size",
    )?;
    match (message.version, &message.contact_link_recipients) {
        (Version::Office8, Some(collection)) => {
            length = add(
                length,
                collection_length(collection, &mut property_count)?,
                "message size",
            )?;
        },
        (Version::Office6, None) => {},
        (Version::Office6, Some(_)) => {
            return Err(corrupted(
                "Office 6 envelopes must not contain contact-link recipients",
            ));
        },
        (Version::Office8, None) => {
            return Err(corrupted(
                "Office 8 envelopes must contain contact-link recipients",
            ));
        },
    }
    length = add(
        length,
        collection_length(&message.recipients, &mut property_count)?,
        "message size",
    )?;
    if message.attachments.len() > MAX_ATTACHMENTS {
        return Err(corrupted("attachment count exceeds the resource cap"));
    }
    length = add(length, 4, "message size")?;
    for attachment in &message.attachments {
        validate_attachment(attachment)?;
        length = add(
            length,
            13usize
                .checked_add(
                    attachment
                        .name
                        .len()
                        .checked_mul(2)
                        .ok_or_else(|| corrupted("attachment name size overflows"))?,
                )
                .and_then(|value| value.checked_add(attachment.data.len()))
                .ok_or_else(|| corrupted("attachment size overflows"))?,
            "message size",
        )?;
    }
    match (message.version, &message.intro_text) {
        (Version::Office8, Some(value)) => {
            validate_utf16(value, "intro text")?;
            if value
                .len()
                .checked_mul(2)
                .is_none_or(|bytes| bytes > u32::MAX as usize)
            {
                return Err(corrupted("intro text exceeds the wire length"));
            }
            length = add(
                length,
                4usize
                    .checked_add(
                        value
                            .len()
                            .checked_mul(2)
                            .ok_or_else(|| corrupted("intro text size overflows"))?,
                    )
                    .ok_or_else(|| corrupted("intro text size overflows"))?,
                "message size",
            )?;
        },
        (Version::Office6, None) => {},
        (Version::Office6, Some(_)) => {
            return Err(corrupted("Office 6 envelopes must not contain intro text"));
        },
        (Version::Office8, None) => {
            return Err(corrupted("Office 8 envelopes must contain intro text"));
        },
    }
    if length > MAX_ENVELOPE_BYTES.saturating_sub(16) {
        return Err(corrupted("MsoEnvelope exceeds the resource cap"));
    }
    Ok(length)
}

fn collection_length(
    collection: &RecipientCollection,
    property_count: &mut usize,
) -> Result<usize> {
    if collection.recipients.len() > MAX_COLLECTION_ITEMS {
        return Err(corrupted("recipient count exceeds the resource cap"));
    }
    let mut length = 12usize; // tag, version, count.
    if collection.recipients.len() > (MAX_ENVELOPE_BYTES - length) / 8 {
        return Err(corrupted(
            "recipient count is too large for the envelope cap",
        ));
    }
    for recipient in &collection.recipients {
        if recipient.properties.len() > MAX_COLLECTION_ITEMS {
            return Err(corrupted("property count exceeds the resource cap"));
        }
        *property_count = property_count
            .checked_add(recipient.properties.len())
            .ok_or_else(|| corrupted("property count overflows"))?;
        if *property_count > MAX_TOTAL_PROPERTIES {
            return Err(corrupted(
                "total recipient properties exceed the resource cap",
            ));
        }
        let mut recipient_length = 8usize; // Count and Ignored.
        for property in &recipient.properties {
            recipient_length = add(
                recipient_length,
                property_length(property)?,
                "recipient size",
            )?;
        }
        length = add(length, recipient_length, "recipient collection size")?;
    }
    Ok(length)
}

fn property_length(property: &RecipientProperty) -> Result<usize> {
    let value_length = match &property.value {
        PropertyValue::Long(_) | PropertyValue::Null(_) | PropertyValue::Error(_) => 4,
        PropertyValue::Boolean(_) => 2,
        PropertyValue::SystemTime { .. } => 8,
        PropertyValue::String8(value) | PropertyValue::Binary(value) => {
            short_blob_length(value, "recipient blob")?
        },
        PropertyValue::Unicode(value) => {
            validate_utf16(value, "recipient Unicode property")?;
            short_utf16_length(value, "recipient Unicode property")?
        },
        PropertyValue::MultiString8(values) | PropertyValue::MultiBinary(values) => {
            if values.len() > MAX_COLLECTION_ITEMS {
                return Err(corrupted(
                    "multi-value property count exceeds the resource cap",
                ));
            }
            let mut length = 4usize;
            for value in values {
                length = add(
                    length,
                    short_blob_length(value, "multi-value property")?,
                    "multi-value property size",
                )?;
            }
            length
        },
    };
    add(4, value_length, "recipient property size")
}

fn validate_attachment(value: &Attachment) -> Result<()> {
    validate_utf16(&value.name, "attachment name")?;
    if value.name.len() > u8::MAX as usize {
        return Err(corrupted(
            "attachment name exceeds the one-byte character count",
        ));
    }
    if value.data.len() > MAX_BLOB_BYTES {
        return Err(corrupted("attachment data exceeds the resource cap"));
    }
    Ok(())
}

fn validate_text(value: &Text, version: Version, name: &str) -> Result<()> {
    match (version, value) {
        (Version::Office6, Text::Ansi(bytes)) => {
            if bytes.len() > u16::MAX as usize {
                return Err(corrupted(format!(
                    "{name} exceeds the 16-bit character count"
                )));
            }
        },
        (Version::Office8, Text::Unicode(units)) => {
            validate_utf16(units, name)?;
            if units.len() > u16::MAX as usize {
                return Err(corrupted(format!(
                    "{name} exceeds the 16-bit character count"
                )));
            }
        },
        (Version::Office6, Text::Unicode(_)) => {
            return Err(corrupted(format!(
                "{name} must use ANSI encoding for Office 6"
            )));
        },
        (Version::Office8, Text::Ansi(_)) => {
            return Err(corrupted(format!(
                "{name} must use Unicode encoding for Office 8"
            )));
        },
    }
    Ok(())
}

fn text_length(value: &Text) -> Result<usize> {
    match value {
        Text::Ansi(bytes) => short_blob_length(bytes, "envelope string"),
        Text::Unicode(units) => short_utf16_length(units, "envelope string"),
    }
}

fn blob16_length(value: &[u8], name: &str) -> Result<usize> {
    short_blob_length(value, name)
}

fn blob32_length(value: &[u8], name: &str) -> Result<usize> {
    if value.len() > MAX_BLOB_BYTES || value.len() > u32::MAX as usize {
        return Err(corrupted(format!("{name} exceeds the 32-bit resource cap")));
    }
    add(4, value.len(), name)
}

fn short_blob_length(value: &[u8], name: &str) -> Result<usize> {
    if value.len() > u16::MAX as usize || value.len() > MAX_BLOB_BYTES {
        return Err(corrupted(format!("{name} exceeds the 16-bit resource cap")));
    }
    add(2, value.len(), name)
}

fn short_utf16_length(value: &[u16], name: &str) -> Result<usize> {
    let bytes = value
        .len()
        .checked_mul(2)
        .ok_or_else(|| corrupted(format!("{name} size overflows")))?;
    if bytes > u16::MAX as usize || bytes > MAX_BLOB_BYTES {
        return Err(corrupted(format!(
            "{name} exceeds the 16-bit byte-size cap"
        )));
    }
    add(2, bytes, name)
}

fn validate_utf16(value: &[u16], name: &str) -> Result<()> {
    if char::decode_utf16(value.iter().copied()).any(|unit| unit.is_err()) {
        return Err(corrupted(format!("{name} contains an unpaired surrogate")));
    }
    Ok(())
}

fn validate_time(value: u32, name: &str) -> Result<()> {
    if value > MAX_MINUTE_TIME {
        Err(corrupted(format!("{name} is outside the defined range")))
    } else {
        Ok(())
    }
}

fn add(left: usize, right: usize, name: &str) -> Result<usize> {
    left.checked_add(right)
        .ok_or_else(|| corrupted(format!("{name} overflows")))
}

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}
