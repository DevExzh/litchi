//! Binary codec for MS-DOC Metadata and Recipient.

use super::model::{
    DeliveryOption, Metadata, NarrowString, Protection, Recipient, validate_short_string,
};
use crate::package::{Error as PackageError, Result};
use crate::parts::fib::FileInformationBlock;

/// FIB index of `fcRouteSlip`/`lcbRouteSlip` in `FibRgFcLcb97`
/// (MS-DOC 2.5.6).
const FIB_INDEX: usize = 70;

const ROUTE_SLIP_FIXED_SIZE: usize = 16;
const MAX_SHORT_STRING_LENGTH: usize = 255;

/// Parse the optional Metadata table range addressed by a FIB.
pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Option<Metadata>> {
    let Some((offset, length)) = fib.get_table_pointer(FIB_INDEX) else {
        return Ok(None);
    };
    if length == 0 {
        return Ok(None);
    }

    let start =
        usize::try_from(offset).map_err(|_| corrupted("Metadata table offset exceeds usize"))?;
    let length =
        usize::try_from(length).map_err(|_| corrupted("Metadata table length exceeds usize"))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted("Metadata table range overflows"))?;
    let data = table_stream
        .get(start..end)
        .ok_or_else(|| corrupted("Metadata extends beyond the table stream"))?;
    parse_bytes(data).map(Some)
}

/// Parse exactly one complete Metadata payload.
pub fn parse_bytes(data: &[u8]) -> Result<Metadata> {
    let mut reader = Reader::new(data);
    let routed = reader.bool16("fRouted")?;
    let return_original = reader.bool16("fReturnOrig")?;
    let track_status = reader.bool16("fTrackStatus")?;
    let dirty = reader.u16("fDirty")?;
    if dirty != 0 {
        return Err(corrupted("Metadata fDirty must be zero"));
    }

    let protection = protection_from_raw(reader.u16("nProtect")?)?;
    let stage = nonnegative_i16(reader.i16("iStage")?, "iStage")?;
    let delivery = delivery_from_raw(reader.i16("delOption")?)?;
    let recipient_count = nonnegative_i16(reader.i16("cRecip")?, "cRecip")?;
    if recipient_count == 0 {
        return Err(corrupted(
            "Metadata cRecip must be greater than zero because iStage is an index",
        ));
    }
    let recipient_count = usize::from(recipient_count);
    if usize::from(stage) >= recipient_count {
        return Err(corrupted("Metadata iStage is outside rgRouteSlips"));
    }

    let subject = reader.narrow_short_string("szSubject")?;
    let message = reader.narrow_short_string("szMessage")?;
    let status = reader.narrow_short_string("szStatus")?;
    let title = reader.narrow_short_string("szTitle")?;

    let mut recipients = Vec::with_capacity(recipient_count);
    for index in 0..recipient_count {
        let (recipient, consumed) = Recipient::parse_prefix(&data[reader.offset..], index)?;
        reader.advance(consumed, "Recipient")?;
        recipients.push(recipient);
    }
    reader.finish()?;

    Metadata::try_new(
        routed,
        return_original,
        track_status,
        protection,
        stage,
        delivery,
        subject,
        message,
        status,
        title,
        recipients,
    )
}

/// Serialize one complete Metadata payload.
pub fn to_bytes(route_slip: &Metadata) -> Result<Vec<u8>> {
    route_slip.validate()?;
    let capacity = serialized_len(route_slip)?;
    let mut data = Vec::with_capacity(capacity);
    data.extend_from_slice(&bool16(route_slip.routed).to_le_bytes());
    data.extend_from_slice(&bool16(route_slip.return_original).to_le_bytes());
    data.extend_from_slice(&bool16(route_slip.track_status).to_le_bytes());
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(&route_slip.protection.raw().to_le_bytes());
    data.extend_from_slice(&route_slip.stage.to_le_bytes());
    data.extend_from_slice(&route_slip.delivery.raw().to_le_bytes());
    let recipient_count = i16::try_from(route_slip.recipients.len())
        .map_err(|_| corrupted("Metadata recipient count exceeds i16::MAX"))?;
    data.extend_from_slice(&recipient_count.to_le_bytes());
    append_short_string(&mut data, &route_slip.subject, "szSubject")?;
    append_short_string(&mut data, &route_slip.message, "szMessage")?;
    append_short_string(&mut data, &route_slip.status, "szStatus")?;
    append_short_string(&mut data, &route_slip.title, "szTitle")?;
    for recipient in &route_slip.recipients {
        append_recipient(&mut data, recipient)?;
    }
    Ok(data)
}

impl Metadata {
    /// Parse the optional Metadata table range addressed by a FIB.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Option<Self>> {
        parse(fib, table_stream)
    }

    /// Parse exactly one complete Metadata payload.
    pub fn parse_bytes(data: &[u8]) -> Result<Self> {
        parse_bytes(data)
    }

    /// Serialize one complete Metadata payload.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        to_bytes(self)
    }
}

impl Recipient {
    /// Parse exactly one complete Recipient payload.
    pub fn parse_bytes(data: &[u8]) -> Result<Self> {
        let (value, consumed) = Self::parse_prefix(data, 0)?;
        if consumed != data.len() {
            return Err(corrupted("Recipient has trailing bytes"));
        }
        Ok(value)
    }

    /// Serialize one complete Recipient payload.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.validate()?;
        let capacity = self
            .entry_id
            .len()
            .checked_add(self.name.len())
            .and_then(|length| length.checked_add(4))
            .ok_or_else(|| corrupted("Recipient serialized length overflows"))?;
        let mut data = Vec::with_capacity(capacity);
        append_recipient(&mut data, self)?;
        Ok(data)
    }

    fn parse_prefix(data: &[u8], index: usize) -> Result<(Self, usize)> {
        let mut reader = Reader::new(data);
        let entry_id_length =
            nonnegative_i16(reader.i16("Recipient cbEntryID")?, "Recipient cbEntryID")?;
        let name_length = positive_i16(reader.i16("Recipient cbszName")?, "Recipient cbszName")?;
        let entry_id = reader
            .take(usize::from(entry_id_length), "Recipient rgbEntryId")?
            .to_vec();
        let name = NarrowString::new(
            reader
                .take(usize::from(name_length), "Recipient szName")?
                .to_vec(),
        );
        let value = Self { entry_id, name };
        value
            .validate()
            .map_err(|error| with_info_context(error, index))?;
        Ok((value, reader.offset))
    }
}

fn append_recipient(data: &mut Vec<u8>, recipient: &Recipient) -> Result<()> {
    recipient.validate()?;
    let entry_id_length = i16::try_from(recipient.entry_id.len())
        .map_err(|_| corrupted("Recipient cbEntryID exceeds i16::MAX"))?;
    let name_length = i16::try_from(recipient.name.len())
        .map_err(|_| corrupted("Recipient cbszName exceeds i16::MAX"))?;
    data.extend_from_slice(&entry_id_length.to_le_bytes());
    data.extend_from_slice(&name_length.to_le_bytes());
    data.extend_from_slice(&recipient.entry_id);
    data.extend_from_slice(recipient.name.as_bytes());
    Ok(())
}

fn append_short_string(data: &mut Vec<u8>, value: &NarrowString, field: &str) -> Result<()> {
    validate_short_string(value, field)?;
    let length =
        u16::try_from(value.len()).map_err(|_| corrupted(format!("{field} length exceeds u16")))?;
    data.extend_from_slice(&length.to_le_bytes());
    data.extend_from_slice(value.as_bytes());
    Ok(())
}

fn serialized_len(route_slip: &Metadata) -> Result<usize> {
    let strings = [
        &route_slip.subject,
        &route_slip.message,
        &route_slip.status,
        &route_slip.title,
    ];
    let mut length = ROUTE_SLIP_FIXED_SIZE;
    for string in strings {
        validate_short_string(string, "Metadata string")?;
        length = length
            .checked_add(2)
            .and_then(|value| value.checked_add(string.len()))
            .ok_or_else(|| corrupted("Metadata serialized length overflows"))?;
    }
    for recipient in &route_slip.recipients {
        recipient.validate()?;
        length = length
            .checked_add(4)
            .and_then(|value| value.checked_add(recipient.entry_id.len()))
            .and_then(|value| value.checked_add(recipient.name.len()))
            .ok_or_else(|| corrupted("Metadata serialized length overflows"))?;
    }
    Ok(length)
}

fn protection_from_raw(value: u16) -> Result<Protection> {
    match value {
        0 => Ok(Protection::Off),
        1 => Ok(Protection::RevisionMark),
        2 => Ok(Protection::Annotation),
        3 => Ok(Protection::Form),
        _ => Err(corrupted(format!(
            "Metadata nProtect has invalid value {value}"
        ))),
    }
}

fn delivery_from_raw(value: i16) -> Result<DeliveryOption> {
    match value {
        0 => Ok(DeliveryOption::Serial),
        1 => Ok(DeliveryOption::Parallel),
        _ => Err(corrupted(format!(
            "Metadata delOption has invalid value {value}"
        ))),
    }
}

fn bool16(value: bool) -> u16 {
    u16::from(value)
}

fn nonnegative_i16(value: i16, field: &str) -> Result<u16> {
    u16::try_from(value).map_err(|_| corrupted(format!("{field} must be nonnegative")))
}

fn positive_i16(value: i16, field: &str) -> Result<u16> {
    if value <= 0 {
        return Err(corrupted(format!("{field} must be greater than zero")));
    }
    Ok(value as u16)
}

fn with_info_context(error: PackageError, index: usize) -> PackageError {
    match error {
        PackageError::Corrupted(message) => corrupted(format!("Recipient {index}: {message}")),
        other => other,
    }
}

struct Reader<'a> {
    data: &'a [u8],
    offset: usize,
}

impl<'a> Reader<'a> {
    const fn new(data: &'a [u8]) -> Self {
        Self { data, offset: 0 }
    }

    fn take(&mut self, length: usize, field: &str) -> Result<&'a [u8]> {
        let end = self
            .offset
            .checked_add(length)
            .ok_or_else(|| corrupted(format!("{field} range overflows")))?;
        let bytes = self
            .data
            .get(self.offset..end)
            .ok_or_else(|| corrupted(format!("{field} is truncated")))?;
        self.offset = end;
        Ok(bytes)
    }

    fn advance(&mut self, length: usize, field: &str) -> Result<()> {
        self.take(length, field).map(|_| ())
    }

    fn u16(&mut self, field: &str) -> Result<u16> {
        let bytes = self.take(2, field)?;
        Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
    }

    fn i16(&mut self, field: &str) -> Result<i16> {
        Ok(self.u16(field)? as i16)
    }

    fn bool16(&mut self, field: &str) -> Result<bool> {
        match self.u16(field)? {
            0 => Ok(false),
            1 => Ok(true),
            value => Err(corrupted(format!(
                "{field} has invalid Bool16 value {value}"
            ))),
        }
    }

    fn narrow_short_string(&mut self, field: &str) -> Result<NarrowString> {
        let length = usize::from(self.u16(&format!("{field} length"))?);
        if length > MAX_SHORT_STRING_LENGTH {
            return Err(corrupted(format!(
                "{field} must contain fewer than 256 ANSI bytes"
            )));
        }
        Ok(NarrowString::new(self.take(length, field)?.to_vec()))
    }

    fn finish(&self) -> Result<()> {
        if self.offset != self.data.len() {
            return Err(corrupted("Metadata has trailing bytes"));
        }
        Ok(())
    }
}

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}
