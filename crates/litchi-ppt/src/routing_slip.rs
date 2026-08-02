//! Inert document routing-slip metadata from legacy PowerPoint files.

use super::package::{PptError, Result};
use super::records::PptRecord;
use crate::consts::PptRecordType;

const RECORD_HEADER_LEN: usize = 8;
const FIXED_PAYLOAD_LEN: usize = 24;
const MIN_TRAILING_UNUSED_LEN: usize = 8;
const FLAG_ONE_AFTER_ANOTHER: u32 = 1 << 0;
const FLAG_RETURN_WHEN_DONE: u32 = 1 << 1;
const FLAG_TRACK_STATUS: u32 = 1 << 2;
const FLAG_DOCUMENT_ROUTED: u32 = 1 << 4;
const FLAG_CYCLE_COMPLETED: u32 = 1 << 5;
const ROUTING_FLAGS_MASK: u32 = FLAG_ONE_AFTER_ANOTHER
    | FLAG_RETURN_WHEN_DONE
    | FLAG_TRACK_STATUS
    | FLAG_DOCUMENT_ROUTED
    | FLAG_CYCLE_COMPLETED;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RoutingSlipText {
    bytes: Vec<u8>,
}

impl RoutingSlipText {
    pub fn from_ansi_bytes(bytes: Vec<u8>) -> Result<Self> {
        if bytes.contains(&0) {
            return Err(corrupted("routing-slip text contains an embedded NUL"));
        }
        if bytes.len() > usize::from(u16::MAX) - 1 {
            return Err(corrupted(
                "routing-slip text exceeds the 16-bit length limit",
            ));
        }
        Ok(Self { bytes })
    }

    pub fn as_ansi_bytes(&self) -> &[u8] {
        &self.bytes
    }

    pub fn to_string_lossy(&self) -> String {
        self.bytes.iter().map(|&byte| char::from(byte)).collect()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RoutingSlipAddress {
    pub text: RoutingSlipText,
    pub trailing_undefined: u8,
}

impl RoutingSlipAddress {
    pub fn new(text: RoutingSlipText) -> Self {
        Self {
            text,
            trailing_undefined: 0,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum RoutingSlipCurrentRecipient {
    OriginatorBeforeRouting,
    Recipient(u32),
    OriginatorAfterRouting,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PowerPointRoutingSlip {
    pub originator: RoutingSlipAddress,
    pub recipients: Vec<RoutingSlipAddress>,
    pub current_recipient: RoutingSlipCurrentRecipient,
    pub subject: RoutingSlipText,
    pub message: RoutingSlipText,
    pub one_after_another: bool,
    pub return_when_done: bool,
    pub track_status: bool,
    pub document_routed: bool,
    pub cycle_completed: bool,
    pub unused1: u32,
    pub unused2: u32,
    pub trailing_undefined: Vec<u8>,
}

impl PowerPointRoutingSlip {
    pub fn new(
        originator: RoutingSlipAddress,
        recipients: Vec<RoutingSlipAddress>,
        subject: RoutingSlipText,
        message: RoutingSlipText,
    ) -> Result<Self> {
        if recipients.len() > u32::MAX as usize {
            return Err(corrupted("routing slip has too many recipients"));
        }
        Ok(Self {
            originator,
            recipients,
            current_recipient: RoutingSlipCurrentRecipient::OriginatorBeforeRouting,
            subject,
            message,
            one_after_another: false,
            return_when_done: false,
            track_status: false,
            document_routed: false,
            cycle_completed: false,
            unused1: 0,
            unused2: 0,
            trailing_undefined: vec![0; MIN_TRAILING_UNUSED_LEN],
        })
    }

    pub fn parse(record: &PptRecord) -> Result<Self> {
        if record.record_type != PptRecordType::DocRoutingSlipAtom
            || record.version != 0
            || record.instance != 0
        {
            return Err(corrupted("DocRoutingSlipAtom has an invalid record header"));
        }
        let data = &record.data;
        require(
            data,
            0,
            FIXED_PAYLOAD_LEN,
            "DocRoutingSlipAtom fixed fields",
        )?;
        let length = usize::try_from(u32_at(data, 0)?)
            .map_err(|_| corrupted("routing-slip length does not fit in memory"))?;
        if length > data.len() || length < RECORD_HEADER_LEN + FIXED_PAYLOAD_LEN {
            return Err(corrupted("DocRoutingSlipAtom has an invalid length"));
        }
        let meaningful_end = length - RECORD_HEADER_LEN;
        if meaningful_end > data.len() || data.len() - meaningful_end < MIN_TRAILING_UNUSED_LEN {
            return Err(corrupted(
                "DocRoutingSlipAtom trailing unused field is truncated",
            ));
        }
        let recipient_count = usize::try_from(u32_at(data, 8)?)
            .map_err(|_| corrupted("routing-slip recipient count does not fit in memory"))?;
        let current_raw = u32_at(data, 12)?;
        let max_current = u32::try_from(recipient_count)
            .ok()
            .and_then(|count| count.checked_add(1))
            .ok_or_else(|| corrupted("routing-slip recipient count overflows"))?;
        if current_raw > max_current {
            return Err(corrupted("routing-slip current recipient is out of range"));
        }
        let flags = u32_at(data, 16)?;
        if flags & !ROUTING_FLAGS_MASK != 0 {
            return Err(corrupted("routing-slip reserved flag bits are not zero"));
        }
        let current_recipient = if current_raw == 0 {
            RoutingSlipCurrentRecipient::OriginatorBeforeRouting
        } else if current_raw == max_current {
            RoutingSlipCurrentRecipient::OriginatorAfterRouting
        } else {
            RoutingSlipCurrentRecipient::Recipient(current_raw)
        };

        let mut offset = FIXED_PAYLOAD_LEN;
        let originator = parse_address(data, &mut offset, meaningful_end, 1)?;
        let minimum_recipient_bytes = recipient_count
            .checked_mul(6)
            .ok_or_else(|| corrupted("routing-slip recipient size overflows"))?;
        if minimum_recipient_bytes > meaningful_end.saturating_sub(offset) {
            return Err(corrupted("routing-slip recipient array is truncated"));
        }
        let mut recipients = Vec::with_capacity(recipient_count);
        for _ in 0..recipient_count {
            recipients.push(parse_address(data, &mut offset, meaningful_end, 2)?);
        }
        let subject = parse_text(data, &mut offset, meaningful_end, 3)?;
        let message = parse_text(data, &mut offset, meaningful_end, 4)?;
        if offset != meaningful_end {
            return Err(corrupted(
                "routing-slip strings do not match the length field",
            ));
        }

        Ok(Self {
            originator,
            recipients,
            current_recipient,
            subject,
            message,
            one_after_another: flags & FLAG_ONE_AFTER_ANOTHER != 0,
            return_when_done: flags & FLAG_RETURN_WHEN_DONE != 0,
            track_status: flags & FLAG_TRACK_STATUS != 0,
            document_routed: flags & FLAG_DOCUMENT_ROUTED != 0,
            cycle_completed: flags & FLAG_CYCLE_COMPLETED != 0,
            unused1: u32_at(data, 4)?,
            unused2: u32_at(data, 20)?,
            trailing_undefined: data[meaningful_end..].to_vec(),
        })
    }

    pub fn to_record(&self) -> Result<PptRecord> {
        let recipient_count = u32::try_from(self.recipients.len())
            .map_err(|_| corrupted("routing slip has too many recipients"))?;
        let current =
            match self.current_recipient {
                RoutingSlipCurrentRecipient::OriginatorBeforeRouting => 0,
                RoutingSlipCurrentRecipient::Recipient(index)
                    if index != 0 && index <= recipient_count =>
                {
                    index
                },
                RoutingSlipCurrentRecipient::Recipient(_) => {
                    return Err(corrupted("routing-slip current recipient is out of range"));
                },
                RoutingSlipCurrentRecipient::OriginatorAfterRouting => recipient_count
                    .checked_add(1)
                    .ok_or_else(|| corrupted("routing-slip recipient count overflows"))?,
            };
        if self.trailing_undefined.len() < MIN_TRAILING_UNUSED_LEN {
            return Err(corrupted("routing-slip trailing unused field is too short"));
        }
        let mut data = vec![0; FIXED_PAYLOAD_LEN];
        data[4..8].copy_from_slice(&self.unused1.to_le_bytes());
        data[8..12].copy_from_slice(&recipient_count.to_le_bytes());
        data[12..16].copy_from_slice(&current.to_le_bytes());
        let flags = if self.one_after_another {
            FLAG_ONE_AFTER_ANOTHER
        } else {
            0
        } | if self.return_when_done {
            FLAG_RETURN_WHEN_DONE
        } else {
            0
        } | if self.track_status {
            FLAG_TRACK_STATUS
        } else {
            0
        } | if self.document_routed {
            FLAG_DOCUMENT_ROUTED
        } else {
            0
        } | if self.cycle_completed {
            FLAG_CYCLE_COMPLETED
        } else {
            0
        };
        data[16..20].copy_from_slice(&flags.to_le_bytes());
        data[20..24].copy_from_slice(&self.unused2.to_le_bytes());
        write_address(&mut data, 1, &self.originator)?;
        for recipient in &self.recipients {
            write_address(&mut data, 2, recipient)?;
        }
        write_text(&mut data, 3, &self.subject)?;
        write_text(&mut data, 4, &self.message)?;
        let length = u32::try_from(RECORD_HEADER_LEN + data.len())
            .map_err(|_| corrupted("routing-slip record is too large"))?;
        data[0..4].copy_from_slice(&length.to_le_bytes());
        data.extend_from_slice(&self.trailing_undefined);
        let data_length =
            u32::try_from(data.len()).map_err(|_| corrupted("routing-slip record is too large"))?;
        Ok(PptRecord {
            record_type: PptRecordType::DocRoutingSlipAtom,
            record_type_raw: 0x0406,
            version: 0,
            instance: 0,
            data_length,
            data,
            children: Vec::new(),
        })
    }
}

fn parse_address(
    data: &[u8],
    offset: &mut usize,
    end: usize,
    kind: u16,
) -> Result<RoutingSlipAddress> {
    let (text, trailing_undefined) = parse_string(data, offset, end, kind, true)?;
    Ok(RoutingSlipAddress {
        text,
        trailing_undefined: trailing_undefined.unwrap_or(0),
    })
}

fn parse_text(data: &[u8], offset: &mut usize, end: usize, kind: u16) -> Result<RoutingSlipText> {
    Ok(parse_string(data, offset, end, kind, false)?.0)
}

fn parse_string(
    data: &[u8],
    offset: &mut usize,
    end: usize,
    expected_kind: u16,
    address: bool,
) -> Result<(RoutingSlipText, Option<u8>)> {
    require_end(data, *offset, 4, end, "routing-slip string header")?;
    let kind = u16_at(data, *offset)?;
    let string_length = usize::from(u16_at(data, *offset + 2)?);
    if kind != expected_kind || (address && string_length == 0) {
        return Err(corrupted(
            "routing-slip string has an invalid type or length",
        ));
    }
    *offset += 4;
    let byte_count = string_length
        .checked_add(1)
        .ok_or_else(|| corrupted("routing-slip string length overflows"))?;
    require_end(data, *offset, byte_count, end, "routing-slip string")?;
    let bytes = &data[*offset..*offset + byte_count];
    let (content, trailing) = if address {
        if bytes[string_length - 1] != 0 {
            return Err(corrupted("routing-slip address is not terminated"));
        }
        (&bytes[..string_length - 1], Some(bytes[string_length]))
    } else {
        if bytes[string_length] != 0 {
            return Err(corrupted("routing-slip text is not terminated"));
        }
        (&bytes[..string_length], None)
    };
    *offset += byte_count;
    Ok((
        RoutingSlipText::from_ansi_bytes(content.to_vec())?,
        trailing,
    ))
}

fn write_address(data: &mut Vec<u8>, kind: u16, value: &RoutingSlipAddress) -> Result<()> {
    let length = value
        .text
        .bytes
        .len()
        .checked_add(1)
        .and_then(|length| u16::try_from(length).ok())
        .ok_or_else(|| corrupted("routing-slip address is too long"))?;
    data.extend_from_slice(&kind.to_le_bytes());
    data.extend_from_slice(&length.to_le_bytes());
    data.extend_from_slice(&value.text.bytes);
    data.push(0);
    data.push(value.trailing_undefined);
    Ok(())
}

fn write_text(data: &mut Vec<u8>, kind: u16, value: &RoutingSlipText) -> Result<()> {
    let length =
        u16::try_from(value.bytes.len()).map_err(|_| corrupted("routing-slip text is too long"))?;
    data.extend_from_slice(&kind.to_le_bytes());
    data.extend_from_slice(&length.to_le_bytes());
    data.extend_from_slice(&value.bytes);
    data.push(0);
    Ok(())
}

fn require(data: &[u8], offset: usize, len: usize, field: &str) -> Result<()> {
    require_end(data, offset, len, data.len(), field)
}

fn require_end(data: &[u8], offset: usize, len: usize, end: usize, field: &str) -> Result<()> {
    if offset
        .checked_add(len)
        .is_none_or(|value| value > end || value > data.len())
    {
        return Err(corrupted(&format!("{field} is truncated")));
    }
    Ok(())
}

fn u16_at(data: &[u8], offset: usize) -> Result<u16> {
    require(data, offset, 2, "routing-slip 16-bit field")?;
    Ok(u16::from_le_bytes([data[offset], data[offset + 1]]))
}

fn u32_at(data: &[u8], offset: usize) -> Result<u32> {
    require(data, offset, 4, "routing-slip 32-bit field")?;
    Ok(u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ]))
}

fn corrupted(message: &str) -> PptError {
    PptError::Corrupted(message.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn text(value: &str) -> RoutingSlipText {
        RoutingSlipText::from_ansi_bytes(value.as_bytes().to_vec()).unwrap()
    }

    #[test]
    fn routing_slip_round_trips_all_typed_state() {
        let mut slip = PowerPointRoutingSlip::new(
            RoutingSlipAddress::new(text("origin")),
            vec![
                RoutingSlipAddress::new(text("first")),
                RoutingSlipAddress::new(text("second")),
            ],
            text("subject"),
            text("message"),
        )
        .unwrap();
        slip.current_recipient = RoutingSlipCurrentRecipient::Recipient(2);
        slip.one_after_another = true;
        slip.return_when_done = true;
        slip.track_status = true;
        slip.document_routed = true;
        slip.cycle_completed = true;
        slip.unused1 = 0x1122_3344;
        slip.unused2 = 0x5566_7788;
        slip.trailing_undefined = vec![0xaa; 11];

        let record = slip.to_record().unwrap();
        assert_eq!(PowerPointRoutingSlip::parse(&record).unwrap(), slip);
    }

    #[test]
    fn routing_slip_rejects_reserved_flags_and_bad_recipients() {
        let slip = PowerPointRoutingSlip::new(
            RoutingSlipAddress::new(text("origin")),
            vec![],
            text(""),
            text(""),
        )
        .unwrap();
        let mut record = slip.to_record().unwrap();
        record.data[16..20].copy_from_slice(&0x08u32.to_le_bytes());
        assert!(PowerPointRoutingSlip::parse(&record).is_err());

        let mut record = slip.to_record().unwrap();
        record.data[12..16].copy_from_slice(&2u32.to_le_bytes());
        assert!(PowerPointRoutingSlip::parse(&record).is_err());
    }
}
