//! Fixed BIFF payload codecs for the `User Names` stream.

use super::model::{UserCheck, UserEntry, UserGuid, UserNames, validate_user_name};
use crate::user_routing::{CUsr, CbUsr, UsrInfo};
use crate::{Error, Result};

pub(crate) const C_USR_RECORD_TYPE: u16 = 401;
pub(crate) const CB_USR_RECORD_TYPE: u16 = 402;
pub(crate) const USR_INFO_RECORD_TYPE: u16 = 403;
pub(crate) const BC_USRS_RECORD_TYPE: u16 = 407;
pub(crate) const USR_CHK_RECORD_TYPE: u16 = 408;
pub(crate) const RECORD_HEADER_LEN: usize = 4;
/// BIFF record data is capped at 8224 bytes (MS-XLS 2.1.4). User Names
/// records are all much smaller, but retaining the stream-wide bound keeps
/// malformed framing from being accepted as a future extension.
pub(crate) const MAX_RECORD_PAYLOAD: usize = 8224;

#[derive(Debug, Clone, Copy)]
pub(crate) struct RecordSpan {
    pub(crate) record_type: u16,
    pub(crate) record_start: usize,
    pub(crate) payload_start: usize,
    pub(crate) payload_end: usize,
}

impl RecordSpan {
    pub(crate) fn payload(self, source: &[u8]) -> &[u8] {
        &source[self.payload_start..self.payload_end]
    }
}

pub(crate) fn frame_record(record_type: u16, payload: &[u8]) -> Result<Vec<u8>> {
    if payload.len() > MAX_RECORD_PAYLOAD {
        return Err(Error::InvalidRecord {
            record_type,
            message: format!(
                "payload has {} bytes; maximum is {MAX_RECORD_PAYLOAD}",
                payload.len()
            ),
        });
    }
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(RECORD_HEADER_LEN + payload.len())
        .map_err(|_error| Error::Allocation("framing User Names record"))?;
    bytes.extend_from_slice(&record_type.to_le_bytes());
    bytes.extend_from_slice(&(payload.len() as u16).to_le_bytes());
    bytes.extend_from_slice(payload);
    Ok(bytes)
}

pub(crate) fn scan_records(source: &[u8], max_stream_bytes: usize) -> Result<Vec<RecordSpan>> {
    if source.len() > max_stream_bytes {
        return Err(Error::UnsafeEdit(format!(
            "User Names stream has {} bytes; maximum is {max_stream_bytes}",
            source.len()
        )));
    }
    let mut records = Vec::new();
    records
        .try_reserve_exact(4 + 255)
        .map_err(|_error| Error::Allocation("retaining User Names record spans"))?;
    let mut offset = 0usize;
    while offset < source.len() {
        if records.len() >= 4 + 255 {
            return Err(Error::InvalidData(
                "User Names stream has more than 255 UsrInfo records".to_string(),
            ));
        }
        let header_end = offset
            .checked_add(RECORD_HEADER_LEN)
            .ok_or(Error::Allocation("framing User Names records"))?;
        let header = source.get(offset..header_end).ok_or_else(|| {
            Error::UnexpectedEndOfStream("truncated User Names record header".to_string())
        })?;
        let record_type = u16::from_le_bytes([header[0], header[1]]);
        let payload_len = usize::from(u16::from_le_bytes([header[2], header[3]]));
        if payload_len > MAX_RECORD_PAYLOAD {
            return Err(Error::InvalidRecord {
                record_type,
                message: format!(
                    "record data has {payload_len} bytes; maximum is {MAX_RECORD_PAYLOAD}"
                ),
            });
        }
        let payload_start = header_end;
        let payload_end = payload_start
            .checked_add(payload_len)
            .ok_or(Error::Allocation("framing User Names record payload"))?;
        if payload_end > source.len() {
            return Err(Error::InvalidRecord {
                record_type,
                message: format!(
                    "payload of {payload_len} bytes is truncated ({} available)",
                    source.len().saturating_sub(payload_start)
                ),
            });
        }
        records.push(RecordSpan {
            record_type,
            record_start: offset,
            payload_start,
            payload_end,
        });
        offset = payload_end;
    }
    Ok(records)
}

pub(crate) fn parse_stream(
    source: &[u8],
    limits: super::model::Limits,
) -> Result<(UserNames, Vec<RecordSpan>)> {
    let limits = limits.validate()?;
    let records = scan_records(source, limits.max_stream_bytes)?;
    if records.len() < 4 {
        return Err(Error::InvalidData(
            "User Names stream must contain CUsr, UsrChk, CbUsr, and BCUsrs".to_string(),
        ));
    }
    let expected = [
        C_USR_RECORD_TYPE,
        USR_CHK_RECORD_TYPE,
        CB_USR_RECORD_TYPE,
        BC_USRS_RECORD_TYPE,
    ];
    for (index, expected_type) in expected.into_iter().enumerate() {
        let found = records[index].record_type;
        if found != expected_type {
            return Err(Error::UnexpectedRecordType {
                expected: expected_type,
                found,
            });
        }
    }

    let cusr = CUsr::parse(records[0].payload(source))?;
    let count = usize::from(cusr.count());
    if count > limits.max_users {
        return Err(Error::UnsafeEdit(format!(
            "User Names stream has {count} users; maximum is {}",
            limits.max_users
        )));
    }
    if records.len() != 4 + count {
        return Err(Error::InvalidData(format!(
            "CUsr declares {count} users but User Names contains {} UsrInfo records",
            records.len().saturating_sub(4)
        )));
    }

    let user_check = parse_user_check(records[1].payload(source))?;
    let cbusr = CbUsr::parse(records[2].payload(source))?;
    let briefcase_user_count = parse_briefcase_count(records[3].payload(source))?;
    let mut users = Vec::new();
    users
        .try_reserve_exact(count)
        .map_err(|_error| Error::Allocation("retaining User Names entries"))?;
    for index in 0..count {
        let span = records[index + 4];
        if span.record_type != USR_INFO_RECORD_TYPE {
            return Err(Error::UnexpectedRecordType {
                expected: USR_INFO_RECORD_TYPE,
                found: span.record_type,
            });
        }
        let payload = span.payload(source);
        let declared_len = usize::from(cbusr.sizes()[index]);
        // CbUsr counts the record-data component. BIFF's four-byte record
        // header is not part of `cb` (MS-XLS 2.1.4), so compare it with the
        // UsrInfo payload rather than the framed record length.
        if declared_len != payload.len() {
            return Err(Error::InvalidData(format!(
                "CbUsr entry {index} declares {declared_len} bytes but UsrInfo has {}",
                payload.len()
            )));
        }
        let user = UserEntry::parse_payload(payload)?;
        if users
            .iter()
            .any(|previous: &UserEntry| previous.user_id == user.user_id)
        {
            return Err(Error::InvalidData(format!(
                "User Names contains duplicate lUsrId {}",
                user.user_id
            )));
        }
        users.push(user);
    }

    Ok((
        UserNames {
            user_check,
            briefcase_user_count,
            user_record_sizes: *cbusr.sizes(),
            users,
        },
        records,
    ))
}

fn parse_user_check(payload: &[u8]) -> Result<UserCheck> {
    if payload.len() != 4 {
        return Err(Error::InvalidLength {
            expected: 4,
            found: payload.len(),
        });
    }
    let version = u16::from_le_bytes([payload[0], payload[1]]);
    if !matches!(version, 0x0200 | 0x0300 | 0x0400 | 0x0500 | 0x0600) {
        return Err(Error::InvalidRecord {
            record_type: USR_CHK_RECORD_TYPE,
            message: format!("UsrChk version 0x{version:04X} is not a BIFF2-BIFF8 value"),
        });
    }
    Ok(UserCheck::new(
        version,
        u16::from_le_bytes([payload[2], payload[3]]),
    ))
}

fn parse_briefcase_count(payload: &[u8]) -> Result<u16> {
    if payload.len() != 2 {
        return Err(Error::InvalidLength {
            expected: 2,
            found: payload.len(),
        });
    }
    Ok(u16::from_le_bytes([payload[0], payload[1]]))
}

impl UserEntry {
    pub(crate) fn parse_payload(payload: &[u8]) -> Result<Self> {
        let parsed = UsrInfo::parse(payload)?;
        let cch = payload.get(28..30).ok_or(Error::InvalidLength {
            expected: 30,
            found: payload.len(),
        })?;
        let character_count = usize::from(u16::from_le_bytes([cch[0], cch[1]]));
        let high_byte = payload[30] & 0x01 != 0;
        let character_bytes = character_count
            .checked_mul(if high_byte { 2 } else { 1 })
            .ok_or(Error::Allocation("sizing UsrInfo characters"))?;
        let unused_offset = 31usize
            .checked_add(character_bytes)
            .ok_or(Error::Allocation("locating UsrInfo unused byte"))?;
        let unused = *payload
            .get(unused_offset)
            .ok_or_else(|| Error::InvalidLength {
                expected: unused_offset + 1,
                found: payload.len(),
            })?;
        Ok(Self {
            user_id: parsed.user_id(),
            guid: *parsed.guid(),
            opened_at: parsed.opened_at(),
            user_name: parsed.user_name().to_owned(),
            string_flags: payload[30],
            unused,
        })
    }

    pub(crate) fn to_payload(&self) -> Result<Vec<u8>> {
        validate_user_name(&self.user_name)?;
        let units: Vec<u16> = self.user_name.encode_utf16().collect();
        let high_byte = self.string_flags & 0x01 != 0 || units.iter().any(|unit| *unit > 0x00FF);
        let string_flags = (self.string_flags & !0x01) | u8::from(high_byte);
        let character_bytes = units
            .len()
            .checked_mul(if high_byte { 2 } else { 1 })
            .ok_or(Error::Allocation("sizing UsrInfo characters"))?;
        let total = 31usize
            .checked_add(character_bytes)
            .and_then(|size| size.checked_add(1))
            .ok_or(Error::Allocation("sizing UsrInfo payload"))?;
        let mut payload = Vec::new();
        payload
            .try_reserve_exact(total)
            .map_err(|_error| Error::Allocation("allocating UsrInfo payload"))?;
        payload.extend_from_slice(&self.user_id.to_le_bytes());
        payload.extend_from_slice(&self.guid);
        payload.extend_from_slice(&self.opened_at.year().to_le_bytes());
        payload.push(self.opened_at.month());
        payload.push(self.opened_at.day());
        payload.push(self.opened_at.hour());
        payload.push(self.opened_at.minute());
        payload.push(self.opened_at.second());
        payload.push(self.opened_at.weekday());
        payload.extend_from_slice(&(units.len() as u16).to_le_bytes());
        payload.push(string_flags);
        if high_byte {
            for unit in units {
                payload.extend_from_slice(&unit.to_le_bytes());
            }
        } else {
            payload.extend(units.into_iter().map(|unit| unit as u8));
        }
        payload.push(self.unused);
        Ok(payload)
    }

    pub(crate) fn guid_matches_any(&self, revision_guids: &[UserGuid]) -> bool {
        revision_guids.iter().any(|guid| guid == &self.guid)
    }
}

pub(crate) fn encode_cusr(count: usize) -> Result<[u8; 2]> {
    let count = u16::try_from(count).map_err(|_error| {
        Error::UnsafeEdit("User Names count does not fit in CUsr.iCount".to_string())
    })?;
    if count > 255 {
        return Err(Error::UnsafeEdit(
            "User Names count exceeds the CUsr.iCount maximum of 255".to_string(),
        ));
    }
    Ok(count.to_le_bytes())
}

pub(crate) fn encode_cbusr(sizes: &[u16; 256]) -> Vec<u8> {
    let mut payload = Vec::with_capacity(512);
    for size in sizes {
        payload.extend_from_slice(&size.to_le_bytes());
    }
    payload
}
