//! Bounded BIFF12 codecs for the shared-workbook owner.

use crate::raw::{Cursor, Kind, Limits, Records, Writer};

use super::model::{
    Guid, Header, Info, MAX_HEADERS, MAX_NAME_UNITS, MAX_PART_BYTES, MAX_RECORDS, MAX_STRING_UNITS,
    RawRecord, RevisionEnvelope, RevisionHeaders, RevisionLog, ShortDateTime, User, UserNames,
};
use super::{invalid, map_raw};
use crate::package::error::Result;

const BEGIN_USERS: u16 = 401;
const COUNT_USERS: u16 = 399;
const USER: u16 = 400;
const INFO: u16 = 398;
const HEADER: u16 = 411;

/// Parse the complete BIFF12 record stream for a user-names part.
pub fn parse_users(data: &[u8]) -> Result<UserNames> {
    let records = parse_records(data, "user-names")?;
    let mut count_slot = None;
    let mut user_slots = Vec::new();
    let mut users = Vec::new();
    let mut begin_slot = None;

    for (index, record) in records.iter().enumerate() {
        match record.kind {
            BEGIN_USERS => {
                if begin_slot.replace(index).is_some() || !record.payload.is_empty() {
                    return Err(invalid("invalid BrtBeginUsers record"));
                }
            },
            COUNT_USERS => {
                if count_slot.replace(index).is_some() {
                    return Err(invalid("duplicate BrtCUsr record"));
                }
                if record.payload.len() != 2 {
                    return Err(invalid("BrtCUsr payload must contain two bytes"));
                }
                let count = u16::from_le_bytes([record.payload[0], record.payload[1]]);
                if usize::from(count) > super::model::MAX_USERS {
                    return Err(invalid("BrtCUsr exceeds the 256-user limit"));
                }
            },
            USER => {
                user_slots.push(index);
                users.push(parse_user(&record.payload)?);
            },
            _ => {},
        }
    }

    let Some(begin_slot) = begin_slot else {
        return Err(invalid("user-names part lacks BrtBeginUsers"));
    };
    let Some(count_slot_value) = count_slot else {
        return Err(invalid("user-names part lacks BrtCUsr"));
    };
    if begin_slot > count_slot_value || user_slots.iter().any(|slot| *slot <= count_slot_value) {
        return Err(invalid("user-names records are out of order"));
    }
    let expected = usize::from(u16::from_le_bytes([
        records[count_slot_value].payload[0],
        records[count_slot_value].payload[1],
    ]));
    if expected != users.len() {
        return Err(invalid(format!(
            "BrtCUsr declares {expected} users but {} BrtUsr records occur",
            users.len()
        )));
    }

    Ok(UserNames {
        relationship_id: String::new(),
        relationship_type: String::new(),
        part_name: String::new(),
        users,
        records,
        count_slot: Some(count_slot_value),
        user_slots,
    })
}

/// Encode a user-names part while retaining unknown records and their order.
pub fn write_users(value: &UserNames) -> Result<Vec<u8>> {
    super::validation::validate_users(value)?;
    let records = if value.records.is_empty() {
        let mut records = Vec::with_capacity(value.users.len().saturating_add(2));
        records.push(RawRecord::new(BEGIN_USERS, Vec::new()));
        records.push(RawRecord::new(
            COUNT_USERS,
            (u16::try_from(value.users.len())
                .map_err(|_| invalid("too many users for BrtCUsr"))?)
            .to_le_bytes()
            .to_vec(),
        ));
        records.extend(
            value
                .users
                .iter()
                .map(|user| encode_user(user).map(|payload| RawRecord::new(USER, payload)))
                .collect::<Result<Vec<_>>>()?,
        );
        records
    } else {
        let Some(count_slot) = value.count_slot else {
            return Err(invalid("source user-names stream lacks BrtCUsr"));
        };
        if value.user_slots.len() > value.records.len() {
            return Err(invalid("source user-names slots are invalid"));
        }
        let mut output = Vec::with_capacity(
            value
                .records
                .len()
                .saturating_add(value.users.len().saturating_sub(value.user_slots.len())),
        );
        let mut user_index = 0usize;
        for (index, record) in value.records.iter().enumerate() {
            if index == count_slot {
                let count = u16::try_from(value.users.len())
                    .map_err(|_| invalid("too many users for BrtCUsr"))?;
                output.push(RawRecord::new(COUNT_USERS, count.to_le_bytes().to_vec()));
            } else if value.user_slots.contains(&index) {
                if let Some(user) = value.users.get(user_index) {
                    output.push(RawRecord::new(USER, encode_user(user)?));
                }
                user_index = user_index.saturating_add(1);
            } else {
                output.push(record.clone());
            }
        }
        while let Some(user) = value.users.get(user_index) {
            output.push(RawRecord::new(USER, encode_user(user)?));
            user_index = user_index.saturating_add(1);
        }
        output
    };
    write_records(&records)
}

/// Parse a revision-headers part, retaining unknown records in source order.
pub fn parse_headers(data: &[u8]) -> Result<RevisionHeaders> {
    let records = parse_records(data, "revision-headers")?;
    let mut info = None;
    let mut info_slot = None;
    let mut headers = Vec::new();
    let mut header_slots = Vec::new();
    for (index, record) in records.iter().enumerate() {
        match record.kind {
            INFO => {
                if info.replace(parse_info(&record.payload)?).is_some() {
                    return Err(invalid("duplicate BrtInfo record"));
                }
                info_slot = Some(index);
            },
            HEADER => {
                headers.push(parse_header(&record.payload)?);
                header_slots.push(index);
            },
            _ => {},
        }
    }
    let Some(info) = info else {
        return Err(invalid("revision-headers part lacks BrtInfo"));
    };
    if headers.is_empty() {
        return Err(invalid("revision-headers part has no BrtRRHeader records"));
    }
    Ok(RevisionHeaders {
        relationship_id: String::new(),
        relationship_type: String::new(),
        part_name: String::new(),
        info,
        headers,
        records,
        info_slot,
        header_slots,
    })
}

/// Encode revision headers while retaining unknown records and their order.
pub fn write_headers(value: &RevisionHeaders) -> Result<Vec<u8>> {
    super::validation::validate_headers(value)?;
    let records = if value.records.is_empty() {
        let mut records = Vec::with_capacity(value.headers.len().saturating_add(1));
        records.push(RawRecord::new(INFO, encode_info(&value.info)?));
        records.extend(
            value
                .headers
                .iter()
                .map(|header| encode_header(header).map(|payload| RawRecord::new(HEADER, payload)))
                .collect::<Result<Vec<_>>>()?,
        );
        records
    } else {
        let Some(info_slot) = value.info_slot else {
            return Err(invalid("source revision-headers stream lacks BrtInfo"));
        };
        let mut output = Vec::with_capacity(
            value
                .records
                .len()
                .saturating_add(value.headers.len().saturating_sub(value.header_slots.len())),
        );
        let mut header_index = 0usize;
        for (index, record) in value.records.iter().enumerate() {
            if index == info_slot {
                output.push(RawRecord::new(INFO, encode_info(&value.info)?));
            } else if value.header_slots.contains(&index) {
                if let Some(header) = value.headers.get(header_index) {
                    output.push(RawRecord::new(HEADER, encode_header(header)?));
                }
                header_index = header_index.saturating_add(1);
            } else {
                output.push(record.clone());
            }
        }
        while let Some(header) = value.headers.get(header_index) {
            output.push(RawRecord::new(HEADER, encode_header(header)?));
            header_index = header_index.saturating_add(1);
        }
        output
    };
    write_records(&records)
}

/// Parse a revision-log part without interpreting or changing its records.
pub fn parse_log(data: &[u8]) -> Result<RevisionLog> {
    Ok(RevisionLog {
        relationship_id: String::new(),
        relationship_type: String::new(),
        part_name: String::new(),
        records: parse_records(data, "revision-log")?,
    })
}

/// Encode an opaque revision-log part in source record order.
pub fn write_log(value: &RevisionLog) -> Result<Vec<u8>> {
    super::validation::validate_log(value)?;
    write_records(&value.records)
}

/// Return the common revision envelope for record kinds that safely begin with
/// `RRd`.  The view is intentionally read-only; the payload remains opaque.
pub(crate) fn revision_envelope(kind: u16, payload: &[u8]) -> Option<RevisionEnvelope> {
    if !is_enveloped_kind(kind) || payload.len() < 14 {
        return None;
    }
    let revision_id = u32::from_le_bytes(payload.get(4..8)?.try_into().ok()?);
    let flags = u16::from_le_bytes(payload.get(10..12)?.try_into().ok()?);
    Some(RevisionEnvelope {
        revision_id,
        revision_type: u16::from_le_bytes(payload.get(8..10)?.try_into().ok()?),
        accepted: flags & 0x0001 != 0,
        undo_action: flags & 0x0002 != 0,
        reserved_one: flags & 0x0004 != 0,
        reserved_two: flags & 0x0008 != 0,
        sheet_id: u16::from_le_bytes(payload.get(12..14)?.try_into().ok()?),
    })
}

fn is_enveloped_kind(kind: u16) -> bool {
    matches!(
        kind,
        771 | 772 | 773 | 774 | 779 | 781 | 782 | 783 | 784 | 785 | 787 | 788
    )
}

fn parse_records(data: &[u8], context: &'static str) -> Result<Vec<RawRecord>> {
    if data.len() > MAX_PART_BYTES {
        return Err(invalid(format!(
            "{context} part exceeds the {} byte bound",
            MAX_PART_BYTES
        )));
    }
    let mut records = Vec::new();
    for record in Records::with_limits(data, Limits::new(MAX_PART_BYTES, MAX_STRING_UNITS)) {
        if records.len() >= MAX_RECORDS {
            return Err(invalid(format!("{context} has too many BIFF12 records")));
        }
        let record = record.map_err(map_raw)?;
        records.push(RawRecord::new(
            record.kind().get(),
            record.payload().to_vec(),
        ));
    }
    Ok(records)
}

fn write_records(records: &[RawRecord]) -> Result<Vec<u8>> {
    if records.len() > MAX_RECORDS {
        return Err(invalid("too many BIFF12 records"));
    }
    let mut output = Vec::new();
    {
        let mut writer = Writer::new(&mut output);
        for record in records {
            let projected = writer
                .get_ref()
                .len()
                .checked_add(record.payload.len())
                .and_then(|length| length.checked_add(6))
                .ok_or_else(|| invalid("BIFF12 part size overflows usize"))?;
            if projected > MAX_PART_BYTES {
                return Err(invalid("BIFF12 part exceeds the configured byte bound"));
            }
            let kind = Kind::new(record.kind).map_err(map_raw)?;
            writer
                .write_record(kind, &record.payload)
                .map_err(map_raw)?;
        }
    }
    Ok(output)
}

fn parse_user(payload: &[u8]) -> Result<User> {
    let mut cursor = Cursor::with_limits(
        payload,
        "BrtUsr",
        Limits::new(MAX_PART_BYTES, MAX_STRING_UNITS),
    );
    let id = cursor.read_u32().map_err(map_raw)?;
    let guid = Guid::from_bytes(
        cursor
            .read_bytes(16)
            .map_err(map_raw)?
            .try_into()
            .map_err(|_| invalid("BrtUsr GUID is not 16 bytes"))?,
    );
    let opened_at = read_date(&mut cursor)?;
    let name = cursor.read_wide_string().map_err(map_raw)?;
    cursor.finish().map_err(map_raw)?;
    Ok(User {
        id,
        guid,
        opened_at,
        name,
    })
}

fn encode_user(user: &User) -> Result<Vec<u8>> {
    let mut payload = Vec::new();
    payload.extend_from_slice(&user.id.to_le_bytes());
    payload.extend_from_slice(&user.guid.as_bytes());
    write_date(&mut payload, user.opened_at);
    write_wide_string(&mut payload, &user.name, MAX_NAME_UNITS)?;
    Ok(payload)
}

fn parse_info(payload: &[u8]) -> Result<Info> {
    let mut cursor = Cursor::new(payload, "BrtInfo");
    let flags = cursor.read_u16().map_err(map_raw)?;
    if flags & !0x000f != 0 || flags & 0x000d != 0x000d {
        return Err(invalid("BrtInfo has invalid reserved flags"));
    }
    let guid = Guid::from_bytes(
        cursor
            .read_bytes(16)
            .map_err(map_raw)?
            .try_into()
            .map_err(|_| invalid("BrtInfo GUID is not 16 bytes"))?,
    );
    let root_guid = Guid::from_bytes(
        cursor
            .read_bytes(16)
            .map_err(map_raw)?
            .try_into()
            .map_err(|_| invalid("BrtInfo root GUID is not 16 bytes"))?,
    );
    let revision_id = cursor.read_u32().map_err(map_raw)?;
    let version = cursor.read_i32().map_err(map_raw)?;
    let history_flags = cursor.read_u16().map_err(map_raw)?;
    if history_flags & !0x0003 != 0 {
        return Err(invalid("BrtInfo has invalid history flags"));
    }
    let revision_history_interval = cursor.read_u16().map_err(map_raw)?;
    cursor.finish().map_err(map_raw)?;
    Ok(Info {
        guid,
        root_guid,
        revision_id,
        version,
        has_revisions: flags & 0x0002 != 0,
        no_revision_history: history_flags & 0x0001 != 0,
        protected: history_flags & 0x0002 != 0,
        revision_history_interval,
    })
}

fn encode_info(info: &Info) -> Result<Vec<u8>> {
    let mut payload = Vec::with_capacity(46);
    let flags = 0x000d_u16 | u16::from(info.has_revisions) << 1;
    payload.extend_from_slice(&flags.to_le_bytes());
    payload.extend_from_slice(&info.guid.as_bytes());
    payload.extend_from_slice(&info.root_guid.as_bytes());
    payload.extend_from_slice(&info.revision_id.to_le_bytes());
    payload.extend_from_slice(&info.version.to_le_bytes());
    let history_flags = u16::from(info.no_revision_history) | (u16::from(info.protected) << 1);
    payload.extend_from_slice(&history_flags.to_le_bytes());
    payload.extend_from_slice(&info.revision_history_interval.to_le_bytes());
    Ok(payload)
}

fn parse_header(payload: &[u8]) -> Result<Header> {
    let mut cursor = Cursor::new(payload, "BrtRRHeader");
    let unused = cursor.read_u32().map_err(map_raw)?;
    let revision_id = cursor.read_u32().map_err(map_raw)?;
    let revision_type = cursor.read_u16().map_err(map_raw)?;
    let flags = cursor.read_u16().map_err(map_raw)?;
    let sheet_id = cursor.read_u16().map_err(map_raw)?;
    if unused != u32::MAX
        || revision_id != 0
        || revision_type != 0x0020
        || flags != 0
        || sheet_id != 0xffff
    {
        return Err(invalid("BrtRRHeader has an invalid RRd envelope"));
    }
    let guid = Guid::from_bytes(
        cursor
            .read_bytes(16)
            .map_err(map_raw)?
            .try_into()
            .map_err(|_| invalid("BrtRRHeader GUID is not 16 bytes"))?,
    );
    let saved_at = read_date(&mut cursor)?;
    let next_sheet_id = cursor.read_u16().map_err(map_raw)?;
    let revision_min = cursor.read_u32().map_err(map_raw)?;
    let revision_max = cursor.read_u32().map_err(map_raw)?;
    let user_name = cursor.read_wide_string().map_err(map_raw)?;
    let relationship_id = cursor.read_wide_string().map_err(map_raw)?;
    let sheet_count = usize::try_from(cursor.read_u32().map_err(map_raw)?)
        .map_err(|_| invalid("BrtRRHeader sheet count overflows usize"))?;
    if sheet_count > MAX_HEADERS {
        return Err(invalid("BrtRRHeader sheet count exceeds the bound"));
    }
    let mut sheet_ids = Vec::with_capacity(sheet_count);
    for _ in 0..sheet_count {
        sheet_ids.push(cursor.read_u16().map_err(map_raw)?);
    }
    let reviewed_count = usize::try_from(cursor.read_u32().map_err(map_raw)?)
        .map_err(|_| invalid("BrtRRHeader reviewed count overflows usize"))?;
    if reviewed_count > MAX_HEADERS {
        return Err(invalid("BrtRRHeader reviewed count exceeds the bound"));
    }
    let mut reviewed = Vec::with_capacity(reviewed_count);
    for _ in 0..reviewed_count {
        reviewed.push(cursor.read_u32().map_err(map_raw)?);
    }
    cursor.finish().map_err(map_raw)?;
    Ok(Header {
        guid,
        saved_at,
        next_sheet_id,
        revision_min,
        revision_max,
        user_name,
        relationship_id,
        sheet_ids,
        reviewed,
    })
}

fn encode_header(header: &Header) -> Result<Vec<u8>> {
    let mut payload = Vec::new();
    payload.extend_from_slice(&u32::MAX.to_le_bytes());
    payload.extend_from_slice(&0_u32.to_le_bytes());
    payload.extend_from_slice(&0x0020_u16.to_le_bytes());
    payload.extend_from_slice(&0_u16.to_le_bytes());
    payload.extend_from_slice(&0xffff_u16.to_le_bytes());
    payload.extend_from_slice(&header.guid.as_bytes());
    write_date(&mut payload, header.saved_at);
    payload.extend_from_slice(&header.next_sheet_id.to_le_bytes());
    payload.extend_from_slice(&header.revision_min.to_le_bytes());
    payload.extend_from_slice(&header.revision_max.to_le_bytes());
    write_wide_string(&mut payload, &header.user_name, MAX_NAME_UNITS)?;
    write_wide_string(&mut payload, &header.relationship_id, MAX_STRING_UNITS)?;
    payload.extend_from_slice(
        &u32::try_from(header.sheet_ids.len())
            .map_err(|_| invalid("too many sheet identifiers"))?
            .to_le_bytes(),
    );
    for sheet_id in &header.sheet_ids {
        payload.extend_from_slice(&sheet_id.to_le_bytes());
    }
    payload.extend_from_slice(
        &u32::try_from(header.reviewed.len())
            .map_err(|_| invalid("too many reviewed revisions"))?
            .to_le_bytes(),
    );
    for revision_id in &header.reviewed {
        payload.extend_from_slice(&revision_id.to_le_bytes());
    }
    Ok(payload)
}

fn read_date(cursor: &mut Cursor<'_>) -> Result<ShortDateTime> {
    Ok(ShortDateTime {
        year: cursor.read_u16().map_err(map_raw)?,
        month: cursor.read_u8().map_err(map_raw)?,
        day: cursor.read_u8().map_err(map_raw)?,
        hour: cursor.read_u8().map_err(map_raw)?,
        minute: cursor.read_u8().map_err(map_raw)?,
        second: cursor.read_u8().map_err(map_raw)?,
        weekday: cursor.read_u8().map_err(map_raw)?,
    })
}

fn write_date(output: &mut Vec<u8>, value: ShortDateTime) {
    output.extend_from_slice(&value.year.to_le_bytes());
    output.extend_from_slice(&[
        value.month,
        value.day,
        value.hour,
        value.minute,
        value.second,
        value.weekday,
    ]);
}

fn write_wide_string(output: &mut Vec<u8>, value: &str, limit: usize) -> Result<()> {
    let units = value.encode_utf16().count();
    if units > limit || units > MAX_STRING_UNITS {
        return Err(invalid(format!(
            "wide string has {units} code units; maximum is {limit}"
        )));
    }
    output.extend_from_slice(
        &u32::try_from(units)
            .map_err(|_| invalid("wide string length overflows BIFF12"))?
            .to_le_bytes(),
    );
    for unit in value.encode_utf16() {
        output.extend_from_slice(&unit.to_le_bytes());
    }
    Ok(())
}
