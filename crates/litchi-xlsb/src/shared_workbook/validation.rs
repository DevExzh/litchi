//! Semantic and package-independent invariants for shared-workbook metadata.

use std::collections::HashSet;

use crate::raw::Kind;

use super::model::{
    Catalog, Header, Info, MAX_HEADERS, MAX_NAME_UNITS, MAX_PART_BYTES, MAX_RECORDS,
    MAX_STRING_UNITS, RawRecord, RevisionHeaders, RevisionLog, ShortDateTime, User, UserNames,
};
use super::{Result, invalid};

/// Validate the complete shared-workbook graph.
pub fn validate_catalog(catalog: &Catalog) -> Result<()> {
    match (&catalog.users, &catalog.headers) {
        (None, None) if catalog.logs.is_empty() => return Ok(()),
        (Some(_), Some(_)) => {},
        _ => {
            return Err(invalid(
                "shared-workbook users and headers must occur together",
            ));
        },
    }
    let users = catalog
        .users
        .as_ref()
        .ok_or_else(|| invalid("shared-workbook users are missing"))?;
    let headers = catalog
        .headers
        .as_ref()
        .ok_or_else(|| invalid("shared-workbook headers are missing"))?;
    validate_users(users)?;
    validate_headers(headers)?;
    if catalog.logs.len() != headers.headers.len() {
        return Err(invalid("revision log/header count mismatch"));
    }

    let header_guids: HashSet<_> = headers.headers.iter().map(|header| header.guid).collect();
    for user in &users.users {
        if !header_guids.contains(&user.guid) {
            return Err(invalid("BrtUsr GUID does not identify a revision header"));
        }
    }
    if headers.info.guid
        != headers
            .headers
            .last()
            .map(|header| header.guid)
            .unwrap_or(headers.info.guid)
    {
        return Err(invalid("BrtInfo GUID does not identify the latest header"));
    }
    if !header_guids.contains(&headers.info.root_guid) {
        return Err(invalid("BrtInfo root GUID does not identify a header"));
    }

    let relationships: HashSet<_> = headers
        .headers
        .iter()
        .map(|header| header.relationship_id.as_str())
        .collect();
    let mut log_relationships = HashSet::new();
    let mut log_parts = HashSet::new();
    for (index, log) in catalog.logs.iter().enumerate() {
        validate_log(log)?;
        if headers
            .headers
            .get(index)
            .is_none_or(|header| header.relationship_id != log.relationship_id)
            || !relationships.contains(log.relationship_id.as_str())
            || !log_relationships.insert(log.relationship_id.as_str())
            || !log_parts.insert(log.part_name.as_str())
        {
            return Err(invalid("revision log identity does not match headers"));
        }
    }
    Ok(())
}

/// Validate user metadata independently of package relationships.
pub fn validate_users(users: &UserNames) -> Result<()> {
    if users.users.len() > super::model::MAX_USERS {
        return Err(invalid("BrtCUsr exceeds the 256-user limit"));
    }
    validate_part_identity(&users.relationship_id, &users.part_name, "user-names")?;
    let mut ids = HashSet::new();
    let mut guids = HashSet::new();
    for user in &users.users {
        validate_user(user)?;
        if !ids.insert(user.id) || !guids.insert(user.guid) {
            return Err(invalid("BrtUsr identifiers and GUIDs must be unique"));
        }
    }
    if users.records.len() > MAX_RECORDS {
        return Err(invalid("user-names record count exceeds the bound"));
    }
    validate_raw_records(&users.records)
}

/// Validate revision-header metadata independently of package relationships.
pub fn validate_headers(headers: &RevisionHeaders) -> Result<()> {
    if headers.headers.is_empty() || headers.headers.len() > MAX_HEADERS {
        return Err(invalid(
            "revision-header count is outside the supported range",
        ));
    }
    validate_part_identity(
        &headers.relationship_id,
        &headers.part_name,
        "revision-headers",
    )?;
    validate_info(&headers.info)?;
    let mut guids = HashSet::new();
    let mut relationships = HashSet::new();
    for header in &headers.headers {
        validate_header(header)?;
        if !guids.insert(header.guid) || !relationships.insert(header.relationship_id.as_str()) {
            return Err(invalid(
                "revision-header GUIDs and relationship IDs must be unique",
            ));
        }
    }
    if headers.records.len() > MAX_RECORDS {
        return Err(invalid("revision-header record count exceeds the bound"));
    }
    validate_raw_records(&headers.records)
}

/// Validate one opaque revision-log part.
pub fn validate_log(log: &RevisionLog) -> Result<()> {
    validate_part_identity(&log.relationship_id, &log.part_name, "revision-log")?;
    if log.records.len() > MAX_RECORDS {
        return Err(invalid("revision-log record count exceeds the bound"));
    }
    validate_raw_records(&log.records)
}

pub(crate) fn validate_local(catalog: &Catalog) -> Result<()> {
    if let Some(users) = &catalog.users {
        validate_users_edit(users)?;
    }
    if let Some(headers) = &catalog.headers {
        validate_headers_edit(headers)?;
    }
    for log in &catalog.logs {
        validate_log_edit(log)?;
    }
    Ok(())
}

fn validate_users_edit(users: &UserNames) -> Result<()> {
    if users.users.len() > super::model::MAX_USERS {
        return Err(invalid("BrtCUsr exceeds the 256-user limit"));
    }
    if !users.relationship_id.is_empty() || !users.part_name.is_empty() {
        validate_part_identity(&users.relationship_id, &users.part_name, "user-names")?;
    }
    let mut ids = HashSet::new();
    let mut guids = HashSet::new();
    for user in &users.users {
        validate_user(user)?;
        if !ids.insert(user.id) || !guids.insert(user.guid) {
            return Err(invalid("BrtUsr identifiers and GUIDs must be unique"));
        }
    }
    validate_raw_records(&users.records)
}

fn validate_headers_edit(headers: &RevisionHeaders) -> Result<()> {
    if headers.headers.len() > MAX_HEADERS {
        return Err(invalid("revision-header count exceeds the bound"));
    }
    if !headers.relationship_id.is_empty() || !headers.part_name.is_empty() {
        validate_part_identity(
            &headers.relationship_id,
            &headers.part_name,
            "revision-headers",
        )?;
    }
    validate_info(&headers.info)?;
    let mut guids = HashSet::new();
    let mut relationships = HashSet::new();
    for header in &headers.headers {
        validate_header(header)?;
        if !guids.insert(header.guid) || !relationships.insert(header.relationship_id.as_str()) {
            return Err(invalid(
                "revision-header GUIDs and relationship IDs must be unique",
            ));
        }
    }
    validate_raw_records(&headers.records)
}

fn validate_log_edit(log: &RevisionLog) -> Result<()> {
    if (!log.relationship_id.is_empty() || !log.part_name.is_empty())
        && (log.relationship_id.is_empty() || log.part_name.is_empty())
    {
        return Err(invalid("revision-log identity must be complete"));
    }
    validate_log_records(&log.records)
}

fn validate_log_records(records: &[RawRecord]) -> Result<()> {
    if records.len() > MAX_RECORDS {
        return Err(invalid("revision-log record count exceeds the bound"));
    }
    validate_raw_records(records)
}

fn validate_raw_records(records: &[RawRecord]) -> Result<()> {
    for record in records {
        if record.kind > Kind::MAX || record.payload.len() > MAX_PART_BYTES {
            return Err(invalid("opaque BIFF12 record exceeds the supported bounds"));
        }
    }
    Ok(())
}

fn validate_info(info: &Info) -> Result<()> {
    if info.version < 1 {
        return Err(invalid("BrtInfo version must be at least one"));
    }
    if info.revision_history_interval > 32_767
        || (info.revision_history_interval == 0 && !info.no_revision_history)
    {
        return Err(invalid("BrtInfo revision-history interval is invalid"));
    }
    Ok(())
}

fn validate_user(user: &User) -> Result<()> {
    validate_date(user.opened_at)?;
    validate_name(&user.name, "BrtUsr name")
}

fn validate_header(header: &Header) -> Result<()> {
    validate_date(header.saved_at)?;
    validate_name(&header.user_name, "BrtRRHeader user")?;
    validate_string(
        &header.relationship_id,
        MAX_STRING_UNITS,
        "BrtRRHeader relationship ID",
    )?;
    if header.relationship_id.is_empty() {
        return Err(invalid("BrtRRHeader relationship ID must not be empty"));
    }
    if header.sheet_ids.is_empty() || header.sheet_ids.len() > 65_535 {
        return Err(invalid("BrtRRHeader sheet count must be 1..65535"));
    }
    let mut sheets = HashSet::new();
    if header.sheet_ids.iter().any(|sheet| !sheets.insert(*sheet)) {
        return Err(invalid("BrtRRHeader sheet identifiers must be unique"));
    }
    match (header.revision_min, header.revision_max) {
        (0, 0) => {
            if !header.reviewed.is_empty() {
                return Err(invalid("reviewed revisions require a non-empty range"));
            }
        },
        (min, max) if min > 0 && max >= min && max < u32::MAX => {
            let span = u64::from(max) - u64::from(min) + 1;
            if u64::try_from(header.reviewed.len()).map_or(true, |count| count > span) {
                return Err(invalid("too many reviewed revisions"));
            }
            let mut reviewed = HashSet::new();
            for revision in &header.reviewed {
                if *revision < min || *revision > max || !reviewed.insert(*revision) {
                    return Err(invalid(
                        "reviewed revision is outside or duplicated in its range",
                    ));
                }
            }
        },
        _ => return Err(invalid("BrtRRHeader revision range is invalid")),
    }
    Ok(())
}

fn validate_part_identity(relationship_id: &str, part_name: &str, label: &str) -> Result<()> {
    if relationship_id.is_empty() || part_name.is_empty() {
        return Err(invalid(format!("{label} package identity is incomplete")));
    }
    validate_string(relationship_id, MAX_STRING_UNITS, "relationship ID")
}

fn validate_name(value: &str, label: &str) -> Result<()> {
    validate_string(value, MAX_NAME_UNITS, label)?;
    if value.is_empty() {
        return Err(invalid(format!("{label} must not be empty")));
    }
    Ok(())
}

fn validate_string(value: &str, limit: usize, label: &str) -> Result<()> {
    let units = value.encode_utf16().count();
    if units > limit || units > MAX_STRING_UNITS {
        return Err(invalid(format!("{label} exceeds its UTF-16 length bound")));
    }
    Ok(())
}

fn validate_date(value: ShortDateTime) -> Result<()> {
    if !(1900..=9999).contains(&value.year)
        || !(1..=12).contains(&value.month)
        || !(0..=23).contains(&value.hour)
        || !(0..=59).contains(&value.minute)
        || !(0..=59).contains(&value.second)
        || !(1..=7).contains(&value.weekday)
    {
        return Err(invalid("ShortDtr scalar is outside its specified range"));
    }
    let max_day = days_in_month(value.year, value.month);
    if value.day == 0 || value.day > max_day {
        return Err(invalid("ShortDtr day is inconsistent with its month"));
    }
    if weekday(value.year, value.month, value.day) != value.weekday {
        return Err(invalid("ShortDtr weekday is inconsistent with its date"));
    }
    Ok(())
}

fn days_in_month(year: u16, month: u8) -> u8 {
    match month {
        2 if year % 4 == 0 && (year % 100 != 0 || year % 400 == 0) => 29,
        2 => 28,
        4 | 6 | 9 | 11 => 30,
        _ => 31,
    }
}

fn weekday(year: u16, month: u8, day: u8) -> u8 {
    // Sakamoto's bounded Gregorian calculation, normalized to Monday=1.
    let mut year = i32::from(year);
    let month = usize::from(month);
    const TABLE: [i32; 12] = [0, 3, 2, 5, 0, 3, 5, 1, 4, 6, 2, 4];
    if month < 3 {
        year -= 1;
    }
    let sunday_zero =
        (year + year / 4 - year / 100 + year / 400 + TABLE[month - 1] + i32::from(day))
            .rem_euclid(7);
    u8::try_from((sunday_zero + 6).rem_euclid(7) + 1).unwrap_or(1)
}
