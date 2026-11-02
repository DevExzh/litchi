//! Exact ZIP-record locality checks of a reassembled candidate (moved from
//! `litchi-numbers`' table-header rewrite, record 0412a): an untouched member
//! keeps its local and central records byte for byte (only the central
//! offset may move); a rewritten member keeps every record field except its
//! CRC, sizes and offset.

/// Whether a central-directory record is unchanged except for its offset.
#[must_use]
pub fn central_record_preserved(source: &[u8], candidate: &[u8]) -> bool {
    const OFFSET: std::ops::Range<usize> = 42..46;
    source.len() == candidate.len()
        && source.len() >= OFFSET.end
        && source[..OFFSET.start] == candidate[..OFFSET.start]
        && source[OFFSET.end..] == candidate[OFFSET.end..]
}

/// Whether an untouched member's records are preserved.
#[must_use]
pub fn package_member_preserved(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    source.raw_name() == candidate.raw_name()
        && source.is_opaque() == candidate.is_opaque()
        && source.raw_record().local_record() == candidate.raw_record().local_record()
        && central_record_preserved(
            source.raw_record().central_directory_record(),
            candidate.raw_record().central_directory_record(),
        )
}

/// Whether a rewritten member keeps every record field but its CRC,
/// sizes and offset.
#[must_use]
pub fn selected_package_member_preserved(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    source.raw_name() == candidate.raw_name()
        && source.is_opaque() == candidate.is_opaque()
        && source.metadata().local() == candidate.metadata().local()
        && source.metadata().central() == candidate.metadata().central()
        && selected_local_record_preserved(source, candidate)
        && selected_central_record_preserved(
            source.raw_record().central_directory_record(),
            candidate.raw_record().central_directory_record(),
        )
}

fn selected_local_record_preserved(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    const CRC_AND_SIZES: std::ops::Range<usize> = 14..26;
    let left = source.raw_record().local_record();
    let right = candidate.raw_record().local_record();
    let (Some(left_header), Some(right_header)) = (
        zip_local_header_length(left),
        zip_local_header_length(right),
    ) else {
        return false;
    };
    if left_header != right_header
        || left[..CRC_AND_SIZES.start] != right[..CRC_AND_SIZES.start]
        || left[CRC_AND_SIZES.end..left_header] != right[CRC_AND_SIZES.end..right_header]
    {
        return false;
    }
    let Some(left_end) = left_header
        .checked_add(source.raw_record().compressed_data().len())
        .filter(|end| *end <= left.len())
    else {
        return false;
    };
    let Some(right_end) = right_header
        .checked_add(candidate.raw_record().compressed_data().len())
        .filter(|end| *end <= right.len())
    else {
        return false;
    };
    selected_local_suffix_preserved(
        source.metadata().local().flags(),
        &left[left_end..],
        &right[right_end..],
    )
}

fn zip_local_header_length(record: &[u8]) -> Option<usize> {
    if record.get(..4)? != b"PK\x03\x04" {
        return None;
    }
    let name = usize::from(u16::from_le_bytes(record.get(26..28)?.try_into().ok()?));
    let extra = usize::from(u16::from_le_bytes(record.get(28..30)?.try_into().ok()?));
    30usize
        .checked_add(name)?
        .checked_add(extra)
        .filter(|length| *length <= record.len())
}

fn selected_local_suffix_preserved(flags: u16, source: &[u8], candidate: &[u8]) -> bool {
    if flags & 0x0008 == 0 {
        return source == candidate;
    }
    let left_prefix = usize::from(source.starts_with(b"PK\x07\x08")) * 4;
    let right_prefix = usize::from(candidate.starts_with(b"PK\x07\x08")) * 4;
    left_prefix == right_prefix
        && source.len() == candidate.len()
        && source.len() >= left_prefix + 12
        && source[..left_prefix] == candidate[..right_prefix]
        && source[left_prefix + 12..] == candidate[right_prefix + 12..]
}

fn selected_central_record_preserved(source: &[u8], candidate: &[u8]) -> bool {
    const CRC_AND_SIZES: std::ops::Range<usize> = 16..28;
    const OFFSET: std::ops::Range<usize> = 42..46;
    source.len() == candidate.len()
        && source.len() >= OFFSET.end
        && source[..CRC_AND_SIZES.start] == candidate[..CRC_AND_SIZES.start]
        && source[CRC_AND_SIZES.end..OFFSET.start] == candidate[CRC_AND_SIZES.end..OFFSET.start]
        && source[OFFSET.end..] == candidate[OFFSET.end..]
}
