//! Shared native XLS F3 fixture builders.
//!
//! This module contains the canonical archive construction used by both the
//! full performance baseline and the isolated F3 qualification package.  The
//! callers add their own corpus manifests and semantic checks, but neither
//! caller owns a second copy of the workload construction.

use std::{error::Error, io::Cursor};

use litchi_cfb::{OleFile, OleWriter};

const COMMENTS_SOURCE_COUNT: usize = 256;
const COMMENTS_OPAQUE_STREAM_COUNT: usize = 8;
const COMMENTS_OPAQUE_STREAM_BYTES: usize = 2 * 1024 * 1024;
const RK_MULRK_OPAQUE_STREAM_COUNT: usize = 2;
const RK_MULRK_OPAQUE_STREAM_BYTES: usize = 96 * 1024;

fn reserve_exact<T>(values: &mut Vec<T>, amount: usize, label: &str) -> Result<(), Box<dyn Error>> {
    values
        .try_reserve_exact(amount)
        .map_err(|error| format!("{label} allocation failed: {error}").into())
}

fn push_bounded<T>(
    values: &mut Vec<T>,
    value: T,
    limit: usize,
    label: &str,
) -> Result<(), Box<dyn Error>> {
    if values.len() >= limit {
        return Err(format!("{label} exceeded its bounded capacity").into());
    }
    values
        .try_reserve(1)
        .map_err(|error| format!("{label} allocation failed: {error}"))?;
    values.push(value);
    Ok(())
}

fn clone_bytes(bytes: &[u8], label: &str) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut cloned = Vec::new();
    reserve_exact(&mut cloned, bytes.len(), label)?;
    cloned.extend_from_slice(bytes);
    Ok(cloned)
}

fn incompressible_payload(index: usize, length: usize) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(length);
    let mut state = (index as u64)
        .wrapping_mul(0x9e37_79b9_7f4a_7c15)
        .wrapping_add(0xd1b5_4a32_d192_ed03);
    for _ in 0..length {
        state ^= state << 13;
        state ^= state >> 7;
        state ^= state << 17;
        bytes.push((state >> 24) as u8);
    }
    bytes
}

/// Builds the exact comments archive used by the F3 number corpus.
pub(crate) fn build_xls_comments_fixture() -> Result<Vec<u8>, Box<dyn Error>> {
    let mut workbook_builder = litchi_xls::writer::Builder::new();
    let comments = workbook_builder.add_worksheet("Comments")?;
    for index in 0..COMMENTS_SOURCE_COUNT {
        let author = format!("Source {index:03}");
        let text = format!("comment source {index:03}");
        workbook_builder.add_comment(
            comments,
            u32::try_from(index)?,
            1,
            author.as_str(),
            text.as_str(),
        )?;
    }
    let untouched = workbook_builder.add_worksheet("Untouched")?;
    workbook_builder.write_number(untouched, 20, 4, 42.0)?;
    workbook_builder.add_comment(untouched, 4, 3, "Sentinel", "untouched sentinel")?;
    let mut workbook_package = Cursor::new(Vec::new());
    workbook_builder.build()?.write_to(&mut workbook_package)?;
    let mut parsed = OleFile::open(
        Cursor::new(workbook_package.into_inner()),
        litchi_cfb::OleFileLimits::default(),
    )?;
    let workbook_stream = parsed.open_stream(&["Workbook"])?;

    let mut writer = OleWriter::new();
    writer.create_stream_owned(&["Workbook"], workbook_stream)?;
    writer.create_storage(&["OpaquePayloads"])?;
    for index in 0..COMMENTS_OPAQUE_STREAM_COUNT {
        let name = format!("Payload{index:03}");
        writer.create_stream_owned(
            &["OpaquePayloads", name.as_str()],
            incompressible_payload(10_000 + index, COMMENTS_OPAQUE_STREAM_BYTES),
        )?;
    }
    writer.create_stream(
        &["OpaqueMetadata"],
        b"litchi-xls-comments-opaque-metadata-v1",
    )?;
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    Ok(output.into_inner())
}

#[derive(Clone, Debug)]
struct RawRecord {
    kind: u16,
    payload: Vec<u8>,
}

/// Builds the exact packed-RK/MulRK archive used by the F3 numeric corpus.
pub(crate) fn build_rk_mulrk_fixture() -> Result<Vec<u8>, Box<dyn Error>> {
    let mut workbook_builder = litchi_xls::writer::Builder::new();
    let packed = workbook_builder.add_worksheet("Packed")?;
    workbook_builder.write_number(packed, 0, 0, 1.0)?;
    workbook_builder.write_number(packed, 1, 0, 2.0)?;
    workbook_builder.write_number(packed, 1, 1, 3.0)?;
    workbook_builder.write_number(packed, 3, 2, 17.0)?;
    let mut package = Cursor::new(Vec::new());
    workbook_builder.build()?.write_to(&mut package)?;

    let mut source_ole = OleFile::open(
        Cursor::new(package.into_inner()),
        litchi_cfb::OleFileLimits::default(),
    )?;
    let source_workbook = source_ole.open_stream(&["Workbook"])?;
    let rebuilt_workbook = convert_numbers_to_rk_mulrk(&source_workbook)?;

    let mut writer = OleWriter::new();
    writer.create_stream_owned(&["Workbook"], rebuilt_workbook)?;
    writer.create_storage(&["OpaquePayloads"])?;
    for index in 0..RK_MULRK_OPAQUE_STREAM_COUNT {
        let name = format!("Payload{index:03}");
        writer.create_stream_owned(
            &["OpaquePayloads", name.as_str()],
            incompressible_payload(70_000 + index, RK_MULRK_OPAQUE_STREAM_BYTES),
        )?;
    }
    writer.create_stream(
        &["OpaqueMetadata"],
        b"litchi-xls-rk-mulrk-opaque-metadata-v1",
    )?;
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    Ok(output.into_inner())
}

fn clone_record(record: &RawRecord, label: &str) -> Result<RawRecord, Box<dyn Error>> {
    Ok(RawRecord {
        kind: record.kind,
        payload: clone_bytes(&record.payload, label)?,
    })
}

fn convert_numbers_to_rk_mulrk(workbook: &[u8]) -> Result<Vec<u8>, Box<dyn Error>> {
    let records = scan_records(workbook)?;
    let mut number_indices = Vec::new();
    reserve_exact(
        &mut number_indices,
        records.len(),
        "XLS RK/MulRK Number-record index",
    )?;
    for (index, record) in records.iter().enumerate() {
        if record.kind == 0x0203 {
            push_bounded(
                &mut number_indices,
                index,
                records.len(),
                "XLS RK/MulRK Number-record index",
            )?;
        }
    }
    let standalone = *number_indices
        .first()
        .ok_or("XLS RK/MulRK source builder emitted no Number records")?;
    let pair = number_indices
        .windows(2)
        .find(|pair| {
            let left = &records[pair[0]].payload;
            let right = &records[pair[1]].payload;
            pair[1] == pair[0] + 1
                && left.get(0..2) == right.get(0..2)
                && left
                    .get(2..4)
                    .and_then(|bytes| bytes.try_into().ok())
                    .map(u16::from_le_bytes)
                    .zip(
                        right
                            .get(2..4)
                            .and_then(|bytes| bytes.try_into().ok())
                            .map(u16::from_le_bytes),
                    )
                    .is_some_and(|(left, right)| right == left + 1)
                && pair[0] != standalone
        })
        .map(|pair| [pair[0], pair[1]])
        .ok_or("XLS RK/MulRK source builder emitted no contiguous Number pair")?;

    let mut transformed = Vec::new();
    reserve_exact(
        &mut transformed,
        records.len(),
        "XLS transformed BIFF records",
    )?;
    for (index, record) in records.iter().enumerate() {
        if index == pair[1] {
            continue;
        }
        if index == standalone {
            let mut payload = clone_bytes(&record.payload, "XLS RK standalone payload")?;
            payload.truncate(10);
            let value = f64::from_le_bytes(
                record
                    .payload
                    .get(6..14)
                    .ok_or("XLS Number record is truncated")?
                    .try_into()?,
            );
            payload[6..10].copy_from_slice(&encode_integer_rk(value)?.to_le_bytes());
            push_bounded(
                &mut transformed,
                RawRecord {
                    kind: 0x027e,
                    payload,
                },
                records.len(),
                "XLS transformed BIFF records",
            )?;
            continue;
        }
        if index == pair[0] {
            let first = records
                .get(pair[0])
                .ok_or("XLS MulRK first Number disappeared")?;
            let second = records
                .get(pair[1])
                .ok_or("XLS MulRK second Number disappeared")?;
            let first_value = f64::from_le_bytes(first.payload[6..14].try_into()?);
            let second_value = f64::from_le_bytes(second.payload[6..14].try_into()?);
            let mut payload = Vec::new();
            reserve_exact(&mut payload, 18, "XLS MulRK payload")?;
            payload.extend_from_slice(&first.payload[0..4]);
            payload.extend_from_slice(&first.payload[4..6]);
            payload.extend_from_slice(&encode_integer_rk(first_value)?.to_le_bytes());
            payload.extend_from_slice(&second.payload[4..6]);
            payload.extend_from_slice(&encode_integer_rk(second_value)?.to_le_bytes());
            payload.extend_from_slice(
                &u16::from_le_bytes(second.payload[2..4].try_into()?).to_le_bytes(),
            );
            push_bounded(
                &mut transformed,
                RawRecord {
                    kind: 0x00bd,
                    payload,
                },
                records.len(),
                "XLS transformed BIFF records",
            )?;
            continue;
        }
        push_bounded(
            &mut transformed,
            clone_record(record, "XLS transformed BIFF payload")?,
            records.len(),
            "XLS transformed BIFF records",
        )?;
    }

    let mut rebuilt = Vec::new();
    reserve_exact(
        &mut rebuilt,
        workbook.len(),
        "XLS transformed BIFF workbook",
    )?;
    for record in transformed {
        rebuilt.extend_from_slice(&record.kind.to_le_bytes());
        rebuilt.extend_from_slice(
            &u16::try_from(record.payload.len())
                .map_err(|_error| "XLS transformed BIFF payload exceeds u16")?
                .to_le_bytes(),
        );
        rebuilt.extend_from_slice(&record.payload);
    }
    Ok(rebuilt)
}

fn scan_records(workbook: &[u8]) -> Result<Vec<RawRecord>, Box<dyn Error>> {
    let mut records = Vec::new();
    reserve_exact(
        &mut records,
        workbook.len().saturating_div(4).saturating_add(1),
        "XLS BIFF record inventory",
    )?;
    let mut offset = 0usize;
    while offset < workbook.len() {
        let header_end = offset
            .checked_add(4)
            .ok_or("XLS BIFF header offset overflow")?;
        let header = workbook
            .get(offset..header_end)
            .ok_or("XLS BIFF record has a truncated header")?;
        let kind = u16::from_le_bytes([header[0], header[1]]);
        let payload_len = usize::from(u16::from_le_bytes([header[2], header[3]]));
        let end = header_end
            .checked_add(payload_len)
            .ok_or("XLS BIFF record length overflow")?;
        let payload = clone_bytes(
            workbook
                .get(header_end..end)
                .ok_or("XLS BIFF record has a truncated payload")?,
            "XLS BIFF record payload",
        )?;
        push_bounded(
            &mut records,
            RawRecord { kind, payload },
            workbook.len().saturating_div(4).saturating_add(1),
            "XLS BIFF record inventory",
        )?;
        offset = end;
    }
    Ok(records)
}

fn encode_integer_rk(value: f64) -> Result<u32, Box<dyn Error>> {
    let integer = i32::try_from(value as i64).map_err(|_error| "XLS RK integer overflow")?;
    if f64::from(integer) != value || !(-(1_i32 << 29)..(1_i32 << 29)).contains(&integer) {
        return Err("XLS deterministic RK corpus value is not an integer RK".into());
    }
    Ok((integer as u32).wrapping_shl(2) | 0x02)
}
