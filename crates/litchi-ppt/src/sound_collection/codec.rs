//! MS-PPT sound-collection record codec and validation.

use super::model::{BuiltinId, Collection, Sound};
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;
use std::collections::HashSet;

const MAX_SOUNDS: usize = 4_096;
const MAX_SOUND_BYTES: usize = 64 * 1_048_576;
const MAX_AGGREGATE_SOUND_BYTES: usize = 256 * 1_048_576;
const MAX_NAME_UNITS: usize = 1_024;
const MAX_ID_DIGITS: usize = 10;

impl<'a> Collection<'a> {
    /// Parse one strict `SoundCollectionContainer` without copying sound data.
    pub fn parse(record: &'a Record) -> Result<Self> {
        require_header(
            record.version,
            record.instance,
            record.record_type_raw,
            0x0f,
            5,
            RecordType::SoundCollection,
            "SoundCollectionContainer",
        )?;
        if usize::try_from(record.data_length).ok() != Some(record.data.len()) {
            return corrupted("SoundCollectionContainer has a truncated payload");
        }
        let mut offset = 0usize;
        let atom = next_record(&record.data, &mut offset, "SoundCollectionContainer")?;
        require_header(
            atom.version,
            atom.instance,
            atom.record_type,
            0,
            0,
            RecordType::SoundCollectionAtom,
            "SoundCollectionAtom",
        )?;
        if atom.data.len() != 4 {
            return corrupted("SoundCollectionAtom must contain exactly four bytes");
        }
        let sound_id_seed = u32::from_le_bytes(
            atom.data
                .try_into()
                .map_err(|_| Error::Corrupted("SoundCollectionAtom is truncated".into()))?,
        );
        if sound_id_seed == 0 {
            return corrupted("SoundCollectionAtom soundIdSeed must be positive");
        }

        let mut sounds = Vec::new();
        let mut ids = HashSet::new();
        let mut aggregate = 0usize;
        while offset < record.data.len() {
            if sounds.len() >= MAX_SOUNDS {
                return corrupted(format!("SoundCollection exceeds {MAX_SOUNDS} sounds"));
            }
            let sound = next_record(&record.data, &mut offset, "SoundCollectionContainer")?;
            require_header(
                sound.version,
                sound.instance,
                sound.record_type,
                0x0f,
                0,
                RecordType::Sound,
                "SoundContainer",
            )?;
            let parsed = parse_sound(sound.data)?;
            if parsed.id > sound_id_seed {
                return corrupted(format!(
                    "sound ID {} exceeds soundIdSeed {sound_id_seed}",
                    parsed.id
                ));
            }
            if !ids.insert(parsed.id) {
                return corrupted(format!("duplicate embedded sound ID {}", parsed.id));
            }
            aggregate = aggregate.checked_add(parsed.data.len()).ok_or_else(|| {
                Error::Corrupted("embedded sound aggregate size overflow".to_string())
            })?;
            if aggregate > MAX_AGGREGATE_SOUND_BYTES {
                return corrupted("embedded sounds exceed 256 MiB aggregate");
            }
            sounds.push(parsed);
        }
        Ok(Self {
            sound_id_seed,
            sounds,
        })
    }
}

#[derive(Clone, Copy)]
struct RecordRef<'a> {
    version: u16,
    instance: u16,
    record_type: u16,
    data: &'a [u8],
}

fn next_record<'a>(data: &'a [u8], offset: &mut usize, context: &str) -> Result<RecordRef<'a>> {
    let header_end = offset
        .checked_add(8)
        .ok_or_else(|| Error::Corrupted(format!("{context} header offset overflow")))?;
    if header_end > data.len() {
        return corrupted(format!("truncated record header in {context}"));
    }
    let version_instance = u16::from_le_bytes([data[*offset], data[*offset + 1]]);
    let record_type = u16::from_le_bytes([data[*offset + 2], data[*offset + 3]]);
    let length = u32::from_le_bytes([
        data[*offset + 4],
        data[*offset + 5],
        data[*offset + 6],
        data[*offset + 7],
    ]);
    let length = usize::try_from(length)
        .map_err(|_| Error::Corrupted(format!("{context} record size overflow")))?;
    let end = header_end
        .checked_add(length)
        .ok_or_else(|| Error::Corrupted(format!("{context} record size overflow")))?;
    if end > data.len() {
        return corrupted(format!("record extends beyond {context}"));
    }
    let record = RecordRef {
        version: version_instance & 0x0f,
        instance: version_instance >> 4,
        record_type,
        data: &data[header_end..end],
    };
    *offset = end;
    Ok(record)
}

fn parse_sound(data: &[u8]) -> Result<Sound<'_>> {
    let mut offset = 0usize;
    let name = parse_cstring(
        next_record(data, &mut offset, "SoundContainer")?,
        0,
        "SoundNameAtom",
        MAX_NAME_UNITS,
        false,
    )?;
    let mut child = next_record(data, &mut offset, "SoundContainer")?;
    let extension = if child.record_type == RecordType::CString as u16 && child.instance == 1 {
        let value = parse_cstring(child, 1, "SoundExtensionAtom", 4, false)?;
        if value.encode_utf16().count() != 4 {
            return corrupted("SoundExtensionAtom must contain exactly four UTF-16 code units");
        }
        child = next_record(data, &mut offset, "SoundContainer")?;
        Some(value)
    } else {
        None
    };
    let id_text = parse_cstring(child, 2, "SoundIdAtom", MAX_ID_DIGITS, false)?;
    if id_text.len() > 1 && id_text.starts_with('0')
        || !id_text.bytes().all(|byte| byte.is_ascii_digit())
    {
        return corrupted("SoundIdAtom must be a canonical positive base-10 integer");
    }
    let id = id_text
        .parse::<u32>()
        .map_err(|_| Error::Corrupted("SoundIdAtom is outside the u32 range".to_string()))?;
    if id == 0 {
        return corrupted("SoundIdAtom must be positive");
    }

    child = next_record(data, &mut offset, "SoundContainer")?;
    let builtin_id = if child.record_type == RecordType::CString as u16 && child.instance == 3 {
        let value = parse_cstring(child, 3, "SoundBuiltinIdAtom", 3, false)?;
        if value.len() != 3 || !value.bytes().all(|byte| byte.is_ascii_digit()) {
            return corrupted("SoundBuiltinIdAtom must be a canonical three-digit integer");
        }
        let value = value.parse::<u16>().map_err(|_| {
            Error::Corrupted("SoundBuiltinIdAtom is outside the u16 range".to_string())
        })?;
        let value = parse_builtin_id(value)?;
        child = next_record(data, &mut offset, "SoundContainer")?;
        Some(value)
    } else {
        None
    };
    require_header(
        child.version,
        child.instance,
        child.record_type,
        0,
        0,
        RecordType::SoundData,
        "SoundDataBlob",
    )?;
    if child.data.len() > MAX_SOUND_BYTES {
        return corrupted("embedded sound exceeds 64 MiB");
    }
    validate_audio_signature(child.data)?;
    if offset != data.len() {
        return corrupted("SoundContainer has trailing or out-of-order child records");
    }
    Ok(Sound {
        id,
        name,
        extension,
        builtin_id,
        data: child.data,
    })
}

fn parse_builtin_id(value: u16) -> Result<BuiltinId> {
    match value {
        100 => Ok(BuiltinId::CashRegister),
        101 => Ok(BuiltinId::Typewriter),
        102 => Ok(BuiltinId::ScreechingBrakes),
        103 => Ok(BuiltinId::Whoosh),
        104 => Ok(BuiltinId::Laser),
        105 => Ok(BuiltinId::Camera),
        106 => Ok(BuiltinId::Chime),
        107 => Ok(BuiltinId::Clapping),
        108 => Ok(BuiltinId::Applause),
        109 => Ok(BuiltinId::DriveBy),
        110 => Ok(BuiltinId::DrumRoll),
        111 => Ok(BuiltinId::Explosion),
        112 => Ok(BuiltinId::BreakingGlass),
        113 => Ok(BuiltinId::Gunshot),
        114 => Ok(BuiltinId::SlideProjector),
        115 => Ok(BuiltinId::Ricochet),
        116 => Ok(BuiltinId::Arrow),
        117 => Ok(BuiltinId::Bomb),
        118 => Ok(BuiltinId::Breeze),
        119 => Ok(BuiltinId::Click),
        120 => Ok(BuiltinId::Coin),
        121 => Ok(BuiltinId::Hammer),
        122 => Ok(BuiltinId::Push),
        123 => Ok(BuiltinId::Suction),
        124 => Ok(BuiltinId::Voltage),
        125 => Ok(BuiltinId::Wind),
        _ => corrupted("SoundBuiltinIdAtom is outside the specified value domain"),
    }
}

fn parse_cstring(
    record: RecordRef<'_>,
    instance: u16,
    context: &str,
    max_units: usize,
    allow_empty: bool,
) -> Result<String> {
    require_header(
        record.version,
        record.instance,
        record.record_type,
        0,
        instance,
        RecordType::CString,
        context,
    )?;
    if !record.data.len().is_multiple_of(2) {
        return corrupted(format!("{context} has odd UTF-16 byte length"));
    }
    let units = record.data.len() / 2;
    if units > max_units {
        return corrupted(format!("{context} exceeds {max_units} UTF-16 code units"));
    }
    let values = record
        .data
        .chunks_exact(2)
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
        .collect::<Vec<_>>();
    let value = String::from_utf16(&values)
        .map_err(|_| Error::Corrupted(format!("{context} contains invalid UTF-16")))?;
    if !allow_empty && value.is_empty() {
        return corrupted(format!("{context} cannot be empty"));
    }
    if value.chars().any(char::is_control) {
        return corrupted(format!("{context} contains non-printable characters"));
    }
    Ok(value)
}

fn validate_audio_signature(data: &[u8]) -> Result<()> {
    let wave = data.len() >= 12 && &data[..4] == b"RIFF" && &data[8..12] == b"WAVE";
    let aiff =
        data.len() >= 12 && &data[..4] == b"FORM" && matches!(&data[8..12], b"AIFF" | b"AIFC");
    if !wave && !aiff {
        return corrupted("SoundDataBlob is not a RIFF/WAVE or FORM/AIFF container");
    }
    Ok(())
}

fn require_header(
    version: u16,
    instance: u16,
    record_type: u16,
    expected_version: u16,
    expected_instance: u16,
    expected_type: RecordType,
    context: &str,
) -> Result<()> {
    if version != expected_version
        || instance != expected_instance
        || record_type != expected_type as u16
    {
        return corrupted(format!("invalid {context} record header"));
    }
    Ok(())
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Corrupted(message.into()))
}
