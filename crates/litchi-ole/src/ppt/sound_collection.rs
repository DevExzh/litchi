//! Strict, inert PowerPoint embedded sound collection reader.

use super::package::{PptError, Result};
use super::records::PptRecord;
use crate::consts::PptRecordType;
use std::collections::HashSet;

const MAX_SOUNDS: usize = 4_096;
const MAX_SOUND_BYTES: usize = 64 * 1_048_576;
const MAX_AGGREGATE_SOUND_BYTES: usize = 256 * 1_048_576;
const MAX_NAME_UNITS: usize = 1_024;
const MAX_ID_DIGITS: usize = 10;

/// One embedded sound. Its media payload borrows the PowerPoint document stream.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct EmbeddedPowerPointSound<'a> {
    pub id: u32,
    pub name: String,
    pub extension: Option<String>,
    pub builtin_id: Option<String>,
    pub data: &'a [u8],
}

/// A validated `SoundCollectionContainer` in source order.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PowerPointSoundCollection<'a> {
    pub sound_id_seed: u32,
    pub sounds: Vec<EmbeddedPowerPointSound<'a>>,
}

impl<'a> PowerPointSoundCollection<'a> {
    pub fn get(&self, id: u32) -> Option<&EmbeddedPowerPointSound<'a>> {
        self.sounds.iter().find(|sound| sound.id == id)
    }

    pub fn parse(record: &'a PptRecord) -> Result<Self> {
        require_header(
            record.version,
            record.instance,
            record.record_type_raw,
            0x0f,
            5,
            PptRecordType::SoundCollection,
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
            PptRecordType::SoundCollectionAtom,
            "SoundCollectionAtom",
        )?;
        if atom.data.len() != 4 {
            return corrupted("SoundCollectionAtom must contain exactly four bytes");
        }
        let sound_id_seed = u32::from_le_bytes(atom.data.try_into().expect("length checked"));
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
                PptRecordType::Sound,
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
                PptError::Corrupted("embedded sound aggregate size overflow".to_string())
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
        .ok_or_else(|| PptError::Corrupted(format!("{context} header offset overflow")))?;
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
        .map_err(|_| PptError::Corrupted(format!("{context} record size overflow")))?;
    let end = header_end
        .checked_add(length)
        .ok_or_else(|| PptError::Corrupted(format!("{context} record size overflow")))?;
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

fn parse_sound(data: &[u8]) -> Result<EmbeddedPowerPointSound<'_>> {
    let mut offset = 0usize;
    let name = parse_cstring(
        next_record(data, &mut offset, "SoundContainer")?,
        0,
        "SoundNameAtom",
        MAX_NAME_UNITS,
        false,
    )?;
    let mut child = next_record(data, &mut offset, "SoundContainer")?;
    let extension = if child.record_type == PptRecordType::CString as u16 && child.instance == 1 {
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
        .map_err(|_| PptError::Corrupted("SoundIdAtom is outside the u32 range".to_string()))?;
    if id == 0 {
        return corrupted("SoundIdAtom must be positive");
    }

    child = next_record(data, &mut offset, "SoundContainer")?;
    let builtin_id = if child.record_type == PptRecordType::CString as u16 && child.instance == 3 {
        let value = parse_cstring(child, 3, "SoundBuiltinIdAtom", MAX_NAME_UNITS, true)?;
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
        PptRecordType::SoundData,
        "SoundDataBlob",
    )?;
    if child.data.len() > MAX_SOUND_BYTES {
        return corrupted("embedded sound exceeds 64 MiB");
    }
    validate_audio_signature(child.data)?;
    if offset != data.len() {
        return corrupted("SoundContainer has trailing or out-of-order child records");
    }
    Ok(EmbeddedPowerPointSound {
        id,
        name,
        extension,
        builtin_id,
        data: child.data,
    })
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
        PptRecordType::CString,
        context,
    )?;
    if record.data.len() % 2 != 0 {
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
        .map_err(|_| PptError::Corrupted(format!("{context} contains invalid UTF-16")))?;
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
    expected_type: PptRecordType,
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
    Err(PptError::Corrupted(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn record(version: u16, instance: u16, kind: PptRecordType, data: &[u8]) -> Vec<u8> {
        let mut output = Vec::with_capacity(8 + data.len());
        output.extend_from_slice(&(version | (instance << 4)).to_le_bytes());
        output.extend_from_slice(&(kind as u16).to_le_bytes());
        output.extend_from_slice(&(data.len() as u32).to_le_bytes());
        output.extend_from_slice(data);
        output
    }

    fn cstring(instance: u16, value: &str) -> Vec<u8> {
        let data = value
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect::<Vec<_>>();
        record(0, instance, PptRecordType::CString, &data)
    }

    fn sound(id: &str) -> Vec<u8> {
        let mut data = cstring(0, "tone.wav");
        data.extend(cstring(1, ".WAV"));
        data.extend(cstring(2, id));
        data.extend(record(
            0,
            0,
            PptRecordType::SoundData,
            b"RIFF\x04\0\0\0WAVE",
        ));
        record(0x0f, 0, PptRecordType::Sound, &data)
    }

    fn collection(seed: u32, sounds: &[Vec<u8>]) -> PptRecord {
        let mut data = record(
            0,
            0,
            PptRecordType::SoundCollectionAtom,
            &seed.to_le_bytes(),
        );
        for sound in sounds {
            data.extend(sound);
        }
        let bytes = record(0x0f, 5, PptRecordType::SoundCollection, &data);
        PptRecord::parse(&bytes, 0).unwrap().0
    }

    #[test]
    fn parses_valid_sound_without_copying_blob() {
        let record = collection(1, &[sound("1")]);
        let parsed = PowerPointSoundCollection::parse(&record).unwrap();
        assert_eq!(parsed.sounds[0].name, "tone.wav");
        assert_eq!(parsed.sounds[0].extension.as_deref(), Some(".WAV"));
        assert_eq!(parsed.sounds[0].data, b"RIFF\x04\0\0\0WAVE");
    }

    #[test]
    fn rejects_duplicate_noncanonical_and_seed_exceeding_ids() {
        assert!(
            PowerPointSoundCollection::parse(&collection(2, &[sound("1"), sound("1")])).is_err()
        );
        assert!(PowerPointSoundCollection::parse(&collection(1, &[sound("01")])).is_err());
        assert!(PowerPointSoundCollection::parse(&collection(1, &[sound("2")])).is_err());
    }
}
