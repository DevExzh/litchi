#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::{BuiltinId, Collection};
use crate::consts::RecordType;
use crate::records::Record;

fn record(version: u16, instance: u16, kind: RecordType, data: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(8 + data.len());
    output.extend_from_slice(&(version | (instance << 4)).to_le_bytes());
    output.extend_from_slice(&(kind as u16).to_le_bytes());
    output.extend_from_slice(&u32::try_from(data.len()).unwrap().to_le_bytes());
    output.extend_from_slice(data);
    output
}

fn cstring(instance: u16, value: &str) -> Vec<u8> {
    let data = value
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect::<Vec<_>>();
    record(0, instance, RecordType::CString, &data)
}

fn sound(id: &str) -> Vec<u8> {
    let mut data = cstring(0, "tone.wav");
    data.extend(cstring(1, ".WAV"));
    data.extend(cstring(2, id));
    data.extend(record(0, 0, RecordType::SoundData, b"RIFF\x04\0\0\0WAVE"));
    record(0x0f, 0, RecordType::Sound, &data)
}

fn sound_with_builtin(id: &str, builtin: &str) -> Vec<u8> {
    let mut data = cstring(0, "tone.wav");
    data.extend(cstring(1, ".WAV"));
    data.extend(cstring(2, id));
    data.extend(cstring(3, builtin));
    data.extend(record(0, 0, RecordType::SoundData, b"RIFF\x04\0\0\0WAVE"));
    record(0x0f, 0, RecordType::Sound, &data)
}

fn collection(seed: u32, sounds: &[Vec<u8>]) -> Record {
    let mut data = record(0, 0, RecordType::SoundCollectionAtom, &seed.to_le_bytes());
    for sound in sounds {
        data.extend(sound);
    }
    let bytes = record(0x0f, 5, RecordType::SoundCollection, &data);
    Record::parse(&bytes, 0).unwrap().0
}

#[test]
fn parses_valid_sound_without_copying_blob() {
    let record = collection(1, &[sound("1")]);
    let parsed = Collection::parse(&record).unwrap();
    assert_eq!(parsed.sounds[0].name, "tone.wav");
    assert_eq!(parsed.sounds[0].extension.as_deref(), Some(".WAV"));
    assert_eq!(parsed.sounds[0].data, b"RIFF\x04\0\0\0WAVE");
}

#[test]
fn rejects_duplicate_noncanonical_and_seed_exceeding_ids() {
    assert!(Collection::parse(&collection(2, &[sound("1"), sound("1")])).is_err());
    assert!(Collection::parse(&collection(1, &[sound("01")])).is_err());
    assert!(Collection::parse(&collection(1, &[sound("2")])).is_err());
    assert!(Collection::parse(&collection(1, &[sound_with_builtin("1", "099")])).is_err());
    let valid_builtin = collection(1, &[sound_with_builtin("1", "119")]);
    let parsed = Collection::parse(&valid_builtin).unwrap();
    assert_eq!(parsed.sounds[0].builtin_id, Some(BuiltinId::Click));
}
