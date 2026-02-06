//! SoundCollection writer for PowerPoint binary format.
//!
//! PowerPoint's `AnimationInfoAtom.nSoundRef` is matched against `CString` instance 2
//! inside each `Sound` container in the `SoundCollection`. The `SoundData` atom MUST
//! contain actual WAV audio data — PowerPoint extracts and plays it directly.
//! An empty `SoundData` results in silence (no sound played).
//!
//! Per LibreOffice `pptin.cxx` `ReadSound()`:
//! 1. Iterate Sound containers in SoundCollection
//! 2. For each, read CString instance 2 (reference ID string)
//! 3. Compare `OUString::number(nSoundRef) == aRefStr`
//! 4. If matched, extract SoundData and play

use super::records::{PptError, RecordBuilder};
use crate::ole::consts::PptRecordType;
use std::collections::{HashMap, HashSet};

/// Built-in PowerPoint sound names per specification.
fn get_builtin_sound_name(id: u32) -> &'static str {
    match id {
        1 => "Applause",
        2 => "Arrow",
        3 => "Bomb",
        4 => "Breeze",
        5 => "Camera",
        6 => "Cash Register",
        7 => "Chime",
        8 => "Click",
        9 => "Coin",
        10 => "Drum Roll",
        11 => "Explosion",
        12 => "Hammer",
        13 => "Laser",
        14 => "Push",
        15 => "Suction",
        16 => "Swoosh",
        17 => "Type",
        18 => "Voltage",
        19 => "Whoosh",
        20 => "Wind",
        _ => "",
    }
}

/// Base frequency (Hz) for each built-in sound, used to generate a distinguishable tone.
fn get_builtin_sound_freq(id: u32) -> f64 {
    match id {
        1 => 523.25,  // Applause - C5
        2 => 554.37,  // Arrow - C#5
        3 => 587.33,  // Bomb - D5
        4 => 622.25,  // Breeze - D#5
        5 => 659.26,  // Camera - E5
        6 => 698.46,  // Cash Register - F5
        7 => 783.99,  // Chime - G5
        8 => 880.00,  // Click - A5
        9 => 987.77,  // Coin - B5
        10 => 261.63, // Drum Roll - C4
        11 => 293.66, // Explosion - D4
        12 => 329.63, // Hammer - E4
        13 => 392.00, // Laser - G4
        14 => 440.00, // Push - A4
        15 => 493.88, // Suction - B4
        16 => 349.23, // Swoosh - F4
        17 => 369.99, // Type - F#4
        18 => 415.30, // Voltage - G#4
        19 => 466.16, // Whoosh - A#4
        20 => 277.18, // Wind - C#4
        _ => 440.0,
    }
}

/// Generate a minimal valid WAV file (PCM, 8-bit, mono, 8000 Hz) with a short tone.
///
/// Each built-in sound gets a unique frequency so they are distinguishable.
/// Duration: ~0.15 seconds (1200 samples at 8000 Hz).
fn generate_wav_tone(freq: f64) -> Vec<u8> {
    const SAMPLE_RATE: u32 = 8000;
    const NUM_SAMPLES: usize = 1200; // 0.15 seconds
    const BITS_PER_SAMPLE: u16 = 8;
    const NUM_CHANNELS: u16 = 1;
    const BYTE_RATE: u32 = SAMPLE_RATE; // mono 8-bit
    const BLOCK_ALIGN: u16 = 1;

    let data_size = NUM_SAMPLES as u32;
    // RIFF header (12) + fmt chunk (24) + data chunk header (8) + data
    let file_size = 4 + 24 + 8 + data_size; // size after "RIFF" + 4-byte size field

    let mut wav = Vec::with_capacity(12 + 24 + 8 + NUM_SAMPLES);

    // RIFF header
    wav.extend_from_slice(b"RIFF");
    wav.extend_from_slice(&file_size.to_le_bytes());
    wav.extend_from_slice(b"WAVE");

    // fmt sub-chunk (16 bytes of format data)
    wav.extend_from_slice(b"fmt ");
    wav.extend_from_slice(&16u32.to_le_bytes()); // sub-chunk size
    wav.extend_from_slice(&1u16.to_le_bytes()); // audio format = PCM
    wav.extend_from_slice(&NUM_CHANNELS.to_le_bytes());
    wav.extend_from_slice(&SAMPLE_RATE.to_le_bytes());
    wav.extend_from_slice(&BYTE_RATE.to_le_bytes());
    wav.extend_from_slice(&BLOCK_ALIGN.to_le_bytes());
    wav.extend_from_slice(&BITS_PER_SAMPLE.to_le_bytes());

    // data sub-chunk
    wav.extend_from_slice(b"data");
    wav.extend_from_slice(&data_size.to_le_bytes());

    // Generate sine wave samples (8-bit unsigned: 0-255, center=128)
    for i in 0..NUM_SAMPLES {
        let t = i as f64 / SAMPLE_RATE as f64;
        // Apply simple fade-in/fade-out envelope to avoid clicks
        let envelope = if i < 100 {
            i as f64 / 100.0
        } else if i > NUM_SAMPLES - 100 {
            (NUM_SAMPLES - i) as f64 / 100.0
        } else {
            1.0
        };
        let sample = 128.0 + 96.0 * envelope * (2.0 * std::f64::consts::PI * freq * t).sin();
        wav.push(sample.clamp(0.0, 255.0) as u8);
    }

    wav
}

/// Write a UTF-16LE CString atom with the given instance number.
fn write_cstring(instance: u16, text: &str) -> Result<Vec<u8>, PptError> {
    let mut atom = RecordBuilder::new(0x00, instance, 0x0FBA);
    for ch in text.encode_utf16() {
        atom.write_data(&ch.to_le_bytes());
    }
    atom.build()
}

/// Build a Sound container with embedded WAV data.
///
/// Structure per LibreOffice `pptexsoundcollection.cxx` `ExSoundEntry::Write`:
/// - Sound container (0x0F, type 0x07E6)
///   - CString (instance 0): sound name (e.g. "Whoosh")
///   - CString (instance 1): extension (".wav")
///   - CString (instance 2): reference ID string — matched by `AnimationInfoAtom.nSoundRef`
///   - SoundData (type 0x07E7): actual WAV binary data
fn build_sound_container(name: &str, ref_id: u32, wav_data: &[u8]) -> Result<Vec<u8>, PptError> {
    let mut children = Vec::new();

    // CString instance 0 — sound name
    children.extend(write_cstring(0, name)?);

    // CString instance 1 — file extension
    children.extend(write_cstring(1, ".WAV")?);

    // CString instance 2 — reference ID (matched against AnimationInfoAtom.nSoundRef)
    children.extend(write_cstring(2, &ref_id.to_string())?);

    // SoundData atom with actual WAV binary data
    let mut data_atom = RecordBuilder::new(0x00, 0, PptRecordType::SoundData as u16);
    data_atom.write_data(wav_data);
    children.extend(data_atom.build()?);

    // Sound container
    let mut container = RecordBuilder::new(0x0F, 0, PptRecordType::Sound as u16);
    container.write_data(&children);
    container.build()
}

/// Build a SoundCollection container with embedded WAV data for each sound.
///
/// `AnimationInfoAtom.nSoundRef` is the 1-based index into this collection, which is
/// matched against CString instance 2 of each Sound container (the reference ID string).
///
/// Returns `(binary_data, mapping)` where mapping is `builtin_sound_id → collection_ref_id`.
pub fn build_sound_collection(
    sound_ids: &HashSet<u32>,
) -> Result<(Vec<u8>, HashMap<u32, u32>), PptError> {
    if sound_ids.is_empty() {
        return Ok((Vec::new(), HashMap::new()));
    }

    // Collect and sort built-in sound IDs (1-20)
    let mut builtin_ids: Vec<u32> = sound_ids
        .iter()
        .filter(|&&id| id > 0 && id <= 20)
        .copied()
        .collect();

    if builtin_ids.is_empty() {
        return Ok((Vec::new(), HashMap::new()));
    }

    builtin_ids.sort_unstable();

    // Build mapping: original sound ID → 1-based collection ref ID
    let mut id_to_ref = HashMap::new();
    for (idx, &id) in builtin_ids.iter().enumerate() {
        id_to_ref.insert(id, (idx + 1) as u32);
    }

    let mut children = Vec::new();

    // SoundCollectionAtom — contains count of sounds
    let mut count_atom = RecordBuilder::new(0x00, 0, PptRecordType::SoundCollectionAtom as u16);
    count_atom.write_data(&(builtin_ids.len() as u32).to_le_bytes());
    children.extend(count_atom.build()?);

    // Sound container for each built-in sound, WITH actual WAV data
    for (idx, &id) in builtin_ids.iter().enumerate() {
        let ref_id = (idx + 1) as u32;
        let name = get_builtin_sound_name(id);
        let freq = get_builtin_sound_freq(id);
        let wav_data = generate_wav_tone(freq);
        children.extend(build_sound_container(name, ref_id, &wav_data)?);
    }

    // SoundCollection container
    let mut container = RecordBuilder::new(0x0F, 0, PptRecordType::SoundCollection as u16);
    container.write_data(&children);

    Ok((container.build()?, id_to_ref))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_empty_sound_collection() {
        let ids = HashSet::new();
        let (data, mapping) = build_sound_collection(&ids).unwrap();
        assert!(data.is_empty());
        assert!(mapping.is_empty());
    }

    #[test]
    fn test_builtin_sound_collection_has_wav_data() {
        let mut ids = HashSet::new();
        ids.insert(19); // Whoosh

        let (data, mapping) = build_sound_collection(&ids).unwrap();
        assert!(!data.is_empty());
        assert_eq!(mapping.get(&19), Some(&1)); // Whoosh → ref 1

        // Verify WAV RIFF header appears somewhere in the output
        let riff = b"RIFF";
        assert!(
            data.windows(4).any(|w| w == riff),
            "SoundData must contain RIFF WAV data"
        );
    }

    #[test]
    fn test_multiple_sounds_mapping() {
        let mut ids = HashSet::new();
        ids.insert(1); // Applause
        ids.insert(8); // Click
        ids.insert(19); // Whoosh

        let (data, mapping) = build_sound_collection(&ids).unwrap();
        assert!(!data.is_empty());
        // IDs are sorted: 1→ref1, 8→ref2, 19→ref3
        assert_eq!(mapping.get(&1), Some(&1));
        assert_eq!(mapping.get(&8), Some(&2));
        assert_eq!(mapping.get(&19), Some(&3));
    }

    #[test]
    fn test_generate_wav_tone_valid() {
        let wav = generate_wav_tone(440.0);
        assert_eq!(&wav[0..4], b"RIFF");
        assert_eq!(&wav[8..12], b"WAVE");
        assert_eq!(&wav[12..16], b"fmt ");
        // Total: 12 (RIFF header) + 24 (fmt chunk) + 8 (data header) + 1200 (samples)
        assert_eq!(wav.len(), 1244);
    }
}
