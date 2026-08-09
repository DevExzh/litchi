//! SoundCollection writer for PowerPoint binary format.
//!
//! PowerPoint sound references are matched against `CString` instance 2
//! inside each `Sound` container in the `SoundCollection`. The `SoundData` atom MUST
//! contain actual WAV or AIFF audio data — PowerPoint extracts and plays it directly.
//! An empty `SoundData` results in silence (no sound played).
//!
//! Per LibreOffice `pptin.cxx` `ReadSound()`:
//! 1. Iterate Sound containers in SoundCollection
//! 2. For each, read CString instance 2 (reference ID string)
//! 3. Compare `OUString::number(nSoundRef) == aRefStr`
//! 4. If matched, extract SoundData and play

use super::records::{Error, RecordBuilder};
use crate::animation::{BuiltinSound, SoundType};
use crate::consts::RecordType;
use std::collections::{BTreeMap, HashMap, HashSet};

/// Resource limits for sound collection authoring.
#[allow(
    clippy::struct_field_names,
    reason = "the shared `max_` prefix documents that every field is a resource bound; renaming would churn sibling modules that reference these fields"
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct SoundCollectionLimits {
    pub max_sounds: usize,
    pub max_sound_bytes: usize,
    pub max_aggregate_sound_bytes: usize,
    pub max_name_units: usize,
}

impl Default for SoundCollectionLimits {
    fn default() -> Self {
        Self {
            max_sounds: 4_096,
            max_sound_bytes: 64 * 1_048_576,
            max_aggregate_sound_bytes: 256 * 1_048_576,
            max_name_units: 1_024,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum PlannedSound<'a> {
    Builtin(BuiltinSound),
    Embedded { name: &'a str, data: &'a [u8] },
}

/// Deterministic, bounded sound-resource planner.
///
/// Input IDs are writer-local handles. `build` maps them to the dense positive
/// identifiers required by `SoundCollectionContainer`.
pub(crate) struct SoundCollectionBuilder<'a> {
    sounds: BTreeMap<u32, PlannedSound<'a>>,
    aggregate_sound_bytes: usize,
    limits: SoundCollectionLimits,
}

impl<'a> SoundCollectionBuilder<'a> {
    pub(crate) fn new(limits: SoundCollectionLimits) -> Self {
        Self {
            sounds: BTreeMap::new(),
            aggregate_sound_bytes: 0,
            limits,
        }
    }

    /// Register an explicit built-in or embedded resource for a writer-local ID.
    pub(crate) fn register(
        &mut self,
        source_id: u32,
        sound_type: &'a SoundType,
    ) -> Result<(), Error> {
        let sound = match sound_type {
            SoundType::Builtin(sound) => PlannedSound::Builtin(*sound),
            SoundType::Embedded { name, data } => {
                self.validate_name(name)?;
                self.validate_audio(data)?;
                PlannedSound::Embedded { name, data }
            },
            SoundType::Linked { .. } => {
                return invalid(
                    "linked animation sounds require external-media records and cannot be embedded in SoundCollection",
                );
            },
        };
        self.insert(source_id, sound)
    }

    /// Register a reference without an explicit resource.
    ///
    /// IDs 1 through 20 use the library's typed built-in sound catalog. Other
    /// IDs must already have an explicit embedded resource.
    pub(crate) fn register_reference(&mut self, source_id: u32) -> Result<(), Error> {
        if source_id == 0 {
            return Ok(());
        }
        if self.sounds.contains_key(&source_id) {
            return Ok(());
        }
        let sound = BuiltinSound::from_id(source_id).ok_or_else(|| {
            std::io::Error::new(
                std::io::ErrorKind::InvalidInput,
                format!("sound reference {source_id} has no registered embedded resource"),
            )
        })?;
        self.insert(source_id, PlannedSound::Builtin(sound))
    }

    /// Serialize the collection and return writer-local ID to file-ID mapping.
    pub(crate) fn build(self) -> Result<(Vec<u8>, HashMap<u32, u32>), Error> {
        if self.sounds.is_empty() {
            return Ok((Vec::new(), HashMap::new()));
        }
        let seed = u32::try_from(self.sounds.len())
            .map_err(|_err| invalid_error("sound collection count exceeds u32"))?;
        let mut mapping = HashMap::with_capacity(self.sounds.len());
        let mut children = Vec::new();
        let mut seed_atom = RecordBuilder::new(0x00, 0, RecordType::SoundCollectionAtom as u16);
        seed_atom.write_data(&seed.to_le_bytes());
        children.extend(seed_atom.build()?);

        for (index, (source_id, sound)) in self.sounds.into_iter().enumerate() {
            let file_id = u32::try_from(index + 1)
                .map_err(|_err| invalid_error("sound identifier exceeds u32"))?;
            mapping.insert(source_id, file_id);
            match sound {
                PlannedSound::Builtin(builtin) => {
                    let wav_data = generate_wav_tone(get_builtin_sound_freq(builtin.id()));
                    children.extend(build_sound_container(
                        builtin.name(),
                        ".WAV",
                        file_id,
                        builtin_description_id(builtin),
                        &wav_data,
                    )?);
                },
                PlannedSound::Embedded { name, data } => {
                    children.extend(build_sound_container(
                        name,
                        audio_extension(data)?,
                        file_id,
                        None,
                        data,
                    )?);
                },
            }
        }

        let mut container = RecordBuilder::new(0x0F, 5, RecordType::SoundCollection as u16);
        container.write_data(&children);
        Ok((container.build()?, mapping))
    }

    fn insert(&mut self, source_id: u32, sound: PlannedSound<'a>) -> Result<(), Error> {
        if source_id == 0 {
            return invalid("sound resource IDs must be positive");
        }
        if let Some(existing) = self.sounds.get(&source_id) {
            if existing != &sound {
                return invalid(format!(
                    "sound resource ID {source_id} has conflicting definitions"
                ));
            }
            return Ok(());
        }
        if self.sounds.len() >= self.limits.max_sounds {
            return invalid("sound collection exceeds the configured sound count");
        }
        if let PlannedSound::Embedded { data, .. } = sound {
            self.aggregate_sound_bytes = self
                .aggregate_sound_bytes
                .checked_add(data.len())
                .ok_or_else(|| invalid_error("sound collection byte count overflows"))?;
            if self.aggregate_sound_bytes > self.limits.max_aggregate_sound_bytes {
                return invalid("sound collection exceeds the configured aggregate byte limit");
            }
        }
        self.sounds.insert(source_id, sound);
        Ok(())
    }

    fn validate_name(&self, name: &str) -> Result<(), Error> {
        let units = name.encode_utf16().count();
        if units == 0 || units > self.limits.max_name_units {
            return invalid("embedded sound name is empty or exceeds the configured limit");
        }
        if name.chars().any(char::is_control) {
            return invalid("embedded sound name contains a control character");
        }
        Ok(())
    }

    fn validate_audio(&self, data: &[u8]) -> Result<(), Error> {
        if data.len() > self.limits.max_sound_bytes {
            return invalid("embedded sound exceeds the configured byte limit");
        }
        let _ = audio_extension(data)?;
        Ok(())
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
        15 => 493.88, // Suction - B4
        16 => 349.23, // Swoosh - F4
        17 => 369.99, // Type - F#4
        18 => 415.30, // Voltage - G#4
        19 => 466.16, // Whoosh - A#4
        20 => 277.18, // Wind - C#4
        // 14 (Push) and any unknown ID fall back to A4
        _ => 440.0,
    }
}

/// Generate a minimal valid WAV file (PCM, 8-bit, mono, 8000 Hz) with a short tone.
///
/// Each built-in sound gets a unique frequency so they are distinguishable.
/// Duration: ~0.15 seconds (1200 samples at 8000 Hz).
#[allow(
    clippy::cast_possible_truncation,
    clippy::cast_precision_loss,
    clippy::cast_sign_loss,
    reason = "sample indices stay below 1200 (exactly representable as f64), and samples are clamped to 0.0..=255.0 so the float-to-u8 truncation toward zero is the intended PCM quantization"
)]
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
        let t = i as f64 / f64::from(SAMPLE_RATE);
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

/// Write a UTF-16LE `CString` atom with the given instance number.
fn write_cstring(instance: u16, text: &str) -> Result<Vec<u8>, Error> {
    let mut atom = RecordBuilder::new(0x00, instance, 0x0FBA);
    for ch in text.encode_utf16() {
        atom.write_data(&ch.to_le_bytes());
    }
    atom.build()
}

/// Build a Sound container with embedded WAV or AIFF data.
///
/// Structure per `LibreOffice` `pptexsoundcollection.cxx` `ExSoundEntry::Write`:
/// - Sound container (0x0F, type 0x07E6)
///   - `CString` (instance 0): sound name (e.g. "Whoosh")
///   - `CString` (instance 1): extension (".wav")
///   - `CString` (instance 2): reference ID string — matched by `AnimationInfoAtom.nSoundRef`
///   - `SoundData` (type 0x07E7): actual WAV binary data
fn build_sound_container(
    name: &str,
    extension: &str,
    ref_id: u32,
    builtin_id: Option<crate::sound_collection::BuiltinId>,
    sound_data: &[u8],
) -> Result<Vec<u8>, Error> {
    let mut children = Vec::new();

    // CString instance 0 — sound name
    children.extend(write_cstring(0, name)?);

    // CString instance 1 — four-code-unit file extension
    children.extend(write_cstring(1, extension)?);

    // CString instance 2 — reference ID (matched against AnimationInfoAtom.nSoundRef)
    children.extend(write_cstring(2, &ref_id.to_string())?);

    if let Some(id) = builtin_id {
        children.extend(write_cstring(3, &id.value().to_string())?);
    }

    // SoundData atom with actual WAV or AIFF binary data
    let mut data_atom = RecordBuilder::new(0x00, 0, RecordType::SoundData as u16);
    data_atom.write_data(sound_data);
    children.extend(data_atom.build()?);

    // Sound container
    let mut container = RecordBuilder::new(0x0F, 0, RecordType::Sound as u16);
    container.write_data(&children);
    container.build()
}

/// Build a `SoundCollection` container for the selected built-in sounds.
///
/// `AnimationInfoAtom.nSoundRef` is the 1-based index into this collection, which is
/// matched against `CString` instance 2 of each Sound container (the reference ID string).
///
/// Returns `(binary_data, mapping)` where mapping is `builtin_sound_id → collection_ref_id`.
///
/// # Errors
///
/// Returns an error if serialization fails or the underlying writer reports an error.
#[allow(
    clippy::implicit_hasher,
    reason = "the public API intentionally accepts `HashSet` with the default hasher; generalizing over hashers would complicate the established signature for no practical benefit"
)]
pub fn build_sound_collection(
    sound_ids: &HashSet<u32>,
) -> Result<(Vec<u8>, HashMap<u32, u32>), Error> {
    let mut builder = SoundCollectionBuilder::new(SoundCollectionLimits::default());
    for sound_id in sound_ids {
        builder.register_reference(*sound_id)?;
    }
    builder.build()
}

fn audio_extension(data: &[u8]) -> Result<&'static str, Error> {
    if data.len() >= 12 && &data[..4] == b"RIFF" && &data[8..12] == b"WAVE" {
        Ok(".WAV")
    } else if data.len() >= 12 && &data[..4] == b"FORM" && matches!(&data[8..12], b"AIFF" | b"AIFC")
    {
        Ok(".AIF")
    } else {
        invalid("embedded sound is not a RIFF/WAVE or FORM/AIFF container")
    }
}

fn builtin_description_id(sound: BuiltinSound) -> Option<crate::sound_collection::BuiltinId> {
    use crate::sound_collection::BuiltinId as Id;
    match sound {
        BuiltinSound::Applause => Some(Id::Applause),
        BuiltinSound::Arrow => Some(Id::Arrow),
        BuiltinSound::Bomb => Some(Id::Bomb),
        BuiltinSound::Breeze => Some(Id::Breeze),
        BuiltinSound::Camera => Some(Id::Camera),
        BuiltinSound::CashRegister => Some(Id::CashRegister),
        BuiltinSound::Chime => Some(Id::Chime),
        BuiltinSound::Click => Some(Id::Click),
        BuiltinSound::Coin => Some(Id::Coin),
        BuiltinSound::DrumRoll => Some(Id::DrumRoll),
        BuiltinSound::Explosion => Some(Id::Explosion),
        BuiltinSound::Hammer => Some(Id::Hammer),
        BuiltinSound::Laser => Some(Id::Laser),
        BuiltinSound::Push => Some(Id::Push),
        BuiltinSound::Suction => Some(Id::Suction),
        BuiltinSound::Swoosh => None,
        BuiltinSound::Typewriter => Some(Id::Typewriter),
        BuiltinSound::Voltage => Some(Id::Voltage),
        BuiltinSound::Whoosh => Some(Id::Whoosh),
        BuiltinSound::Wind => Some(Id::Wind),
    }
}

fn invalid<T>(message: impl Into<String>) -> Result<T, Error> {
    Err(invalid_error(message))
}

fn invalid_error(message: impl Into<String>) -> Error {
    std::io::Error::new(std::io::ErrorKind::InvalidInput, message.into())
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]
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

    #[test]
    fn planner_preserves_embedded_aiff_and_rejects_conflicts_and_limits() {
        let aiff = b"FORM\x00\x00\x00\x04AIFF";
        let embedded = SoundType::Embedded {
            name: "Custom tone".to_string(),
            data: aiff.to_vec(),
        };
        let conflicting = SoundType::Builtin(BuiltinSound::Click);
        let mut builder = SoundCollectionBuilder::new(SoundCollectionLimits::default());
        builder.register(42, &embedded).unwrap();
        assert!(builder.register(42, &conflicting).is_err());

        let mut output_builder = SoundCollectionBuilder::new(SoundCollectionLimits::default());
        output_builder.register(42, &embedded).unwrap();
        let (bytes, mapping) = output_builder.build().unwrap();
        let (record, consumed) = crate::Record::parse(&bytes, 0).unwrap();
        assert_eq!(consumed, bytes.len());
        let collection = crate::sound_collection::Collection::parse(&record).unwrap();
        assert_eq!(collection.sounds[0].id, mapping[&42]);
        assert_eq!(collection.sounds[0].name, "Custom tone");
        assert_eq!(collection.sounds[0].extension.as_deref(), Some(".AIF"));
        assert_eq!(collection.sounds[0].data, aiff);

        let mut limited = SoundCollectionBuilder::new(SoundCollectionLimits {
            max_sound_bytes: aiff.len() - 1,
            ..Default::default()
        });
        assert!(limited.register(42, &embedded).is_err());
        assert!(limited.register_reference(42).is_err());
    }
}
