//! Strict, inert legacy PowerPoint external-media metadata.

use super::hyperlink::PowerPointHyperlinks;
use super::package::{PptError, Result};
use super::records::PptRecord;
use super::sound_collection::PowerPointSoundCollection;
use crate::consts::PptRecordType;
use std::collections::HashSet;

const MEDIA_FLAGS_MASK: u16 = 0x0007;
const MAX_PATH_UNITS: usize = 32_768;
const MAX_EXTERNAL_MEDIA_OBJECTS: usize = 4_096;

/// The common eight-byte payload in an `ExMediaAtom`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct PowerPointExternalMedia {
    pub id: u32,
    pub loop_playback: bool,
    pub rewind_after_playing: bool,
    pub narration: bool,
    /// Undefined source bytes preserved for record roundtrips.
    pub unused: [u8; 2],
}

impl PowerPointExternalMedia {
    pub fn parse(record: &PptRecord) -> Result<Self> {
        if record.version != 0
            || record.instance != 0
            || record.record_type_raw != PptRecordType::ExternalMediaAtom.as_u16()
            || record.data.len() != 8
            || record.data_length != 8
        {
            return corrupted("ExMediaAtom has an invalid header or size");
        }
        let id = u32::from_le_bytes(record.data[0..4].try_into().expect("fixed slice"));
        if id == 0 {
            return corrupted("ExMediaAtom external object ID must be positive");
        }
        let flags = u16::from_le_bytes([record.data[4], record.data[5]]);
        if flags & !MEDIA_FLAGS_MASK != 0 {
            return corrupted("ExMediaAtom has nonzero reserved flag bits");
        }
        Ok(Self {
            id,
            loop_playback: flags & 1 != 0,
            rewind_after_playing: flags & 2 != 0,
            narration: flags & 4 != 0,
            unused: [record.data[6], record.data[7]],
        })
    }

    pub fn to_record(&self) -> Result<PptRecord> {
        Ok(PptRecord::parse(&self.to_record_bytes()?, 0)?.0)
    }

    pub fn to_record_bytes(&self) -> Result<[u8; 16]> {
        if self.id == 0 {
            return corrupted("ExMediaAtom external object ID must be positive");
        }
        let flags = self.loop_playback as u16
            | (self.rewind_after_playing as u16) << 1
            | (self.narration as u16) << 2;
        let mut bytes = [0; 16];
        bytes[2..4].copy_from_slice(&PptRecordType::ExternalMediaAtom.as_u16().to_le_bytes());
        bytes[4..8].copy_from_slice(&8u32.to_le_bytes());
        bytes[8..12].copy_from_slice(&self.id.to_le_bytes());
        bytes[12..14].copy_from_slice(&flags.to_le_bytes());
        bytes[14..16].copy_from_slice(&self.unused);
        Ok(bytes)
    }
}

/// The shared `ExVideoContainer` nested by AVI and MCI movie records.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PowerPointExternalVideo {
    pub media: PowerPointExternalMedia,
    /// An inert UNC or local path. Parsing never accesses this path.
    pub path: Option<String>,
}

impl PowerPointExternalVideo {
    pub fn parse(record: &PptRecord) -> Result<Self> {
        if record.version != 0x0f
            || record.instance != 0
            || record.record_type_raw != PptRecordType::ExternalVideo.as_u16()
        {
            return corrupted("ExVideoContainer has an invalid header");
        }
        let children = PptRecord::parse_sequence_strict(&record.data, "ExVideoContainer")?;
        if !(1..=2).contains(&children.len()) {
            return corrupted("ExVideoContainer must contain media and optional path atoms");
        }
        let media = PowerPointExternalMedia::parse(&children[0])?;
        if media.narration {
            return corrupted("video ExMediaAtom cannot have the narration flag set");
        }
        let path = children.get(1).map(parse_path).transpose()?;
        Ok(Self { media, path })
    }

    pub fn to_record(&self) -> Result<PptRecord> {
        Ok(PptRecord::parse(&self.to_record_bytes()?, 0)?.0)
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        if self.media.narration {
            return corrupted("video ExMediaAtom cannot have the narration flag set");
        }
        let mut children = self.media.to_record_bytes()?.to_vec();
        if let Some(path) = &self.path {
            children.extend_from_slice(&record_bytes(
                0,
                0,
                PptRecordType::CString.as_u16(),
                &encode_path(path)?,
            )?);
        }
        record_bytes(0x0f, 0, PptRecordType::ExternalVideo.as_u16(), &children)
    }
}

/// The external movie container family that selects an AVI or MCI player.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PowerPointExternalMovieKind {
    Avi,
    Mci,
}

impl PowerPointExternalMovieKind {
    fn record_type(self) -> PptRecordType {
        match self {
            Self::Avi => PptRecordType::ExternalAviMovie,
            Self::Mci => PptRecordType::ExternalMciMovie,
        }
    }

    fn from_record_type(record_type: u16) -> Result<Self> {
        match record_type {
            value if value == PptRecordType::ExternalAviMovie.as_u16() => Ok(Self::Avi),
            value if value == PptRecordType::ExternalMciMovie.as_u16() => Ok(Self::Mci),
            _ => corrupted("external movie container has an invalid record type"),
        }
    }
}

/// A validated `ExAviMovieContainer` or `ExMCIMovieContainer`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PowerPointExternalMovie {
    pub kind: PowerPointExternalMovieKind,
    pub video: PowerPointExternalVideo,
}

impl PowerPointExternalMovie {
    pub fn parse(record: &PptRecord) -> Result<Self> {
        if record.version != 0x0f || record.instance != 0 {
            return corrupted("external movie container has an invalid header");
        }
        let kind = PowerPointExternalMovieKind::from_record_type(record.record_type_raw)?;
        let children = PptRecord::parse_sequence_strict(&record.data, "external movie container")?;
        if children.len() != 1 {
            return corrupted("external movie container must contain exactly one ExVideoContainer");
        }
        Ok(Self {
            kind,
            video: PowerPointExternalVideo::parse(&children[0])?,
        })
    }

    pub fn to_record(&self) -> Result<PptRecord> {
        Ok(PptRecord::parse(&self.to_record_bytes()?, 0)?.0)
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        record_bytes(
            0x0f,
            0,
            self.kind.record_type().as_u16(),
            &self.video.to_record_bytes()?,
        )
    }
}

/// The linked external-audio container family.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PowerPointLinkedAudioKind {
    Midi,
    Wav,
}

impl PowerPointLinkedAudioKind {
    fn record_type(self) -> PptRecordType {
        match self {
            Self::Midi => PptRecordType::ExternalMidiAudio,
            Self::Wav => PptRecordType::ExternalWavAudioLink,
        }
    }

    fn from_record_type(record_type: u16) -> Result<Self> {
        match record_type {
            value if value == PptRecordType::ExternalMidiAudio.as_u16() => Ok(Self::Midi),
            value if value == PptRecordType::ExternalWavAudioLink.as_u16() => Ok(Self::Wav),
            _ => corrupted("linked audio container has an invalid record type"),
        }
    }
}

/// A validated `ExMIDIAudioContainer` or `ExWAVAudioLinkContainer`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PowerPointLinkedAudio {
    pub kind: PowerPointLinkedAudioKind,
    pub media: PowerPointExternalMedia,
    /// An inert UNC or local path. Parsing never accesses this path.
    pub path: Option<String>,
}

impl PowerPointLinkedAudio {
    pub fn parse(record: &PptRecord) -> Result<Self> {
        if record.version != 0x0f || record.instance != 0 {
            return corrupted("linked audio container has an invalid header");
        }
        let kind = PowerPointLinkedAudioKind::from_record_type(record.record_type_raw)?;
        let children = PptRecord::parse_sequence_strict(&record.data, "linked audio container")?;
        if !(1..=2).contains(&children.len()) {
            return corrupted("linked audio container must contain media and optional path atoms");
        }
        Ok(Self {
            kind,
            media: PowerPointExternalMedia::parse(&children[0])?,
            path: children.get(1).map(parse_path).transpose()?,
        })
    }

    pub fn to_record(&self) -> Result<PptRecord> {
        Ok(PptRecord::parse(&self.to_record_bytes()?, 0)?.0)
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        let mut children = self.media.to_record_bytes()?.to_vec();
        if let Some(path) = &self.path {
            children.extend_from_slice(&record_bytes(
                0,
                0,
                PptRecordType::CString.as_u16(),
                &encode_path(path)?,
            )?);
        }
        record_bytes(0x0f, 0, self.kind.record_type().as_u16(), &children)
    }
}

/// A validated `ExWAVAudioEmbeddedContainer`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct PowerPointEmbeddedWav {
    pub media: PowerPointExternalMedia,
    /// A null reference is represented by `None`.
    pub sound_id: Option<u32>,
    pub duration_ms: u32,
}

impl PowerPointEmbeddedWav {
    pub fn parse(record: &PptRecord) -> Result<Self> {
        if record.version != 0x0f
            || record.instance != 0
            || record.record_type_raw != PptRecordType::ExternalWavAudioEmbedded.as_u16()
        {
            return corrupted("ExWAVAudioEmbeddedContainer has an invalid header");
        }
        let children =
            PptRecord::parse_sequence_strict(&record.data, "ExWAVAudioEmbeddedContainer")?;
        if children.len() != 2 {
            return corrupted(
                "ExWAVAudioEmbeddedContainer must contain media and embedded-audio atoms",
            );
        }
        let media = PowerPointExternalMedia::parse(&children[0])?;
        let atom = &children[1];
        if atom.version != 1
            || atom.instance != 1
            || atom.record_type_raw != PptRecordType::ExternalWavAudioEmbeddedAtom.as_u16()
            || atom.data.len() != 8
            || atom.data_length != 8
        {
            return corrupted("ExWAVAudioEmbeddedAtom has an invalid header or size");
        }
        let sound_id = u32::from_le_bytes(atom.data[..4].try_into().expect("fixed slice"));
        let duration_ms = i32::from_le_bytes(atom.data[4..].try_into().expect("fixed slice"));
        if duration_ms < 0 {
            return corrupted("ExWAVAudioEmbeddedAtom duration cannot be negative");
        }
        Ok(Self {
            media,
            sound_id: (sound_id != 0).then_some(sound_id),
            duration_ms: duration_ms as u32,
        })
    }

    pub fn validate_sound_collection(&self, sounds: &PowerPointSoundCollection<'_>) -> Result<()> {
        if let Some(id) = self.sound_id
            && sounds.get(id).is_none()
        {
            return corrupted(format!("embedded WAV references missing sound ID {id}"));
        }
        Ok(())
    }

    pub fn to_record(&self) -> Result<PptRecord> {
        Ok(PptRecord::parse(&self.to_record_bytes()?, 0)?.0)
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        if self.duration_ms > i32::MAX as u32 {
            return corrupted("embedded WAV duration exceeds the signed 32-bit range");
        }
        let mut children = self.media.to_record_bytes()?.to_vec();
        let mut atom = [0; 8];
        atom[..4].copy_from_slice(&self.sound_id.unwrap_or(0).to_le_bytes());
        atom[4..].copy_from_slice(&self.duration_ms.to_le_bytes());
        children.extend_from_slice(&record_bytes(
            1,
            1,
            PptRecordType::ExternalWavAudioEmbeddedAtom.as_u16(),
            &atom,
        )?);
        record_bytes(
            0x0f,
            0,
            PptRecordType::ExternalWavAudioEmbedded.as_u16(),
            &children,
        )
    }
}

/// One audio or video definition from the document `ExObjListContainer`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum PowerPointExternalMediaObject {
    Movie(PowerPointExternalMovie),
    LinkedAudio(PowerPointLinkedAudio),
    CdAudio(PowerPointCdAudio),
    EmbeddedWav(PowerPointEmbeddedWav),
}

impl PowerPointExternalMediaObject {
    pub fn id(&self) -> u32 {
        match self {
            Self::Movie(value) => value.video.media.id,
            Self::LinkedAudio(value) => value.media.id,
            Self::CdAudio(value) => value.media.id,
            Self::EmbeddedWav(value) => value.media.id,
        }
    }
}

/// Strict audio/video definitions discovered in a document external-object list.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PowerPointExternalMediaCollection {
    pub id_seed: u32,
    pub objects: Vec<PowerPointExternalMediaObject>,
}

impl PowerPointExternalMediaCollection {
    /// Discover the single `ExObjListContainer`, if present.
    pub fn parse(root: &PptRecord) -> Result<Option<Self>> {
        let mut lists = Vec::new();
        collect_external_object_lists(root, &mut lists);
        if lists.len() > 1 {
            return corrupted("record tree contains multiple external-object lists");
        }
        let Some(record) = lists.first() else {
            return Ok(None);
        };
        let result = Self::parse_list(record)?;
        let hyperlinks = PowerPointHyperlinks::parse(root)?;
        if hyperlinks
            .hyperlinks
            .iter()
            .any(|hyperlink| result.get(hyperlink.id).is_some())
        {
            return corrupted("external-object list reuses an ID for media and a hyperlink");
        }
        Ok(Some(result))
    }

    fn parse_list(record: &PptRecord) -> Result<Self> {
        if record.version != 0x0f
            || record.instance != 0
            || record.record_type_raw != PptRecordType::ExObjList.as_u16()
        {
            return corrupted("ExObjListContainer has an invalid header");
        }
        let children = PptRecord::parse_sequence_strict(&record.data, "ExObjListContainer")?;
        let Some(atom) = children.first() else {
            return corrupted("ExObjListContainer is missing ExObjListAtom");
        };
        if atom.version != 0
            || atom.instance != 0
            || atom.record_type_raw != PptRecordType::ExObjListAtom.as_u16()
            || atom.data.len() != 4
            || atom.data_length != 4
        {
            return corrupted("ExObjListAtom has an invalid header or size");
        }
        let signed_seed = i32::from_le_bytes(atom.data[..4].try_into().expect("fixed slice"));
        if signed_seed < 1 {
            return corrupted("ExObjListAtom identifier seed must be positive");
        }
        let id_seed = signed_seed as u32;
        let mut ids = HashSet::new();
        let mut objects = Vec::new();
        for child in &children[1..] {
            let object = match child.record_type {
                PptRecordType::ExternalAviMovie | PptRecordType::ExternalMciMovie => Some(
                    PowerPointExternalMediaObject::Movie(PowerPointExternalMovie::parse(child)?),
                ),
                PptRecordType::ExternalMidiAudio | PptRecordType::ExternalWavAudioLink => {
                    Some(PowerPointExternalMediaObject::LinkedAudio(
                        PowerPointLinkedAudio::parse(child)?,
                    ))
                },
                PptRecordType::ExternalCdAudio => Some(PowerPointExternalMediaObject::CdAudio(
                    PowerPointCdAudio::parse(child)?,
                )),
                PptRecordType::ExternalWavAudioEmbedded => {
                    Some(PowerPointExternalMediaObject::EmbeddedWav(
                        PowerPointEmbeddedWav::parse(child)?,
                    ))
                },
                _ => None,
            };
            let Some(object) = object else { continue };
            if objects.len() >= MAX_EXTERNAL_MEDIA_OBJECTS {
                return corrupted(format!(
                    "external-object list exceeds {MAX_EXTERNAL_MEDIA_OBJECTS} media objects"
                ));
            }
            let id = object.id();
            if id > id_seed {
                return corrupted(format!(
                    "external media ID {id} exceeds ExObjList seed {id_seed}"
                ));
            }
            if !ids.insert(id) {
                return corrupted(format!(
                    "external-object list contains duplicate media ID {id}"
                ));
            }
            objects.push(object);
        }
        Ok(Self { id_seed, objects })
    }

    pub fn get(&self, id: u32) -> Option<&PowerPointExternalMediaObject> {
        self.objects.iter().find(|object| object.id() == id)
    }

    /// Validate every non-null embedded WAV reference without decoding sound data.
    pub fn validate_sound_collection(
        &self,
        sounds: Option<&PowerPointSoundCollection<'_>>,
    ) -> Result<()> {
        for object in &self.objects {
            let PowerPointExternalMediaObject::EmbeddedWav(value) = object else {
                continue;
            };
            if value.sound_id.is_some() {
                let sounds = sounds.ok_or_else(|| {
                    PptError::Corrupted(
                        "embedded WAV references a sound but the document has no SoundCollection"
                            .to_string(),
                    )
                })?;
                value.validate_sound_collection(sounds)?;
            }
        }
        Ok(())
    }
}

fn collect_external_object_lists<'a>(record: &'a PptRecord, lists: &mut Vec<&'a PptRecord>) {
    if record.record_type == PptRecordType::ExObjList {
        lists.push(record);
    }
    for child in &record.children {
        collect_external_object_lists(child, lists);
    }
}

fn parse_path(record: &PptRecord) -> Result<String> {
    if record.version != 0
        || record.instance != 0
        || record.record_type_raw != PptRecordType::CString.as_u16()
        || record.data.len() % 2 != 0
        || record.data.len() / 2 > MAX_PATH_UNITS
    {
        return corrupted("UncOrLocalPathAtom has an invalid header or size");
    }
    let units = record
        .data
        .chunks_exact(2)
        .map(|bytes| u16::from_le_bytes([bytes[0], bytes[1]]))
        .collect::<Vec<_>>();
    if units.contains(&0) {
        return corrupted("UncOrLocalPathAtom contains an embedded null");
    }
    String::from_utf16(&units)
        .map_err(|_| PptError::Corrupted("UncOrLocalPathAtom contains invalid UTF-16".to_string()))
}

fn encode_path(path: &str) -> Result<Vec<u8>> {
    let units = path.encode_utf16().collect::<Vec<_>>();
    if units.len() > MAX_PATH_UNITS {
        return corrupted(format!(
            "UncOrLocalPathAtom exceeds {MAX_PATH_UNITS} UTF-16 units"
        ));
    }
    if units.contains(&0) {
        return corrupted("UncOrLocalPathAtom contains an embedded null");
    }
    Ok(units.into_iter().flat_map(u16::to_le_bytes).collect())
}

/// CD audio time in track/minute/second/frame form.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
pub struct PowerPointCdTime {
    pub track: u8,
    pub minute: u8,
    pub second: u8,
    pub frame: u8,
}

impl PowerPointCdTime {
    pub fn new(track: u8, minute: u8, second: u8, frame: u8) -> Result<Self> {
        let value = Self {
            track,
            minute,
            second,
            frame,
        };
        value.validate()?;
        Ok(value)
    }

    fn parse(data: &[u8]) -> Result<Self> {
        Self::new(data[0], data[1], data[2], data[3])
    }

    fn validate(self) -> Result<()> {
        if !(1..=100).contains(&self.track)
            || self.minute > 60
            || self.second >= 60
            || self.frame >= 74
        {
            return corrupted("TmsfTimeStruct contains an out-of-range component");
        }
        Ok(())
    }

    fn bytes(self) -> [u8; 4] {
        [self.track, self.minute, self.second, self.frame]
    }
}

/// A validated `ExCDAudioContainer`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct PowerPointCdAudio {
    pub media: PowerPointExternalMedia,
    pub start: PowerPointCdTime,
    pub end: PowerPointCdTime,
}

impl PowerPointCdAudio {
    pub fn parse(record: &PptRecord) -> Result<Self> {
        if record.version != 0x0f
            || record.instance != 0
            || record.record_type_raw != PptRecordType::ExternalCdAudio.as_u16()
        {
            return corrupted("ExCDAudioContainer has an invalid header");
        }
        let children = PptRecord::parse_sequence_strict(&record.data, "ExCDAudioContainer")?;
        if children.len() != 2 {
            return corrupted("ExCDAudioContainer must contain media and CD-audio atoms");
        }
        let media = PowerPointExternalMedia::parse(&children[0])?;
        let atom = &children[1];
        if atom.version != 0
            || atom.instance != 0
            || atom.record_type_raw != PptRecordType::ExternalCdAudioAtom.as_u16()
            || atom.data.len() != 8
            || atom.data_length != 8
        {
            return corrupted("ExCDAudioAtom has an invalid header or size");
        }
        let start = PowerPointCdTime::parse(&atom.data[..4])?;
        let end = PowerPointCdTime::parse(&atom.data[4..])?;
        if start > end {
            return corrupted("ExCDAudioAtom start must not be later than end");
        }
        Ok(Self { media, start, end })
    }

    pub fn to_record(&self) -> Result<PptRecord> {
        Ok(PptRecord::parse(&self.to_record_bytes()?, 0)?.0)
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        self.start.validate()?;
        self.end.validate()?;
        if self.start > self.end {
            return corrupted("ExCDAudioAtom start must not be later than end");
        }
        let mut children = self.media.to_record_bytes()?.to_vec();
        let mut times = [0; 8];
        times[..4].copy_from_slice(&self.start.bytes());
        times[4..].copy_from_slice(&self.end.bytes());
        children.extend_from_slice(&record_bytes(
            0,
            0,
            PptRecordType::ExternalCdAudioAtom.as_u16(),
            &times,
        )?);
        record_bytes(0x0f, 0, PptRecordType::ExternalCdAudio.as_u16(), &children)
    }
}

fn record_bytes(version: u16, instance: u16, record_type: u16, data: &[u8]) -> Result<Vec<u8>> {
    let length = u32::try_from(data.len())
        .map_err(|_| PptError::Corrupted("PowerPoint record payload exceeds u32".to_string()))?;
    let mut bytes = Vec::with_capacity(8usize.saturating_add(data.len()));
    bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    bytes.extend_from_slice(&record_type.to_le_bytes());
    bytes.extend_from_slice(&length.to_le_bytes());
    bytes.extend_from_slice(data);
    Ok(bytes)
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(PptError::Corrupted(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn cd_audio() -> PowerPointCdAudio {
        PowerPointCdAudio {
            media: PowerPointExternalMedia {
                id: 17,
                loop_playback: true,
                rewind_after_playing: true,
                narration: true,
                unused: [0xaa, 0x55],
            },
            start: PowerPointCdTime::new(1, 2, 3, 4).unwrap(),
            end: PowerPointCdTime::new(2, 0, 0, 0).unwrap(),
        }
    }

    #[test]
    fn protocol_shaped_cd_audio_roundtrips_losslessly() {
        let expected = cd_audio();
        let parsed = PowerPointCdAudio::parse(&expected.to_record().unwrap()).unwrap();
        assert_eq!(parsed, expected);
        assert_eq!(
            parsed.to_record_bytes().unwrap(),
            expected.to_record_bytes().unwrap()
        );
    }

    #[test]
    fn rejects_hostile_media_flags_times_order_and_headers() {
        for components in [
            (0, 0, 0, 0),
            (101, 0, 0, 0),
            (1, 61, 0, 0),
            (1, 0, 60, 0),
            (1, 0, 0, 74),
        ] {
            assert!(
                PowerPointCdTime::new(components.0, components.1, components.2, components.3)
                    .is_err()
            );
        }
        let mut value = cd_audio();
        value.end = PowerPointCdTime::new(1, 0, 0, 0).unwrap();
        assert!(value.to_record_bytes().is_err());
        let mut bytes = cd_audio().to_record_bytes().unwrap();
        bytes[20] = 0x08;
        let record = PptRecord::parse(&bytes, 0).unwrap().0;
        assert!(PowerPointCdAudio::parse(&record).is_err());
        bytes = cd_audio().to_record_bytes().unwrap();
        bytes[0] = 0;
        let record = PptRecord::parse(&bytes, 0).unwrap().0;
        assert!(PowerPointCdAudio::parse(&record).is_err());
    }

    #[test]
    fn external_video_roundtrips_optional_inert_paths() {
        for path in [None, Some(r"\\server\share\movie.avi".to_string())] {
            let expected = PowerPointExternalVideo {
                media: PowerPointExternalMedia {
                    id: 23,
                    loop_playback: false,
                    rewind_after_playing: true,
                    narration: false,
                    unused: [3, 4],
                },
                path,
            };
            let parsed = PowerPointExternalVideo::parse(&expected.to_record().unwrap()).unwrap();
            assert_eq!(parsed, expected);
        }
    }

    #[test]
    fn external_video_rejects_narration_and_hostile_paths() {
        let mut video = PowerPointExternalVideo {
            media: cd_audio().media,
            path: Some("movie.avi".into()),
        };
        assert!(video.to_record_bytes().is_err());
        video.media.narration = false;
        video.path = Some("bad\0path".into());
        assert!(video.to_record_bytes().is_err());
        video.path = Some("x".repeat(MAX_PATH_UNITS + 1));
        assert!(video.to_record_bytes().is_err());
    }

    #[test]
    fn avi_and_mci_movie_containers_roundtrip_canonically() {
        for kind in [
            PowerPointExternalMovieKind::Avi,
            PowerPointExternalMovieKind::Mci,
        ] {
            let expected = PowerPointExternalMovie {
                kind,
                video: PowerPointExternalVideo {
                    media: PowerPointExternalMedia {
                        id: 41,
                        loop_playback: true,
                        rewind_after_playing: false,
                        narration: false,
                        unused: [0x12, 0x34],
                    },
                    path: Some(r"\\server\share\movie.avi".into()),
                },
            };
            let parsed = PowerPointExternalMovie::parse(&expected.to_record().unwrap()).unwrap();
            assert_eq!(parsed, expected);
            assert_eq!(
                parsed.to_record_bytes().unwrap(),
                expected.to_record_bytes().unwrap()
            );
        }
    }

    #[test]
    fn movie_containers_reject_wrong_headers_and_extra_children() {
        let movie = PowerPointExternalMovie {
            kind: PowerPointExternalMovieKind::Avi,
            video: PowerPointExternalVideo {
                media: PowerPointExternalMedia {
                    id: 42,
                    loop_playback: false,
                    rewind_after_playing: false,
                    narration: false,
                    unused: [0, 0],
                },
                path: None,
            },
        };
        let mut bytes = movie.to_record_bytes().unwrap();
        bytes[0] = 0;
        assert!(PowerPointExternalMovie::parse(&PptRecord::parse(&bytes, 0).unwrap().0).is_err());

        let child = movie.video.to_record_bytes().unwrap();
        let doubled = [child.as_slice(), child.as_slice()].concat();
        let bytes =
            record_bytes(0x0f, 0, PptRecordType::ExternalAviMovie.as_u16(), &doubled).unwrap();
        assert!(PowerPointExternalMovie::parse(&PptRecord::parse(&bytes, 0).unwrap().0).is_err());
    }

    #[test]
    fn midi_and_linked_wav_roundtrip_optional_inert_paths() {
        for kind in [
            PowerPointLinkedAudioKind::Midi,
            PowerPointLinkedAudioKind::Wav,
        ] {
            let expected = PowerPointLinkedAudio {
                kind,
                media: cd_audio().media,
                path: Some(r"C:\media\theme.mid".into()),
            };
            assert_eq!(
                PowerPointLinkedAudio::parse(&expected.to_record().unwrap()).unwrap(),
                expected
            );
        }
    }

    #[test]
    fn linked_audio_rejects_misplaced_children_and_hostile_paths() {
        let mut audio = PowerPointLinkedAudio {
            kind: PowerPointLinkedAudioKind::Wav,
            media: cd_audio().media,
            path: Some("bad\0path".into()),
        };
        assert!(audio.to_record_bytes().is_err());
        audio.path = None;
        let child = audio.media.to_record_bytes().unwrap();
        let payload = [child.as_slice(), child.as_slice()].concat();
        let bytes = record_bytes(
            0x0f,
            0,
            PptRecordType::ExternalWavAudioLink.as_u16(),
            &payload,
        )
        .unwrap();
        assert!(PowerPointLinkedAudio::parse(&PptRecord::parse(&bytes, 0).unwrap().0).is_err());
    }

    #[test]
    fn embedded_wav_roundtrips_nullable_sound_reference() {
        for sound_id in [None, Some(7)] {
            let expected = PowerPointEmbeddedWav {
                media: cd_audio().media,
                sound_id,
                duration_ms: 90_000,
            };
            assert_eq!(
                PowerPointEmbeddedWav::parse(&expected.to_record().unwrap()).unwrap(),
                expected
            );
        }
    }

    #[test]
    fn embedded_wav_rejects_negative_or_overflowing_duration() {
        let value = PowerPointEmbeddedWav {
            media: cd_audio().media,
            sound_id: Some(1),
            duration_ms: i32::MAX as u32 + 1,
        };
        assert!(value.to_record_bytes().is_err());

        let mut bytes = PowerPointEmbeddedWav {
            duration_ms: 1,
            ..value
        }
        .to_record_bytes()
        .unwrap();
        let duration_offset = 8 + 16 + 8 + 4;
        bytes[duration_offset..duration_offset + 4].copy_from_slice(&(-1i32).to_le_bytes());
        assert!(PowerPointEmbeddedWav::parse(&PptRecord::parse(&bytes, 0).unwrap().0).is_err());
    }

    fn external_object_list(seed: i32, objects: &[Vec<u8>]) -> PptRecord {
        let mut payload = record_bytes(
            0,
            0,
            PptRecordType::ExObjListAtom.as_u16(),
            &seed.to_le_bytes(),
        )
        .unwrap();
        for object in objects {
            payload.extend_from_slice(object);
        }
        let bytes = record_bytes(0x0f, 0, PptRecordType::ExObjList.as_u16(), &payload).unwrap();
        PptRecord::parse(&bytes, 0).unwrap().0
    }

    fn hyperlink(id: u32) -> Vec<u8> {
        let unicode = |value: &str| {
            value
                .encode_utf16()
                .flat_map(u16::to_le_bytes)
                .collect::<Vec<_>>()
        };
        let mut payload = record_bytes(
            0,
            0,
            PptRecordType::ExternalHyperlinkAtom.as_u16(),
            &id.to_le_bytes(),
        )
        .unwrap();
        payload.extend(
            record_bytes(
                0,
                0,
                PptRecordType::CString.as_u16(),
                &unicode("Media link"),
            )
            .unwrap(),
        );
        payload.extend(
            record_bytes(
                0,
                1,
                PptRecordType::CString.as_u16(),
                &unicode("https://example.test"),
            )
            .unwrap(),
        );
        payload.extend(
            record_bytes(0, 3, PptRecordType::CString.as_u16(), &unicode("slide")).unwrap(),
        );
        record_bytes(0x0f, 0, PptRecordType::ExternalHyperlink.as_u16(), &payload).unwrap()
    }

    #[test]
    fn media_collection_discovers_typed_objects_and_resolves_ids() {
        let movie = PowerPointExternalMovie {
            kind: PowerPointExternalMovieKind::Mci,
            video: PowerPointExternalVideo {
                media: PowerPointExternalMedia {
                    id: 3,
                    loop_playback: false,
                    rewind_after_playing: false,
                    narration: false,
                    unused: [0, 0],
                },
                path: None,
            },
        };
        let audio = PowerPointLinkedAudio {
            kind: PowerPointLinkedAudioKind::Midi,
            media: PowerPointExternalMedia {
                id: 7,
                ..cd_audio().media
            },
            path: Some("theme.mid".into()),
        };
        let root = external_object_list(
            7,
            &[
                movie.to_record_bytes().unwrap(),
                audio.to_record_bytes().unwrap(),
            ],
        );
        let parsed = PowerPointExternalMediaCollection::parse(&root)
            .unwrap()
            .unwrap();
        assert_eq!(parsed.id_seed, 7);
        assert_eq!(parsed.objects.len(), 2);
        assert!(parsed.get(3).is_some());
        assert!(parsed.get(7).is_some());
    }

    #[test]
    fn media_collection_rejects_duplicate_ids_and_low_seed() {
        let first = PowerPointLinkedAudio {
            kind: PowerPointLinkedAudioKind::Midi,
            media: PowerPointExternalMedia {
                id: 9,
                ..cd_audio().media
            },
            path: None,
        }
        .to_record_bytes()
        .unwrap();
        let second = PowerPointLinkedAudio {
            kind: PowerPointLinkedAudioKind::Wav,
            media: PowerPointExternalMedia {
                id: 9,
                ..cd_audio().media
            },
            path: None,
        }
        .to_record_bytes()
        .unwrap();
        assert!(
            PowerPointExternalMediaCollection::parse(&external_object_list(
                9,
                &[first.clone(), second]
            ))
            .is_err()
        );
        assert!(
            PowerPointExternalMediaCollection::parse(&external_object_list(8, &[first])).is_err()
        );
    }

    #[test]
    fn media_collection_rejects_cross_family_hyperlink_id_collisions() {
        let media = PowerPointLinkedAudio {
            kind: PowerPointLinkedAudioKind::Midi,
            media: PowerPointExternalMedia {
                id: 11,
                ..cd_audio().media
            },
            path: None,
        }
        .to_record_bytes()
        .unwrap();
        let root = external_object_list(11, &[media, hyperlink(11)]);
        assert!(PowerPointExternalMediaCollection::parse(&root).is_err());
    }
}
