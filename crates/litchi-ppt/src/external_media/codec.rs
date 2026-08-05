//! Record codecs for strict, inert legacy PowerPoint external-media metadata.

use super::model::{
    CdAudio, CdTime, Collection, EmbeddedWav, LinkedAudio, LinkedAudioKind, Media, Movie,
    MovieKind, Object, UnknownRecord, Video,
};
use crate::consts::RecordType;
use crate::hyperlink::Hyperlinks;
use crate::package::{Error, Result};
use crate::records::Record;
use crate::sound_collection::Collection as SoundCollection;
use std::collections::HashSet;

pub(crate) const MEDIA_FLAGS_MASK: u16 = 0x0007;
pub(crate) const MAX_PATH_UNITS: usize = 32_768;
pub(crate) const MAX_EXTERNAL_MEDIA_OBJECTS: usize = 4_096;

impl Media {
    pub fn parse(record: &Record) -> Result<Self> {
        if record.version != 0
            || record.instance != 0
            || record.record_type_raw != RecordType::ExternalMediaAtom.as_u16()
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

    pub fn to_record(&self) -> Result<Record> {
        Ok(Record::parse(&self.to_record_bytes()?, 0)?.0)
    }

    pub fn to_record_bytes(&self) -> Result<[u8; 16]> {
        if self.id == 0 {
            return corrupted("ExMediaAtom external object ID must be positive");
        }
        let flags = self.loop_playback as u16
            | (self.rewind_after_playing as u16) << 1
            | (self.narration as u16) << 2;
        let mut bytes = [0; 16];
        bytes[2..4].copy_from_slice(&RecordType::ExternalMediaAtom.as_u16().to_le_bytes());
        bytes[4..8].copy_from_slice(&8u32.to_le_bytes());
        bytes[8..12].copy_from_slice(&self.id.to_le_bytes());
        bytes[12..14].copy_from_slice(&flags.to_le_bytes());
        bytes[14..16].copy_from_slice(&self.unused);
        Ok(bytes)
    }
}

impl Video {
    pub fn parse(record: &Record) -> Result<Self> {
        if record.version != 0x0f
            || record.instance != 0
            || record.record_type_raw != RecordType::ExternalVideo.as_u16()
        {
            return corrupted("ExVideoContainer has an invalid header");
        }
        let children = Record::parse_sequence_strict(&record.data, "ExVideoContainer")?;
        if !(1..=2).contains(&children.len()) {
            return corrupted("ExVideoContainer must contain media and optional path atoms");
        }
        let media = Media::parse(&children[0])?;
        if media.narration {
            return corrupted("video ExMediaAtom cannot have the narration flag set");
        }
        let path = children.get(1).map(parse_path).transpose()?;
        Ok(Self { media, path })
    }

    pub fn to_record(&self) -> Result<Record> {
        Ok(Record::parse(&self.to_record_bytes()?, 0)?.0)
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
                RecordType::CString.as_u16(),
                &encode_path(path)?,
            )?);
        }
        record_bytes(0x0f, 0, RecordType::ExternalVideo.as_u16(), &children)
    }
}

impl MovieKind {
    fn record_type(self) -> RecordType {
        match self {
            Self::Avi => RecordType::ExternalAviMovie,
            Self::Mci => RecordType::ExternalMciMovie,
        }
    }

    fn from_record_type(record_type: u16) -> Result<Self> {
        match record_type {
            value if value == RecordType::ExternalAviMovie.as_u16() => Ok(Self::Avi),
            value if value == RecordType::ExternalMciMovie.as_u16() => Ok(Self::Mci),
            _ => corrupted("external movie container has an invalid record type"),
        }
    }
}

impl Movie {
    pub fn parse(record: &Record) -> Result<Self> {
        if record.version != 0x0f || record.instance != 0 {
            return corrupted("external movie container has an invalid header");
        }
        let kind = MovieKind::from_record_type(record.record_type_raw)?;
        let children = Record::parse_sequence_strict(&record.data, "external movie container")?;
        if children.len() != 1 {
            return corrupted("external movie container must contain exactly one ExVideoContainer");
        }
        Ok(Self {
            kind,
            video: Video::parse(&children[0])?,
        })
    }

    pub fn to_record(&self) -> Result<Record> {
        Ok(Record::parse(&self.to_record_bytes()?, 0)?.0)
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

impl LinkedAudioKind {
    fn record_type(self) -> RecordType {
        match self {
            Self::Midi => RecordType::ExternalMidiAudio,
            Self::Wav => RecordType::ExternalWavAudioLink,
        }
    }

    fn from_record_type(record_type: u16) -> Result<Self> {
        match record_type {
            value if value == RecordType::ExternalMidiAudio.as_u16() => Ok(Self::Midi),
            value if value == RecordType::ExternalWavAudioLink.as_u16() => Ok(Self::Wav),
            _ => corrupted("linked audio container has an invalid record type"),
        }
    }
}

impl LinkedAudio {
    pub fn parse(record: &Record) -> Result<Self> {
        if record.version != 0x0f || record.instance != 0 {
            return corrupted("linked audio container has an invalid header");
        }
        let kind = LinkedAudioKind::from_record_type(record.record_type_raw)?;
        let children = Record::parse_sequence_strict(&record.data, "linked audio container")?;
        if !(1..=2).contains(&children.len()) {
            return corrupted("linked audio container must contain media and optional path atoms");
        }
        Ok(Self {
            kind,
            media: Media::parse(&children[0])?,
            path: children.get(1).map(parse_path).transpose()?,
        })
    }

    pub fn to_record(&self) -> Result<Record> {
        Ok(Record::parse(&self.to_record_bytes()?, 0)?.0)
    }

    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        let mut children = self.media.to_record_bytes()?.to_vec();
        if let Some(path) = &self.path {
            children.extend_from_slice(&record_bytes(
                0,
                0,
                RecordType::CString.as_u16(),
                &encode_path(path)?,
            )?);
        }
        record_bytes(0x0f, 0, self.kind.record_type().as_u16(), &children)
    }
}

impl EmbeddedWav {
    pub fn parse(record: &Record) -> Result<Self> {
        if record.version != 0x0f
            || record.instance != 0
            || record.record_type_raw != RecordType::ExternalWavAudioEmbedded.as_u16()
        {
            return corrupted("ExWAVAudioEmbeddedContainer has an invalid header");
        }
        let children = Record::parse_sequence_strict(&record.data, "ExWAVAudioEmbeddedContainer")?;
        if children.len() != 2 {
            return corrupted(
                "ExWAVAudioEmbeddedContainer must contain media and embedded-audio atoms",
            );
        }
        let media = Media::parse(&children[0])?;
        let atom = &children[1];
        if atom.version != 1
            || atom.instance != 1
            || atom.record_type_raw != RecordType::ExternalWavAudioEmbeddedAtom.as_u16()
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

    pub fn validate_sound_collection(&self, sounds: &SoundCollection<'_>) -> Result<()> {
        if let Some(id) = self.sound_id
            && sounds.get(id).is_none()
        {
            return corrupted(format!("embedded WAV references missing sound ID {id}"));
        }
        Ok(())
    }

    pub fn to_record(&self) -> Result<Record> {
        Ok(Record::parse(&self.to_record_bytes()?, 0)?.0)
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
            RecordType::ExternalWavAudioEmbeddedAtom.as_u16(),
            &atom,
        )?);
        record_bytes(
            0x0f,
            0,
            RecordType::ExternalWavAudioEmbedded.as_u16(),
            &children,
        )
    }
}

impl Object {
    pub fn id(&self) -> u32 {
        match self {
            Self::Movie(value) => value.video.media.id,
            Self::LinkedAudio(value) => value.media.id,
            Self::CdAudio(value) => value.media.id,
            Self::EmbeddedWav(value) => value.media.id,
        }
    }
}

impl Collection {
    /// Discover the single `ExObjListContainer`, if present.
    pub fn parse(root: &Record) -> Result<Option<Self>> {
        let mut lists = Vec::new();
        collect_external_object_lists(root, &mut lists);
        if lists.len() > 1 {
            return corrupted("record tree contains multiple external-object lists");
        }
        let Some(record) = lists.first() else {
            return Ok(None);
        };
        let result = Self::parse_list(record)?;
        let hyperlinks = Hyperlinks::parse(root)?;
        if hyperlinks
            .hyperlinks
            .iter()
            .any(|hyperlink| result.get(hyperlink.id).is_some())
        {
            return corrupted("external-object list reuses an ID for media and a hyperlink");
        }
        Ok(Some(result))
    }

    fn parse_list(record: &Record) -> Result<Self> {
        if record.version != 0x0f
            || record.instance != 0
            || record.record_type_raw != RecordType::ExObjList.as_u16()
        {
            return corrupted("ExObjListContainer has an invalid header");
        }
        let children = Record::parse_sequence_strict(&record.data, "ExObjListContainer")?;
        let Some(atom) = children.first() else {
            return corrupted("ExObjListContainer is missing ExObjListAtom");
        };
        if atom.version != 0
            || atom.instance != 0
            || atom.record_type_raw != RecordType::ExObjListAtom.as_u16()
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
        let mut unknown_records = Vec::new();
        for child in &children[1..] {
            let object = match child.record_type {
                RecordType::ExternalAviMovie | RecordType::ExternalMciMovie => {
                    Some(Object::Movie(Movie::parse(child)?))
                },
                RecordType::ExternalMidiAudio | RecordType::ExternalWavAudioLink => {
                    Some(Object::LinkedAudio(LinkedAudio::parse(child)?))
                },
                RecordType::ExternalCdAudio => Some(Object::CdAudio(CdAudio::parse(child)?)),
                RecordType::ExternalWavAudioEmbedded => {
                    Some(Object::EmbeddedWav(EmbeddedWav::parse(child)?))
                },
                _ => {
                    unknown_records.push(UnknownRecord::from_record(child, objects.len()));
                    None
                },
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
        Ok(Self {
            id_seed,
            objects,
            unknown_records,
        })
    }

    pub fn get(&self, id: u32) -> Option<&Object> {
        self.objects.iter().find(|object| object.id() == id)
    }

    /// Unmodeled `ExObjList` children retained in source order.
    pub fn unknown_records(&self) -> &[UnknownRecord] {
        &self.unknown_records
    }

    /// Validate every non-null embedded WAV reference without decoding sound data.
    pub fn validate_sound_collection(&self, sounds: Option<&SoundCollection<'_>>) -> Result<()> {
        for object in &self.objects {
            let Object::EmbeddedWav(value) = object else {
                continue;
            };
            if value.sound_id.is_some() {
                let sounds = sounds.ok_or_else(|| {
                    Error::Corrupted(
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

impl UnknownRecord {
    /// The original raw record type value.
    pub fn record_type(&self) -> u16 {
        self.record.record_type_raw
    }

    /// The original record version.
    pub fn version(&self) -> u16 {
        self.record.version
    }

    /// The original record instance.
    pub fn instance(&self) -> u16 {
        self.record.instance
    }

    /// The original record payload, borrowed from this collection snapshot.
    pub fn data(&self) -> &[u8] {
        &self.record.data
    }

    /// Reconstruct the exact header and payload retained by this record.
    pub fn to_record_bytes(&self) -> Result<Vec<u8>> {
        record_bytes(
            self.record.version,
            self.record.instance,
            self.record.record_type_raw,
            &self.record.data,
        )
    }

    pub(crate) fn from_record(record: &Record, object_index: usize) -> Self {
        Self {
            record: record.clone(),
            object_index,
        }
    }
}

fn collect_external_object_lists<'a>(record: &'a Record, lists: &mut Vec<&'a Record>) {
    if record.record_type == RecordType::ExObjList {
        lists.push(record);
    }
    for child in &record.children {
        collect_external_object_lists(child, lists);
    }
}

fn parse_path(record: &Record) -> Result<String> {
    if record.version != 0
        || record.instance != 0
        || record.record_type_raw != RecordType::CString.as_u16()
        || !record.data.len().is_multiple_of(2)
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
        .map_err(|_| Error::Corrupted("UncOrLocalPathAtom contains invalid UTF-16".to_string()))
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

impl CdTime {
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

impl CdAudio {
    pub fn parse(record: &Record) -> Result<Self> {
        if record.version != 0x0f
            || record.instance != 0
            || record.record_type_raw != RecordType::ExternalCdAudio.as_u16()
        {
            return corrupted("ExCDAudioContainer has an invalid header");
        }
        let children = Record::parse_sequence_strict(&record.data, "ExCDAudioContainer")?;
        if children.len() != 2 {
            return corrupted("ExCDAudioContainer must contain media and CD-audio atoms");
        }
        let media = Media::parse(&children[0])?;
        let atom = &children[1];
        if atom.version != 0
            || atom.instance != 0
            || atom.record_type_raw != RecordType::ExternalCdAudioAtom.as_u16()
            || atom.data.len() != 8
            || atom.data_length != 8
        {
            return corrupted("ExCDAudioAtom has an invalid header or size");
        }
        let start = CdTime::parse(&atom.data[..4])?;
        let end = CdTime::parse(&atom.data[4..])?;
        if start > end {
            return corrupted("ExCDAudioAtom start must not be later than end");
        }
        Ok(Self { media, start, end })
    }

    pub fn to_record(&self) -> Result<Record> {
        Ok(Record::parse(&self.to_record_bytes()?, 0)?.0)
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
            RecordType::ExternalCdAudioAtom.as_u16(),
            &times,
        )?);
        record_bytes(0x0f, 0, RecordType::ExternalCdAudio.as_u16(), &children)
    }
}

pub(crate) fn record_bytes(
    version: u16,
    instance: u16,
    record_type: u16,
    data: &[u8],
) -> Result<Vec<u8>> {
    let length = u32::try_from(data.len())
        .map_err(|_| Error::Corrupted("PowerPoint record payload exceeds u32".to_string()))?;
    let mut bytes = Vec::with_capacity(8usize.saturating_add(data.len()));
    bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    bytes.extend_from_slice(&record_type.to_le_bytes());
    bytes.extend_from_slice(&length.to_le_bytes());
    bytes.extend_from_slice(data);
    Ok(bytes)
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Corrupted(message.into()))
}
