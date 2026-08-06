//! Contextual PowerPoint external-media values.

use crate::records::Record;

/// Resource ceilings for one complete legacy PowerPoint record tree.
///
/// The limits apply before any record payload is decoded. External paths are
/// also bounded by the format-level ceiling in the codec.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Limits {
    /// Maximum serialized root record size.
    pub max_root_bytes: usize,
    /// Maximum payload retained for one parsed record.
    pub max_record_bytes: usize,
    /// Maximum record nesting visited while checking owners.
    pub max_depth: usize,
    /// Maximum records visited while checking owners.
    pub max_records: usize,
    /// Maximum `ExObjRefAtom` records retained from the owner tree.
    pub max_owner_references: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_root_bytes: 64 * 1024 * 1024,
            max_record_bytes: 64 * 1024 * 1024,
            max_depth: 128,
            max_records: 1_000_000,
            max_owner_references: 65_536,
        }
    }
}

/// Playback flags shared by every `ExMediaAtom`-backed media object.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Playback {
    pub loop_playback: bool,
    pub rewind_after_playing: bool,
    pub narration: bool,
}

impl Playback {
    pub const fn new(loop_playback: bool, rewind_after_playing: bool, narration: bool) -> Self {
        Self {
            loop_playback,
            rewind_after_playing,
            narration,
        }
    }
}

/// The eight-byte semantic payload carried by an `ExMediaAtom`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Media {
    pub id: u32,
    pub loop_playback: bool,
    pub rewind_after_playing: bool,
    pub narration: bool,
    /// Undefined source bytes preserved for record roundtrips.
    pub unused: [u8; 2],
}

/// An external video definition nested by an AVI or MCI movie.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Video {
    pub media: Media,
    /// An inert UNC or local path. Parsing never accesses this path.
    pub path: Option<String>,
}

/// The external movie container family.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum MovieKind {
    Avi,
    Mci,
}

/// A validated `ExAviMovieContainer` or `ExMCIMovieContainer`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Movie {
    pub kind: MovieKind,
    pub video: Video,
}

/// The linked external-audio container family.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum LinkedAudioKind {
    Midi,
    Wav,
}

/// A validated `ExMIDIAudioContainer` or `ExWAVAudioLinkContainer`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct LinkedAudio {
    pub kind: LinkedAudioKind,
    pub media: Media,
    /// An inert UNC or local path. Parsing never accesses this path.
    pub path: Option<String>,
}

/// A validated `ExWAVAudioEmbeddedContainer`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct EmbeddedWav {
    pub media: Media,
    /// A null reference is represented by `None`.
    pub sound_id: Option<u32>,
    pub duration_ms: u32,
}

/// CD audio time in track/minute/second/frame form.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
pub struct CdTime {
    pub track: u8,
    pub minute: u8,
    pub second: u8,
    pub frame: u8,
}

/// A validated `ExCDAudioContainer`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct CdAudio {
    pub media: Media,
    pub start: CdTime,
    pub end: CdTime,
}

/// One typed audio/video definition from the document `ExObjListContainer`.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Object {
    Movie(Movie),
    LinkedAudio(LinkedAudio),
    CdAudio(CdAudio),
    EmbeddedWav(EmbeddedWav),
}

/// A bounded, lossless child of `ExObjList` that is not modeled as media.
///
/// The original record header fields and payload are retained so reading a
/// media collection never discards OLE, hyperlink, or future-version
/// children. The payload is exposed only by borrow or an exact reconstruction.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct UnknownRecord {
    pub(crate) record: Record,
    pub(crate) object_index: usize,
}

/// Strict audio/video definitions discovered in a document external-object list.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Collection {
    pub id_seed: u32,
    pub objects: Vec<Object>,
    pub(crate) unknown_records: Vec<UnknownRecord>,
}

impl Collection {
    /// Return the typed playback flags for one media object.
    pub fn playback(&self, id: u32) -> Option<Playback> {
        self.get(id).map(Object::playback)
    }
}

impl Object {
    /// Return the playback flags carried by this media definition.
    pub fn playback(&self) -> Playback {
        let media = match self {
            Self::Movie(value) => value.video.media,
            Self::LinkedAudio(value) => value.media,
            Self::CdAudio(value) => value.media,
            Self::EmbeddedWav(value) => value.media,
        };
        Playback {
            loop_playback: media.loop_playback,
            rewind_after_playing: media.rewind_after_playing,
            narration: media.narration,
        }
    }
}
