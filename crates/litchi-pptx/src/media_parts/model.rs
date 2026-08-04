use super::*;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Conformance {
    Transitional,
    Strict,
}

impl Conformance {
    pub(super) fn pml(self) -> &'static str {
        match self {
            Self::Transitional => PML,
            Self::Strict => STRICT_PML,
        }
    }
    pub(super) fn dml(self) -> &'static str {
        match self {
            Self::Transitional => DML,
            Self::Strict => STRICT_DML,
        }
    }
    pub(super) fn rel(self) -> &'static str {
        match self {
            Self::Transitional => REL,
            Self::Strict => STRICT_REL,
        }
    }
    pub(super) fn image_rel(self) -> &'static str {
        match self {
            Self::Transitional => rt::IMAGE,
            Self::Strict => rt::STRICT_IMAGE,
        }
    }
    pub(super) fn media_rel(self, kind: Kind) -> &'static str {
        match (self, kind) {
            (Self::Transitional, Kind::Audio) => rt::AUDIO,
            (Self::Transitional, Kind::Video) => rt::VIDEO,
            (Self::Strict, Kind::Audio) => STRICT_AUDIO_REL,
            (Self::Strict, Kind::Video) => STRICT_VIDEO_REL,
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kind {
    Audio,
    Video,
}

/// Immutable media bytes with copy-free clones and slice-style access.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Data(Arc<Vec<u8>>);

impl Data {
    pub(super) fn from_shared(data: Arc<Vec<u8>>) -> Self {
        Self(data)
    }

    pub(super) fn into_shared(self) -> Arc<Vec<u8>> {
        self.0
    }

    /// Borrow the inert payload bytes.
    #[must_use]
    pub fn as_slice(&self) -> &[u8] {
        self.0.as_slice()
    }

    /// Return the payload length in bytes.
    #[must_use]
    pub fn len(&self) -> usize {
        self.0.len()
    }

    /// Return whether the payload is empty.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.0.is_empty()
    }

    /// Recover the backing vector without copying when this is its sole owner.
    ///
    /// On sharing, ownership of `self` is returned so callers can keep borrowing
    /// the bytes or deliberately choose a copy.
    pub fn try_into_vec(self) -> std::result::Result<Vec<u8>, Self> {
        Arc::try_unwrap(self.0).map_err(Self)
    }

    /// Return whether two values share the same immutable allocation.
    #[must_use]
    pub fn shares_with(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.0, &other.0)
    }
}

impl From<Vec<u8>> for Data {
    fn from(value: Vec<u8>) -> Self {
        Self(Arc::new(value))
    }
}

impl AsRef<[u8]> for Data {
    fn as_ref(&self) -> &[u8] {
        self.as_slice()
    }
}

impl std::ops::Deref for Data {
    type Target = [u8];

    fn deref(&self) -> &Self::Target {
        self.as_slice()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Resource {
    pub part_name: String,
    pub content_type: String,
    /// Stored and returned verbatim. The payload is never decoded or executed.
    /// Clones share these immutable bytes instead of copying large media parts.
    pub data: Data,
}

impl Resource {
    /// Construct an inert resource while moving its payload into shared storage.
    #[must_use]
    pub fn new(
        part_name: impl Into<String>,
        content_type: impl Into<String>,
        data: impl Into<Data>,
    ) -> Self {
        Self {
            part_name: part_name.into(),
            content_type: content_type.into(),
            data: data.into(),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Poster {
    pub relationship_id: String,
    pub resource: Option<Resource>,
}

/// A checked DrawingML transform for a media picture.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Transform {
    pub(super) x: Coordinate,
    pub(super) y: Coordinate,
    pub(super) width: Extent,
    pub(super) height: Extent,
}

impl Transform {
    /// Construct a transform from schema-checked DrawingML values.
    #[must_use]
    pub fn new(x: Coordinate, y: Coordinate, width: Extent, height: Extent) -> Self {
        Self {
            x,
            y,
            width,
            height,
        }
    }

    /// Construct a transform from EMUs with all schema bounds checked.
    pub fn emu(x: i64, y: i64, width: i64, height: i64) -> Result<Self> {
        Ok(Self::new(
            Coordinate::emu(x).map_err(|error| coordinate_error(error, "x"))?,
            Coordinate::emu(y).map_err(|error| coordinate_error(error, "y"))?,
            Extent::emu(width).map_err(|error| coordinate_error(error, "width"))?,
            Extent::emu(height).map_err(|error| coordinate_error(error, "height"))?,
        ))
    }

    /// Borrow the horizontal offset.
    pub const fn x(&self) -> &Coordinate {
        &self.x
    }

    /// Borrow the vertical offset.
    pub const fn y(&self) -> &Coordinate {
        &self.y
    }

    /// Borrow the schema-checked horizontal extent.
    pub const fn width(&self) -> &Extent {
        &self.width
    }

    /// Borrow the schema-checked vertical extent.
    pub const fn height(&self) -> &Extent {
        &self.height
    }
}

/// Typed amounts removed from the beginning and end of media playback.
///
/// The media-length-dependent sum constraint is intentionally local to
/// media-aware validation; an [`Offset`] only validates its own exact value.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Trim {
    /// Authored `st`; absence has an effective value of zero.
    pub start: Option<Offset>,
    /// Authored `end`; absence has an effective value of zero.
    pub end: Option<Offset>,
}

impl Trim {
    /// Borrow the effective start offset without erasing authored absence.
    pub fn start(&self) -> &Offset {
        self.start.as_ref().unwrap_or(&Offset::ZERO)
    }

    /// Borrow the effective end offset without erasing authored absence.
    pub fn end(&self) -> &Offset {
        self.end.as_ref().unwrap_or(&Offset::ZERO)
    }
}

/// Typed fade durations at the beginning and end of media playback.
///
/// The combined-duration constraint requires the media length and is therefore
/// kept out of the reusable [`Offset`] value.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Fade {
    /// Authored `in`; absence has an effective value of zero.
    pub fade_in: Option<Offset>,
    /// Authored `out`; absence has an effective value of zero.
    pub fade_out: Option<Offset>,
}

impl Fade {
    /// Borrow the effective fade-in duration without erasing authored absence.
    pub fn fade_in(&self) -> &Offset {
        self.fade_in.as_ref().unwrap_or(&Offset::ZERO)
    }

    /// Borrow the effective fade-out duration without erasing authored absence.
    pub fn fade_out(&self) -> &Offset {
        self.fade_out.as_ref().unwrap_or(&Offset::ZERO)
    }
}

/// A named point on a media timeline.
///
/// Time uniqueness is semantic. The upper bound against the actual media
/// length remains a media-aware check because payloads are not decoded here.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct Bookmark {
    pub name: Option<String>,
    pub time: Option<Offset>,
}

/// A bounded, canonical `p:extLst` fragment retained without interpretation.
///
/// The wrapper is validated and canonicalized while retaining QName prefixes
/// and their bindings. Extension payloads remain inert and are never loaded,
/// dispatched, or executed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ExtensionList {
    pub(crate) xml: Box<str>,
}

impl AsRef<str> for ExtensionList {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl TryFrom<&str> for ExtensionList {
    type Error = Error;

    fn try_from(value: &str) -> Result<Self> {
        Self::parse(value.as_bytes())
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Extension {
    pub embed_relationship_id: Option<String>,
    pub link_relationship_id: Option<String>,
    pub trim: Option<Trim>,
    pub fade: Option<Fade>,
    pub bookmarks: Vec<Bookmark>,
    /// Optional opaque PresentationML extension metadata, ordered last by XSD.
    pub extensions: Option<ExtensionList>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Picture {
    pub shape_id: u32,
    pub name: String,
    pub kind: Kind,
    /// The ISO/IEC 29500 `a:audioFile` or `a:videoFile` relationship identifier.
    pub relationship_id: String,
    /// Filled by package loading and required by package storage.
    pub resource: Option<Resource>,
    pub poster: Option<Poster>,
    pub transform: Option<Transform>,
    pub office_extension: Option<Extension>,
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct List {
    pub pictures: Vec<Picture>,
}
