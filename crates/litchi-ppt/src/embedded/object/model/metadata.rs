//! Fixed-width OLE object metadata and semantic enum domains.

pub(crate) const MAX_OLE_NAME_UNITS: usize = 32_768;
pub(crate) const MAX_METAFILE_BYTES: usize = 64 * 1_048_576;
pub(crate) const MAX_OLE_OBJECTS: usize = 4_096;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[repr(u32)]
pub enum DrawAspect {
    Content = 1,
    Thumbnail = 2,
    Icon = 4,
    DocumentPrint = 8,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[repr(u32)]
pub enum ObjectType {
    Embedded = 0,
    Linked = 1,
    ActiveXControl = 2,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[repr(u32)]
pub enum ObjectSubtype {
    Default = 0,
    ClipArtGallery = 1,
    WordTable = 2,
    Excel = 3,
    Graph = 4,
    OrganizationChart = 5,
    Equation = 6,
    WordArt = 7,
    Sound = 8,
    Image = 9,
    Presentation = 10,
    Slide = 11,
    Project = 12,
    NoteIt = 13,
    ExcelChart = 14,
    MediaPlayer = 15,
}

/// The exact 24-byte payload of an `ExOleObjAtom`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Metadata {
    pub draw_aspect: DrawAspect,
    pub object_type: ObjectType,
    pub id: u32,
    pub subtype: ObjectSubtype,
    pub persist_id: u32,
    pub unused: [u8; 4],
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[repr(u32)]
pub enum ColorFollow {
    None = 0,
    EntireScheme = 1,
    TextAndBackground = 2,
}

/// The recommendation-level dimension policy preserves producer-defined bytes.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DimensionPolicy {
    Send,
    Omit,
    ProducerDefined(u8),
}

/// The exact eight-byte payload of an `ExOleEmbedAtom`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct EmbedPreferences {
    pub color_follow: ColorFollow,
    pub cannot_lock_server: bool,
    pub dimension_policy: DimensionPolicy,
    pub is_word_table: bool,
    pub unused: u8,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[repr(u32)]
pub enum UpdateMode {
    Always = 0,
    OnCall = 1,
}

/// Inert link metadata. No link is followed by this type.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct LinkInfo {
    pub slide_id: Option<u32>,
    pub update_mode: UpdateMode,
    pub unused: [u8; 4],
}
