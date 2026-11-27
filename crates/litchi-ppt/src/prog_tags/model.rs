/// Resource limits for document/slide programmable-tag parsing and serialization.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ProgTagLimits {
    /// Maximum `DocProgTagsContainer`/`SlideProgTagsContainer` payload size.
    pub max_container_bytes: usize,
    /// Maximum number of direct string or binary tags.
    pub max_tags: usize,
    /// Maximum payload size of one `ProgStringTag` or `ProgBinaryTag` container.
    pub max_tag_bytes: usize,
    /// Maximum number of UTF-16 code units in one tag name or value.
    pub max_string_code_units: usize,
    /// Maximum payload size of one `BinaryTagDataBlob`.
    pub max_binary_payload_bytes: usize,
    /// Maximum number of records inside one versioned `BinaryTagDataBlob`.
    pub max_binary_records: usize,
}

impl Default for ProgTagLimits {
    fn default() -> Self {
        Self {
            max_container_bytes: 16 * 1024 * 1024,
            max_tags: 1024,
            max_tag_bytes: 8 * 1024 * 1024,
            max_string_code_units: 64 * 1024,
            max_binary_payload_bytes: 8 * 1024 * 1024,
            max_binary_records: 64 * 1024,
        }
    }
}

/// The record family a `ProgTags` container belongs to.
///
/// The record type is identical (`RT_ProgTags`) in both scopes, but the set of
/// assigned versioned binary-tag names differs: document tags assign
/// `___PPT9` through `___PPT12` (section 2.4.23.4) while slide tags assign
/// only `___PPT9`, `___PPT10`, and `___PPT12` (section 2.5.22). Any other
/// name, including `___PPT11` at slide scope, is an `UnknownBinaryTag`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProgTagScope {
    /// `DocProgTagsContainer` inside the `DocumentContainer` (section 2.4.23.1).
    Document,
    /// `SlideProgTagsContainer` inside a slide, notes, handout, or main-master
    /// container (section 2.5.19).
    Slide,
}

/// Discriminant of a document/slide binary programmable tag.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProgBinaryTagVersion {
    /// `___PPT9` / `PP9DocBinaryTagExtension` or `PP9SlideBinaryTagExtension`.
    PowerPoint9,
    /// `___PPT10` / `PP10DocBinaryTagExtension` or `PP10SlideBinaryTagExtension`.
    PowerPoint10,
    /// `___PPT11` / `PP11DocBinaryTagExtension` (document scope only).
    PowerPoint11,
    /// `___PPT12` / `PP12DocBinaryTagExtension` or `PP12SlideBinaryTagExtension`.
    PowerPoint12,
    /// Any tag name not assigned by section 2.4.23.4 or 2.5.22 for the scope.
    Unknown,
}

/// One `ProgStringTagContainer` (section 2.11.30) and its name/value pair.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProgStringTag {
    /// Decoded tag name, excluding an optional terminating NUL.
    pub name: String,
    /// Optional decoded Unicode value.
    pub value: Option<String>,
    pub(super) name_units: Vec<u16>,
    pub(super) value_units: Option<Vec<u16>>,
}

/// One `DocProgBinaryTagContainer`/`SlideProgBinaryTagContainer` record pair.
///
/// The `BinaryTagDataBlob` payload is retained byte-for-byte in `payload`.
/// For versioned tags the payload is validated as a strict record sequence at
/// parse time; use [`ProgBinaryTag::records`] to decode it into
/// typed records. Unknown tags are preserved without any interpretation, as
/// required by sections 2.4.23.4 and 2.5.22.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProgBinaryTag {
    /// Decoded tag name, excluding an optional terminating NUL.
    pub name: String,
    /// Typed tag-name discriminant for the container scope.
    pub version: ProgBinaryTagVersion,
    /// Raw `BinaryTagDataBlob` payload, preserved for byte-exact serialization.
    pub payload: Vec<u8>,
    pub(super) name_units: Vec<u16>,
}

/// Direct child of a `DocProgTagsContainer`/`SlideProgTagsContainer`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ProgTag {
    /// Unicode name/value programmable tag.
    String(ProgStringTag),
    /// Binary programmable tag.
    Binary(ProgBinaryTag),
}

/// Typed programmable tags of one document- or slide-level `ProgTags` container.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ProgTags {
    /// The record family this container belongs to.
    pub scope: ProgTagScope,
    /// Original container record instance. Sections 2.4.23.1 and 2.5.19 say
    /// this SHOULD be zero, so a nonzero value is preserved rather than rejected.
    pub instance: u16,
    /// Direct tags in file order.
    pub tags: Vec<ProgTag>,
}
