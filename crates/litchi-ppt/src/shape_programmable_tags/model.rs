use crate::text_extensions::{TextStyleExtension9, TextStyleExtension10, TextStyleExtension11};

/// Resource limits for shape programmable-tag parsing and serialization.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ShapeProgrammableTagLimits {
    /// Maximum `ShapeProgTagsContainer` payload size.
    pub max_container_bytes: usize,
    /// Maximum number of direct string or binary tags.
    pub max_tags: usize,
    /// Maximum payload size of one `ProgStringTag` or `ProgBinaryTag`.
    pub max_tag_bytes: usize,
    /// Maximum number of UTF-16 code units in one tag name or value.
    pub max_string_code_units: usize,
    /// Maximum known PP9/PP10/PP11 style-atom payload size.
    pub max_style_payload_bytes: usize,
    /// Maximum decoded style runs in one known extension.
    pub max_style_runs: usize,
    /// Maximum opaque payload size for an unknown binary tag.
    pub max_unknown_binary_bytes: usize,
}

impl Default for ShapeProgrammableTagLimits {
    fn default() -> Self {
        Self {
            max_container_bytes: 4 * 1024 * 1024,
            max_tags: 1024,
            max_tag_bytes: 1024 * 1024,
            max_string_code_units: 64 * 1024,
            max_style_payload_bytes: 1024 * 1024,
            max_style_runs: 64 * 1024,
            max_unknown_binary_bytes: 1024 * 1024,
        }
    }
}

/// Raw atom retained alongside a typed versioned style payload.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ShapeStyleAtom {
    /// Original record version.
    pub version: u16,
    /// Original record instance.
    pub instance: u16,
    /// Original numeric record type.
    pub record_type: u16,
    /// Original atom payload.
    pub data: Vec<u8>,
}

/// Discriminant of a shape binary programmable tag.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ShapeBinaryTagVersion {
    /// `___PPT9` / `PP9ShapeBinaryTagExtension`.
    PowerPoint9,
    /// `___PPT10` / `PP10ShapeBinaryTagExtension`.
    PowerPoint10,
    /// `___PPT11` / `PP11ShapeBinaryTagExtension`.
    PowerPoint11,
    /// Any tag name not assigned by sections 2.7.18 through 2.7.20.
    Unknown,
}

/// Typed or preserved payload of a shape binary programmable tag.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ShapeBinaryTagPayload {
    /// A `StyleTextProp9Atom` and its typed text-style data.
    PowerPoint9 {
        /// Decoded style data.
        style: TextStyleExtension9,
        /// Original style atom retained for byte-exact serialization.
        atom: ShapeStyleAtom,
    },
    /// A `StyleTextProp10Atom` and its typed text-style data.
    PowerPoint10 {
        /// Decoded style data.
        style: TextStyleExtension10,
        /// Original style atom retained for byte-exact serialization.
        atom: ShapeStyleAtom,
    },
    /// A `StyleTextProp11Atom` and its typed text-style data.
    PowerPoint11 {
        /// Decoded style data.
        style: TextStyleExtension11,
        /// Original style atom retained for byte-exact serialization.
        atom: ShapeStyleAtom,
    },
    /// An unassigned binary tag value, preserved without interpretation.
    Unknown(Vec<u8>),
}

/// One `ShapeProgBinaryTagContainer` and its CString/data-blob pair.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ShapeBinaryTag {
    /// Decoded tag name, excluding an optional terminating NUL.
    pub name: String,
    /// Typed tag-name discriminant.
    pub version: ShapeBinaryTagVersion,
    /// Typed or preserved value.
    pub payload: ShapeBinaryTagPayload,
    pub(super) name_units: Vec<u16>,
}

/// One `ProgStringTagContainer` allowed inside shape programmable tags.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ShapeStringTag {
    /// Decoded tag name, excluding an optional terminating NUL.
    pub name: String,
    /// Optional decoded Unicode value.
    pub value: Option<String>,
    pub(super) name_units: Vec<u16>,
    pub(super) value_units: Option<Vec<u16>>,
}

/// Direct child of a `ShapeProgTagsContainer`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ShapeProgrammableTag {
    /// Unicode name/value programmable tag.
    String(ShapeStringTag),
    /// Binary programmable tag.
    Binary(ShapeBinaryTag),
}

/// Typed shape programmable tags retained from one `OfficeArt` `ClientData`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ShapeProgrammableTags {
    /// Original `ShapeProgTagsContainer` record instance. Section 2.7.14 says
    /// this SHOULD be zero, so a nonzero value is preserved rather than rejected.
    pub instance: u16,
    /// Direct tags in file order.
    pub tags: Vec<ShapeProgrammableTag>,
}

/// Shape-level result returned by [`crate::slide::Slide`].
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ShapeProgrammableTagsEntry {
    /// `OfficeArt` shape identifier.
    pub shape_id: u32,
    /// Programmable tags owned by that shape.
    pub programmable_tags: ShapeProgrammableTags,
}

/// Presentation-level shape programmable-tag result.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PresentationShapeProgrammableTagsEntry {
    /// One-based slide number.
    pub slide_number: usize,
    /// `OfficeArt` shape identifier.
    pub shape_id: u32,
    /// Programmable tags owned by that shape.
    pub programmable_tags: ShapeProgrammableTags,
}

impl ShapeProgrammableTags {
    /// Return the decoded `PowerPoint` 9 style payload, when present.
    #[must_use]
    pub fn powerpoint9(&self) -> Option<&TextStyleExtension9> {
        self.tags.iter().find_map(|tag| match tag {
            ShapeProgrammableTag::Binary(ShapeBinaryTag {
                payload: ShapeBinaryTagPayload::PowerPoint9 { style, .. },
                ..
            }) => Some(style),
            ShapeProgrammableTag::String(_) | ShapeProgrammableTag::Binary(_) => None,
        })
    }

    /// Return the decoded `PowerPoint` 10 style payload, when present.
    #[must_use]
    pub fn powerpoint10(&self) -> Option<&TextStyleExtension10> {
        self.tags.iter().find_map(|tag| match tag {
            ShapeProgrammableTag::Binary(ShapeBinaryTag {
                payload: ShapeBinaryTagPayload::PowerPoint10 { style, .. },
                ..
            }) => Some(style),
            ShapeProgrammableTag::String(_) | ShapeProgrammableTag::Binary(_) => None,
        })
    }

    /// Return the decoded `PowerPoint` 11 style payload, when present.
    #[must_use]
    pub fn powerpoint11(&self) -> Option<&TextStyleExtension11> {
        self.tags.iter().find_map(|tag| match tag {
            ShapeProgrammableTag::Binary(ShapeBinaryTag {
                payload: ShapeBinaryTagPayload::PowerPoint11 { style, .. },
                ..
            }) => Some(style),
            ShapeProgrammableTag::String(_) | ShapeProgrammableTag::Binary(_) => None,
        })
    }
}
