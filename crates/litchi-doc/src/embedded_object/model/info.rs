//! Typed `ObjInfo` metadata for one embedded OLE object.

/// Typed MS-DOC `ObjInfo` metadata for an embedded-object storage.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Info {
    /// Whether the optional `ODTPersist2` member was present in the stream.
    ///
    /// The three `*_enhanced_metafile` fields below remain flattened for API
    /// compatibility with the original `Info` facade. This flag preserves a
    /// present-but-zero `ODTPersist2`, which is distinct from a four-byte ODT.
    pub persist2_present: bool,
    pub default_handler: bool,
    pub linked: bool,
    pub display_as_icon: bool,
    pub ole1: bool,
    pub manual_update: bool,
    pub recompose_on_resize: bool,
    pub activex: bool,
    pub stream_control: bool,
    pub view_object: bool,
    pub enhanced_metafile: bool,
    pub queried_enhanced_metafile: bool,
    pub stored_as_enhanced_metafile: bool,
    pub clipboard_format: u16,
    /// Undefined, ignorable bits retained from `ODTPersist1`.
    ///
    /// The MUST-be-zero bits 10 and 11 are validated separately and are not
    /// included here.
    pub reserved_persist1: u16,
    /// Undefined, ignorable bits retained from the optional `ODTPersist2`.
    ///
    /// Bit 1 is a MUST-be-zero bit and is validated separately.
    pub reserved_persist2: u16,
}
