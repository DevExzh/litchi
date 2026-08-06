//! Typed semantic values for a DOC embedded OLE object.

use litchi_ole_common::object::{Editor as ObjectEditor, Limits};

/// Typed MS-DOC `ObjInfo` metadata for an embedded-object storage.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Info {
    /// Whether the optional `ODTPersist2` member was present in the stream.
    ///
    /// The three `*_enhanced_metafile` fields below remain flattened for API
    /// compatibility with the original `Info` facade.  This flag preserves a
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

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct WriteOptions {
    pub storage_id: u32,
    pub instruction: String,
    /// Complete PICFAndOfficeArtData block for the Data stream.
    pub picture_data: Vec<u8>,
    /// Standalone CFB to install as `ObjectPool/_<storage_id>`.
    pub compound_file: Vec<u8>,
}

impl WriteOptions {
    pub fn new(storage_id: u32, compound_file: Vec<u8>, picture_data: Vec<u8>) -> Self {
        Self {
            storage_id,
            instruction: format!(" EMBED LITCHI_OBJECT _{storage_id} "),
            picture_data,
            compound_file,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Reference {
    pub storage_id: u32,
    pub storage_name: String,
    pub start_cp: u32,
    pub separator_cp: u32,
    pub end_cp: u32,
    pub data_offset: u32,
}

#[derive(Clone, Debug)]
pub(super) struct RawPiece {
    pub(super) start: u32,
    pub(super) end: u32,
    pub(super) fc: u32,
    pub(super) unicode: bool,
    pub(super) pcd_prefix: [u8; 2],
    pub(super) prm: [u8; 2],
}

#[derive(Clone, Debug)]
pub(super) struct FieldMarker {
    pub(super) cp: u32,
    pub(super) descriptor: [u8; 2],
}

/// Transactional editor for the DOC field and ObjectPool owner.
#[derive(Clone)]
pub struct Editor {
    pub(super) package: ObjectEditor,
    pub(super) object_pool_exists: bool,
    pub(super) limits: Limits,
    pub(super) word_path: Vec<String>,
    pub(super) table_path: Vec<String>,
    pub(super) data_path: Vec<String>,
    pub(super) word: Vec<u8>,
    pub(super) table: Vec<u8>,
    pub(super) data: Vec<u8>,
    pub(super) pieces: Vec<RawPiece>,
    pub(super) fields: Vec<FieldMarker>,
    pub(super) main_ccp: u32,
    pub(super) changed: bool,
}
