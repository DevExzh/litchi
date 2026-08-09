//! Internal snapshot state for the transactional embedded-object editor.

use super::super::Limits;
use litchi_ole_common::object::Editor as ObjectEditor;

#[derive(Clone, Debug)]
pub(in crate::embedded_object) struct RawPiece {
    pub(in crate::embedded_object) start: u32,
    pub(in crate::embedded_object) end: u32,
    pub(in crate::embedded_object) fc: u32,
    pub(in crate::embedded_object) unicode: bool,
    pub(in crate::embedded_object) pcd_prefix: [u8; 2],
    pub(in crate::embedded_object) prm: [u8; 2],
}

#[derive(Clone, Debug)]
pub(in crate::embedded_object) struct FieldMarker {
    pub(in crate::embedded_object) cp: u32,
    pub(in crate::embedded_object) descriptor: [u8; 2],
}

/// Transactional editor for the DOC field and `ObjectPool` owner.
#[derive(Clone)]
pub struct Editor {
    pub(in crate::embedded_object) package: ObjectEditor,
    pub(in crate::embedded_object) object_pool_exists: bool,
    pub(in crate::embedded_object) limits: Limits,
    pub(in crate::embedded_object) word_path: Vec<String>,
    pub(in crate::embedded_object) table_path: Vec<String>,
    pub(in crate::embedded_object) data_path: Vec<String>,
    pub(in crate::embedded_object) word: Vec<u8>,
    pub(in crate::embedded_object) table: Vec<u8>,
    pub(in crate::embedded_object) data: Vec<u8>,
    pub(in crate::embedded_object) pieces: Vec<RawPiece>,
    pub(in crate::embedded_object) fields: Vec<FieldMarker>,
    pub(in crate::embedded_object) main_ccp: u32,
    pub(in crate::embedded_object) changed: bool,
}
