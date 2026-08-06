use super::super::{AnimationInfo, SlideAnimationExtension};
use crate::embedded::object::editor::Editor as ObjectEditor;
use std::collections::BTreeSet;

pub(super) const ESCHER_SP_CONTAINER: u16 = 0xF004;
pub(super) const ESCHER_SP: u16 = 0xF00A;
pub(super) const ESCHER_CLIENT_DATA: u16 = 0xF011;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Scope {
    Slide,
    MainMaster,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct EditorLimits {
    pub max_persist_records: usize,
    pub max_record_bytes: usize,
    pub max_timeline_nodes: usize,
    pub max_timeline_depth: usize,
    pub max_build_entries: usize,
    pub max_shapes: usize,
}

impl Default for EditorLimits {
    fn default() -> Self {
        Self {
            max_persist_records: 65_536,
            max_record_bytes: 64 * 1024 * 1024,
            max_timeline_nodes: 65_536,
            max_timeline_depth: 128,
            max_build_entries: 65_536,
            max_shapes: 65_536,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct Timeline {
    pub persist_id: u32,
    pub scope: Scope,
    pub extension: SlideAnimationExtension,
}

#[derive(Clone, Debug)]
pub struct LegacyShapeAnimation {
    pub persist_id: u32,
    pub scope: Scope,
    pub shape_id: u32,
    pub animation: AnimationInfo,
}

#[derive(Clone)]
pub(super) struct PersistAnimation {
    pub(super) persist_id: u32,
    pub(super) scope: Scope,
    pub(super) record: Vec<u8>,
    pub(super) extension_payload: Option<Vec<u8>>,
    pub(super) extension: SlideAnimationExtension,
    pub(super) shape_ids: BTreeSet<u32>,
    pub(super) legacy: Vec<LegacyShapeAnimation>,
}

#[derive(Clone)]
pub struct Editor {
    pub(super) package: ObjectEditor,
    pub(super) entries: Vec<PersistAnimation>,
    pub(super) limits: EditorLimits,
    pub(super) changed: bool,
}
