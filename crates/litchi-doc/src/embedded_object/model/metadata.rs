//! Contextual metadata for one DOC embedded-object storage.

use super::{CompObj, Info, Ole, Unknown};

/// Passive metadata discovered below one managed DOC embedded object.
///
/// Each recognized stream is decoded independently. A malformed recognized
/// stream is retained as [`Unknown`] instead of being activated or discarded;
/// unrelated streams are retained the same way.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Metadata {
    pub(crate) class_id: Option<String>,
    pub(crate) comp_obj: Option<CompObj>,
    pub(crate) ole: Option<Ole>,
    pub(crate) obj_info: Option<Info>,
    pub(crate) unknown: Vec<Unknown>,
}

impl Metadata {
    /// The `ObjectPool` storage CLSID, when the CFB directory declared one.
    #[must_use]
    pub fn class_id(&self) -> Option<&str> {
        self.class_id.as_deref()
    }

    /// Typed `\x01CompObj` metadata, when present and valid.
    #[must_use]
    pub const fn comp_obj(&self) -> Option<&CompObj> {
        self.comp_obj.as_ref()
    }

    /// Typed `\x01Ole` metadata, when present and valid.
    #[must_use]
    pub const fn ole(&self) -> Option<&Ole> {
        self.ole.as_ref()
    }

    /// Typed `\x03ObjInfo` / `ODT` metadata, when present and valid.
    #[must_use]
    pub const fn obj_info(&self) -> Option<&Info> {
        self.obj_info.as_ref()
    }

    /// Whether the DOC `ObjInfo` flags identify an `ActiveX` control.
    #[must_use]
    pub fn is_activex(&self) -> bool {
        self.obj_info.as_ref().is_some_and(|info| info.activex)
    }

    /// Whether the OLE metadata identifies a linked object.
    #[must_use]
    pub fn is_linked(&self) -> bool {
        self.obj_info.as_ref().is_some_and(|info| info.linked)
            || self.ole.as_ref().is_some_and(|ole| ole.kind().is_linked())
    }

    /// Unknown or malformed descendant streams, in stable CFB discovery
    /// order.
    #[must_use]
    pub fn unknown(&self) -> &[Unknown] {
        &self.unknown
    }

    /// Whether any stream bytes remain outside the typed views.
    #[must_use]
    pub fn has_unknown(&self) -> bool {
        !self.unknown.is_empty()
            || self
                .comp_obj
                .as_ref()
                .is_some_and(|value| !value.trailing().is_empty())
            || self
                .ole
                .as_ref()
                .is_some_and(|value| !value.trailing().is_empty())
    }

    pub(in crate::embedded_object) fn from_parts(
        class_id: Option<String>,
        comp_obj: Option<CompObj>,
        ole: Option<Ole>,
        obj_info: Option<Info>,
        unknown: Vec<Unknown>,
    ) -> Self {
        Self {
            class_id,
            comp_obj,
            ole,
            obj_info,
            unknown,
        }
    }
}
