//! Bounded authoring for existing embedded-OLE `Obj` metadata.

use crate::error::Result;

use super::{FtCmo, FtPioGrbit, OleObjectRecord};

/// A typed edit to the identity and picture flags of one existing OLE `Obj`.
///
/// The edit never creates an `Obj`, changes its payload, or changes its
/// storage reference. Unset fields retain their source values, including
/// undefined/reserved bits in the supplied flag words.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct ObjectMetadataEdit {
    object_id: Option<u16>,
    common_flags: Option<u16>,
    picture_flags: Option<FtPioGrbit>,
}

impl ObjectMetadataEdit {
    /// Starts an edit that leaves every metadata field unchanged.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            object_id: None,
            common_flags: None,
            picture_flags: None,
        }
    }

    /// Sets the `FtCmo.id` identity, preserving all other fields.
    #[must_use]
    pub const fn with_object_id(mut self, object_id: u16) -> Self {
        self.object_id = Some(object_id);
        self
    }

    /// Sets the complete `FtCmo` flags word, including caller-supplied
    /// currently undefined bits.
    #[must_use]
    pub const fn with_common_flags(mut self, flags: u16) -> Self {
        self.common_flags = Some(flags);
        self
    }

    /// Sets the complete typed `FtPioGrbit` word, including unknown bits.
    #[must_use]
    pub const fn with_picture_flags(mut self, flags: FtPioGrbit) -> Self {
        self.picture_flags = Some(flags);
        self
    }

    /// Whether this edit is a semantic no-op.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.object_id.is_none() && self.common_flags.is_none() && self.picture_flags.is_none()
    }

    pub(crate) fn apply(self, object: &mut OleObjectRecord) -> Result<()> {
        if let Some(object_id) = self.object_id {
            common_mut(object)?.object_id = object_id;
        }
        if let Some(flags) = self.common_flags {
            common_mut(object)?.flags = flags;
        }
        if let Some(picture_flags) = self.picture_flags {
            let picture = object
                .subrecords
                .iter_mut()
                .find_map(|value| match value {
                    super::ObjSubrecord::PictureFlags(value) => Some(value),
                    _ => None,
                })
                .ok_or_else(|| invalid("OLE Obj has no FtPioGrbit"))?;
            *picture = picture_flags;
        }
        object.validate()
    }
}

fn common_mut(object: &mut OleObjectRecord) -> Result<&mut FtCmo> {
    object
        .subrecords
        .iter_mut()
        .find_map(|value| match value {
            super::ObjSubrecord::Common(value) => Some(value),
            _ => None,
        })
        .ok_or_else(|| invalid("OLE Obj has no FtCmo"))
}

fn invalid(message: &str) -> crate::error::Error {
    crate::error::Error::InvalidRecord {
        record_type: super::super::OBJ,
        message: message.into(),
    }
}
