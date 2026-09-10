//! Contextual, format-level metadata for Obj records and worksheet controls.

mod control;
mod metadata;
mod obj;
mod payload;

pub use control::{
    CheckState, DropDownStyle, EditBoxValidation, FormControl, FtCblsData, FtEdoData, FtGboData,
    FtLbsData, FtRboData, FtSbs, LbsDropData, LbsItem, ListBehaviorClass, ListSelectionType,
};
pub use metadata::ObjectMetadataEdit;
pub use obj::{FtCf, FtCmo, FtPictFmla, FtPioGrbit, ObjSubrecord, ObjectType, OleObjectRecord};
pub(crate) use payload::validate_compound_file_for_publication;
pub use payload::{EmbeddedObjectDraft, EmbeddedPayload};
