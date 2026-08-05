//! BIFF8 internal defined-name (`Lbl`) parsing and public models.

mod codec;
mod model;

#[cfg(test)]
mod tests;

pub(crate) const LBL_RECORD_TYPE: u16 = 0x0018;
pub(crate) const NAME_PUBLISH_RECORD_TYPE: u16 = 0x0893;
pub(crate) const NAME_CMT_RECORD_TYPE: u16 = 0x0894;
pub(crate) const NAME_FN_GRP12_RECORD_TYPE: u16 = 0x0899;

const LBL_HEADER_LEN: usize = 14;
const FLAG_HIDDEN: u16 = 0x0001;
const FLAG_FUNCTION: u16 = 0x0002;
const FLAG_VBA: u16 = 0x0004;
const FLAG_PROCEDURE: u16 = 0x0008;
const FLAG_BUILT_IN: u16 = 0x0020;
const RESERVED_FLAG_MASK: u16 = 0x9000;

pub(crate) use codec::unicode_name_eq;
pub(crate) use model::DefinedNameSlot;
pub use model::{
    BuiltInName, DefinedName, DefinedNameFutureRecords, DefinedNameKind, NameFnGrp12, NamePublish,
    NameScope,
};
