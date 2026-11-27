//! Legacy Word form-field binary data.
//!
//! This facade keeps the public API focused on typed, inert metadata while
//! the MS-DOC wire representation, semantic validation, and focused tests
//! live in separate layers.
//!
//! [MS-DOC] section 2.9.158 defines `NilPICFAndBinData`, the binary payload
//! stored in the Data stream for the picture character (U+0001) of a
//! FORMTEXT, FORMCHECKBOX, or FORMDROPDOWN field. Its `FFData` payload is
//! decoded without resolving or invoking stored macro names, filling forms,
//! changing field state, or refreshing fields.
//!
//! The drop-down item list is always the inline hsttbDropList STTB inside
//! `FFData`. The FIB fcFormFldSttbs/lcbFormFldSttbs pair (table pointer index
//! 45) is undefined and MUST be ignored, with a mandated length of zero, so
//! no table-stream fallback path exists here.

mod codec;
mod model;
mod validation;

#[cfg(test)]
mod tests;

pub use model::NilPicfAndBinData;
pub use model::{CheckBoxState, FormFieldData, FormFieldDataKind, FormFieldTextKind};
