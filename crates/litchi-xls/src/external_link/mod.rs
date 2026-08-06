//! Layered BIFF8 supporting-book links and external cell caches.
//!
//! The semantic model is kept separate from BIFF payload codecs and the
//! workbook-level record sequence collector. Targets remain inert metadata;
//! no external path is resolved or opened.

mod codec;
mod edit;
mod model;
mod package;
mod validation;

#[cfg(test)]
mod tests;

pub(crate) const EXTERN_SHEET_RECORD_TYPE: u16 = 0x0017;
pub(crate) const XCT_RECORD_TYPE: u16 = 0x0059;
pub(crate) const CRN_RECORD_TYPE: u16 = 0x005a;
pub(crate) const SUP_BOOK_RECORD_TYPE: u16 = 0x01ae;
pub(crate) const EXTERN_NAME_RECORD_TYPE: u16 = 0x0023;
pub(crate) const CONTINUE_RECORD_TYPE: u16 = 0x003c;

pub(super) const MAX_SUPPORTING_BOOKS: usize = 1024;
pub(super) const MAX_EXTERNAL_SHEETS: usize = 256;
pub(super) const MAX_EXTERNAL_REFERENCES: usize = 1370;
pub(super) const MAX_CACHED_CELLS: usize = 65_536;
pub(super) const MAX_EXTERNAL_NAMES: usize = 4096;
pub(super) const MAX_EXTERNAL_NAME_BYTES: usize = 1_048_576;
pub(super) const MAX_DDE_OLE_VALUES: usize = 65_536;

#[allow(unused_imports, unreachable_pub)]
pub use edit::{Commit, Patch, Snapshot, Transaction};
pub use model::{
    CacheRow, CachedValue, ClipboardFormat, ErrorValue, Links, Name, NameBody, Sheet,
    SheetReference, SupportingBook, ValueMatrix, Workbook,
};
pub(crate) use package::ExternalLinkCollector;
