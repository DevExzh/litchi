//! Layered BIFF8 `WebPub` owner.
//!
//! `WebPub` records appear in the workbook globals substream and in
//! worksheet substreams (`*WebPub` in both `WORKBOOKCONTENT` and
//! `WORKSHEETCONTENT`, MS-XLS 2.1). The semantic model, binary codec, and
//! regression coverage are kept in separate layers while this module keeps
//! the existing crate-level facade intact.
//!
//! Everything in this owner is inert: URLs and file paths are stored
//! verbatim and are never opened, resolved, or fetched.

mod codec;
mod edit;
mod model;

#[cfg(test)]
mod tests;

use crate::Error;

/// Record type of the `WebPub` record (MS-XLS 2.4.344).
pub(crate) const WEB_PUB_RECORD_TYPE: u16 = 0x0801;

// `FrtRefHeaderU.grbitFrt` bits (MS-XLS 2.5.135 FrtFlags).
const FRT_REF: u16 = 0x0001;

// Flag word bits following `tws`/`twd`.
const AUTO_REPUBLISH: u16 = 0x0002;
const MHTML: u16 = 0x0008;

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type: WEB_PUB_RECORD_TYPE,
        message: message.into(),
    }
}

#[allow(unused_imports, unreachable_pub)]
pub use edit::{Commit, Patch, Snapshot, Transaction};
pub use model::{WebPageType, WebPub, WebPubRange, WebSourceType};
