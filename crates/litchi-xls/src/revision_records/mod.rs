//! Layered BIFF8 revision-record facade.
//!
//! The model owner keeps the contextual typed record views and their parsing
//! entry points, while the binary and validation seams own wire primitives and
//! shared invariants. This keeps the crate-private revision-log integration
//! stable without exposing implementation modules as part of the API.

mod codec;
mod model;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{
    FileLock, FileLockPurpose, RevisionCellContent, RevisionCellLocation, RevisionCellRange,
    RevisionRecordHeader, RevisionType, RrInsertSh, RrTabId, RrdChgCell, RrdConflict, RrdHead,
    RrdInfo, RrdInsDel, RrdMove, RrdRenSheet, RrdUserView, ShortDtr, UsrExcl,
};
pub use transaction::{Commit, Patch, RevisionFlags, Snapshot, Transaction};

pub(crate) use codec::{
    CONTINUE_RECORD_TYPE, EOF_RECORD_TYPE, FILE_LOCK_RECORD_TYPE, NOTE_RECORD_TYPE,
    RECORD_HEADER_LEN, RR_AUTO_FMT_RECORD_TYPE, RR_FORMAT_RECORD_TYPE, RR_INSERT_SH_RECORD_TYPE,
    RR_TAB_ID_RECORD_TYPE, RRD_CHG_CELL_RECORD_TYPE, RRD_CONFLICT_RECORD_TYPE,
    RRD_DEF_NAME_RECORD_TYPE, RRD_HEAD_RECORD_TYPE, RRD_INFO_RECORD_TYPE,
    RRD_INS_DEL_BEGIN_RECORD_TYPE, RRD_INS_DEL_END_RECORD_TYPE, RRD_INS_DEL_RECORD_TYPE,
    RRD_MOVE_BEGIN_RECORD_TYPE, RRD_MOVE_END_RECORD_TYPE, RRD_MOVE_RECORD_TYPE,
    RRD_REN_SHEET_RECORD_TYPE, RRD_RST_ETXP_RECORD_TYPE, RRD_TQSIF_RECORD_TYPE,
    RRD_USER_VIEW_RECORD_TYPE, USR_EXCL_RECORD_TYPE,
};

pub(crate) use validation::validate_empty_marker;
