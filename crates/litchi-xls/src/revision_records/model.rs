//! Typed, inert readers for the BIFF8 shared-workbook revision records.
//!
//! These records live in the `Revision Log` stream (MS-XLS 2.1.7.14) of a
//! shared workbook. The model types below retain the existing public API;
//! wire primitives and shared invariants live in the sibling codec and
//! validation seams.

use super::codec::GUID_LEN;

/// Revision kind stored in `RRD.revt` (MS-XLS 2.5.212 `RevisionType`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum RevisionType {
    InsertRow,
    InsertColumn,
    DeleteRow,
    DeleteColumn,
    CellMove,
    InsertSheet,
    Sort,
    ChangeCell,
    RenameSheet,
    DefineName,
    Format,
    AutoFormat,
    Note,
    Header,
    Conflict,
    AddView,
    DeleteView,
    TrashQueryTableField,
}

/// Date and time of a revision action (MS-XLS 2.5.239 `ShortDTR`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ShortDtr {
    pub(super) year: u16,
    pub(super) month: u8,
    pub(super) day: u8,
    pub(super) hour: u8,
    pub(super) minute: u8,
    pub(super) second: u8,
    pub(super) weekday: u8,
}

/// Cell range used by insert/delete and move revisions (MS-XLS 2.5.209 `Ref8U`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct RevisionCellRange {
    pub(super) first_row: u16,
    pub(super) last_row: u16,
    pub(super) first_column: u16,
    pub(super) last_column: u16,
}

/// Location of a changed cell (MS-XLS 2.5.198.109 `RgceLoc`).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct RevisionCellLocation {
    pub(super) row: u16,
    pub(super) column_flags: u16,
}

/// The fixed RRD structure shared by all revision records (MS-XLS 2.5.220).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct RevisionRecordHeader {
    pub(super) memory_size: u32,
    pub(super) revision_id: i32,
    pub(super) revision_type: RevisionType,
    pub(super) accepted: bool,
    pub(super) undo_action: bool,
    pub(super) deleted_at_edge_of_sort: bool,
    pub(super) tab_id: u16,
}

/// MS-XLS 2.4.227 `RRDInfo`: shared-workbook revision-tracking state.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RrdInfo {
    pub(super) biff_version: u16,
    pub(super) shared: bool,
    pub(super) disk_has_revisions: bool,
    pub(super) auto_delete_revisions: bool,
    pub(super) track_revisions: bool,
    pub(super) exclusive: bool,
    pub(super) guid: [u8; GUID_LEN],
    pub(super) root_guid: [u8; GUID_LEN],
    pub(super) revision_id: i32,
    pub(super) version: u32,
    pub(super) history_preserved_off: bool,
    pub(super) history_protected: bool,
    pub(super) history_interval_days: u16,
}

/// MS-XLS 2.4.116 `FileLock`: a lock held on the shared workbook.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum FileLockPurpose {
    NotLocked,
    WritingUserInfo,
    MergingRevisions,
    MakeExclusive,
    DeleteOrRename,
}

/// MS-XLS 2.4.116 `FileLock` record. Inert: reading it never acquires or
/// releases any lock.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FileLock {
    pub(super) purpose: FileLockPurpose,
    pub(super) user_name: String,
    pub(super) unused: Vec<u8>,
}

/// MS-XLS 2.4.339 `UsrExcl`: an exclusive lock on the shared workbook.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UsrExcl {
    pub(super) exclusive: bool,
    pub(super) date_time: ShortDtr,
    pub(super) user_name: String,
}

/// MS-XLS 2.4.226 `RRDHead`: metadata for one user's set of revisions.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RrdHead {
    pub(super) guid: [u8; GUID_LEN],
    pub(super) code_page: u16,
    pub(super) user_name: String,
    pub(super) saved_at: ShortDtr,
    pub(super) next_tab_id: i16,
}

/// MS-XLS 2.4.241 `RRTabId`: sheet identifiers in `BoundSheet8` order.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RrTabId {
    pub(super) sheet_ids: Vec<u16>,
}

/// MS-XLS 2.4.234 `RRDRenSheet`: old and new names of a renamed sheet.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RrdRenSheet {
    pub(super) header: RevisionRecordHeader,
    pub(super) old_name: String,
    pub(super) new_name: String,
}

/// MS-XLS 2.4.228 `RRDInsDel`: an insertion or deletion of rows or columns.
///
/// The `Ducr` undo array is preserved raw; applying it would replay formula
/// edits, which an inert reader must not do.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RrdInsDel {
    pub(super) header: RevisionRecordHeader,
    pub(super) end_of_list: bool,
    pub(super) range: RevisionCellRange,
    pub(super) undo_count: u32,
    pub(super) undo_data: Vec<u8>,
}

/// MS-XLS 2.4.231 `RRDMove`: a moved cell range.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RrdMove {
    pub(super) header: RevisionRecordHeader,
    pub(super) source: RevisionCellRange,
    pub(super) destination: RevisionCellRange,
    pub(super) source_tab_id: u16,
    pub(super) undo_count: u32,
    pub(super) undo_data: Vec<u8>,
}

/// MS-XLS 2.4.239 `RRInsertSh`: an inserted sheet.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RrInsertSh {
    pub(super) header: RevisionRecordHeader,
    pub(super) position: u16,
    pub(super) name: String,
}

/// Kind of cell contents recorded by an `RRDChgCell` revision.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum RevisionCellContent {
    Blank,
    RkNumber,
    Xnum,
    RichExtendedString,
    BoolError,
    Formula,
}

/// MS-XLS 2.4.223 `RRDChgCell`: a cell-change revision.
///
/// The variable tail — optional DXFN differential formats followed by the old
/// and new cell values (RK numbers, Xnum doubles, rich extended strings,
/// Boolean/error values, or parsed formulas) — is preserved raw. `cbOldVal`
/// is validated against the MS-XLS size table, but the values themselves are
/// not decoded: rich strings and `CellParsedFormula` token arrays are
/// variable-length structures specified outside the revision record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RrdChgCell {
    pub(super) header: RevisionRecordHeader,
    pub(super) new_content: RevisionCellContent,
    pub(super) old_content: RevisionCellContent,
    pub(super) lotus_prefix: bool,
    pub(super) has_old_format: bool,
    pub(super) old_format_empty: bool,
    pub(super) reset_to_style_format: bool,
    pub(super) clear_style_format: bool,
    pub(super) has_new_format: bool,
    pub(super) new_format_empty: bool,
    pub(super) display_format: u8,
    pub(super) phonetic_shown: bool,
    pub(super) old_phonetic_shown: bool,
    pub(super) formula_adjusted: bool,
    pub(super) location: RevisionCellLocation,
    pub(super) old_value_size: u32,
    pub(super) formatting_run_count: u16,
    pub(super) tail: Vec<u8>,
}

/// MS-XLS 2.4.224 `RRDConflict`: resolution of a conflict between revisions.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RrdConflict {
    pub(super) header: RevisionRecordHeader,
}

/// MS-XLS 2.4.237 `RRDUserView`: a custom-view revision.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RrdUserView {
    pub(super) header: RevisionRecordHeader,
    pub(super) guid: [u8; GUID_LEN],
}
