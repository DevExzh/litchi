//! Binary BIFF8 codecs for typed shared-workbook revision records.

use crate::{Error, Result};

use super::{
    model::{
        FileLock, FileLockPurpose, RevisionCellContent, RevisionCellLocation, RevisionCellRange,
        RevisionRecordHeader, RevisionType, RrInsertSh, RrTabId, RrdChgCell, RrdConflict, RrdHead,
        RrdInfo, RrdInsDel, RrdMove, RrdRenSheet, RrdUserView, ShortDtr, UsrExcl,
    },
    validation::validate_sheet_name_chars,
};

/// MS-XLS 2.4.226 `RRDHead` record type (record enumeration value 312).
pub(crate) const RRD_HEAD_RECORD_TYPE: u16 = 0x0138;
/// MS-XLS 2.4.228 `RRDInsDel` record type (record enumeration value 311).
pub(crate) const RRD_INS_DEL_RECORD_TYPE: u16 = 0x0137;
/// MS-XLS 2.4.231 `RRDMove` record type (record enumeration value 320).
pub(crate) const RRD_MOVE_RECORD_TYPE: u16 = 0x0140;
/// MS-XLS 2.4.223 `RRDChgCell` record type (record enumeration value 315).
pub(crate) const RRD_CHG_CELL_RECORD_TYPE: u16 = 0x013B;
/// MS-XLS 2.4.241 `RRTabId` record type (record enumeration value 317).
pub(crate) const RR_TAB_ID_RECORD_TYPE: u16 = 0x013D;
/// MS-XLS 2.4.234 `RRDRenSheet` record type (record enumeration value 318).
pub(crate) const RRD_REN_SHEET_RECORD_TYPE: u16 = 0x013E;
/// MS-XLS 2.4.238 `RRFormat` record type (record enumeration value 330).
pub(crate) const RR_FORMAT_RECORD_TYPE: u16 = 0x014A;
/// MS-XLS 2.4.222 `RRAutoFmt` record type (record enumeration value 331).
pub(crate) const RR_AUTO_FMT_RECORD_TYPE: u16 = 0x014B;
/// MS-XLS 2.4.239 `RRInsertSh` record type (record enumeration value 333).
pub(crate) const RR_INSERT_SH_RECORD_TYPE: u16 = 0x014D;
/// MS-XLS 2.4.232 `RRDMoveBegin` record type (record enumeration value 334).
pub(crate) const RRD_MOVE_BEGIN_RECORD_TYPE: u16 = 0x014E;
/// MS-XLS 2.4.233 `RRDMoveEnd` record type (record enumeration value 335).
pub(crate) const RRD_MOVE_END_RECORD_TYPE: u16 = 0x014F;
/// MS-XLS 2.4.229 `RRDInsDelBegin` record type (record enumeration value 336).
pub(crate) const RRD_INS_DEL_BEGIN_RECORD_TYPE: u16 = 0x0150;
/// MS-XLS 2.4.230 `RRDInsDelEnd` record type (record enumeration value 337).
pub(crate) const RRD_INS_DEL_END_RECORD_TYPE: u16 = 0x0151;
/// MS-XLS 2.4.224 `RRDConflict` record type (record enumeration value 338).
pub(crate) const RRD_CONFLICT_RECORD_TYPE: u16 = 0x0152;
/// MS-XLS 2.4.225 `RRDDefName` record type (record enumeration value 339).
pub(crate) const RRD_DEF_NAME_RECORD_TYPE: u16 = 0x0153;
/// MS-XLS 2.4.235 `RRDRstEtxp` record type (record enumeration value 340).
pub(crate) const RRD_RST_ETXP_RECORD_TYPE: u16 = 0x0154;
/// MS-XLS 2.4.339 `UsrExcl` record type (record enumeration value 404).
pub(crate) const USR_EXCL_RECORD_TYPE: u16 = 0x0194;
/// MS-XLS 2.4.116 `FileLock` record type (record enumeration value 405).
pub(crate) const FILE_LOCK_RECORD_TYPE: u16 = 0x0195;
/// MS-XLS 2.4.227 `RRDInfo` record type (record enumeration value 406).
pub(crate) const RRD_INFO_RECORD_TYPE: u16 = 0x0196;
/// MS-XLS 2.4.237 `RRDUserView` record type (record enumeration value 428).
pub(crate) const RRD_USER_VIEW_RECORD_TYPE: u16 = 0x01AC;
/// MS-XLS 2.4.236 `RRDTQSIF` record type (record enumeration value 2056).
pub(crate) const RRD_TQSIF_RECORD_TYPE: u16 = 0x0808;
/// MS-XLS 2.4.179 `Note` record type (record enumeration value 28).
pub(crate) const NOTE_RECORD_TYPE: u16 = 0x001C;
/// MS-XLS 2.4.92 `EOF` record type (record enumeration value 10).
pub(crate) const EOF_RECORD_TYPE: u16 = 0x000A;
/// MS-XLS 2.4.58 `Continue` record type (record enumeration value 60).
pub(crate) const CONTINUE_RECORD_TYPE: u16 = 0x003C;

/// Byte length of the BIFF record header (`rt` + `cb`).
pub(crate) const RECORD_HEADER_LEN: usize = 4;
/// Byte length of the fixed RRD structure (MS-XLS 2.5.220).
pub(crate) const RRD_LEN: usize = 14;
/// Minimum legal `RRD.cbMemory` value (MS-XLS 2.5.220).
pub(crate) const RRD_MIN_MEMORY_SIZE: u32 = 26;
/// `RRDHead.rrd.cbMemory` is fixed to this sentinel and MUST be ignored.
pub(crate) const RRD_HEAD_MEMORY_SENTINEL: u32 = 0xFFFF_FFFF;
/// `RRD.tabid` value marking a revision that belongs to no specific sheet.
pub(crate) const NO_SHEET_TAB_ID: u16 = 0xFFFF;
/// Byte length of a `Ref8U` structure (MS-XLS 2.5.209).
pub(crate) const REF8U_LEN: usize = 8;
/// Maximum `Ref8U` column index (MS-XLS 2.5.209 `ColU` constraint).
pub(crate) const REF8U_MAX_COLUMN: u16 = 0x00FF;
/// Byte length of a `ShortDTR` structure (MS-XLS 2.5.239).
pub(crate) const SHORT_DTR_LEN: usize = 8;
/// Byte length of a GUID as stored in these records (MS-DTYP 2.3.4).
pub(crate) const GUID_LEN: usize = 16;
/// Maximum characters of the `RRDHead.stUser` name.
pub(crate) const RRD_HEAD_MAX_USER_CHARS: usize = 54;
/// Byte length of the fixed `RRDHead.stUser` field.
pub(crate) const RRD_HEAD_USER_FIELD_LEN: usize = 114;
/// Maximum characters of the `UsrExcl.stUser` name.
pub(crate) const USR_EXCL_MAX_USER_CHARS: usize = 0x0036;
/// Fixed character count of the `UsrExcl.stUser` field.
pub(crate) const USR_EXCL_USER_FIELD_CHARS: usize = 147;
/// Byte length of the `UsrExcl` fixed part before `stUser`.
pub(crate) const USR_EXCL_PREFIX_LEN: usize = 4 + SHORT_DTR_LEN + 2;
/// Maximum characters of the `FileLock.stUsrName` name.
pub(crate) const FILE_LOCK_MAX_USER_CHARS: usize = 52;
/// Total byte length of the `FileLock` record payload.
pub(crate) const FILE_LOCK_PAYLOAD_LEN: usize = 162;
/// Byte length of the fixed `RRDRenSheet` sheet-name fields.
pub(crate) const REN_SHEET_NAME_FIELD_LEN: usize = 255;
/// Maximum `RRDRenSheet`/`RRInsertSh` name characters when compressed.
pub(crate) const REN_SHEET_MAX_COMPRESSED_CHARS: u16 = 227;
/// Maximum `RRDRenSheet`/`RRInsertSh` name characters when UTF-16.
pub(crate) const REN_SHEET_MAX_UTF16_CHARS: u16 = 127;
/// Byte length of the fixed `RRInsertSh.stName` field.
pub(crate) const INSERT_SH_NAME_FIELD_LEN: usize = 256;
/// Maximum `RRTabId` sheet identifiers (above this the record is absent).
pub(crate) const MAX_TAB_ID_COUNT: usize = 4112;
/// Fixed `RRDInfo` payload length.
pub(crate) const RRD_INFO_PAYLOAD_LEN: usize = 50;
/// Fixed `RRDUserView` payload length.
pub(crate) const RRD_USER_VIEW_PAYLOAD_LEN: usize = RRD_LEN + GUID_LEN;
/// `RRDChgCell` fixed part: RRD + 4 flag bytes + `RgceLoc` + cbOldVal + cetxpRst.
pub(crate) const RRD_CHG_CELL_FIXED_LEN: usize = RRD_LEN + 4 + 4 + 4 + 2;
/// Minimum byte length of a `CellParsedFormula` old cell value.
pub(crate) const MIN_FORMULA_VALUE_LEN: u32 = 0x18;
/// `XLUnicodeStringNoCch` option bit selecting UTF-16 characters.
pub(crate) const STRING_HIGH_BYTE: u8 = 0x01;

fn invalid(record_type: u16, message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

fn read_u16(data: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([data[offset], data[offset + 1]])
}

fn read_i16(data: &[u8], offset: usize) -> i16 {
    i16::from_le_bytes([data[offset], data[offset + 1]])
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}

fn read_i32(data: &[u8], offset: usize) -> i32 {
    i32::from_le_bytes([
        data[offset],
        data[offset + 1],
        data[offset + 2],
        data[offset + 3],
    ])
}
/// Decode an `XLUnicodeStringNoCch` inside a fixed-size field of `field`
/// bytes: one option-flags byte followed by `cch` characters. Characters past
/// `cch` are ignored per MS-XLS 2.5.294.
fn decode_fixed_string(
    record_type: u16,
    field: &[u8],
    cch: usize,
    context: &str,
) -> Result<String> {
    let Some((&flags, characters)) = field.split_first() else {
        return Err(invalid(record_type, format!("{context} field is empty")));
    };
    if flags & !STRING_HIGH_BYTE != 0 {
        return Err(invalid(
            record_type,
            format!("{context} contains reserved string option bits"),
        ));
    }
    let wide = flags & STRING_HIGH_BYTE != 0;
    let byte_count = cch
        .checked_mul(if wide { 2 } else { 1 })
        .ok_or_else(|| invalid(record_type, format!("{context} length overflows")))?;
    let bytes = characters.get(..byte_count).ok_or_else(|| {
        invalid(
            record_type,
            format!("{context} characters exceed the fixed field"),
        )
    })?;
    if wide {
        let units = bytes
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect::<Vec<_>>();
        String::from_utf16(&units)
            .map_err(|_error| invalid(record_type, format!("{context} contains invalid UTF-16")))
    } else {
        Ok(bytes.iter().map(|&byte| char::from(byte)).collect())
    }
}

/// Validate the MS-XLS name-length split for `RRDRenSheet` and `RRInsertSh`.

impl RevisionType {
    /// Decode the MS-XLS 2.5.212 enumeration value.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn from_u16(record_type: u16, value: u16) -> Result<Self> {
        match value {
            0x0000 => Ok(Self::InsertRow),
            0x0001 => Ok(Self::InsertColumn),
            0x0002 => Ok(Self::DeleteRow),
            0x0003 => Ok(Self::DeleteColumn),
            0x0004 => Ok(Self::CellMove),
            0x0005 => Ok(Self::InsertSheet),
            0x0007 => Ok(Self::Sort),
            0x0008 => Ok(Self::ChangeCell),
            0x0009 => Ok(Self::RenameSheet),
            0x000A => Ok(Self::DefineName),
            0x000B => Ok(Self::Format),
            0x000C => Ok(Self::AutoFormat),
            0x000D => Ok(Self::Note),
            0x0020 => Ok(Self::Header),
            0x0025 => Ok(Self::Conflict),
            0x002B => Ok(Self::AddView),
            0x002C => Ok(Self::DeleteView),
            0x002E => Ok(Self::TrashQueryTableField),
            other => Err(invalid(
                record_type,
                format!("unknown revision type 0x{other:04X}"),
            )),
        }
    }

    /// The MS-XLS 2.5.212 enumeration value.
    #[must_use]
    pub fn to_u16(self) -> u16 {
        match self {
            Self::InsertRow => 0x0000,
            Self::InsertColumn => 0x0001,
            Self::DeleteRow => 0x0002,
            Self::DeleteColumn => 0x0003,
            Self::CellMove => 0x0004,
            Self::InsertSheet => 0x0005,
            Self::Sort => 0x0007,
            Self::ChangeCell => 0x0008,
            Self::RenameSheet => 0x0009,
            Self::DefineName => 0x000A,
            Self::Format => 0x000B,
            Self::AutoFormat => 0x000C,
            Self::Note => 0x000D,
            Self::Header => 0x0020,
            Self::Conflict => 0x0025,
            Self::AddView => 0x002B,
            Self::DeleteView => 0x002C,
            Self::TrashQueryTableField => 0x002E,
        }
    }
}

impl ShortDtr {
    /// Parse the fixed 8-byte structure with Gregorian calendar validation.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse(record_type: u16, data: &[u8]) -> Result<Self> {
        if data.len() != SHORT_DTR_LEN {
            return Err(invalid(
                record_type,
                format!(
                    "ShortDTR has {} bytes; expected {SHORT_DTR_LEN}",
                    data.len()
                ),
            ));
        }
        let value = Self {
            year: read_u16(data, 0),
            month: data[2],
            day: data[3],
            hour: data[4],
            minute: data[5],
            second: data[6],
            weekday: data[7],
        };
        if !(1900..=9999).contains(&value.year) {
            return Err(invalid(record_type, "ShortDTR year is out of range"));
        }
        if !(1..=12).contains(&value.month) {
            return Err(invalid(record_type, "ShortDTR month is out of range"));
        }
        if !(1..=31).contains(&value.day) {
            return Err(invalid(record_type, "ShortDTR day is out of range"));
        }
        if value.hour > 23 {
            return Err(invalid(record_type, "ShortDTR hour is out of range"));
        }
        if value.minute > 59 {
            return Err(invalid(record_type, "ShortDTR minute is out of range"));
        }
        if value.second > 59 {
            return Err(invalid(record_type, "ShortDTR second is out of range"));
        }
        if value.weekday > 7 {
            return Err(invalid(record_type, "ShortDTR weekday is out of range"));
        }
        Ok(value)
    }

    #[must_use]
    pub fn year(&self) -> u16 {
        self.year
    }
    #[must_use]
    pub fn month(&self) -> u8 {
        self.month
    }
    #[must_use]
    pub fn day(&self) -> u8 {
        self.day
    }
    #[must_use]
    pub fn hour(&self) -> u8 {
        self.hour
    }
    #[must_use]
    pub fn minute(&self) -> u8 {
        self.minute
    }
    #[must_use]
    pub fn second(&self) -> u8 {
        self.second
    }
    /// Weekday: 0 = unspecified, 1 = Monday, ..., 7 = Sunday.
    #[must_use]
    pub fn weekday(&self) -> u8 {
        self.weekday
    }
}

impl RevisionCellRange {
    /// Parse the fixed 8-byte structure, enforcing the ordering constraints.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse(record_type: u16, data: &[u8]) -> Result<Self> {
        if data.len() != REF8U_LEN {
            return Err(invalid(
                record_type,
                format!("Ref8U has {} bytes; expected {REF8U_LEN}", data.len()),
            ));
        }
        let value = Self {
            first_row: read_u16(data, 0),
            last_row: read_u16(data, 2),
            first_column: read_u16(data, 4),
            last_column: read_u16(data, 6),
        };
        if value.first_row > value.last_row {
            return Err(invalid(record_type, "Ref8U first row exceeds last row"));
        }
        if value.first_column > value.last_column {
            return Err(invalid(
                record_type,
                "Ref8U first column exceeds last column",
            ));
        }
        if value.last_column > REF8U_MAX_COLUMN {
            return Err(invalid(record_type, "Ref8U column exceeds 0x00FF"));
        }
        Ok(value)
    }

    #[must_use]
    pub fn first_row(&self) -> u16 {
        self.first_row
    }
    #[must_use]
    pub fn last_row(&self) -> u16 {
        self.last_row
    }
    #[must_use]
    pub fn first_column(&self) -> u16 {
        self.first_column
    }
    #[must_use]
    pub fn last_column(&self) -> u16 {
        self.last_column
    }
}

impl RevisionCellLocation {
    fn parse(data: &[u8]) -> Self {
        Self {
            row: read_u16(data, 0),
            column_flags: read_u16(data, 2),
        }
    }

    /// Zero-based row coordinate.
    #[must_use]
    pub fn row(&self) -> u16 {
        self.row
    }

    /// Zero-based column coordinate (the low 14 bits of the `ColRelU`).
    #[must_use]
    pub fn column(&self) -> u16 {
        self.column_flags & 0x3FFF
    }

    /// Whether the column coordinate is a relative reference.
    #[must_use]
    pub fn is_column_relative(&self) -> bool {
        self.column_flags & 0x4000 != 0
    }

    /// Whether the row coordinate is a relative reference.
    #[must_use]
    pub fn is_row_relative(&self) -> bool {
        self.column_flags & 0x8000 != 0
    }
}

impl RevisionRecordHeader {
    /// Parse the 14-byte structure. `is_head` selects the RRDHead-specific
    /// `cbMemory` rule (fixed sentinel instead of the >= 26 minimum).
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse(record_type: u16, data: &[u8], is_head: bool) -> Result<Self> {
        if data.len() < RRD_LEN {
            return Err(invalid(
                record_type,
                format!("RRD has {} bytes; expected at least {RRD_LEN}", data.len()),
            ));
        }
        let memory_size = read_u32(data, 0);
        let revision_id = read_i32(data, 4);
        let revision_type = RevisionType::from_u16(record_type, read_u16(data, 8))?;
        let flags = read_u16(data, 10);
        if flags & 0xFFF4 != 0 {
            return Err(invalid(record_type, "RRD contains reserved flag bits"));
        }
        if !is_head && memory_size < RRD_MIN_MEMORY_SIZE {
            return Err(invalid(
                record_type,
                format!("RRD cbMemory is {memory_size}; minimum is {RRD_MIN_MEMORY_SIZE}"),
            ));
        }
        if is_head && memory_size != RRD_HEAD_MEMORY_SENTINEL {
            return Err(invalid(
                record_type,
                "RRDHead cbMemory is not the 0xFFFFFFFF sentinel",
            ));
        }
        if revision_id < 0 {
            return Err(invalid(record_type, "RRD revid is negative"));
        }
        Ok(Self {
            memory_size,
            revision_id,
            revision_type,
            accepted: flags & 0x0001 != 0,
            undo_action: flags & 0x0002 != 0,
            deleted_at_edge_of_sort: flags & 0x0008 != 0,
            tab_id: read_u16(data, 12),
        })
    }

    /// In-memory size of the revision record structure (ignored for `RRDHead`).
    #[must_use]
    pub fn memory_size(&self) -> u32 {
        self.memory_size
    }
    #[must_use]
    pub fn revision_id(&self) -> i32 {
        self.revision_id
    }
    #[must_use]
    pub fn revision_type(&self) -> RevisionType {
        self.revision_type
    }
    #[must_use]
    pub fn is_accepted(&self) -> bool {
        self.accepted
    }
    #[must_use]
    pub fn is_undo_action(&self) -> bool {
        self.undo_action
    }
    #[must_use]
    pub fn is_deleted_at_edge_of_sort(&self) -> bool {
        self.deleted_at_edge_of_sort
    }
    /// Sheet the revision belongs to; `None` for the 0xFFFF "no sheet" marker.
    #[must_use]
    pub fn tab_id(&self) -> Option<u16> {
        if self.tab_id == NO_SHEET_TAB_ID {
            None
        } else {
            Some(self.tab_id)
        }
    }

    fn require_reviewable(&self, record_type: u16) -> Result<()> {
        if self.revision_id <= 0 {
            return Err(invalid(
                record_type,
                "reviewable revision has a non-positive revid",
            ));
        }
        Ok(())
    }

    fn require_sheet(&self, record_type: u16) -> Result<()> {
        if self.tab_id == NO_SHEET_TAB_ID {
            return Err(invalid(record_type, "revision does not specify a sheet"));
        }
        Ok(())
    }
}

impl RrdInfo {
    /// Parse the fixed 50-byte record payload.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse_payload(data: &[u8]) -> Result<Self> {
        if data.len() != RRD_INFO_PAYLOAD_LEN {
            return Err(invalid(
                RRD_INFO_RECORD_TYPE,
                format!(
                    "RRDInfo payload has {} bytes; expected {RRD_INFO_PAYLOAD_LEN}",
                    data.len()
                ),
            ));
        }
        if read_u16(data, 2) != 0 {
            return Err(invalid(
                RRD_INFO_RECORD_TYPE,
                "RRDInfo reserved1 is nonzero",
            ));
        }
        let flags = read_u16(data, 4);
        if flags & 0xFFE0 != 0 {
            return Err(invalid(
                RRD_INFO_RECORD_TYPE,
                "RRDInfo contains reserved flag bits",
            ));
        }
        let flags2 = read_u16(data, 44);
        if flags2 & 0xFFFC != 0 {
            return Err(invalid(
                RRD_INFO_RECORD_TYPE,
                "RRDInfo contains reserved flag bits",
            ));
        }
        let mut guid = [0u8; GUID_LEN];
        guid.copy_from_slice(&data[6..22]);
        let mut root_guid = [0u8; GUID_LEN];
        root_guid.copy_from_slice(&data[22..38]);
        let value = Self {
            biff_version: read_u16(data, 0),
            shared: flags & 0x0001 != 0,
            disk_has_revisions: flags & 0x0002 != 0,
            auto_delete_revisions: flags & 0x0004 != 0,
            track_revisions: flags & 0x0008 != 0,
            exclusive: flags & 0x0010 != 0,
            guid,
            root_guid,
            revision_id: read_i32(data, 38),
            version: read_u32(data, 42),
            history_preserved_off: flags2 & 0x0001 != 0,
            history_protected: flags2 & 0x0002 != 0,
            history_interval_days: read_u16(data, 46),
        };
        value.validate()?;
        Ok(value)
    }

    fn validate(&self) -> Result<()> {
        let invalid_info = |message: &str| invalid(RRD_INFO_RECORD_TYPE, message);
        if self.shared && self.exclusive {
            return Err(invalid_info("RRDInfo is both shared and exclusive"));
        }
        if self.track_revisions && !self.shared {
            return Err(invalid_info("RRDInfo tracks revisions without sharing"));
        }
        if self.disk_has_revisions && !self.track_revisions {
            return Err(invalid_info(
                "RRDInfo has on-disk revisions without revision tracking",
            ));
        }
        if self.auto_delete_revisions && !self.track_revisions {
            return Err(invalid_info(
                "RRDInfo auto-deletes revisions without revision tracking",
            ));
        }
        if self.revision_id < 0 {
            return Err(invalid_info("RRDInfo revid is negative"));
        }
        if self.history_interval_days > 0x7FFF {
            return Err(invalid_info("RRDInfo history interval exceeds 0x7FFF"));
        }
        if self.history_preserved_off {
            if self.history_interval_days != 0 {
                return Err(invalid_info(
                    "RRDInfo disables history with a nonzero interval",
                ));
            }
            if !self.shared {
                return Err(invalid_info("RRDInfo disables history without sharing"));
            }
        } else if !self.exclusive && self.history_interval_days == 0 {
            return Err(invalid_info(
                "RRDInfo preserves history with a zero interval",
            ));
        }
        if self.history_protected && !self.shared {
            return Err(invalid_info("RRDInfo protects history without sharing"));
        }
        Ok(())
    }

    /// Major BIFF version last used to save the shared workbook.
    #[must_use]
    pub fn biff_version(&self) -> u16 {
        self.biff_version
    }
    #[must_use]
    pub fn is_shared(&self) -> bool {
        self.shared
    }
    #[must_use]
    pub fn disk_has_revisions(&self) -> bool {
        self.disk_has_revisions
    }
    #[must_use]
    pub fn auto_delete_revisions(&self) -> bool {
        self.auto_delete_revisions
    }
    #[must_use]
    pub fn track_revisions(&self) -> bool {
        self.track_revisions
    }
    #[must_use]
    pub fn is_exclusive(&self) -> bool {
        self.exclusive
    }
    /// GUID of the most recent revision header (or all-zero).
    #[must_use]
    pub fn guid(&self) -> &[u8; GUID_LEN] {
        &self.guid
    }
    /// GUID of the last saved revision header (or all-zero).
    #[must_use]
    pub fn root_guid(&self) -> &[u8; GUID_LEN] {
        &self.root_guid
    }
    #[must_use]
    pub fn revision_id(&self) -> i32 {
        self.revision_id
    }
    #[must_use]
    pub fn version(&self) -> u32 {
        self.version
    }
    /// Whether revision history is discarded (`fNoRevHist`).
    #[must_use]
    pub fn history_disabled(&self) -> bool {
        self.history_preserved_off
    }
    #[must_use]
    pub fn history_protected(&self) -> bool {
        self.history_protected
    }
    /// Days for which revision history is kept; ignored when exclusive.
    #[must_use]
    pub fn history_interval_days(&self) -> u16 {
        self.history_interval_days
    }
}

impl FileLockPurpose {
    fn from_u32(value: u32) -> Result<Self> {
        match value {
            0x0000_0000 => Ok(Self::NotLocked),
            0x0001_0001 => Ok(Self::WritingUserInfo),
            0x0001_0002 => Ok(Self::MergingRevisions),
            0x0001_0004 => Ok(Self::MakeExclusive),
            0x0001_0008 => Ok(Self::DeleteOrRename),
            other => Err(invalid(
                FILE_LOCK_RECORD_TYPE,
                format!("FileLock has unknown purpose 0x{other:08X}"),
            )),
        }
    }
}

impl FileLock {
    /// Parse the fixed 162-byte record payload.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse_payload(data: &[u8]) -> Result<Self> {
        if data.len() != FILE_LOCK_PAYLOAD_LEN {
            return Err(invalid(
                FILE_LOCK_RECORD_TYPE,
                format!(
                    "FileLock payload has {} bytes; expected {FILE_LOCK_PAYLOAD_LEN}",
                    data.len()
                ),
            ));
        }
        let purpose = FileLockPurpose::from_u32(read_u32(data, 0))?;
        let cch = usize::from(read_u16(data, 4));
        if cch > FILE_LOCK_MAX_USER_CHARS {
            return Err(invalid(
                FILE_LOCK_RECORD_TYPE,
                format!("FileLock user name has {cch} characters; maximum is 52"),
            ));
        }
        let user_name =
            decode_fixed_string(FILE_LOCK_RECORD_TYPE, &data[6..], cch, "FileLock stUsrName")?;
        // stUsrName spans cch (2 bytes) + option flags (1 byte) + characters.
        let string_end = 6
            + 1
            + if data[6] & STRING_HIGH_BYTE != 0 {
                2 * cch
            } else {
                cch
            };
        Ok(Self {
            purpose,
            user_name,
            unused: data[string_end..].to_vec(),
        })
    }

    #[must_use]
    pub fn purpose(&self) -> FileLockPurpose {
        self.purpose
    }
    #[must_use]
    pub fn user_name(&self) -> &str {
        &self.user_name
    }
    #[must_use]
    pub fn unused_bytes(&self) -> &[u8] {
        &self.unused
    }
}

impl UsrExcl {
    /// Parse the record payload: `fExclusive`, `sdtr`, `cchUser`, and the
    /// fixed 147-character `stUser` field.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse_payload(data: &[u8]) -> Result<Self> {
        if data.len() < USR_EXCL_PREFIX_LEN + 1 + USR_EXCL_USER_FIELD_CHARS {
            return Err(invalid(
                USR_EXCL_RECORD_TYPE,
                format!(
                    "UsrExcl payload has {} bytes; expected at least {}",
                    data.len(),
                    USR_EXCL_PREFIX_LEN + 1 + USR_EXCL_USER_FIELD_CHARS
                ),
            ));
        }
        let exclusive = match read_u32(data, 0) {
            0 => false,
            1 => true,
            other => {
                return Err(invalid(
                    USR_EXCL_RECORD_TYPE,
                    format!("UsrExcl fExclusive is 0x{other:08X}; expected a Boolean"),
                ));
            },
        };
        let date_time = ShortDtr::parse(USR_EXCL_RECORD_TYPE, &data[4..12])?;
        let cch = read_u16(data, 12);
        if usize::from(cch) > USR_EXCL_MAX_USER_CHARS {
            return Err(invalid(
                USR_EXCL_RECORD_TYPE,
                format!("UsrExcl user name has {cch} characters; maximum is 54"),
            ));
        }
        let field = &data[14..];
        let wide = field[0] & STRING_HIGH_BYTE != 0;
        let field_len = 1 + USR_EXCL_USER_FIELD_CHARS * if wide { 2 } else { 1 };
        if field.len() != field_len {
            return Err(invalid(
                USR_EXCL_RECORD_TYPE,
                format!(
                    "UsrExcl stUser field has {} bytes; expected {field_len}",
                    field.len()
                ),
            ));
        }
        let user_name = decode_fixed_string(
            USR_EXCL_RECORD_TYPE,
            field,
            usize::from(cch),
            "UsrExcl stUser",
        )?;
        Ok(Self {
            exclusive,
            date_time,
            user_name,
        })
    }

    #[must_use]
    pub fn is_exclusive(&self) -> bool {
        self.exclusive
    }
    #[must_use]
    pub fn date_time(&self) -> ShortDtr {
        self.date_time
    }
    #[must_use]
    pub fn user_name(&self) -> &str {
        &self.user_name
    }
}

impl RrdHead {
    /// Parse the fixed 158-byte record payload.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse_payload(data: &[u8]) -> Result<Self> {
        const PAYLOAD_LEN: usize =
            RRD_LEN + GUID_LEN + 2 + 2 + RRD_HEAD_USER_FIELD_LEN + SHORT_DTR_LEN + 2;
        if data.len() != PAYLOAD_LEN {
            return Err(invalid(
                RRD_HEAD_RECORD_TYPE,
                format!(
                    "RRDHead payload has {} bytes; expected {PAYLOAD_LEN}",
                    data.len()
                ),
            ));
        }
        let header = RevisionRecordHeader::parse(RRD_HEAD_RECORD_TYPE, data, true)?;
        if header.revision_type != RevisionType::Header {
            return Err(invalid(
                RRD_HEAD_RECORD_TYPE,
                "RRDHead revision type is not REVTHEADER",
            ));
        }
        if header.revision_id != 0 {
            return Err(invalid(RRD_HEAD_RECORD_TYPE, "RRDHead has a nonzero revid"));
        }
        let mut guid = [0u8; GUID_LEN];
        guid.copy_from_slice(&data[RRD_LEN..RRD_LEN + GUID_LEN]);
        let cch_offset = RRD_LEN + GUID_LEN + 2;
        let cch = read_u16(data, cch_offset);
        if usize::from(cch) > RRD_HEAD_MAX_USER_CHARS {
            return Err(invalid(
                RRD_HEAD_RECORD_TYPE,
                format!("RRDHead user name has {cch} characters; maximum is 54"),
            ));
        }
        let field_offset = cch_offset + 2;
        let user_name = decode_fixed_string(
            RRD_HEAD_RECORD_TYPE,
            &data[field_offset..field_offset + RRD_HEAD_USER_FIELD_LEN],
            usize::from(cch),
            "RRDHead stUser",
        )?;
        let dtr_offset = field_offset + RRD_HEAD_USER_FIELD_LEN;
        let saved_at = ShortDtr::parse(
            RRD_HEAD_RECORD_TYPE,
            &data[dtr_offset..dtr_offset + SHORT_DTR_LEN],
        )?;
        let next_tab_id = read_i16(data, dtr_offset + SHORT_DTR_LEN);
        if next_tab_id < -1 {
            return Err(invalid(
                RRD_HEAD_RECORD_TYPE,
                "RRDHead tabidMac is less than -1",
            ));
        }
        Ok(Self {
            guid,
            code_page: read_u16(data, RRD_LEN + GUID_LEN),
            user_name,
            saved_at,
            next_tab_id,
        })
    }

    /// Identifier of this set of revisions.
    #[must_use]
    pub fn guid(&self) -> &[u8; GUID_LEN] {
        &self.guid
    }
    /// Sheet code page; 1200 means Unicode.
    #[must_use]
    pub fn code_page(&self) -> u16 {
        self.code_page
    }
    #[must_use]
    pub fn user_name(&self) -> &str {
        &self.user_name
    }
    #[must_use]
    pub fn saved_at(&self) -> ShortDtr {
        self.saved_at
    }
    /// Next available sheet identifier (`tabidMac`).
    #[must_use]
    pub fn next_tab_id(&self) -> i16 {
        self.next_tab_id
    }
}

impl RrTabId {
    /// Parse the record payload, an array of 2-byte sheet identifiers.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse_payload(data: &[u8]) -> Result<Self> {
        if !data.len().is_multiple_of(2) {
            return Err(invalid(
                RR_TAB_ID_RECORD_TYPE,
                "RRTabId payload has an odd byte count",
            ));
        }
        let sheet_ids = data
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect::<Vec<_>>();
        if sheet_ids.len() > MAX_TAB_ID_COUNT {
            return Err(invalid(
                RR_TAB_ID_RECORD_TYPE,
                format!(
                    "RRTabId has {} sheet identifiers; maximum is {MAX_TAB_ID_COUNT}",
                    sheet_ids.len()
                ),
            ));
        }
        Ok(Self { sheet_ids })
    }

    #[must_use]
    pub fn sheet_ids(&self) -> &[u16] {
        &self.sheet_ids
    }
}

impl RrdRenSheet {
    /// Parse the fixed 528-byte record payload.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse_payload(data: &[u8]) -> Result<Self> {
        const PAYLOAD_LEN: usize =
            RRD_LEN + 2 + REN_SHEET_NAME_FIELD_LEN + 2 + REN_SHEET_NAME_FIELD_LEN;
        if data.len() != PAYLOAD_LEN {
            return Err(invalid(
                RRD_REN_SHEET_RECORD_TYPE,
                format!(
                    "RRDRenSheet payload has {} bytes; expected {PAYLOAD_LEN}",
                    data.len()
                ),
            ));
        }
        let header = RevisionRecordHeader::parse(RRD_REN_SHEET_RECORD_TYPE, data, false)?;
        header.require_reviewable(RRD_REN_SHEET_RECORD_TYPE)?;
        if header.revision_type != RevisionType::RenameSheet {
            return Err(invalid(
                RRD_REN_SHEET_RECORD_TYPE,
                "RRDRenSheet revision type is not REVTRENSHEET",
            ));
        }
        header.require_sheet(RRD_REN_SHEET_RECORD_TYPE)?;
        let old_cch = read_u16(data, RRD_LEN);
        let old_field = &data[RRD_LEN + 2..RRD_LEN + 2 + REN_SHEET_NAME_FIELD_LEN];
        validate_sheet_name_chars(
            RRD_REN_SHEET_RECORD_TYPE,
            old_field,
            old_cch,
            "RRDRenSheet stOldName",
        )?;
        let old_name = decode_fixed_string(
            RRD_REN_SHEET_RECORD_TYPE,
            old_field,
            usize::from(old_cch),
            "RRDRenSheet stOldName",
        )?;
        let new_offset = RRD_LEN + 2 + REN_SHEET_NAME_FIELD_LEN;
        let new_cch = read_u16(data, new_offset);
        let new_field = &data[new_offset + 2..new_offset + 2 + REN_SHEET_NAME_FIELD_LEN];
        validate_sheet_name_chars(
            RRD_REN_SHEET_RECORD_TYPE,
            new_field,
            new_cch,
            "RRDRenSheet stNewName",
        )?;
        let new_name = decode_fixed_string(
            RRD_REN_SHEET_RECORD_TYPE,
            new_field,
            usize::from(new_cch),
            "RRDRenSheet stNewName",
        )?;
        Ok(Self {
            header,
            old_name,
            new_name,
        })
    }

    #[must_use]
    pub fn header(&self) -> &RevisionRecordHeader {
        &self.header
    }
    #[must_use]
    pub fn old_name(&self) -> &str {
        &self.old_name
    }
    #[must_use]
    pub fn new_name(&self) -> &str {
        &self.new_name
    }
}

impl RrdInsDel {
    /// Parse the record payload.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse_payload(data: &[u8]) -> Result<Self> {
        const FIXED_LEN: usize = RRD_LEN + 2 + REF8U_LEN + 4;
        if data.len() < FIXED_LEN {
            return Err(invalid(
                RRD_INS_DEL_RECORD_TYPE,
                format!(
                    "RRDInsDel payload has {} bytes; expected at least {FIXED_LEN}",
                    data.len()
                ),
            ));
        }
        let header = RevisionRecordHeader::parse(RRD_INS_DEL_RECORD_TYPE, data, false)?;
        header.require_reviewable(RRD_INS_DEL_RECORD_TYPE)?;
        match header.revision_type {
            RevisionType::InsertRow
            | RevisionType::InsertColumn
            | RevisionType::DeleteRow
            | RevisionType::DeleteColumn => {},
            _ => {
                return Err(invalid(
                    RRD_INS_DEL_RECORD_TYPE,
                    "RRDInsDel revision type is not an insert/delete",
                ));
            },
        }
        header.require_sheet(RRD_INS_DEL_RECORD_TYPE)?;
        let flags = read_u16(data, RRD_LEN);
        if flags & 0xFFFE != 0 {
            return Err(invalid(
                RRD_INS_DEL_RECORD_TYPE,
                "RRDInsDel contains reserved flag bits",
            ));
        }
        let end_of_list = flags & 0x0001 != 0;
        if end_of_list && header.revision_type != RevisionType::InsertRow {
            return Err(invalid(
                RRD_INS_DEL_RECORD_TYPE,
                "RRDInsDel fEndOfList is set for a non row-insert",
            ));
        }
        let range = RevisionCellRange::parse(
            RRD_INS_DEL_RECORD_TYPE,
            &data[RRD_LEN + 2..RRD_LEN + 2 + REF8U_LEN],
        )?;
        let undo_count = read_u32(data, RRD_LEN + 2 + REF8U_LEN);
        let undo_data = data[FIXED_LEN..].to_vec();
        if undo_count == 0 && !undo_data.is_empty() {
            return Err(invalid(
                RRD_INS_DEL_RECORD_TYPE,
                "RRDInsDel has undo bytes but a zero undo count",
            ));
        }
        if undo_count > 0 && undo_data.is_empty() {
            return Err(invalid(
                RRD_INS_DEL_RECORD_TYPE,
                "RRDInsDel is missing its Ducr undo data",
            ));
        }
        Ok(Self {
            header,
            end_of_list,
            range,
            undo_count,
            undo_data,
        })
    }

    #[must_use]
    pub fn header(&self) -> &RevisionRecordHeader {
        &self.header
    }
    #[must_use]
    pub fn is_end_of_list(&self) -> bool {
        self.end_of_list
    }
    #[must_use]
    pub fn range(&self) -> RevisionCellRange {
        self.range
    }
    #[must_use]
    pub fn undo_count(&self) -> u32 {
        self.undo_count
    }
    /// Raw `Ducr` undo array (MS-XLS 2.5.71), preserved without interpreting.
    #[must_use]
    pub fn undo_data(&self) -> &[u8] {
        &self.undo_data
    }
}

impl RrdMove {
    /// Parse the record payload.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse_payload(data: &[u8]) -> Result<Self> {
        const FIXED_LEN: usize = RRD_LEN + 2 * REF8U_LEN + 2 + 4;
        if data.len() < FIXED_LEN {
            return Err(invalid(
                RRD_MOVE_RECORD_TYPE,
                format!(
                    "RRDMove payload has {} bytes; expected at least {FIXED_LEN}",
                    data.len()
                ),
            ));
        }
        let header = RevisionRecordHeader::parse(RRD_MOVE_RECORD_TYPE, data, false)?;
        header.require_reviewable(RRD_MOVE_RECORD_TYPE)?;
        if header.revision_type != RevisionType::CellMove {
            return Err(invalid(
                RRD_MOVE_RECORD_TYPE,
                "RRDMove revision type is not REVTMOVE",
            ));
        }
        header.require_sheet(RRD_MOVE_RECORD_TYPE)?;
        let source =
            RevisionCellRange::parse(RRD_MOVE_RECORD_TYPE, &data[RRD_LEN..RRD_LEN + REF8U_LEN])?;
        let destination = RevisionCellRange::parse(
            RRD_MOVE_RECORD_TYPE,
            &data[RRD_LEN + REF8U_LEN..RRD_LEN + 2 * REF8U_LEN],
        )?;
        let source_tab_id = read_u16(data, RRD_LEN + 2 * REF8U_LEN);
        let undo_count = read_u32(data, RRD_LEN + 2 * REF8U_LEN + 2);
        let undo_data = data[FIXED_LEN..].to_vec();
        if undo_count == 0 && !undo_data.is_empty() {
            return Err(invalid(
                RRD_MOVE_RECORD_TYPE,
                "RRDMove has undo bytes but a zero undo count",
            ));
        }
        if undo_count > 0 && undo_data.is_empty() {
            return Err(invalid(
                RRD_MOVE_RECORD_TYPE,
                "RRDMove is missing its Ducr undo data",
            ));
        }
        Ok(Self {
            header,
            source,
            destination,
            source_tab_id,
            undo_count,
            undo_data,
        })
    }

    #[must_use]
    pub fn header(&self) -> &RevisionRecordHeader {
        &self.header
    }
    #[must_use]
    pub fn source(&self) -> RevisionCellRange {
        self.source
    }
    #[must_use]
    pub fn destination(&self) -> RevisionCellRange {
        self.destination
    }
    /// Sheet on which the source range resides.
    #[must_use]
    pub fn source_tab_id(&self) -> u16 {
        self.source_tab_id
    }
    #[must_use]
    pub fn undo_count(&self) -> u32 {
        self.undo_count
    }
    /// Raw `Ducr` undo array (MS-XLS 2.5.71), preserved without interpreting.
    #[must_use]
    pub fn undo_data(&self) -> &[u8] {
        &self.undo_data
    }
}

impl RrInsertSh {
    /// Parse the fixed 276-byte record payload.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse_payload(data: &[u8]) -> Result<Self> {
        const PAYLOAD_LEN: usize = RRD_LEN + 2 + 2 + 2 + INSERT_SH_NAME_FIELD_LEN;
        if data.len() != PAYLOAD_LEN {
            return Err(invalid(
                RR_INSERT_SH_RECORD_TYPE,
                format!(
                    "RRInsertSh payload has {} bytes; expected {PAYLOAD_LEN}",
                    data.len()
                ),
            ));
        }
        let header = RevisionRecordHeader::parse(RR_INSERT_SH_RECORD_TYPE, data, false)?;
        header.require_reviewable(RR_INSERT_SH_RECORD_TYPE)?;
        if header.revision_type != RevisionType::InsertSheet {
            return Err(invalid(
                RR_INSERT_SH_RECORD_TYPE,
                "RRInsertSh revision type is not REVTINSERTSH",
            ));
        }
        header.require_sheet(RR_INSERT_SH_RECORD_TYPE)?;
        if read_u16(data, RRD_LEN + 2) != 0 {
            return Err(invalid(
                RR_INSERT_SH_RECORD_TYPE,
                "RRInsertSh reserved field is nonzero",
            ));
        }
        let cch = read_u16(data, RRD_LEN + 4);
        let name_field = &data[RRD_LEN + 6..RRD_LEN + 6 + INSERT_SH_NAME_FIELD_LEN];
        validate_sheet_name_chars(
            RR_INSERT_SH_RECORD_TYPE,
            name_field,
            cch,
            "RRInsertSh stName",
        )?;
        let name = decode_fixed_string(
            RR_INSERT_SH_RECORD_TYPE,
            name_field,
            usize::from(cch),
            "RRInsertSh stName",
        )?;
        Ok(Self {
            header,
            position: read_u16(data, RRD_LEN),
            name,
        })
    }

    #[must_use]
    pub fn header(&self) -> &RevisionRecordHeader {
        &self.header
    }
    /// Position of the new sheet in the workbook (`itabPos`).
    #[must_use]
    pub fn position(&self) -> u16 {
        self.position
    }
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }
}

impl RevisionCellContent {
    fn from_bits(record_type: u16, bits: u32, context: &str) -> Result<Self> {
        match bits {
            0x0 => Ok(Self::Blank),
            0x1 => Ok(Self::RkNumber),
            0x2 => Ok(Self::Xnum),
            0x3 => Ok(Self::RichExtendedString),
            0x4 => Ok(Self::BoolError),
            0x5 => Ok(Self::Formula),
            other => Err(invalid(
                record_type,
                format!("RRDChgCell {context} has unknown content type {other}"),
            )),
        }
    }
}

impl RrdChgCell {
    /// Parse the record payload.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse_payload(data: &[u8]) -> Result<Self> {
        if data.len() < RRD_CHG_CELL_FIXED_LEN {
            return Err(invalid(
                RRD_CHG_CELL_RECORD_TYPE,
                format!(
                    "RRDChgCell payload has {} bytes; expected at least {RRD_CHG_CELL_FIXED_LEN}",
                    data.len()
                ),
            ));
        }
        let header = RevisionRecordHeader::parse(RRD_CHG_CELL_RECORD_TYPE, data, false)?;
        if header.revision_type != RevisionType::ChangeCell {
            return Err(invalid(
                RRD_CHG_CELL_RECORD_TYPE,
                "RRDChgCell revision type is not REVTCHANGECELL",
            ));
        }
        if header.is_deleted_at_edge_of_sort() {
            return Err(invalid(
                RRD_CHG_CELL_RECORD_TYPE,
                "RRDChgCell sets fDelAtEdgeOfSort",
            ));
        }
        header.require_sheet(RRD_CHG_CELL_RECORD_TYPE)?;
        let flags = read_u32(data, RRD_LEN);
        if flags & 0xF800_C000 != 0 {
            return Err(invalid(
                RRD_CHG_CELL_RECORD_TYPE,
                "RRDChgCell contains reserved flag bits",
            ));
        }
        let new_content =
            RevisionCellContent::from_bits(RRD_CHG_CELL_RECORD_TYPE, flags & 0x7, "vt")?;
        let old_content =
            RevisionCellContent::from_bits(RRD_CHG_CELL_RECORD_TYPE, (flags >> 3) & 0x7, "vtOld")?;
        let location = RevisionCellLocation::parse(&data[RRD_LEN + 4..RRD_LEN + 8]);
        let old_value_size = read_u32(data, RRD_LEN + 8);
        let expected_old_size = match old_content {
            RevisionCellContent::Blank => Some(0),
            RevisionCellContent::RkNumber => Some(4),
            RevisionCellContent::Xnum => Some(8),
            RevisionCellContent::BoolError => Some(2),
            RevisionCellContent::RichExtendedString | RevisionCellContent::Formula => None,
        };
        if let Some(expected) = expected_old_size
            && old_value_size != expected
        {
            return Err(invalid(
                RRD_CHG_CELL_RECORD_TYPE,
                format!(
                    "RRDChgCell cbOldVal is {old_value_size}; expected {expected} for the old content type"
                ),
            ));
        }
        if old_content == RevisionCellContent::Formula && old_value_size < MIN_FORMULA_VALUE_LEN {
            return Err(invalid(
                RRD_CHG_CELL_RECORD_TYPE,
                "RRDChgCell old formula is smaller than 24 bytes",
            ));
        }
        let formatting_run_count = read_u16(data, RRD_LEN + 12);
        let tail = data[RRD_CHG_CELL_FIXED_LEN..].to_vec();
        let old_value_size_usize = usize::try_from(old_value_size).map_err(|_error| {
            invalid(
                RRD_CHG_CELL_RECORD_TYPE,
                "RRDChgCell cbOldVal overflows usize",
            )
        })?;
        if old_value_size_usize > tail.len() {
            return Err(invalid(
                RRD_CHG_CELL_RECORD_TYPE,
                "RRDChgCell old value extends past the record payload",
            ));
        }
        Ok(Self {
            header,
            new_content,
            old_content,
            lotus_prefix: flags & 0x0040 != 0,
            has_old_format: flags & 0x0100 != 0,
            old_format_empty: flags & 0x0200 != 0,
            reset_to_style_format: flags & 0x0400 != 0,
            clear_style_format: flags & 0x0800 != 0,
            has_new_format: flags & 0x1000 != 0,
            new_format_empty: flags & 0x2000 != 0,
            display_format: ((flags >> 16) & 0xFF) as u8,
            phonetic_shown: flags & 0x0100_0000 != 0,
            old_phonetic_shown: flags & 0x0200_0000 != 0,
            formula_adjusted: flags & 0x0400_0000 != 0,
            location,
            old_value_size,
            formatting_run_count,
            tail,
        })
    }

    #[must_use]
    pub fn header(&self) -> &RevisionRecordHeader {
        &self.header
    }
    #[must_use]
    pub fn new_content(&self) -> RevisionCellContent {
        self.new_content
    }
    #[must_use]
    pub fn old_content(&self) -> RevisionCellContent {
        self.old_content
    }
    /// Whether Lotus 1-2-3 prefix characters are present (`f123Prefix`).
    #[must_use]
    pub fn has_lotus_prefix(&self) -> bool {
        self.lotus_prefix
    }
    #[must_use]
    pub fn has_old_format(&self) -> bool {
        self.has_old_format
    }
    #[must_use]
    pub fn old_format_empty(&self) -> bool {
        self.old_format_empty
    }
    /// `fXfDxf`: reset the cell format to the cell style before applying `dxf`.
    #[must_use]
    pub fn resets_to_style_format(&self) -> bool {
        self.reset_to_style_format
    }
    /// `fStyXfDxf`: clear the cell format before applying `dxf`.
    #[must_use]
    pub fn clears_style_format(&self) -> bool {
        self.clear_style_format
    }
    #[must_use]
    pub fn has_new_format(&self) -> bool {
        self.has_new_format
    }
    #[must_use]
    pub fn new_format_empty(&self) -> bool {
        self.new_format_empty
    }
    /// Number format used to display the new cell contents (`ifmtDisp`).
    #[must_use]
    pub fn display_format(&self) -> u8 {
        self.display_format
    }
    #[must_use]
    pub fn phonetic_shown(&self) -> bool {
        self.phonetic_shown
    }
    #[must_use]
    pub fn old_phonetic_shown(&self) -> bool {
        self.old_phonetic_shown
    }
    /// Whether the change came from a formula adjustment (`fEOLFmlaUpdate`).
    #[must_use]
    pub fn formula_adjusted(&self) -> bool {
        self.formula_adjusted
    }
    #[must_use]
    pub fn location(&self) -> RevisionCellLocation {
        self.location
    }
    /// Byte size of the old cell contents (`cbOldVal`).
    #[must_use]
    pub fn old_value_size(&self) -> u32 {
        self.old_value_size
    }
    /// Number of `RRDRstEtxp` records that follow (`cetxpRst`).
    #[must_use]
    pub fn formatting_run_count(&self) -> u16 {
        self.formatting_run_count
    }
    /// Raw tail: optional DXFN formats followed by the old and new values.
    #[must_use]
    pub fn tail(&self) -> &[u8] {
        &self.tail
    }
}

impl RrdConflict {
    /// Parse the record payload, which is the RRD structure alone.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse_payload(data: &[u8]) -> Result<Self> {
        if data.len() != RRD_LEN {
            return Err(invalid(
                RRD_CONFLICT_RECORD_TYPE,
                format!(
                    "RRDConflict payload has {} bytes; expected {RRD_LEN}",
                    data.len()
                ),
            ));
        }
        let header = RevisionRecordHeader::parse(RRD_CONFLICT_RECORD_TYPE, data, false)?;
        header.require_reviewable(RRD_CONFLICT_RECORD_TYPE)?;
        if header.revision_type != RevisionType::Conflict {
            return Err(invalid(
                RRD_CONFLICT_RECORD_TYPE,
                "RRDConflict revision type is not REVTCONFLICT",
            ));
        }
        Ok(Self { header })
    }

    #[must_use]
    pub fn header(&self) -> &RevisionRecordHeader {
        &self.header
    }
}

impl RrdUserView {
    /// Parse the fixed 30-byte record payload.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn parse_payload(data: &[u8]) -> Result<Self> {
        if data.len() != RRD_USER_VIEW_PAYLOAD_LEN {
            return Err(invalid(
                RRD_USER_VIEW_RECORD_TYPE,
                format!(
                    "RRDUserView payload has {} bytes; expected {RRD_USER_VIEW_PAYLOAD_LEN}",
                    data.len()
                ),
            ));
        }
        let header = RevisionRecordHeader::parse(RRD_USER_VIEW_RECORD_TYPE, data, false)?;
        if header.revision_id != 0 {
            return Err(invalid(
                RRD_USER_VIEW_RECORD_TYPE,
                "RRDUserView has a nonzero revid",
            ));
        }
        if !matches!(
            header.revision_type,
            RevisionType::AddView | RevisionType::DeleteView
        ) {
            return Err(invalid(
                RRD_USER_VIEW_RECORD_TYPE,
                "RRDUserView revision type is not a view revision",
            ));
        }
        if header.tab_id().is_some() {
            return Err(invalid(
                RRD_USER_VIEW_RECORD_TYPE,
                "RRDUserView corresponds to a specific sheet",
            ));
        }
        let mut guid = [0u8; GUID_LEN];
        guid.copy_from_slice(&data[RRD_LEN..RRD_LEN + GUID_LEN]);
        Ok(Self { header, guid })
    }

    #[must_use]
    pub fn header(&self) -> &RevisionRecordHeader {
        &self.header
    }
    /// Identifier of the custom view whose revision this record describes.
    #[must_use]
    pub fn guid(&self) -> &[u8; GUID_LEN] {
        &self.guid
    }
}
