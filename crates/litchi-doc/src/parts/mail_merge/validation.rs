//! Raw-value validation for mail-merge wire fields.
//! This layer turns specification values into typed semantic enums and flags.

use super::model::{
    FilterComparison, FilterCondition, MailMergeDestination, MailMergeDocumentType, MailMergeType,
    MergeDataSourceKind, MergeErrorCheck, Rfs, SortDirection, Wpms,
};
use crate::package::{Error as PackageError, Result};

pub(super) const FC_PMS: usize = 44;
/// Table-pointer index of `fcPmsNew`/`lcbPmsNew`.
pub(super) const FC_PMS_NEW: usize = 126;
/// Table-pointer index of `fcODSO`/`lcbODSO`.
pub(super) const FC_ODSO: usize = 127;

/// Fixed `Pms` prefix size in bytes, through `cblszSqlStr` (MS-DOC 2.9.205).
pub(super) const PMS_HEADER_LEN: usize = 30;
/// Size in bytes of one `Pmfs` element (MS-DOC 2.9.204).
pub(super) const PMFS_LEN: usize = 8;
/// `Pms.iRecCur` nil value: no current record.
pub(super) const IREC_NIL: u32 = 0xFFFF_FFFF;
/// Largest valid `Pms.iRecCur` record index.
pub(super) const IREC_MAX: u32 = 0xFFFF_FFF0;
/// Maximum byte length of `Pms.lxszSqlStr` (MS-DOC 2.9.205).
pub(super) const SQL_MAX_BYTES: u16 = 512;
/// A present `lxszSqlStr` must hold at least one character plus its null
/// terminator, so the minimum byte length is four.
pub(super) const SQL_MIN_BYTES: u16 = 4;

/// `Wpms` bit layout (MS-DOC 2.9.347).
pub(super) const WPMS_MAIN_DOC: u16 = 0x0001;
pub(super) const WPMS_DATA_SOURCE: u16 = 0x0002;
pub(super) const WPMS_HEADER_FILE: u16 = 0x0004;
pub(super) const WPMS_TYPE_SHIFT: u16 = 3;
pub(super) const WPMS_TYPE_MASK: u16 = 0x000F;
pub(super) const WPMS_AUTO: u16 = 0x0100;
pub(super) const WPMS_SUPPRESS_BLANK: u16 = 0x0400;
pub(super) const WPMS_REC_SELECT: u16 = 0x0800;
pub(super) const WPMS_DEST_SHIFT: u16 = 13;
pub(super) const WPMS_DEST_MASK: u16 = 0x0007;

/// `Wpmsdt.docType` bit mask (MS-DOC 2.9.348).
pub(super) const WPMSDT_DOC_TYPE_MASK: u32 = 0x0000_003F;

/// `Pmfs` flag bits in its second byte (MS-DOC 2.9.204).
pub(super) const PMFS_LINK_TO_FILE: u8 = 0x01;
pub(super) const PMFS_LINK_TO_CONNECTION: u8 = 0x02;
pub(super) const PMFS_NO_PROMPT_QT: u8 = 0x04;
pub(super) const PMFS_QUERY: u8 = 0x08;

/// `FNPI` bit layout (MS-DOC 2.9.93): `fnpt` in the low 4 bits.
pub(super) const FNPI_TYPE_MASK: u16 = 0x000F;
pub(super) const FNPI_IDENTIFIER_SHIFT: u16 = 4;
/// `FNPI.fnpt` value for a mail merge data source file.
pub(super) const FNPI_TYPE_MAIL_MERGE: u16 = 0x3;

/// `Rfs` flag bits in its first byte (MS-DOC 2.9.227).
pub(super) const RFS_SHOW_DATA: u32 = 0x01;
pub(super) const RFS_CHECK_ERROR_SHIFT: u32 = 1;
pub(super) const RFS_CHECK_ERROR_MASK: u32 = 0x03;
pub(super) const RFS_MAN_DOC_SETUP: u32 = 0x08;
pub(super) const RFS_MAIL_AS_TEXT: u32 = 0x10;
pub(super) const RFS_DEFAULT_SQL: u32 = 0x40;
pub(super) const RFS_MAIL_AS_HTML: u32 = 0x80;
/// Bit position of `Rfs.hsttbRfs` within the 4-byte structure.
pub(super) const RFS_HSTTB_SHIFT: u32 = 16;

/// `SttbfRfs` markers (MS-DOC 2.9.289).
pub(super) const STTB_F_EXTEND: u16 = 0xFFFF;
pub(super) const STTBF_RFS_CB_EXTRA: u16 = 0;
pub(super) const STTBF_RFS_MIN_STRINGS: u16 = 4;
pub(super) const STTBF_RFS_MAX_STRINGS: u16 = 5;
pub(super) const STTBF_RFS_MAX_CHARS: u16 = 0x00FF;

/// `ODSOPropertyBase.cb` value that introduces an `ODSOPropertyLarge`.
pub(super) const ODSO_LARGE: u16 = 0xFFFF;

/// `ODSOPropertyBase.id` values (MS-DOC 2.9.162).
pub(super) const ODSO_ID_CONNECTION_STRING: u16 = 0x0000;
pub(super) const ODSO_ID_DATA_TABLE: u16 = 0x0001;
pub(super) const ODSO_ID_DATA_SOURCE_FILE: u16 = 0x0002;
pub(super) const ODSO_ID_CONNECTION_TYPE: u16 = 0x0010;
pub(super) const ODSO_ID_COLUMN_DELIMITER: u16 = 0x0011;
pub(super) const ODSO_ID_FIRST_ROW_IS_HEADER: u16 = 0x0012;
pub(super) const ODSO_ID_RECIPIENT_FILTERS: u16 = 0x0013;
pub(super) const ODSO_ID_SORT_ORDER: u16 = 0x0014;
pub(super) const ODSO_ID_RECIPIENTS: u16 = 0x0015;
pub(super) const ODSO_ID_FIELD_MAP: u16 = 0x0016;
pub(super) const ODSO_ID_WIZARD_STEP: u16 = 0x0017;

/// Mail-merge wizard steps are numbered 1 through 6 (MS-DOC 2.9.162).
pub(super) const WIZARD_STEP_MIN: u16 = 1;
pub(super) const WIZARD_STEP_MAX: u16 = 6;

/// `FilterDataItem` fixed prefix: `cbItem`, `iColumn`, `iComparisonOperator`,
/// and `iCondition` (MS-DOC 2.9.87).
pub(super) const FILTER_ITEM_HEADER_LEN: u32 = 16;
/// Largest database column index a filter or sort key may reference.
pub(super) const MAX_COLUMN_INDEX: u32 = 254;
/// Maximum character count of a `FilterDataItem` comparison string.
pub(super) const MAX_FILTER_CHARS: usize = 212;
/// Maximum number of `SortColumnAndDirection` items (MS-DOC 2.9.162).
pub(super) const MAX_SORT_KEYS: usize = 3;
/// Size in bytes of one `SortColumnAndDirection` (MS-DOC 2.9.252).
pub(super) const SORT_KEY_LEN: usize = 8;

/// `RecipientDataItem`/`FieldMapDataItem` ids (MS-DOC 2.9.224, 2.9.84).
pub(super) const ITEM_TERMINATOR: u16 = 0x0000;
pub(super) const RECIPIENT_INCLUDED: u16 = 0x0001;
pub(super) const RECIPIENT_UNIQUE_COLUMN: u16 = 0x0002;
pub(super) const RECIPIENT_HASH: u16 = 0x0003;
pub(super) const RECIPIENT_UNIQUE_VALUE: u16 = 0x0004;
pub(super) const FIELD_MAP_MAPPED: u16 = 0x0001;
pub(super) const FIELD_MAP_COLUMN_NAME: u16 = 0x0002;
pub(super) const FIELD_MAP_FIELD_NAME: u16 = 0x0003;
pub(super) const FIELD_MAP_COLUMN_INDEX: u16 = 0x0004;
/// `FieldMapDataItem` column index meaning "not mapped" (MS-DOC 2.9.84).
pub(super) const FIELD_MAP_COLUMN_NIL: u32 = 0xFFFF_FFFF;
/// The mandated value of a `FieldMapDataItem` mapped flag (MS-DOC 2.9.84).
pub(super) const FIELD_MAP_MAPPED_VALUE: u32 = 0x0000_0001;

/// Shared count/size markers of `RecipientInfo` and `FieldMapInfo`
/// (MS-DOC 2.9.225, 2.9.85).
pub(super) const COUNT_MARKER: u16 = 0x0000;
pub(super) const CB_COUNT: u16 = 0x0004;
pub(super) const LIST_SIZE_MARKER: u16 = 0x0001;
pub(super) const LIST_SIZE_OVERFLOW: u16 = 0xFFFF;

/// Number of standard mail merge address fields in a `FieldMapInfo`
/// (MS-DOC 2.9.162).
pub(super) const FIELD_MAP_COUNT: u32 = 30;

pub(super) fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}

impl MailMergeType {
    pub(super) fn parse(raw: u16) -> Result<Self> {
        match raw {
            0x0 => Ok(Self::None),
            0x1 => Ok(Self::Letters),
            0x2 => Ok(Self::Labels),
            0x4 => Ok(Self::Envelopes),
            0x8 => Ok(Self::Catalog),
            _ => Err(corrupted("Wpms.wpmsType is not a defined merge type")),
        }
    }
}

impl MailMergeDestination {
    pub(super) fn parse(raw: u16) -> Result<Self> {
        match raw {
            0x0 => Ok(Self::None),
            0x1 => Ok(Self::Printer),
            0x2 => Ok(Self::Email),
            0x4 => Ok(Self::Fax),
            _ => Err(corrupted("Wpms.wpmsDest is not a defined destination")),
        }
    }
}

impl MailMergeDocumentType {
    pub(super) fn parse(raw: u32) -> Result<Self> {
        match raw {
            0x00 => Ok(Self::None),
            0x01 => Ok(Self::Letters),
            0x02 => Ok(Self::Labels),
            0x04 => Ok(Self::Envelopes),
            0x08 => Ok(Self::Catalog),
            0x10 => Ok(Self::Email),
            0x20 => Ok(Self::Fax),
            _ => Err(corrupted("Wpmsdt.docType is not a defined document type")),
        }
    }
}

impl Wpms {
    pub(super) fn parse(raw: u16) -> Result<Self> {
        Ok(Wpms {
            main_document: raw & WPMS_MAIN_DOC != 0,
            data_source: raw & WPMS_DATA_SOURCE != 0,
            header_file: raw & WPMS_HEADER_FILE != 0,
            merge_type: MailMergeType::parse((raw >> WPMS_TYPE_SHIFT) & WPMS_TYPE_MASK)?,
            is_automatic: raw & WPMS_AUTO != 0,
            suppress_blank_lines: raw & WPMS_SUPPRESS_BLANK != 0,
            record_selection: raw & WPMS_REC_SELECT != 0,
            destination: MailMergeDestination::parse((raw >> WPMS_DEST_SHIFT) & WPMS_DEST_MASK)?,
        })
    }
}

impl MergeDataSourceKind {
    pub(super) fn parse(raw: u8) -> Result<Self> {
        match raw {
            0xFF => Ok(Self::None),
            0x00 => Ok(Self::DataFile),
            0x01 => Ok(Self::Access),
            0x02 => Ok(Self::Excel),
            0x03 => Ok(Self::Query),
            0x04 => Ok(Self::Odbc),
            0x05 => Ok(Self::Odso),
            _ => Err(corrupted("Pmfs.ipfnpmf is not a defined data source kind")),
        }
    }
}

impl MergeErrorCheck {
    pub(super) fn parse(raw: u32) -> Result<Self> {
        match raw {
            0 => Ok(Self::SimulateAndReport),
            1 => Ok(Self::PauseAndReport),
            2 => Ok(Self::ReportInNewDocument),
            _ => Err(corrupted("Rfs.grfChkErr is not a defined setting")),
        }
    }
}

impl Rfs {
    pub(super) fn parse(raw: u32) -> Result<Self> {
        Ok(Rfs {
            show_data: raw & RFS_SHOW_DATA != 0,
            error_checking: MergeErrorCheck::parse(
                (raw >> RFS_CHECK_ERROR_SHIFT) & RFS_CHECK_ERROR_MASK,
            )?,
            manual_doc_setup: raw & RFS_MAN_DOC_SETUP != 0,
            mail_as_text: raw & RFS_MAIL_AS_TEXT != 0,
            default_sql: raw & RFS_DEFAULT_SQL != 0,
            mail_as_html: raw & RFS_MAIL_AS_HTML != 0,
            has_string_table: (raw >> RFS_HSTTB_SHIFT) != 0,
        })
    }
}

impl FilterComparison {
    pub(super) fn parse(raw: u32) -> Result<Self> {
        match raw {
            0 => Ok(Self::Equal),
            1 => Ok(Self::NotEqual),
            2 => Ok(Self::LessThan),
            3 => Ok(Self::GreaterThan),
            4 => Ok(Self::LessThanOrEqual),
            5 => Ok(Self::GreaterThanOrEqual),
            6 => Ok(Self::Empty),
            7 => Ok(Self::NotEmpty),
            _ => Err(corrupted(
                "FilterDataItem.iComparisonOperator is not defined",
            )),
        }
    }
}

impl FilterCondition {
    pub(super) fn parse(raw: u32) -> Result<Self> {
        match raw {
            0 => Ok(Self::And),
            1 => Ok(Self::Or),
            _ => Err(corrupted("FilterDataItem.iCondition is not 0 or 1")),
        }
    }
}

impl SortDirection {
    pub(super) fn parse(raw: u32) -> Result<Self> {
        match raw {
            0 => Ok(Self::Ascending),
            1 => Ok(Self::Descending),
            _ => Err(corrupted("SortColumnAndDirection.iDirection is not 0 or 1")),
        }
    }
}
