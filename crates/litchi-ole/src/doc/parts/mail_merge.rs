//! Legacy Word mail-merge data-source state (`Pms` and the ODSO property set).
//!
//! The Word 97-era `Pms` structure (MS-DOC 2.9.205) records the print/mail
//! merge state: the merge flags (`Wpms`, 2.9.347), two data-source connection
//! descriptors (`Pmfs`, 2.9.204), record filtering (`Rfs`, 2.9.227), a stored
//! SQL query, the connection string table (`SttbfRfs`, 2.9.289), and the merge
//! document type (`Wpmsdt`, 2.9.348). Word 2002+ stores the current merge
//! state in a second `Pms` addressed by `fcPmsNew` (MS-DOC 2.5.8). The Word
//! 2002+ Office Data Source Object (`fcODSO`, MS-DOC 2.5) is a sequence of
//! variable-length `ODSOPropertyBase`
//! items (2.9.162) covering the connection string, data table, source file,
//! recipient filters (`FilterDataItem`, 2.9.87), sort keys
//! (`SortColumnAndDirection`, 2.9.252), recipient inclusion (`RecipientInfo`,
//! 2.9.225), and column-to-address-field mappings (`FieldMapInfo`, 2.9.85).
//!
//! Everything here is parsed as inert metadata: data-source paths, connection
//! strings, and SQL queries are stored verbatim and are never opened,
//! resolved, contacted, or executed, and no merge is ever performed.

use super::fib::{FileInformationBlock, WORD_97_NFIB};
use crate::doc::package::{DocError, Result};

/// Table-pointer index of `fcPms`/`lcbPms`.
const FC_PMS: usize = 44;
/// Table-pointer index of `fcPmsNew`/`lcbPmsNew`.
const FC_PMS_NEW: usize = 126;
/// Table-pointer index of `fcODSO`/`lcbODSO`.
const FC_ODSO: usize = 127;

/// Fixed `Pms` prefix size in bytes, through `cblszSqlStr` (MS-DOC 2.9.205).
const PMS_HEADER_LEN: usize = 30;
/// Size in bytes of one `Pmfs` element (MS-DOC 2.9.204).
const PMFS_LEN: usize = 8;
/// `Pms.iRecCur` nil value: no current record.
const IREC_NIL: u32 = 0xFFFF_FFFF;
/// Largest valid `Pms.iRecCur` record index.
const IREC_MAX: u32 = 0xFFFF_FFF0;
/// Maximum byte length of `Pms.lxszSqlStr` (MS-DOC 2.9.205).
const SQL_MAX_BYTES: u16 = 512;
/// A present `lxszSqlStr` must hold at least one character plus its null
/// terminator, so the minimum byte length is four.
const SQL_MIN_BYTES: u16 = 4;

/// `Wpms` bit layout (MS-DOC 2.9.347).
const WPMS_MAIN_DOC: u16 = 0x0001;
const WPMS_DATA_SOURCE: u16 = 0x0002;
const WPMS_HEADER_FILE: u16 = 0x0004;
const WPMS_TYPE_SHIFT: u16 = 3;
const WPMS_TYPE_MASK: u16 = 0x000F;
const WPMS_AUTO: u16 = 0x0100;
const WPMS_SUPPRESS_BLANK: u16 = 0x0400;
const WPMS_REC_SELECT: u16 = 0x0800;
const WPMS_DEST_SHIFT: u16 = 13;
const WPMS_DEST_MASK: u16 = 0x0007;

/// `Wpmsdt.docType` bit mask (MS-DOC 2.9.348).
const WPMSDT_DOC_TYPE_MASK: u32 = 0x0000_003F;

/// `Pmfs` flag bits in its second byte (MS-DOC 2.9.204).
const PMFS_LINK_TO_FILE: u8 = 0x01;
const PMFS_LINK_TO_CONNECTION: u8 = 0x02;
const PMFS_NO_PROMPT_QT: u8 = 0x04;
const PMFS_QUERY: u8 = 0x08;

/// `FNPI` bit layout (MS-DOC 2.9.93): `fnpt` in the low 4 bits.
const FNPI_TYPE_MASK: u16 = 0x000F;
const FNPI_IDENTIFIER_SHIFT: u16 = 4;
/// `FNPI.fnpt` value for a mail merge data source file.
const FNPI_TYPE_MAIL_MERGE: u16 = 0x3;

/// `Rfs` flag bits in its first byte (MS-DOC 2.9.227).
const RFS_SHOW_DATA: u32 = 0x01;
const RFS_CHECK_ERROR_SHIFT: u32 = 1;
const RFS_CHECK_ERROR_MASK: u32 = 0x03;
const RFS_MAN_DOC_SETUP: u32 = 0x08;
const RFS_MAIL_AS_TEXT: u32 = 0x10;
const RFS_DEFAULT_SQL: u32 = 0x40;
const RFS_MAIL_AS_HTML: u32 = 0x80;
/// Bit position of `Rfs.hsttbRfs` within the 4-byte structure.
const RFS_HSTTB_SHIFT: u32 = 16;

/// `SttbfRfs` markers (MS-DOC 2.9.289).
const STTB_F_EXTEND: u16 = 0xFFFF;
const STTBF_RFS_CB_EXTRA: u16 = 0;
const STTBF_RFS_MIN_STRINGS: u16 = 4;
const STTBF_RFS_MAX_STRINGS: u16 = 5;
const STTBF_RFS_MAX_CHARS: u16 = 0x00FF;

/// `ODSOPropertyBase.cb` value that introduces an `ODSOPropertyLarge`.
const ODSO_LARGE: u16 = 0xFFFF;

/// `ODSOPropertyBase.id` values (MS-DOC 2.9.162).
const ODSO_ID_CONNECTION_STRING: u16 = 0x0000;
const ODSO_ID_DATA_TABLE: u16 = 0x0001;
const ODSO_ID_DATA_SOURCE_FILE: u16 = 0x0002;
const ODSO_ID_CONNECTION_TYPE: u16 = 0x0010;
const ODSO_ID_COLUMN_DELIMITER: u16 = 0x0011;
const ODSO_ID_FIRST_ROW_IS_HEADER: u16 = 0x0012;
const ODSO_ID_RECIPIENT_FILTERS: u16 = 0x0013;
const ODSO_ID_SORT_ORDER: u16 = 0x0014;
const ODSO_ID_RECIPIENTS: u16 = 0x0015;
const ODSO_ID_FIELD_MAP: u16 = 0x0016;
const ODSO_ID_WIZARD_STEP: u16 = 0x0017;

/// Mail-merge wizard steps are numbered 1 through 6 (MS-DOC 2.9.162).
const WIZARD_STEP_MIN: u16 = 1;
const WIZARD_STEP_MAX: u16 = 6;

/// `FilterDataItem` fixed prefix: `cbItem`, `iColumn`, `iComparisonOperator`,
/// and `iCondition` (MS-DOC 2.9.87).
const FILTER_ITEM_HEADER_LEN: u32 = 16;
/// Largest database column index a filter or sort key may reference.
const MAX_COLUMN_INDEX: u32 = 254;
/// Maximum character count of a `FilterDataItem` comparison string.
const MAX_FILTER_CHARS: usize = 212;
/// Maximum number of `SortColumnAndDirection` items (MS-DOC 2.9.162).
const MAX_SORT_KEYS: usize = 3;
/// Size in bytes of one `SortColumnAndDirection` (MS-DOC 2.9.252).
const SORT_KEY_LEN: usize = 8;

/// `RecipientDataItem`/`FieldMapDataItem` ids (MS-DOC 2.9.224, 2.9.84).
const ITEM_TERMINATOR: u16 = 0x0000;
const RECIPIENT_INCLUDED: u16 = 0x0001;
const RECIPIENT_UNIQUE_COLUMN: u16 = 0x0002;
const RECIPIENT_HASH: u16 = 0x0003;
const RECIPIENT_UNIQUE_VALUE: u16 = 0x0004;
const FIELD_MAP_MAPPED: u16 = 0x0001;
const FIELD_MAP_COLUMN_NAME: u16 = 0x0002;
const FIELD_MAP_FIELD_NAME: u16 = 0x0003;
const FIELD_MAP_COLUMN_INDEX: u16 = 0x0004;
/// `FieldMapDataItem` column index meaning "not mapped" (MS-DOC 2.9.84).
const FIELD_MAP_COLUMN_NIL: u32 = 0xFFFF_FFFF;
/// The mandated value of a `FieldMapDataItem` mapped flag (MS-DOC 2.9.84).
const FIELD_MAP_MAPPED_VALUE: u32 = 0x0000_0001;

/// Shared count/size markers of `RecipientInfo` and `FieldMapInfo`
/// (MS-DOC 2.9.225, 2.9.85).
const COUNT_MARKER: u16 = 0x0000;
const CB_COUNT: u16 = 0x0004;
const LIST_SIZE_MARKER: u16 = 0x0001;
const LIST_SIZE_OVERFLOW: u16 = 0xFFFF;

/// Number of standard mail merge address fields in a `FieldMapInfo`
/// (MS-DOC 2.9.162).
const FIELD_MAP_COUNT: u32 = 30;

fn corrupted(message: impl Into<String>) -> DocError {
    DocError::Corrupted(message.into())
}

/// A bounds-checked cursor over one structure's byte range.
struct Reader<'a> {
    data: &'a [u8],
    position: usize,
    context: &'static str,
}

impl<'a> Reader<'a> {
    fn new(data: &'a [u8], context: &'static str) -> Self {
        Reader {
            data,
            position: 0,
            context,
        }
    }

    fn remaining(&self) -> usize {
        self.data.len() - self.position
    }

    fn bytes(&mut self, count: usize) -> Result<&'a [u8]> {
        let end = self
            .position
            .checked_add(count)
            .ok_or_else(|| corrupted(format!("{} range overflows", self.context)))?;
        let slice = self
            .data
            .get(self.position..end)
            .ok_or_else(|| corrupted(format!("{} is truncated", self.context)))?;
        self.position = end;
        Ok(slice)
    }

    fn u8(&mut self) -> Result<u8> {
        Ok(self.bytes(1)?[0])
    }

    fn u16(&mut self) -> Result<u16> {
        let bytes = self.bytes(2)?;
        Ok(u16::from_le_bytes([bytes[0], bytes[1]]))
    }

    fn u32(&mut self) -> Result<u32> {
        let bytes = self.bytes(4)?;
        Ok(u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]))
    }

    /// Require that the whole range was consumed exactly.
    fn finish(&self) -> Result<()> {
        if self.remaining() != 0 {
            return Err(corrupted(format!("{} has trailing bytes", self.context)));
        }
        Ok(())
    }
}

/// Decode a UTF-16LE string that must occupy `bytes` exactly.
fn decode_utf16(bytes: &[u8], context: &str) -> Result<String> {
    if bytes.len() % 2 != 0 {
        return Err(corrupted(format!("{context} has an odd byte length")));
    }
    let units: Vec<u16> = bytes
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
        .collect();
    String::from_utf16(&units).map_err(|_| corrupted(format!("{context} is not valid UTF-16")))
}

/// The document type of a mail merge (`Wpms.wpmsType`, MS-DOC 2.9.347).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MailMergeType {
    /// No mail merge.
    None,
    /// Form letters.
    Letters,
    /// Mailing labels.
    Labels,
    /// Envelopes.
    Envelopes,
    /// Catalog or directory.
    Catalog,
}

impl MailMergeType {
    fn parse(raw: u16) -> Result<Self> {
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

/// The destination of a mail merge (`Wpms.wpmsDest`, MS-DOC 2.9.347).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MailMergeDestination {
    /// No destination selected.
    None,
    /// A printer.
    Printer,
    /// E-mail messages.
    Email,
    /// Fax.
    Fax,
}

impl MailMergeDestination {
    fn parse(raw: u16) -> Result<Self> {
        match raw {
            0x0 => Ok(Self::None),
            0x1 => Ok(Self::Printer),
            0x2 => Ok(Self::Email),
            0x4 => Ok(Self::Fax),
            _ => Err(corrupted("Wpms.wpmsDest is not a defined destination")),
        }
    }
}

/// The document type of a mail merge (`Wpmsdt.docType`, MS-DOC 2.9.348).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MailMergeDocumentType {
    /// No mail merge.
    None,
    /// Form letters.
    Letters,
    /// Mailing labels.
    Labels,
    /// Envelopes.
    Envelopes,
    /// Catalog or directory.
    Catalog,
    /// E-mail messages.
    Email,
    /// Fax.
    Fax,
}

impl MailMergeDocumentType {
    fn parse(raw: u32) -> Result<Self> {
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

/// The current state of a mail merge (`Wpms`, MS-DOC 2.9.347).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Wpms {
    /// Whether the main document was selected for the merge.
    pub main_document: bool,
    /// Whether a data source was selected for the merge.
    pub data_source: bool,
    /// Whether merge field names come from a header file.
    pub header_file: bool,
    /// The document type of the merge.
    pub merge_type: MailMergeType,
    /// Whether this is an automatic label or envelope merge.
    pub is_automatic: bool,
    /// Whether blank lines in the data files are suppressed.
    pub suppress_blank_lines: bool,
    /// Whether record selection is enabled.
    pub record_selection: bool,
    /// The merge destination.
    pub destination: MailMergeDestination,
}

impl Wpms {
    fn parse(raw: u16) -> Result<Self> {
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

/// The kind of a mail merge data source (`Pmfs.ipfnpmf`, MS-DOC 2.9.204).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MergeDataSourceKind {
    /// No data source.
    None,
    /// A data file.
    DataFile,
    /// A Microsoft Access database.
    Access,
    /// A Microsoft Excel file.
    Excel,
    /// A Microsoft Query database.
    Query,
    /// ODBC.
    Odbc,
    /// An Office Data Source Object (ODSO).
    Odso,
}

impl MergeDataSourceKind {
    fn parse(raw: u8) -> Result<Self> {
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

/// A field/record separator token for data-file sources (`Pmfs.tkField` and
/// `Pmfs.tkRec`, MS-DOC 2.9.204).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MergeFileToken {
    /// No separator.
    None,
    /// Enter (paragraph).
    Enter,
    /// Tabulation.
    Tab,
    /// `,`
    Comma,
    /// `.`
    Period,
    /// `!`
    Exclamation,
    /// `#`
    Hash,
    /// `$`
    Dollar,
    /// `%`
    Percent,
    /// `&`
    Ampersand,
    /// `(`
    LeftParenthesis,
    /// `)`
    RightParenthesis,
    /// `*`
    Asterisk,
    /// `+`
    Plus,
    /// `-`
    Minus,
    /// `/`
    Slash,
    /// `:`
    Colon,
    /// `;`
    Semicolon,
    /// `<`
    LessThan,
    /// `=`
    Equals,
    /// `>`
    GreaterThan,
    /// `?`
    Question,
    /// `@`
    At,
    /// `[`
    LeftBracket,
    /// `]`
    RightBracket,
    /// `^`
    Caret,
    /// `_`
    Underscore,
    /// `` ` ``
    Backtick,
    /// `{`
    LeftBrace,
    /// `}`
    RightBrace,
    /// `|`
    Pipe,
    /// `~`
    Tilde,
    /// End-of-field marker.
    FieldEnd,
    /// Table-cell marker.
    TableCell,
    /// Table-row marker.
    TableRow,
}

impl MergeFileToken {
    /// Map a raw `tkField`/`tkRec` value to its token, when it is defined.
    pub fn from_raw(raw: i16) -> Option<Self> {
        Some(match raw {
            0x00 => Self::None,
            0x02 => Self::Enter,
            0x06 => Self::Tab,
            0x0A => Self::Comma,
            0x0B => Self::Period,
            0x0C => Self::Exclamation,
            0x0D => Self::Hash,
            0x0E => Self::Dollar,
            0x0F => Self::Percent,
            0x10 => Self::Ampersand,
            0x11 => Self::LeftParenthesis,
            0x12 => Self::RightParenthesis,
            0x13 => Self::Asterisk,
            0x14 => Self::Plus,
            0x15 => Self::Minus,
            0x16 => Self::Slash,
            0x17 => Self::Colon,
            0x18 => Self::Semicolon,
            0x19 => Self::LessThan,
            0x1A => Self::Equals,
            0x1B => Self::GreaterThan,
            0x1C => Self::Question,
            0x1D => Self::At,
            0x1E => Self::LeftBracket,
            0x1F => Self::RightBracket,
            0x21 => Self::Caret,
            0x22 => Self::Underscore,
            0x23 => Self::Backtick,
            0x24 => Self::LeftBrace,
            0x25 => Self::RightBrace,
            0x26 => Self::Pipe,
            0x27 => Self::Tilde,
            0x46 => Self::FieldEnd,
            0x47 => Self::TableCell,
            0x48 => Self::TableRow,
            _ => return None,
        })
    }
}

/// A file-name type and identifier (`FNPI`, MS-DOC 2.9.93) referencing the
/// data file in the document's `SttbFnm` table.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Fnpi {
    raw: u16,
}

impl Fnpi {
    /// Wrap a raw 2-byte `FNPI` value (used by other table parsers).
    pub(crate) fn from_raw(raw: u16) -> Self {
        Fnpi { raw }
    }

    /// `fnpt`: the type of the referenced file name.
    pub fn file_type(&self) -> u8 {
        (self.raw & FNPI_TYPE_MASK) as u8
    }

    /// `fnpd`: the raw 12-bit identifier of the file name within its type.
    pub fn identifier(&self) -> u16 {
        self.raw >> FNPI_IDENTIFIER_SHIFT
    }

    /// Whether this references a mail merge data source file (`fnpt` = 3).
    pub fn is_mail_merge_source(&self) -> bool {
        self.raw & FNPI_TYPE_MASK == FNPI_TYPE_MAIL_MERGE
    }
}

/// One mail merge data source connection descriptor (`Pmfs`, MS-DOC 2.9.204).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Pmfs {
    /// The type of the data source.
    pub source_kind: MergeDataSourceKind,
    /// For data-file sources, whether the file must exist as named.
    pub link_to_file: bool,
    /// Whether an extra string is appended to the DDE connection string.
    pub link_to_connection: bool,
    /// Whether the user was already asked about editing ODBC via Microsoft
    /// Query.
    pub no_prompt_query_tools: bool,
    /// Whether the merge uses a query instead of reading the data file
    /// directly.
    pub uses_query: bool,
    /// Raw `tkField` token separating fields in a data file; meaningful only
    /// for `MergeDataSourceKind::DataFile`.
    pub field_token: i16,
    /// Raw `tkRec` token separating records in a data file; meaningful only
    /// for `MergeDataSourceKind::DataFile`.
    pub record_token: i16,
    /// Reference to the data file name in the document's `SttbFnm` table.
    pub file_name: Fnpi,
}

impl Pmfs {
    /// The typed field separator token, when the raw value is defined.
    pub fn field_separator(&self) -> Option<MergeFileToken> {
        MergeFileToken::from_raw(self.field_token)
    }

    /// The typed record separator token, when the raw value is defined.
    pub fn record_separator(&self) -> Option<MergeFileToken> {
        MergeFileToken::from_raw(self.record_token)
    }

    fn parse(data: &[u8]) -> Result<Self> {
        debug_assert_eq!(data.len(), PMFS_LEN);
        let flags = data[1];
        Ok(Pmfs {
            source_kind: MergeDataSourceKind::parse(data[0])?,
            link_to_file: flags & PMFS_LINK_TO_FILE != 0,
            link_to_connection: flags & PMFS_LINK_TO_CONNECTION != 0,
            no_prompt_query_tools: flags & PMFS_NO_PROMPT_QT != 0,
            uses_query: flags & PMFS_QUERY != 0,
            field_token: i16::from_le_bytes([data[2], data[3]]),
            record_token: i16::from_le_bytes([data[4], data[5]]),
            file_name: Fnpi {
                raw: u16::from_le_bytes([data[6], data[7]]),
            },
        })
    }
}

/// Error checking and reporting settings for a merge (`Rfs.grfChkErr`,
/// MS-DOC 2.9.227).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MergeErrorCheck {
    /// Simulate the merge and report errors in a new document.
    SimulateAndReport,
    /// Complete the merge and pause to report errors.
    PauseAndReport,
    /// Complete the merge and report errors in a new document.
    ReportInNewDocument,
}

impl MergeErrorCheck {
    fn parse(raw: u32) -> Result<Self> {
        match raw {
            0 => Ok(Self::SimulateAndReport),
            1 => Ok(Self::PauseAndReport),
            2 => Ok(Self::ReportInNewDocument),
            _ => Err(corrupted("Rfs.grfChkErr is not a defined setting")),
        }
    }
}

/// Record filtering and related merge properties (`Rfs`, MS-DOC 2.9.227).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Rfs {
    /// Whether data values are shown in the merged fields instead of field
    /// names.
    pub show_data: bool,
    /// The error checking and reporting setting.
    pub error_checking: MergeErrorCheck,
    /// Whether the main document envelope or mailing labels are set up.
    pub manual_doc_setup: bool,
    /// Whether a merge e-mail message is in plain text format.
    pub mail_as_text: bool,
    /// Whether the default SQL query string is `SELECT * FROM x`.
    pub default_sql: bool,
    /// Whether a merge e-mail message is in HTML format.
    pub mail_as_html: bool,
    /// Whether a `SttbfRfs` string table follows in the `Pms`.
    pub has_string_table: bool,
}

impl Rfs {
    fn parse(raw: u32) -> Result<Self> {
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

/// The mail merge connection and record filtering string table (`SttbfRfs`,
/// MS-DOC 2.9.289).
///
/// Connection strings are stored verbatim and are never opened or contacted.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SttbfRfs {
    strings: Vec<String>,
}

impl SttbfRfs {
    /// All strings in table order (4 or 5 entries).
    pub fn strings(&self) -> &[String] {
        &self.strings
    }

    /// `Data0`: the connection string to the merge data source.
    pub fn connection_string(&self) -> &str {
        &self.strings[0]
    }

    /// `Data1`: the connection string to the field-name source; empty when
    /// field names come from the same source as the data.
    pub fn header_connection_string(&self) -> &str {
        &self.strings[1]
    }

    /// `Data2`: the e-mail subject line for e-mail merges.
    pub fn email_subject(&self) -> &str {
        &self.strings[2]
    }

    /// `Data3`: the data column holding e-mail addresses or fax numbers.
    pub fn address_column(&self) -> &str {
        &self.strings[3]
    }

    fn parse(reader: &mut Reader<'_>) -> Result<Self> {
        if reader.u16()? != STTB_F_EXTEND {
            return Err(corrupted("SttbfRfs.fExtend is not 0xFFFF"));
        }
        let count = reader.u16()?;
        if !(STTBF_RFS_MIN_STRINGS..=STTBF_RFS_MAX_STRINGS).contains(&count) {
            return Err(corrupted("SttbfRfs.cData is not 4 or 5"));
        }
        if reader.u16()? != STTBF_RFS_CB_EXTRA {
            return Err(corrupted("SttbfRfs.cbExtra is not zero"));
        }
        let mut strings = Vec::with_capacity(usize::from(count));
        for _ in 0..count {
            let chars = reader.u16()?;
            if chars > STTBF_RFS_MAX_CHARS {
                return Err(corrupted("SttbfRfs string exceeds 255 characters"));
            }
            let raw = reader.bytes(usize::from(chars) * 2)?;
            strings.push(decode_utf16(raw, "SttbfRfs string")?);
        }
        Ok(SttbfRfs { strings })
    }
}

/// The print/mail merge state of a document (`Pms`, MS-DOC 2.9.205).
///
/// The stored SQL query and connection strings are inert: they are exposed
/// verbatim and never executed, resolved, or contacted.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Pms {
    /// The merge state flags.
    pub state: Wpms,
    /// `ipmfMF`: index into `sources` of the header field source supplying
    /// the merge column names.
    pub header_source_index: u8,
    /// `ipmfFetch`: index into `sources` of the data fetch source.
    pub fetch_source_index: u8,
    /// `iRecCur`: index of the current merge record, or `None` when nil.
    pub current_record: Option<u32>,
    /// `rgpmfs`: the two data source connection descriptors.
    pub sources: [Pmfs; 2],
    /// Record filtering and related merge properties.
    pub filter: Rfs,
    /// The stored SQL query string (without its null terminator), when
    /// present. Never executed.
    pub sql_query: Option<String>,
    /// The connection and record filtering string table, when present.
    pub strings: Option<SttbfRfs>,
    /// The merge document type (`Wpmsdt`), when the trailing field exists.
    pub document_type: Option<MailMergeDocumentType>,
}

impl Pms {
    fn parse(data: &[u8]) -> Result<Self> {
        if data.len() < PMS_HEADER_LEN {
            return Err(corrupted("Pms is truncated"));
        }
        let mut reader = Reader::new(data, "Pms");
        let state = Wpms::parse(reader.u16()?)?;
        let header_source_index = reader.u8()?;
        let fetch_source_index = reader.u8()?;
        if header_source_index > 1 || fetch_source_index > 1 {
            return Err(corrupted("Pms source index is not 0 or 1"));
        }
        let current_record = match reader.u32()? {
            IREC_NIL => None,
            value if value <= IREC_MAX => Some(value),
            _ => return Err(corrupted("Pms.iRecCur is out of range")),
        };
        let sources = [
            Pmfs::parse(reader.bytes(PMFS_LEN)?)?,
            Pmfs::parse(reader.bytes(PMFS_LEN)?)?,
        ];
        let filter = Rfs::parse(reader.u32()?)?;
        let sql_length = reader.u16()?;
        let sql_query = if sql_length == 0 {
            None
        } else {
            if sql_length % 2 != 0 {
                return Err(corrupted("Pms.cblszSqlStr is not even"));
            }
            if !(SQL_MIN_BYTES..=SQL_MAX_BYTES).contains(&sql_length) {
                return Err(corrupted("Pms.cblszSqlStr is out of range"));
            }
            let raw = reader.bytes(usize::from(sql_length))?;
            let (text, terminator) = raw.split_at(raw.len() - 2);
            if terminator != [0, 0] {
                return Err(corrupted("Pms.lxszSqlStr is not null-terminated"));
            }
            Some(decode_utf16(text, "Pms.lxszSqlStr")?)
        };
        let strings = if filter.has_string_table {
            Some(SttbfRfs::parse(&mut reader)?)
        } else {
            None
        };
        let document_type = match reader.remaining() {
            0 => None,
            4 => Some(MailMergeDocumentType::parse(
                reader.u32()? & WPMSDT_DOC_TYPE_MASK,
            )?),
            _ => return Err(corrupted("Pms has a partial Wpmsdt")),
        };
        reader.finish()?;
        Ok(Pms {
            state,
            header_source_index,
            fetch_source_index,
            current_record,
            sources,
            filter,
            sql_query,
            strings,
            document_type,
        })
    }
}

/// A comparison operator of a recipient filter (`FilterDataItem.
/// iComparisonOperator`, MS-DOC 2.9.87).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FilterComparison {
    /// Equal.
    Equal,
    /// Not equal.
    NotEqual,
    /// Less than.
    LessThan,
    /// Greater than.
    GreaterThan,
    /// Less than or equal.
    LessThanOrEqual,
    /// Greater than or equal.
    GreaterThanOrEqual,
    /// Empty.
    Empty,
    /// Not empty.
    NotEmpty,
}

impl FilterComparison {
    fn parse(raw: u32) -> Result<Self> {
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

/// How one filter comparison combines with the others
/// (`FilterDataItem.iCondition`, MS-DOC 2.9.87).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FilterCondition {
    /// Logical AND.
    And,
    /// Logical OR.
    Or,
}

impl FilterCondition {
    fn parse(raw: u32) -> Result<Self> {
        match raw {
            0 => Ok(Self::And),
            1 => Ok(Self::Or),
            _ => Err(corrupted("FilterDataItem.iCondition is not 0 or 1")),
        }
    }
}

/// One recipient filter (`FilterDataItem`, MS-DOC 2.9.87).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FilterDataItem {
    /// Zero-based index of the data source column being filtered.
    pub column: u32,
    /// The comparison operator.
    pub comparison: FilterComparison,
    /// How this comparison combines with the others.
    pub condition: FilterCondition,
    /// The comparison value (without its null terminator).
    pub value: String,
}

impl FilterDataItem {
    fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::new(data, "FilterDataItem");
        let column = reader.u32()?;
        if column > MAX_COLUMN_INDEX {
            return Err(corrupted("FilterDataItem.iColumn is out of range"));
        }
        let comparison = FilterComparison::parse(reader.u32()?)?;
        let condition = FilterCondition::parse(reader.u32()?)?;
        let raw = reader.bytes(reader.remaining())?;
        if raw.len() < 2 {
            return Err(corrupted("FilterDataItem string is truncated"));
        }
        let (text, terminator) = raw.split_at(raw.len() - 2);
        if terminator != [0, 0] {
            return Err(corrupted("FilterDataItem string is not null-terminated"));
        }
        let value = decode_utf16(text, "FilterDataItem string")?;
        if value.chars().count() > MAX_FILTER_CHARS {
            return Err(corrupted("FilterDataItem string exceeds 212 characters"));
        }
        reader.finish()?;
        Ok(FilterDataItem {
            column,
            comparison,
            condition,
            value,
        })
    }

    fn parse_list(data: &[u8]) -> Result<Vec<Self>> {
        let mut reader = Reader::new(data, "recipient filter list");
        let mut filters = Vec::new();
        while reader.remaining() > 0 {
            let item_size = usize::try_from(reader.u32()?)
                .map_err(|_| corrupted("FilterDataItem.cbItem exceeds usize"))?;
            if item_size < FILTER_ITEM_HEADER_LEN as usize + 2 {
                return Err(corrupted("FilterDataItem.cbItem is too small"));
            }
            filters.push(Self::parse(reader.bytes(item_size - 4)?)?);
        }
        Ok(filters)
    }
}

/// A recipient sort direction (`SortColumnAndDirection.iDirection`,
/// MS-DOC 2.9.252).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SortDirection {
    /// Ascending order.
    Ascending,
    /// Descending order.
    Descending,
}

impl SortDirection {
    fn parse(raw: u32) -> Result<Self> {
        match raw {
            0 => Ok(Self::Ascending),
            1 => Ok(Self::Descending),
            _ => Err(corrupted("SortColumnAndDirection.iDirection is not 0 or 1")),
        }
    }
}

/// One recipient sort key (`SortColumnAndDirection`, MS-DOC 2.9.252).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SortColumnAndDirection {
    /// Zero-based index of the data source column to sort by.
    pub column: u32,
    /// The sort direction.
    pub direction: SortDirection,
}

impl SortColumnAndDirection {
    fn parse_list(data: &[u8]) -> Result<Vec<Self>> {
        if data.len() % SORT_KEY_LEN != 0 {
            return Err(corrupted("sort key list has a partial item"));
        }
        if data.len() / SORT_KEY_LEN > MAX_SORT_KEYS {
            return Err(corrupted("sort key list exceeds three items"));
        }
        let mut keys = Vec::with_capacity(data.len() / SORT_KEY_LEN);
        for chunk in data.chunks_exact(SORT_KEY_LEN) {
            let column = u32::from_le_bytes([chunk[0], chunk[1], chunk[2], chunk[3]]);
            if column > MAX_COLUMN_INDEX {
                return Err(corrupted("SortColumnAndDirection.iColumn is out of range"));
            }
            let direction =
                SortDirection::parse(u32::from_le_bytes([chunk[4], chunk[5], chunk[6], chunk[7]]))?;
            keys.push(SortColumnAndDirection { column, direction });
        }
        Ok(keys)
    }
}

/// One recipient of a mail merge (`RecipientBase`, MS-DOC 2.9.224/2.9.225).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct RecipientEntry {
    /// Whether the recipient is included in the merge (defaults to `true`
    /// when the file stores no status item).
    pub included: bool,
    /// Zero-based index of the column uniquely identifying the recipient.
    pub unique_column: Option<u32>,
    /// Hash uniquely identifying the recipient when no unique column exists.
    pub record_hash: Option<u32>,
    /// Contents of the column uniquely identifying the recipient.
    pub unique_value: Option<String>,
}

/// The recipient inclusion list of a mail merge (`RecipientInfo`,
/// MS-DOC 2.9.225).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RecipientInfo {
    /// The recipients in data source order.
    pub recipients: Vec<RecipientEntry>,
}

impl RecipientInfo {
    fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::new(data, "RecipientInfo");
        if reader.u16()? != COUNT_MARKER {
            return Err(corrupted("RecipientInfo.countMarker is not zero"));
        }
        if reader.u16()? != CB_COUNT {
            return Err(corrupted("RecipientInfo.cbCount is not 4"));
        }
        let count = usize::try_from(reader.u32()?)
            .map_err(|_| corrupted("RecipientInfo.cRecipients exceeds usize"))?;
        if reader.u16()? != LIST_SIZE_MARKER {
            return Err(corrupted("RecipientInfo list size marker is not 1"));
        }
        let short_size = reader.u16()?;
        let list_size = if short_size == LIST_SIZE_OVERFLOW {
            usize::try_from(reader.u32()?)
                .map_err(|_| corrupted("RecipientInfo list size exceeds usize"))?
        } else {
            usize::from(short_size)
        };
        let list = reader.bytes(list_size)?;
        reader.finish()?;
        let mut items = Reader::new(list, "RecipientInfo recipients");
        let mut recipients = Vec::with_capacity(count);
        for _ in 0..count {
            let mut recipient = RecipientEntry {
                included: true,
                ..RecipientEntry::default()
            };
            loop {
                let id = items.u16()?;
                let size = usize::from(items.u16()?);
                if id == ITEM_TERMINATOR {
                    if size != 0 {
                        return Err(corrupted("RecipientTerminator has data"));
                    }
                    break;
                }
                let value = items.bytes(size)?;
                match id {
                    RECIPIENT_INCLUDED => {
                        if size != 4 {
                            return Err(corrupted("recipient inclusion item is not 4 bytes"));
                        }
                        recipient.included =
                            match u32::from_le_bytes([value[0], value[1], value[2], value[3]]) {
                                0 => false,
                                1 => true,
                                _ => {
                                    return Err(corrupted(
                                        "recipient inclusion value is not 0 or 1",
                                    ));
                                },
                            };
                    },
                    RECIPIENT_UNIQUE_COLUMN | RECIPIENT_HASH => {
                        if size != 4 {
                            return Err(corrupted("recipient integer item is not 4 bytes"));
                        }
                        let number = u32::from_le_bytes([value[0], value[1], value[2], value[3]]);
                        if id == RECIPIENT_UNIQUE_COLUMN {
                            recipient.unique_column = Some(number);
                        } else {
                            recipient.record_hash = Some(number);
                        }
                    },
                    RECIPIENT_UNIQUE_VALUE => {
                        recipient.unique_value =
                            Some(decode_utf16(value, "recipient unique value")?);
                    },
                    _ => return Err(corrupted("recipient item id is not defined")),
                }
            }
            recipients.push(recipient);
        }
        items.finish()?;
        Ok(RecipientInfo { recipients })
    }
}

/// How one standard address field maps to a data source column
/// (`FieldMapBase`, MS-DOC 2.9.83/2.9.84).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct FieldMapping {
    /// Zero-based index of the mapped data source column, when mapped.
    pub column_index: Option<u32>,
    /// Name of the mapped data source column, when stored.
    pub column_name: Option<String>,
}

/// Column-to-address-field mappings of a mail merge (`FieldMapInfo`,
/// MS-DOC 2.9.85).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FieldMapInfo {
    /// One mapping per standard address field, in
    /// [`FieldMapInfo::STANDARD_ADDRESS_FIELDS`] order.
    pub mappings: Vec<FieldMapping>,
}

impl FieldMapInfo {
    /// The 30 standard mail merge address fields, in `FieldMapDataItem`
    /// order (MS-DOC 2.9.162).
    pub const STANDARD_ADDRESS_FIELDS: [&'static str; 30] = [
        "Unique Identifier",
        "Courtesy Title",
        "First Name",
        "Middle Name",
        "Last Name",
        "Suffix",
        "Nickname",
        "Job Title",
        "Company",
        "Address 1",
        "Address 2",
        "City",
        "State",
        "Postal Code",
        "Country or Region",
        "Business Phone",
        "Business Fax",
        "Home Phone",
        "Home Fax",
        "E-mail Address",
        "Web Page",
        "Spouse Courtesy Title",
        "Spouse First Name",
        "Spouse Middle Name",
        "Spouse Last Name",
        "Spouse Nickname",
        "Phonetic Guide for First Name",
        "Phonetic Guide for Last Name",
        "Address 3",
        "Department",
    ];

    fn parse(data: &[u8]) -> Result<Self> {
        let mut reader = Reader::new(data, "FieldMapInfo");
        if reader.u16()? != COUNT_MARKER {
            return Err(corrupted("FieldMapInfo.countMarker is not zero"));
        }
        if reader.u16()? != CB_COUNT {
            return Err(corrupted("FieldMapInfo.cbCount is not 4"));
        }
        if reader.u32()? != FIELD_MAP_COUNT {
            return Err(corrupted("FieldMapInfo.cFields is not 30"));
        }
        if reader.u16()? != LIST_SIZE_MARKER {
            return Err(corrupted("FieldMapInfo list size marker is not 1"));
        }
        let short_size = reader.u16()?;
        let list_size = if short_size == LIST_SIZE_OVERFLOW {
            usize::try_from(reader.u32()?)
                .map_err(|_| corrupted("FieldMapInfo list size exceeds usize"))?
        } else {
            usize::from(short_size)
        };
        let list = reader.bytes(list_size)?;
        reader.finish()?;
        let mut items = Reader::new(list, "FieldMapInfo mappings");
        let mut mappings = Vec::with_capacity(FIELD_MAP_COUNT as usize);
        for _ in 0..FIELD_MAP_COUNT {
            let mut mapping = FieldMapping::default();
            loop {
                let id = items.u16()?;
                let size = usize::from(items.u16()?);
                if id == ITEM_TERMINATOR {
                    if size != 0 {
                        return Err(corrupted("FieldMapTerminator has data"));
                    }
                    break;
                }
                let value = items.bytes(size)?;
                match id {
                    FIELD_MAP_MAPPED => {
                        if size != 4
                            || u32::from_le_bytes([value[0], value[1], value[2], value[3]])
                                != FIELD_MAP_MAPPED_VALUE
                        {
                            return Err(corrupted("field map mapped flag is not 1"));
                        }
                    },
                    FIELD_MAP_COLUMN_NAME => {
                        mapping.column_name = Some(decode_utf16(value, "field map column name")?);
                    },
                    FIELD_MAP_FIELD_NAME => {
                        // The standard field name is ignored by definition
                        // (MS-DOC 2.9.84); only its framing is validated.
                    },
                    FIELD_MAP_COLUMN_INDEX => {
                        if size != 4 {
                            return Err(corrupted("field map column index is not 4 bytes"));
                        }
                        let index = u32::from_le_bytes([value[0], value[1], value[2], value[3]]);
                        if index != FIELD_MAP_COLUMN_NIL {
                            mapping.column_index = Some(index);
                        }
                    },
                    _ => return Err(corrupted("field map item id is not defined")),
                }
            }
            mappings.push(mapping);
        }
        items.finish()?;
        Ok(FieldMapInfo { mappings })
    }
}

/// One Office Data Source Object property (`ODSOPropertyBase`, MS-DOC
/// 2.9.162).
///
/// Strings such as connection strings and file names are stored verbatim and
/// are never opened, resolved, or contacted.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OdsoProperty {
    /// `0x0000`: the data source connection string (UDL). Never contacted.
    ConnectionString(String),
    /// `0x0001`: the data set (for example a table name) to use.
    DataTable(String),
    /// `0x0002`: name of the file used as the data source. Never opened.
    DataSourceFile(String),
    /// `0x0010`: the stored data source connection type; the value is not
    /// interpreted, as applications reset it on load.
    ConnectionType(u32),
    /// `0x0011`: the Unicode code point used as a text data source column
    /// delimiter.
    ColumnDelimiter(u16),
    /// `0x0012`: whether the first row holds column names.
    FirstRowIsHeader(bool),
    /// `0x0013`: the recipient filters.
    RecipientFilters(Vec<FilterDataItem>),
    /// `0x0014`: up to three recipient sort keys.
    SortOrder(Vec<SortColumnAndDirection>),
    /// `0x0015`: the recipient inclusion list.
    Recipients(RecipientInfo),
    /// `0x0016`: the column-to-address-field mappings.
    FieldMap(FieldMapInfo),
    /// `0x0017`: the last mail merge wizard step shown (1-6).
    WizardStep(u16),
    /// A property whose `id` is not defined by MS-DOC 2.9.162, preserved
    /// verbatim.
    Unknown {
        /// The raw `ODSOPropertyBase.id`.
        id: u16,
        /// The raw property value bytes.
        data: Vec<u8>,
    },
}

impl OdsoProperty {
    fn decode(id: u16, value: &[u8]) -> Result<Self> {
        Ok(match id {
            ODSO_ID_CONNECTION_STRING => {
                Self::ConnectionString(decode_utf16(value, "ODSO connection string")?)
            },
            ODSO_ID_DATA_TABLE => Self::DataTable(decode_utf16(value, "ODSO data table")?),
            ODSO_ID_DATA_SOURCE_FILE => {
                Self::DataSourceFile(decode_utf16(value, "ODSO data source file")?)
            },
            ODSO_ID_CONNECTION_TYPE => Self::ConnectionType(expect_u32(value, "ODSO property")?),
            ODSO_ID_COLUMN_DELIMITER => {
                if value.len() != 2 {
                    return Err(corrupted("ODSO column delimiter is not 2 bytes"));
                }
                Self::ColumnDelimiter(u16::from_le_bytes([value[0], value[1]]))
            },
            ODSO_ID_FIRST_ROW_IS_HEADER => match expect_u32(value, "ODSO property")? {
                0 => Self::FirstRowIsHeader(false),
                1 => Self::FirstRowIsHeader(true),
                _ => return Err(corrupted("ODSO first-row flag is not 0 or 1")),
            },
            ODSO_ID_RECIPIENT_FILTERS => Self::RecipientFilters(FilterDataItem::parse_list(value)?),
            ODSO_ID_SORT_ORDER => Self::SortOrder(SortColumnAndDirection::parse_list(value)?),
            ODSO_ID_RECIPIENTS => Self::Recipients(RecipientInfo::parse(value)?),
            ODSO_ID_FIELD_MAP => Self::FieldMap(FieldMapInfo::parse(value)?),
            ODSO_ID_WIZARD_STEP => {
                if value.len() != 2 {
                    return Err(corrupted("ODSO wizard step is not 2 bytes"));
                }
                let step = u16::from_le_bytes([value[0], value[1]]);
                if !(WIZARD_STEP_MIN..=WIZARD_STEP_MAX).contains(&step) {
                    return Err(corrupted("ODSO wizard step is not between 1 and 6"));
                }
                Self::WizardStep(step)
            },
            _ => Self::Unknown {
                id,
                data: value.to_vec(),
            },
        })
    }
}

fn expect_u32(value: &[u8], context: &str) -> Result<u32> {
    if value.len() != 4 {
        return Err(corrupted(format!("{context} is not 4 bytes")));
    }
    Ok(u32::from_le_bytes([value[0], value[1], value[2], value[3]]))
}

/// The mail-merge data-source state of a document: the Word 97 `Pms` and the
/// Word 2002+ ODSO property set, addressed by `fcPms` and `fcODSO`.
///
/// All state is inert: data-source paths, connection strings, and SQL queries
/// are stored verbatim and never opened, resolved, contacted, or executed,
/// and no merge is performed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentMailMerge {
    state: Option<Pms>,
    new_state: Option<Pms>,
    odso_properties: Vec<OdsoProperty>,
}

impl DocumentMailMerge {
    /// The Word 97 merge state (`Pms`), when the document carries one.
    pub fn state(&self) -> Option<&Pms> {
        self.state.as_ref()
    }

    /// The Word 2002+ merge state (`fcPmsNew`, a new `Pms` recording the
    /// current state of the print merge operation), when the document
    /// carries one.
    pub fn new_state(&self) -> Option<&Pms> {
        self.new_state.as_ref()
    }

    /// The Word 2002+ ODSO properties in storage order (empty when the
    /// document carries no ODSO data).
    pub fn odso_properties(&self) -> &[OdsoProperty] {
        &self.odso_properties
    }

    /// Parse the mail-merge state addressed by the FIB, or `None` when the
    /// document carries neither a `Pms` nor ODSO data.
    pub fn parse(
        fib: &FileInformationBlock,
        table_stream: &[u8],
    ) -> Result<Option<DocumentMailMerge>> {
        // The Word 6/95 FIB table-pointer layout assigns these indices to
        // unrelated structures, so they only carry merge state from Word 97
        // on.
        if fib.version() < WORD_97_NFIB {
            return Ok(None);
        }
        let state = optional_slice(fib, table_stream, FC_PMS, "Pms")?
            .map(Pms::parse)
            .transpose()?;
        let new_state = optional_slice(fib, table_stream, FC_PMS_NEW, "PmsNew")?
            .map(Pms::parse)
            .transpose()?;
        let odso_properties = match optional_slice(fib, table_stream, FC_ODSO, "ODSO data")? {
            Some(data) => parse_odso_properties(data)?,
            None => Vec::new(),
        };
        if state.is_none() && new_state.is_none() && odso_properties.is_empty() {
            return Ok(None);
        }
        Ok(Some(DocumentMailMerge {
            state,
            new_state,
            odso_properties,
        }))
    }
}

/// Parse the ODSO property bag, which is a sequence of variable-length
/// `ODSOPropertyBase` items filling the byte range exactly.
fn parse_odso_properties(data: &[u8]) -> Result<Vec<OdsoProperty>> {
    let mut reader = Reader::new(data, "ODSO property set");
    let mut properties = Vec::new();
    while reader.remaining() > 0 {
        let id = reader.u16()?;
        let size = reader.u16()?;
        let value = if size == ODSO_LARGE {
            let large_size = usize::try_from(reader.u32()?)
                .map_err(|_| corrupted("ODSO property size exceeds usize"))?;
            reader.bytes(large_size)?
        } else {
            reader.bytes(usize::from(size))?
        };
        properties.push(OdsoProperty::decode(id, value)?);
    }
    Ok(properties)
}

fn optional_slice<'a>(
    fib: &FileInformationBlock,
    table_stream: &'a [u8],
    index: usize,
    name: &str,
) -> Result<Option<&'a [u8]>> {
    let Some((offset, length)) = fib.get_table_pointer(index) else {
        return Ok(None);
    };
    if length == 0 {
        return Ok(None);
    }
    let start =
        usize::try_from(offset).map_err(|_| corrupted(format!("{name} offset exceeds usize")))?;
    let length =
        usize::try_from(length).map_err(|_| corrupted(format!("{name} length exceeds usize")))?;
    let end = start
        .checked_add(length)
        .ok_or_else(|| corrupted(format!("{name} range overflows")))?;
    table_stream
        .get(start..end)
        .map(Some)
        .ok_or_else(|| corrupted(format!("{name} extends beyond the table stream")))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn utf16(text: &str) -> Vec<u8> {
        text.encode_utf16().flat_map(u16::to_le_bytes).collect()
    }

    fn pmfs_bytes(kind: u8, flags: u8, tk_field: i16, tk_rec: i16, fnpi: u16) -> [u8; PMFS_LEN] {
        let mut bytes = [0u8; PMFS_LEN];
        bytes[0] = kind;
        bytes[1] = flags;
        bytes[2..4].copy_from_slice(&tk_field.to_le_bytes());
        bytes[4..6].copy_from_slice(&tk_rec.to_le_bytes());
        bytes[6..8].copy_from_slice(&fnpi.to_le_bytes());
        bytes
    }

    struct PmsBuilder {
        wpms: u16,
        ipmf_mf: u8,
        ipmf_fetch: u8,
        irec_cur: u32,
        rfs: u32,
        sql: Option<String>,
        sttbf: Option<Vec<Vec<u8>>>,
        wpmsdt: Option<u32>,
    }

    impl PmsBuilder {
        fn new() -> Self {
            PmsBuilder {
                wpms: 0x0409,
                ipmf_mf: 0,
                ipmf_fetch: 0,
                irec_cur: IREC_NIL,
                rfs: 0,
                sql: None,
                sttbf: None,
                wpmsdt: None,
            }
        }

        fn build(&self) -> Vec<u8> {
            let mut data = Vec::new();
            data.extend_from_slice(&self.wpms.to_le_bytes());
            data.push(self.ipmf_mf);
            data.push(self.ipmf_fetch);
            data.extend_from_slice(&self.irec_cur.to_le_bytes());
            data.extend_from_slice(&pmfs_bytes(0x00, 0, 0, 0, 0xFFF3));
            data.extend_from_slice(&pmfs_bytes(0xFF, 0, 0, 0, 0xFFF3));
            data.extend_from_slice(&self.rfs.to_le_bytes());
            match &self.sql {
                Some(sql) => {
                    let mut encoded = utf16(sql);
                    encoded.extend_from_slice(&[0, 0]);
                    data.extend_from_slice(&(encoded.len() as u16).to_le_bytes());
                    data.extend_from_slice(&encoded);
                },
                None => data.extend_from_slice(&0u16.to_le_bytes()),
            }
            if let Some(strings) = &self.sttbf {
                data.extend_from_slice(&STTB_F_EXTEND.to_le_bytes());
                data.extend_from_slice(&(strings.len() as u16).to_le_bytes());
                data.extend_from_slice(&0u16.to_le_bytes());
                for string in strings {
                    data.extend_from_slice(&((string.len() / 2) as u16).to_le_bytes());
                    data.extend_from_slice(string);
                }
            }
            if let Some(doc_type) = self.wpmsdt {
                data.extend_from_slice(&doc_type.to_le_bytes());
            }
            data
        }
    }

    #[test]
    fn parses_full_pms() {
        let mut builder = PmsBuilder::new();
        builder.ipmf_mf = 1;
        builder.ipmf_fetch = 0;
        builder.irec_cur = 1;
        builder.sql = Some("SELECT * FROM [myTable] WHERE x".to_string());
        builder.wpmsdt = Some(0x01);
        let pms = Pms::parse(&builder.build()).unwrap();
        assert_eq!(pms.state.merge_type, MailMergeType::Letters);
        assert!(pms.state.main_document);
        assert!(!pms.state.data_source);
        assert!(pms.state.suppress_blank_lines);
        assert_eq!(pms.state.destination, MailMergeDestination::None);
        assert_eq!(pms.header_source_index, 1);
        assert_eq!(pms.fetch_source_index, 0);
        assert_eq!(pms.current_record, Some(1));
        assert_eq!(pms.sources[0].source_kind, MergeDataSourceKind::DataFile);
        assert_eq!(pms.sources[1].source_kind, MergeDataSourceKind::None);
        assert!(pms.sources[0].file_name.is_mail_merge_source());
        assert_eq!(pms.sources[0].file_name.identifier(), 0x0FFF);
        assert_eq!(
            pms.sql_query.as_deref(),
            Some("SELECT * FROM [myTable] WHERE x")
        );
        assert!(pms.strings.is_none());
        assert_eq!(pms.document_type, Some(MailMergeDocumentType::Letters));
    }

    #[test]
    fn parses_minimal_pms() {
        let pms = Pms::parse(&PmsBuilder::new().build()).unwrap();
        assert_eq!(pms.current_record, None);
        assert_eq!(pms.sql_query, None);
        assert_eq!(pms.document_type, None);
    }

    #[test]
    fn parses_pms_with_sttbf_rfs() {
        let mut builder = PmsBuilder::new();
        builder.rfs = 0x0001_0000 | RFS_SHOW_DATA; // hsttbRfs nonzero
        builder.sttbf = Some(vec![
            utf16("DSN=mailmerge;"),
            utf16(""),
            utf16("Your order"),
            utf16("Email"),
            utf16("ignored"),
        ]);
        builder.wpmsdt = Some(0x10);
        let pms = Pms::parse(&builder.build()).unwrap();
        assert!(pms.filter.show_data);
        let strings = pms.strings.unwrap();
        assert_eq!(strings.strings().len(), 5);
        assert_eq!(strings.connection_string(), "DSN=mailmerge;");
        assert_eq!(strings.header_connection_string(), "");
        assert_eq!(strings.email_subject(), "Your order");
        assert_eq!(strings.address_column(), "Email");
        assert_eq!(pms.document_type, Some(MailMergeDocumentType::Email));
    }

    #[test]
    fn parses_pmfs_flags_and_tokens() {
        let pmfs = Pmfs::parse(&pmfs_bytes(0x00, 0x0F, 0x06, 0x02, 0xFFF3)).unwrap();
        assert!(pmfs.link_to_file);
        assert!(pmfs.link_to_connection);
        assert!(pmfs.no_prompt_query_tools);
        assert!(pmfs.uses_query);
        assert_eq!(pmfs.field_separator(), Some(MergeFileToken::Tab));
        assert_eq!(pmfs.record_separator(), Some(MergeFileToken::Enter));
        assert_eq!(
            MergeFileToken::from_raw(0x48),
            Some(MergeFileToken::TableRow)
        );
        assert_eq!(MergeFileToken::from_raw(0x30), None);
    }

    #[test]
    fn parses_rfs_flags() {
        // byte0 = 0x82: grfChkErr=1, fMailAsHtml; hsttbRfs = 0.
        let rfs = Rfs::parse(0x0082).unwrap();
        assert!(!rfs.show_data);
        assert_eq!(rfs.error_checking, MergeErrorCheck::PauseAndReport);
        assert!(rfs.mail_as_html);
        assert!(!rfs.mail_as_text);
        assert!(!rfs.has_string_table);
        assert!(Rfs::parse(0x0006).is_err()); // grfChkErr = 3
    }

    #[test]
    fn rejects_malformed_pms() {
        let good = PmsBuilder::new().build();
        // Truncated header.
        assert!(Pms::parse(&good[..PMS_HEADER_LEN - 1]).is_err());
        // ipmfMF out of range.
        let mut bad = good.clone();
        bad[2] = 2;
        assert!(Pms::parse(&bad).is_err());
        // ipmfFetch out of range.
        let mut bad = good.clone();
        bad[3] = 2;
        assert!(Pms::parse(&bad).is_err());
        // iRecCur out of range.
        let mut builder = PmsBuilder::new();
        builder.irec_cur = IREC_MAX + 1;
        assert!(Pms::parse(&builder.build()).is_err());
        // Undefined wpmsType.
        let mut builder = PmsBuilder::new();
        builder.wpms = 0x0003 << WPMS_TYPE_SHIFT;
        assert!(Pms::parse(&builder.build()).is_err());
        // Undefined wpmsDest.
        let mut builder = PmsBuilder::new();
        builder.wpms = 0x0003 << WPMS_DEST_SHIFT;
        assert!(Pms::parse(&builder.build()).is_err());
        // Undefined data source kind.
        let mut bad = good.clone();
        bad[8] = 0x06;
        assert!(Pms::parse(&bad).is_err());
        // Odd SQL length.
        let mut bad = PmsBuilder::new().build();
        bad[28] = 3;
        bad.extend_from_slice(&[0, 0, 0]);
        assert!(Pms::parse(&bad).is_err());
        // SQL length too small (null terminator only).
        let mut bad = PmsBuilder::new().build();
        bad[28] = 2;
        bad.extend_from_slice(&[0, 0]);
        assert!(Pms::parse(&bad).is_err());
        // SQL length too large.
        let mut builder = PmsBuilder::new();
        builder.sql = Some("x".repeat(300));
        assert!(Pms::parse(&builder.build()).is_err());
        // SQL missing its null terminator.
        let mut bad = PmsBuilder::new().build();
        bad[28] = 4;
        bad.extend_from_slice(&utf16("xy"));
        assert!(Pms::parse(&bad).is_err());
        // Declared string table missing.
        let mut builder = PmsBuilder::new();
        builder.rfs = 0x0001_0000;
        assert!(Pms::parse(&builder.build()).is_err());
        // Partial trailing Wpmsdt.
        let mut bad = PmsBuilder::new().build();
        bad.extend_from_slice(&[0, 0]);
        assert!(Pms::parse(&bad).is_err());
        // Undefined Wpmsdt document type.
        let mut builder = PmsBuilder::new();
        builder.wpmsdt = Some(0x03);
        assert!(Pms::parse(&builder.build()).is_err());
    }

    #[test]
    fn rejects_malformed_sttbf_rfs() {
        let with_table = |strings: Vec<Vec<u8>>, f_extend: u16, cb_extra: u16| {
            let mut builder = PmsBuilder::new();
            builder.rfs = 0x0001_0000;
            builder.sttbf = Some(strings);
            let mut data = builder.build();
            if f_extend != STTB_F_EXTEND {
                let at = PMS_HEADER_LEN;
                data[at..at + 2].copy_from_slice(&f_extend.to_le_bytes());
            }
            if cb_extra != 0 {
                let at = PMS_HEADER_LEN + 4;
                data[at..at + 2].copy_from_slice(&cb_extra.to_le_bytes());
            }
            data
        };
        let strings = || vec![utf16("a"), utf16(""), utf16(""), utf16(""), utf16("")];
        // Bad fExtend.
        assert!(Pms::parse(&with_table(strings(), 0x0000, 0)).is_err());
        // Bad cbExtra.
        assert!(Pms::parse(&with_table(strings(), STTB_F_EXTEND, 2)).is_err());
        // Too few strings (cData = 3): truncated table.
        assert!(Pms::parse(&with_table(strings()[..3].to_vec(), STTB_F_EXTEND, 0)).is_err());
        // String exceeding 255 characters.
        let mut oversized = strings();
        oversized[0] = utf16(&"x".repeat(256));
        assert!(Pms::parse(&with_table(oversized, STTB_F_EXTEND, 0)).is_err());
    }

    fn odso_item(id: u16, value: &[u8]) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&id.to_le_bytes());
        if value.len() >= ODSO_LARGE as usize {
            data.extend_from_slice(&ODSO_LARGE.to_le_bytes());
            data.extend_from_slice(&(value.len() as u32).to_le_bytes());
        } else {
            data.extend_from_slice(&(value.len() as u16).to_le_bytes());
        }
        data.extend_from_slice(value);
        data
    }

    fn recipient_info_bytes(recipients: &[Vec<(u16, Vec<u8>)>]) -> Vec<u8> {
        let mut list = Vec::new();
        for items in recipients {
            for (id, value) in items {
                list.extend_from_slice(&id.to_le_bytes());
                list.extend_from_slice(&(value.len() as u16).to_le_bytes());
                list.extend_from_slice(value);
            }
            list.extend_from_slice(&[0, 0, 0, 0]);
        }
        let mut data = Vec::new();
        data.extend_from_slice(&COUNT_MARKER.to_le_bytes());
        data.extend_from_slice(&CB_COUNT.to_le_bytes());
        data.extend_from_slice(&(recipients.len() as u32).to_le_bytes());
        data.extend_from_slice(&LIST_SIZE_MARKER.to_le_bytes());
        data.extend_from_slice(&(list.len() as u16).to_le_bytes());
        data.extend_from_slice(&list);
        data
    }

    fn field_map_info_bytes(mappings: &[Vec<(u16, Vec<u8>)>]) -> Vec<u8> {
        let mut list = Vec::new();
        for items in mappings {
            for (id, value) in items {
                list.extend_from_slice(&id.to_le_bytes());
                list.extend_from_slice(&(value.len() as u16).to_le_bytes());
                list.extend_from_slice(value);
            }
            list.extend_from_slice(&[0, 0, 0, 0]);
        }
        let mut data = Vec::new();
        data.extend_from_slice(&COUNT_MARKER.to_le_bytes());
        data.extend_from_slice(&CB_COUNT.to_le_bytes());
        data.extend_from_slice(&FIELD_MAP_COUNT.to_le_bytes());
        data.extend_from_slice(&LIST_SIZE_MARKER.to_le_bytes());
        data.extend_from_slice(&(list.len() as u16).to_le_bytes());
        data.extend_from_slice(&list);
        data
    }

    #[test]
    fn parses_odso_scalar_properties() {
        let mut bag = Vec::new();
        bag.extend_from_slice(&odso_item(
            ODSO_ID_CONNECTION_STRING,
            &utf16("Provider=SQLOLEDB;Data Source=srv;"),
        ));
        bag.extend_from_slice(&odso_item(ODSO_ID_DATA_TABLE, &utf16("Customers")));
        bag.extend_from_slice(&odso_item(
            ODSO_ID_DATA_SOURCE_FILE,
            &utf16("C:\\data\\customers.mdb"),
        ));
        bag.extend_from_slice(&odso_item(ODSO_ID_CONNECTION_TYPE, &5u32.to_le_bytes()));
        bag.extend_from_slice(&odso_item(ODSO_ID_COLUMN_DELIMITER, &0x2Cu16.to_le_bytes()));
        bag.extend_from_slice(&odso_item(ODSO_ID_FIRST_ROW_IS_HEADER, &1u32.to_le_bytes()));
        bag.extend_from_slice(&odso_item(ODSO_ID_WIZARD_STEP, &3u16.to_le_bytes()));
        let properties = parse_odso_properties(&bag).unwrap();
        assert_eq!(
            properties,
            vec![
                OdsoProperty::ConnectionString("Provider=SQLOLEDB;Data Source=srv;".into()),
                OdsoProperty::DataTable("Customers".into()),
                OdsoProperty::DataSourceFile("C:\\data\\customers.mdb".into()),
                OdsoProperty::ConnectionType(5),
                OdsoProperty::ColumnDelimiter(0x2C),
                OdsoProperty::FirstRowIsHeader(true),
                OdsoProperty::WizardStep(3),
            ]
        );
    }

    #[test]
    fn parses_odso_large_and_unknown_properties() {
        let large_value = utf16(&"x".repeat(0x1_0000));
        let mut bag = odso_item(ODSO_ID_CONNECTION_STRING, &large_value);
        bag.extend_from_slice(&odso_item(0x0099, &[1, 2, 3]));
        let properties = parse_odso_properties(&bag).unwrap();
        assert_eq!(properties.len(), 2);
        assert_eq!(
            properties[1],
            OdsoProperty::Unknown {
                id: 0x0099,
                data: vec![1, 2, 3],
            }
        );
    }

    #[test]
    fn rejects_malformed_odso_bags() {
        // Partial property header.
        assert!(parse_odso_properties(&[0, 0, 4]).is_err());
        // Value overruns the bag.
        let mut bag = odso_item(ODSO_ID_DATA_TABLE, &utf16("abc"));
        bag.truncate(bag.len() - 2);
        assert!(parse_odso_properties(&bag).is_err());
        // Large property with an overrunning size.
        let mut bag = Vec::new();
        bag.extend_from_slice(&0u16.to_le_bytes());
        bag.extend_from_slice(&ODSO_LARGE.to_le_bytes());
        bag.extend_from_slice(&100u32.to_le_bytes());
        assert!(parse_odso_properties(&bag).is_err());
        // Odd-length Unicode string.
        assert!(parse_odso_properties(&odso_item(ODSO_ID_DATA_TABLE, b"a")).is_err());
        // Wrong scalar sizes.
        assert!(parse_odso_properties(&odso_item(ODSO_ID_CONNECTION_TYPE, &[0; 3])).is_err());
        assert!(parse_odso_properties(&odso_item(ODSO_ID_COLUMN_DELIMITER, &[0; 4])).is_err());
        assert!(parse_odso_properties(&odso_item(ODSO_ID_WIZARD_STEP, &[0; 4])).is_err());
        // Out-of-range scalar values.
        assert!(
            parse_odso_properties(&odso_item(ODSO_ID_FIRST_ROW_IS_HEADER, &2u32.to_le_bytes()))
                .is_err()
        );
        assert!(
            parse_odso_properties(&odso_item(ODSO_ID_WIZARD_STEP, &7u16.to_le_bytes())).is_err()
        );
        assert!(
            parse_odso_properties(&odso_item(ODSO_ID_WIZARD_STEP, &0u16.to_le_bytes())).is_err()
        );
    }

    fn filter_item_bytes(column: u32, comparison: u32, condition: u32, value: &str) -> Vec<u8> {
        let mut body = Vec::new();
        body.extend_from_slice(&column.to_le_bytes());
        body.extend_from_slice(&comparison.to_le_bytes());
        body.extend_from_slice(&condition.to_le_bytes());
        body.extend_from_slice(&utf16(value));
        body.extend_from_slice(&[0, 0]);
        let mut item = Vec::new();
        item.extend_from_slice(&((body.len() + 4) as u32).to_le_bytes());
        item.extend_from_slice(&body);
        item
    }

    #[test]
    fn parses_odso_recipient_filters() {
        let mut value = filter_item_bytes(2, 3, 0, "smith");
        value.extend_from_slice(&filter_item_bytes(0, 7, 1, ""));
        let bag = odso_item(ODSO_ID_RECIPIENT_FILTERS, &value);
        let properties = parse_odso_properties(&bag).unwrap();
        assert_eq!(
            properties,
            vec![OdsoProperty::RecipientFilters(vec![
                FilterDataItem {
                    column: 2,
                    comparison: FilterComparison::GreaterThan,
                    condition: FilterCondition::And,
                    value: "smith".into(),
                },
                FilterDataItem {
                    column: 0,
                    comparison: FilterComparison::NotEmpty,
                    condition: FilterCondition::Or,
                    value: String::new(),
                },
            ])]
        );
    }

    #[test]
    fn rejects_malformed_filters() {
        let wrap =
            |value: &[u8]| parse_odso_properties(&odso_item(ODSO_ID_RECIPIENT_FILTERS, value));
        // cbItem smaller than the fixed prefix.
        assert!(wrap(&4u32.to_le_bytes()).is_err());
        // cbItem overrunning the value.
        let mut bad = filter_item_bytes(0, 0, 0, "a");
        bad[0] = 0xFF;
        assert!(wrap(&bad).is_err());
        // Column index out of range.
        assert!(wrap(&filter_item_bytes(255, 0, 0, "a")).is_err());
        // Undefined comparison operator.
        assert!(wrap(&filter_item_bytes(0, 8, 0, "a")).is_err());
        // Undefined condition.
        assert!(wrap(&filter_item_bytes(0, 0, 2, "a")).is_err());
        // Missing null terminator: trim the terminator and adjust cbItem.
        let mut bad = filter_item_bytes(0, 0, 0, "a");
        bad.truncate(bad.len() - 2);
        let size = (bad.len() as u32).to_le_bytes();
        bad[..4].copy_from_slice(&size);
        assert!(wrap(&bad).is_err());
        // Comparison string exceeding 212 characters.
        assert!(wrap(&filter_item_bytes(0, 0, 0, &"x".repeat(213))).is_err());
    }

    #[test]
    fn parses_odso_sort_order() {
        let mut value = Vec::new();
        value.extend_from_slice(&1u32.to_le_bytes());
        value.extend_from_slice(&0u32.to_le_bytes());
        value.extend_from_slice(&2u32.to_le_bytes());
        value.extend_from_slice(&1u32.to_le_bytes());
        let bag = odso_item(ODSO_ID_SORT_ORDER, &value);
        let properties = parse_odso_properties(&bag).unwrap();
        assert_eq!(
            properties,
            vec![OdsoProperty::SortOrder(vec![
                SortColumnAndDirection {
                    column: 1,
                    direction: SortDirection::Ascending,
                },
                SortColumnAndDirection {
                    column: 2,
                    direction: SortDirection::Descending,
                },
            ])]
        );
        // Partial item.
        assert!(parse_odso_properties(&odso_item(ODSO_ID_SORT_ORDER, &[0; 4])).is_err());
        // More than three keys.
        let mut too_many = Vec::new();
        for _ in 0..4 {
            too_many.extend_from_slice(&0u32.to_le_bytes());
            too_many.extend_from_slice(&0u32.to_le_bytes());
        }
        assert!(parse_odso_properties(&odso_item(ODSO_ID_SORT_ORDER, &too_many)).is_err());
        // Column out of range.
        let mut bad_column = Vec::new();
        bad_column.extend_from_slice(&255u32.to_le_bytes());
        bad_column.extend_from_slice(&0u32.to_le_bytes());
        assert!(parse_odso_properties(&odso_item(ODSO_ID_SORT_ORDER, &bad_column)).is_err());
        // Undefined direction.
        let mut bad_direction = Vec::new();
        bad_direction.extend_from_slice(&0u32.to_le_bytes());
        bad_direction.extend_from_slice(&2u32.to_le_bytes());
        assert!(parse_odso_properties(&odso_item(ODSO_ID_SORT_ORDER, &bad_direction)).is_err());
    }

    #[test]
    fn parses_odso_recipient_info() {
        let first = vec![
            (RECIPIENT_INCLUDED, 0u32.to_le_bytes().to_vec()),
            (RECIPIENT_UNIQUE_COLUMN, 4u32.to_le_bytes().to_vec()),
            (RECIPIENT_UNIQUE_VALUE, utf16("key-1")),
        ];
        let second = vec![(RECIPIENT_HASH, 0xDEAD_BEEFu32.to_le_bytes().to_vec())];
        let bag = odso_item(ODSO_ID_RECIPIENTS, &recipient_info_bytes(&[first, second]));
        let properties = parse_odso_properties(&bag).unwrap();
        let [OdsoProperty::Recipients(info)] = properties.as_slice() else {
            panic!("expected recipient info");
        };
        assert_eq!(info.recipients.len(), 2);
        assert!(!info.recipients[0].included);
        assert_eq!(info.recipients[0].unique_column, Some(4));
        assert_eq!(info.recipients[0].unique_value.as_deref(), Some("key-1"));
        // Inclusion defaults to true when no status item is stored.
        assert!(info.recipients[1].included);
        assert_eq!(info.recipients[1].record_hash, Some(0xDEAD_BEEF));
    }

    #[test]
    fn rejects_malformed_recipient_info() {
        let wrap = |value: &[u8]| parse_odso_properties(&odso_item(ODSO_ID_RECIPIENTS, value));
        // Wrong count marker.
        let mut bad = recipient_info_bytes(&[]);
        bad[1] = 1;
        assert!(wrap(&bad).is_err());
        // Wrong cbCount.
        let mut bad = recipient_info_bytes(&[]);
        bad[3] = 8;
        assert!(wrap(&bad).is_err());
        // Wrong list size marker.
        let mut bad = recipient_info_bytes(&[]);
        bad[9] = 2;
        assert!(wrap(&bad).is_err());
        // List size overrun.
        let mut bad = recipient_info_bytes(&[]);
        bad[10] = 4;
        assert!(wrap(&bad).is_err());
        // Undefined item id.
        let bad = recipient_info_bytes(&[vec![(0x0009, 0u32.to_le_bytes().to_vec())]]);
        assert!(wrap(&bad).is_err());
        // Terminator carrying data.
        let bad = recipient_info_bytes(&[vec![(ITEM_TERMINATOR, vec![0, 0, 0, 0])]]);
        // The terminator is emitted after the items, so inject a bad one.
        let mut with_bad_terminator = bad.clone();
        let at = 12; // start of the list
        with_bad_terminator[at + 2] = 4;
        assert!(wrap(&with_bad_terminator).is_err());
        // Inclusion value other than 0/1.
        let bad = recipient_info_bytes(&[vec![(RECIPIENT_INCLUDED, 2u32.to_le_bytes().to_vec())]]);
        assert!(wrap(&bad).is_err());
        // Missing terminator: declared list ends mid-recipient.
        let mut good =
            recipient_info_bytes(&[vec![(RECIPIENT_UNIQUE_COLUMN, 4u32.to_le_bytes().to_vec())]]);
        let total = good.len();
        good[10] = (total - 12 - 4) as u8; // drop the terminator from the size
        good.truncate(total - 4);
        assert!(wrap(&good).is_err());
    }

    #[test]
    fn parses_odso_field_map_info() {
        let mut mappings: Vec<Vec<(u16, Vec<u8>)>> = vec![Vec::new(); FIELD_MAP_COUNT as usize];
        mappings[2] = vec![
            (
                FIELD_MAP_MAPPED,
                FIELD_MAP_MAPPED_VALUE.to_le_bytes().to_vec(),
            ),
            (FIELD_MAP_COLUMN_NAME, utf16("GivenName")),
            (FIELD_MAP_FIELD_NAME, utf16("First Name")),
            (FIELD_MAP_COLUMN_INDEX, 3u32.to_le_bytes().to_vec()),
        ];
        mappings[19] = vec![(
            FIELD_MAP_COLUMN_INDEX,
            FIELD_MAP_COLUMN_NIL.to_le_bytes().to_vec(),
        )];
        let bag = odso_item(ODSO_ID_FIELD_MAP, &field_map_info_bytes(&mappings));
        let properties = parse_odso_properties(&bag).unwrap();
        let [OdsoProperty::FieldMap(info)] = properties.as_slice() else {
            panic!("expected field map info");
        };
        assert_eq!(info.mappings.len(), FIELD_MAP_COUNT as usize);
        assert_eq!(info.mappings[2].column_index, Some(3));
        assert_eq!(info.mappings[2].column_name.as_deref(), Some("GivenName"));
        // 0xFFFFFFFF means "not mapped".
        assert_eq!(info.mappings[19].column_index, None);
        assert_eq!(FieldMapInfo::STANDARD_ADDRESS_FIELDS[2], "First Name");
        assert_eq!(FieldMapInfo::STANDARD_ADDRESS_FIELDS[29], "Department");
    }

    #[test]
    fn rejects_malformed_field_map_info() {
        let empty = || vec![Vec::new(); FIELD_MAP_COUNT as usize];
        let wrap = |value: &[u8]| parse_odso_properties(&odso_item(ODSO_ID_FIELD_MAP, value));
        // Wrong count marker.
        let mut bad = field_map_info_bytes(&empty());
        bad[1] = 1;
        assert!(wrap(&bad).is_err());
        // Wrong field count.
        let mut bad = field_map_info_bytes(&empty());
        bad[4] = 29;
        assert!(wrap(&bad).is_err());
        // Wrong list size marker.
        let mut bad = field_map_info_bytes(&empty());
        bad[9] = 2;
        assert!(wrap(&bad).is_err());
        // Mapped flag other than 1.
        let mut mappings = empty();
        mappings[0] = vec![(FIELD_MAP_MAPPED, 2u32.to_le_bytes().to_vec())];
        assert!(wrap(&field_map_info_bytes(&mappings)).is_err());
        // Undefined item id.
        let mut mappings = empty();
        mappings[0] = vec![(0x0005, 0u32.to_le_bytes().to_vec())];
        assert!(wrap(&field_map_info_bytes(&mappings)).is_err());
        // Missing terminator on the last mapping.
        let mut good = field_map_info_bytes(&empty());
        let total = good.len();
        good[10] = (total - 12 - 4) as u8;
        good.truncate(total - 4);
        assert!(wrap(&good).is_err());
    }

    #[test]
    fn parses_pms_new_through_the_fib() {
        const FIB_POINTERS: usize = 145;

        fn set_fib_pointer(fib: &mut [u8], index: usize, offset: u32, length: u32) {
            let declared = u16::from_le_bytes([fib[152], fib[153]]);
            let count = declared.max(u16::try_from(index + 1).unwrap());
            fib[152..154].copy_from_slice(&count.to_le_bytes());
            let start = 154 + index * 8;
            fib[start..start + 4].copy_from_slice(&offset.to_le_bytes());
            fib[start + 4..start + 8].copy_from_slice(&length.to_le_bytes());
        }

        let mut fib_data = vec![0; 154 + FIB_POINTERS * 8];
        fib_data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
        fib_data[2..4].copy_from_slice(&0x0101u16.to_le_bytes());
        fib_data[152..154].copy_from_slice(&(FIB_POINTERS as u16).to_le_bytes());

        let mut builder = PmsBuilder::new();
        builder.irec_cur = 7;
        let pms_new = builder.build();
        set_fib_pointer(&mut fib_data, FC_PMS_NEW, 0, pms_new.len() as u32);
        let fib = FileInformationBlock::parse(&fib_data).unwrap();

        let mail_merge = DocumentMailMerge::parse(&fib, &pms_new)
            .unwrap()
            .expect("merge state present");
        assert!(mail_merge.state().is_none());
        assert_eq!(
            mail_merge.new_state().and_then(|pms| pms.current_record),
            Some(7)
        );

        // A malformed PmsNew is reported, not ignored.
        let mut fib_data = fib.raw_data().to_vec();
        set_fib_pointer(&mut fib_data, FC_PMS_NEW, 0, 1);
        let fib = FileInformationBlock::parse(&fib_data).unwrap();
        assert!(DocumentMailMerge::parse(&fib, &pms_new).is_err());
    }
}

