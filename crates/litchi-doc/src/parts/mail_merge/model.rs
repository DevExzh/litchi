//! Semantic mail-merge models for the legacy Word `Pms` and ODSO parts.
//! Stored paths, connection strings, and SQL remain inert metadata.

use super::validation::{FNPI_IDENTIFIER_SHIFT, FNPI_TYPE_MAIL_MERGE, FNPI_TYPE_MASK};

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
    pub(super) raw: u16,
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

/// The mail merge connection and record filtering string table (`SttbfRfs`,
/// MS-DOC 2.9.289).
///
/// Connection strings are stored verbatim and are never opened or contacted.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SttbfRfs {
    pub(super) strings: Vec<String>,
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

/// How one filter comparison combines with the others
/// (`FilterDataItem.iCondition`, MS-DOC 2.9.87).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FilterCondition {
    /// Logical AND.
    And,
    /// Logical OR.
    Or,
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

/// A recipient sort direction (`SortColumnAndDirection.iDirection`,
/// MS-DOC 2.9.252).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SortDirection {
    /// Ascending order.
    Ascending,
    /// Descending order.
    Descending,
}

/// One recipient sort key (`SortColumnAndDirection`, MS-DOC 2.9.252).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SortColumnAndDirection {
    /// Zero-based index of the data source column to sort by.
    pub column: u32,
    /// The sort direction.
    pub direction: SortDirection,
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

/// The mail-merge data-source state of a document: the Word 97 `Pms` and the
/// Word 2002+ ODSO property set, addressed by `fcPms` and `fcODSO`.
///
/// All state is inert: data-source paths, connection strings, and SQL queries
/// are stored verbatim and never opened, resolved, contacted, or executed,
/// and no merge is performed.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DocumentMailMerge {
    pub(super) state: Option<Pms>,
    pub(super) new_state: Option<Pms>,
    pub(super) odso_properties: Vec<OdsoProperty>,
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
}
