//! BIFF8 `QUERYTABLE` record sequence (MS-XLS 2.1.7.20.5): typed, inert
//! reading of worksheet query tables and their external data connections.
//!
//! A query table is described by the record sequence
//! `Qsi DBQUERY QsiSXTag DBQUERYEXT [SXADDLQSI] [QSIR] [SORTDATA12]`
//! (MS-XLS 2.1.7.20.5), where `DBQUERY = DbOrParamQry ...` (DbQuery and
//! ParamQry share record type 0x00DC and are disambiguated by the preceding
//! record, MS-XLS 2.4.79) and
//! `DBQUERYEXT = DBQueryExt [ExtString] *4[OleDbConn *ExtString] [TxtQry *ExtString]`.
//!
//! Everything in this module is INERT: connection strings, SQL command text,
//! Web query URLs, post data, and text-source file paths are stored verbatim
//! and are never opened, resolved, contacted, refreshed, or executed.
//!
//! `SORTDATA12` interaction: the optional trailing `SortData` record
//! (0x0895, MS-XLS 2.4.264) shares its record type with the sheet-level
//! `SORTANDFILTER` sort definition. The worksheet walker feeds this
//! collector before [`super::sort_data::SortDataCollector`]; while a
//! `QUERYTABLE` sequence is open, the `SortData` record and its declared
//! `ContinueFrt12` records are consumed here (payloads preserved verbatim in
//! [`QueryTable::sort_data_bytes`]) so they are never mis-attributed as
//! the worksheet sort. `SXADDLQSI` (`SXAddl` records of the `SxcQsi` class)
//! and `QSIR` (`Qsir`/`Qsif` formatting records) are consumed but not
//! interpreted. Malformed records never abort worksheet parsing: a broken
//! core record (Qsi/DbQuery/DBQueryExt) drops the in-progress sequence and a
//! broken optional record is ignored.

use litchi_core::binary;

use super::pivot_table::parse_qsi_sx_tag;
use super::records::Encoding;
use super::utils::parse_string_record;

/// `Qsi` record type (MS-XLS 2.4.208).
pub(crate) const QSI_RECORD_TYPE: u16 = 0x01AD;
/// `DbOrParamQry` record type: DbQuery or ParamQry (MS-XLS 2.4.79).
pub(crate) const DB_OR_PARAM_QRY_RECORD_TYPE: u16 = 0x00DC;
/// `SXString` record type (MS-XLS 2.4.304).
pub(crate) const SX_STRING_RECORD_TYPE: u16 = 0x00CD;
/// `QsiSXTag` record type (MS-XLS 2.4.211).
pub(crate) const QSI_SX_TAG_RECORD_TYPE: u16 = 0x0802;
/// `DBQueryExt` record type (MS-XLS 2.4.81).
pub(crate) const DB_QUERY_EXT_RECORD_TYPE: u16 = 0x0803;
/// `ExtString` record type (MS-XLS 2.4.108).
pub(crate) const EXT_STRING_RECORD_TYPE: u16 = 0x0804;
/// `TxtQry` record type (MS-XLS 2.4.330).
pub(crate) const TXT_QRY_RECORD_TYPE: u16 = 0x0805;
/// `Qsir` record type (MS-XLS 2.4.210).
pub(crate) const QSIR_RECORD_TYPE: u16 = 0x0806;
/// `Qsif` record type (MS-XLS 2.4.209).
pub(crate) const QSIF_RECORD_TYPE: u16 = 0x0807;
/// `OleDbConn` record type (MS-XLS 2.4.186).
pub(crate) const OLE_DB_CONN_RECORD_TYPE: u16 = 0x080A;

/// `SXAddl` record type (MS-XLS 2.4.273.1); only the `SxcQsi` class belongs
/// to a `QUERYTABLE` sequence.
const SXADDL_RECORD_TYPE: u16 = 0x0864;
/// `SortData` record type (MS-XLS 2.4.264); trailing `SORTDATA12` member.
const SORT_DATA_RECORD_TYPE: u16 = 0x0895;
/// `ContinueFrt12` record type (MS-XLS 2.4.62); continues a `SortData`.
const CONTINUE_FRT12_RECORD_TYPE: u16 = 0x087F;
/// `SXAddl` class of the `SxcQsi` class records (MS-XLS 2.2.5.1.1).
const SXC_QSI_CLASS: u8 = 0x05;
/// `SXAddl` record kind ending a class sequence.
const SXD_END: u8 = 0xFF;
/// Maximum number of `TxtWf` field descriptors in a `TxtQry` record.
const MAX_TXT_FIELDS: usize = 256;

// Qsi flag bits (first flag word).
const QSI_TITLES: u16 = 0x0001;
const QSI_ROW_NUMS: u16 = 0x0002;
const QSI_DISABLE_REFRESH: u16 = 0x0004;
const QSI_ASYNC: u16 = 0x0008;
const QSI_NEW_ASYNC: u16 = 0x0010;
const QSI_AUTO_REFRESH: u16 = 0x0020;
const QSI_SHRINK: u16 = 0x0040;
const QSI_FILL: u16 = 0x0080;
const QSI_AUTO_FORMAT: u16 = 0x0100;
const QSI_SAVE_DATA: u16 = 0x0200;
const QSI_DISABLE_EDIT: u16 = 0x0400;
const QSI_OVERWRITE: u16 = 0x2000;

// Qsi AutoFormat attribute bits (second flag word).
const QSI_ATR_NUM: u16 = 0x0001;
const QSI_ATR_FNT: u16 = 0x0002;
const QSI_ATR_ALC: u16 = 0x0004;
const QSI_ATR_BDR: u16 = 0x0008;
const QSI_ATR_PAT: u16 = 0x0010;
const QSI_ATR_PROT: u16 = 0x0020;

// DbQuery flag bits.
const DBQUERY_DBT_MASK: u16 = 0x0007;
const DBQUERY_ODBC_CONN: u16 = 0x0008;
const DBQUERY_SQL: u16 = 0x0010;
const DBQUERY_SQL_SAV: u16 = 0x0020;
const DBQUERY_WEB: u16 = 0x0040;
const DBQUERY_SAVE_PWD: u16 = 0x0080;
const DBQUERY_TABLES_ONLY_HTML: u16 = 0x0100;

// DBQueryExt flag bits.
const DBEXT_MAINTAIN: u16 = 0x0001;
const DBEXT_NEW_QUERY: u16 = 0x0002;
const DBEXT_IMPORT_XML_SOURCE: u16 = 0x0004;
const DBEXT_SP_LIST_SRC: u16 = 0x0008;
const DBEXT_SP_LIST_REINIT_CACHE: u16 = 0x0010;
const DBEXT_SRC_IS_XML: u16 = 0x0080;

// DBQueryExt trailing flag bits.
const DBEXT_TABLE_NAMES: u16 = 0x0002;

// TxtQry flag bits (first flag word).
const TXT_FILE: u16 = 0x0001;
const TXT_DELIMITED: u16 = 0x0002;
const TXT_CPID_SHIFT: u16 = 2;
const TXT_CPID_MASK: u16 = 0x0003;
const TXT_PROMPT_FOR_FILE: u16 = 0x0010;
const TXT_USE_NEW_CPID: u16 = 0x8000;

// TxtQry delimiter flag bits (second flag byte).
const TXT_DELIM_TAB: u8 = 0x01;
const TXT_DELIM_SPACE: u8 = 0x02;
const TXT_DELIM_COMMA: u8 = 0x04;
const TXT_DELIM_SEMICOLON: u8 = 0x08;
const TXT_DELIM_CUSTOM: u8 = 0x10;
const TXT_DELIM_CONSECUTIVE: u8 = 0x20;
const TXT_TEXT_DELIM_SHIFT: u8 = 6;
const TXT_TEXT_DELIM_MASK: u8 = 0x03;

// OleDbConn flag bits.
const OLECONN_PASSWD_STRIPPED: u16 = 0x0001;
const OLECONN_LOCAL: u16 = 0x0002;

// ParamQry fixed flags.
const PARAMQRY_PBT_MASK: u16 = 0x0003;
const PARAMQRY_NON_DEFAULT_NAME: u16 = 0x0004;

/// Data source type of an external connection (MS-XLS 2.5.64
/// `DataSourceType`, also the 3-bit `dbt` field of DbQuery, MS-XLS 2.4.80).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum QuerySource {
    /// ODBC-based source (`DBT_ODBC`).
    Odbc,
    /// DAO record set (`DBT_DAO`).
    Dao,
    /// Web query (`DBT_WEB`); the command text is a URL.
    Web,
    /// OLE DB-based source (`DBT_OLEDB`).
    OleDb,
    /// Text-based source created via a text query (`DBT_TXT`).
    Text,
    /// ADO record set (`DBT_ADO`).
    Ado,
    /// A value outside the `DataSourceType` enumeration.
    Unknown(u16),
}

impl QuerySource {
    fn from_dbt(dbt: u16) -> Self {
        match dbt {
            0x0001 => Self::Odbc,
            0x0002 => Self::Dao,
            0x0004 => Self::Web,
            0x0005 => Self::OleDb,
            0x0006 => Self::Text,
            0x0007 => Self::Ado,
            other => Self::Unknown(other),
        }
    }
}

/// Parameter kind of a parameterized query (MS-XLS 2.5.197
/// `PARAMQRY_Fixed.pbt`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum QueryParameterType {
    /// The user is prompted for the value of the parameter.
    Prompt,
    /// The parameter value is specified in the query.
    Value,
    /// The parameter value is specified in a cell.
    CellReference,
    /// A value outside the `pbt` enumeration.
    Unknown(u16),
}

/// A parameter of a parameterized query (SXString name followed by a
/// ParamQry record, MS-XLS 2.4.190).
#[derive(Debug, Clone, PartialEq)]
pub struct QueryParameter {
    /// Parameter name from the preceding SXString record.
    pub name: String,
    /// Parameter kind (`PARAMQRY_Fixed.pbt`).
    pub parameter_type: QueryParameterType,
    /// ODBC SQL data type of the parameter (`PARAMQRY_Fixed.wTypeSql`).
    pub sql_type: u16,
    /// Whether a non-default prompt is stored (`fNonDefaultName`).
    pub non_default_name: bool,
    /// Type selector of the stored parameter value (`PARAMQRY_Fixed.grbit`).
    pub value_type: u16,
    /// Prompt string stored for prompt parameters (`pbt` = Prompt), stored
    /// verbatim and never shown to anyone by this library.
    pub prompt: Option<String>,
}

/// Code page of a text-query source file (`TxtQry.iCpid`, MS-XLS 2.4.330).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextCodePage {
    /// Macintosh code page.
    Macintosh,
    /// Windows (ANSI) code page.
    WindowsAnsi,
    /// MS-DOS (PC-8) code page.
    MsDos,
    /// A value outside the `iCpid` enumeration.
    Unknown(u16),
}

/// Text qualifier of a delimited text query (`TxtQry.iTextDelm`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextDelimiter {
    /// Quotation mark.
    QuotationMark,
    /// Apostrophe.
    Apostrophe,
    /// No text delimiter.
    None,
    /// A value outside the `iTextDelm` enumeration.
    Unknown(u8),
}

/// Column import format of a text-query field (`TxtWf.fieldType`, MS-XLS
/// 2.5.273).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum TextFieldFormat {
    /// General.
    General,
    /// Text.
    Text,
    /// Date in month, day, year order.
    DateMdy,
    /// Date in day, month, year order.
    DateDmy,
    /// Date in year, month, day order.
    DateYmd,
    /// Date in month, year, day order.
    DateMyd,
    /// Date in day, year, month order.
    DateDym,
    /// Date in year, day, month order.
    DateYdm,
    /// Skip importing the field.
    Skip,
    /// A value outside the `fieldType` enumeration.
    Unknown(u32),
}

/// One field of a text query (`TxtWf`, MS-XLS 2.5.273).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextField {
    /// Column import format.
    pub format: TextFieldFormat,
    /// Zero-based character position of the field.
    pub start: i32,
}

/// Text query settings (`TxtQry`, MS-XLS 2.4.330). The source file path is
/// stored verbatim and is never opened.
#[derive(Debug, Clone, PartialEq)]
pub struct TextQuery {
    /// Whether the source data is delimited (false = fixed-width fields).
    pub delimited: bool,
    /// Code page of the source file.
    pub codepage: TextCodePage,
    /// Raw application-specific code page hint (`iCpidNew`).
    pub new_codepage: u16,
    /// Whether `new_codepage` supersedes `codepage` (`fUseNewiCpid`).
    pub use_new_codepage: bool,
    /// Whether a file name is prompted for on refresh.
    pub prompt_for_file: bool,
    /// Row in the source file where the query begins.
    pub row_start_at: i32,
    /// Tab is a column delimiter.
    pub tab: bool,
    /// Space is a column delimiter.
    pub space: bool,
    /// Comma is a column delimiter.
    pub comma: bool,
    /// Semicolon is a column delimiter.
    pub semicolon: bool,
    /// Custom delimiter character (`chCustom`), when `fCustom` is set.
    pub custom_delimiter: Option<char>,
    /// Consecutive delimiters are treated as one.
    pub consecutive: bool,
    /// Text qualifier for delimited fields.
    pub text_delimiter: TextDelimiter,
    /// Decimal separator of the source file.
    pub decimal_separator: char,
    /// Thousands separator of the source file.
    pub thousands_separator: char,
    /// Per-field widths/formats (`rgtxtwf`).
    pub fields: Vec<TextField>,
    /// Path of the source text file (`rgchFile`), stored verbatim and never
    /// opened.
    pub file: String,
    /// Connection string chunks following the `TxtQry` record, concatenated.
    pub connection_string: String,
}

/// An OLE DB connection of an external connection (OleDbConn followed by its
/// ExtString records, MS-XLS 2.4.186). The connection string is stored
/// verbatim and is never used.
#[derive(Debug, Clone, PartialEq)]
pub struct OleDbConnection {
    /// Whether the password was stripped from the connection string
    /// (`fPasswd`).
    pub password_stripped: bool,
    /// Whether this is the main (false) or an alternate (true) connection
    /// string (`fLocal`).
    pub local: bool,
    /// Concatenated ExtString chunks of the connection string.
    pub connection_string: String,
}

/// HTML formatting applied to imported Web query data (`DBQueryExt.wHtmlFmt`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HtmlFormatting {
    /// No formatting is applied.
    None,
    /// Rich text formatting only.
    RichText,
    /// Full HTML formatting, including cell formatting.
    Full,
    /// A value outside the `wHtmlFmt` enumeration.
    Unknown(u16),
}

/// A worksheet query table: the typed, inert result of a `QUERYTABLE`
/// record sequence (MS-XLS 2.1.7.20.5).
///
/// All connection strings, command text, URLs, and file paths are stored
/// verbatim and are never opened, resolved, contacted, refreshed, or
/// executed.
#[derive(Debug, Clone, PartialEq)]
pub struct QueryTable {
    /// Query table name (`Qsi.rgchName`).
    pub name: String,
    /// First row of the query table contains column titles.
    pub titles: bool,
    /// First column displays row numbers.
    pub row_numbers: bool,
    /// The query table cannot be refreshed (`fDisableRefresh`).
    pub disable_refresh: bool,
    /// Refresh of the query table is enabled, per the bound `QsiSXTag`
    /// (`fEnableRefresh`); `None` when no matching tag was present.
    pub enable_refresh: Option<bool>,
    /// The query table refreshes data asynchronously.
    pub async_refresh: bool,
    /// The first background refresh had not finished when the file was saved
    /// (`fNewAsync`).
    pub first_refresh_pending: bool,
    /// Refresh automatically when the document is opened.
    pub auto_refresh: bool,
    /// Unused cells are deleted (rather than cleared) on refresh.
    pub shrink: bool,
    /// Formulas in adjacent columns are filled down on refresh.
    pub fill: bool,
    /// Query table data is saved with the document.
    pub save_data: bool,
    /// The query table content is locked against editing.
    pub disable_edit: bool,
    /// New data overwrites existing cells (rather than inserting).
    pub overwrite: bool,
    /// AutoFormat flag (`fAutoFormat`; unused per MS-XLS).
    pub auto_format: bool,
    /// AutoFormat table index (`itblAutoFmt`).
    pub auto_format_index: u16,
    /// AutoFormat applies to numeric cell data.
    pub auto_format_number: bool,
    /// AutoFormat applies to cell text.
    pub auto_format_font: bool,
    /// AutoFormat applies to cell text alignment.
    pub auto_format_alignment: bool,
    /// AutoFormat applies to borders.
    pub auto_format_border: bool,
    /// AutoFormat applies to patterns.
    pub auto_format_pattern: bool,
    /// AutoFormat applies to cell protection.
    pub auto_format_protection: bool,
    /// Additional option flags from the bound `QsiSXTag` (`dwQsiFuture`),
    /// preserved verbatim; zero when no tag was bound.
    pub qsi_future: u32,
    /// Data source type of the external connection.
    pub source: QuerySource,
    /// The database connection remains open once established.
    pub maintain_connection: bool,
    /// The connection was never refreshed (`fNewQuery`).
    pub new_query: bool,
    /// The underlying XML source (rather than the Web page table) is
    /// imported (Web queries only).
    pub import_xml_source: bool,
    /// The connection uses the Web-based data provider (`fSPListSrc`).
    pub sharepoint_list_source: bool,
    /// Web-based data is reinitialized rather than refreshed.
    pub sharepoint_list_reinit: bool,
    /// The external connection source is XML (`fSrcIsXml`).
    pub source_is_xml: bool,
    /// The password is kept in the ODBC connection string (`fSavePwd`).
    pub save_password: bool,
    /// Web queries only work on HTML tables (`fTablesOnlyHTML`).
    pub tables_only_html: bool,
    /// Command text: the SQL statement or the Web query URL, concatenated
    /// from its SXString chunks. Stored verbatim, never executed or
    /// contacted.
    pub command_text: Option<String>,
    /// ODBC connection string, concatenated from its SXString chunks.
    /// Stored verbatim, never used.
    pub connection_string: Option<String>,
    /// Web query post statement, concatenated from its SXString chunks.
    pub web_post: Option<String>,
    /// SQL statement for server-based fields (`cstSQLSav` chunks).
    pub sql_server_fields: Option<String>,
    /// Query parameters with their prompts, in record order.
    pub parameters: Vec<QueryParameter>,
    /// Comma-delimited list of table names to import (ExtString after
    /// DBQueryExt when `fTableNames` is set).
    pub table_names: Option<String>,
    /// Raw `ConnGrbitDbt` flags of the DBQueryExt record (`grbitDbt`).
    pub connection_flags: u16,
    /// Data functionality level the connection was last edited with.
    pub edited_version: u8,
    /// Data functionality level the connection was last refreshed with.
    pub refreshed_version: u8,
    /// Minimum data functionality level able to refresh the connection.
    pub refreshable_min_version: u8,
    /// Minutes between automatic refreshes; 0 disables timed refresh.
    pub refresh_interval: u16,
    /// HTML formatting applied to imported Web query data.
    pub html_formatting: HtmlFormatting,
    /// Raw `PBT` items describing the query parameters (`rgPbt`).
    pub parameter_flags: Vec<u16>,
    /// Text query settings, when this is a text query.
    pub text_query: Option<Box<TextQuery>>,
    /// OLE DB connections, in record order.
    pub ole_db_connections: Vec<OleDbConnection>,
    /// `rgbFutureBytes` of the DBQueryExt record, preserved verbatim.
    pub future_bytes: Vec<u8>,
    /// Concatenated payloads of the trailing `SORTDATA12` member (`SortData`
    /// plus its `ContinueFrt12` records), preserved verbatim; empty when the
    /// sequence carried no sort definition.
    pub sort_data_bytes: Vec<u8>,
}

impl Default for QueryTable {
    fn default() -> Self {
        Self {
            name: String::new(),
            titles: false,
            row_numbers: false,
            disable_refresh: false,
            enable_refresh: None,
            async_refresh: false,
            first_refresh_pending: false,
            auto_refresh: false,
            shrink: false,
            fill: false,
            save_data: false,
            disable_edit: false,
            overwrite: false,
            auto_format: false,
            auto_format_index: 0,
            auto_format_number: false,
            auto_format_font: false,
            auto_format_alignment: false,
            auto_format_border: false,
            auto_format_pattern: false,
            auto_format_protection: false,
            qsi_future: 0,
            source: QuerySource::Unknown(0),
            maintain_connection: false,
            new_query: false,
            import_xml_source: false,
            sharepoint_list_source: false,
            sharepoint_list_reinit: false,
            source_is_xml: false,
            save_password: false,
            tables_only_html: false,
            command_text: None,
            connection_string: None,
            web_post: None,
            sql_server_fields: None,
            parameters: Vec::new(),
            table_names: None,
            connection_flags: 0,
            edited_version: 0,
            refreshed_version: 0,
            refreshable_min_version: 0,
            refresh_interval: 0,
            html_formatting: HtmlFormatting::None,
            parameter_flags: Vec::new(),
            text_query: None,
            ole_db_connections: Vec::new(),
            future_bytes: Vec::new(),
            sort_data_bytes: Vec::new(),
        }
    }
}

/// Which trailing ExtString record is expected next within a `DBQUERYEXT`
/// collection.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ExtContext {
    None,
    TableNames,
    OleDb(usize),
    TextQuery,
}

/// In-progress assembly of one `QUERYTABLE` sequence.
#[derive(Debug)]
struct QueryTableBuild {
    table: QueryTable,
    dbquery_seen: bool,
    /// Previous record was an SXString or a ParamQry: a following 0x00DC is
    /// a ParamQry rather than a DbQuery (MS-XLS 2.4.79).
    last_string_or_param: bool,
    remaining_query: u16,
    remaining_odbc_conn: u16,
    remaining_web_post: u16,
    remaining_sql_sav: u16,
    pending_param_name: Option<String>,
    query_chunks: Vec<String>,
    odbc_conn_chunks: Vec<String>,
    web_post_chunks: Vec<String>,
    sql_sav_chunks: Vec<String>,
    ext_context: ExtContext,
    ole_db_remaining: u16,
    in_sxaddl_qsi: bool,
    sort_data_remaining: Option<u32>,
}

impl QueryTableBuild {
    fn new(table: QueryTable) -> Self {
        Self {
            table,
            dbquery_seen: false,
            last_string_or_param: false,
            remaining_query: 0,
            remaining_odbc_conn: 0,
            remaining_web_post: 0,
            remaining_sql_sav: 0,
            pending_param_name: None,
            query_chunks: Vec::new(),
            odbc_conn_chunks: Vec::new(),
            web_post_chunks: Vec::new(),
            sql_sav_chunks: Vec::new(),
            ext_context: ExtContext::None,
            ole_db_remaining: 0,
            in_sxaddl_qsi: false,
            sort_data_remaining: None,
        }
    }

    fn finish(mut self) -> QueryTable {
        if !self.query_chunks.is_empty() {
            self.table.command_text = Some(self.query_chunks.concat());
        }
        if !self.odbc_conn_chunks.is_empty() {
            self.table.connection_string = Some(self.odbc_conn_chunks.concat());
        }
        if !self.web_post_chunks.is_empty() {
            self.table.web_post = Some(self.web_post_chunks.concat());
        }
        if !self.sql_sav_chunks.is_empty() {
            self.table.sql_server_fields = Some(self.sql_sav_chunks.concat());
        }
        self.table
    }
}

fn parse_qsi(data: &[u8]) -> Option<QueryTable> {
    if data.len() < 13 {
        return None;
    }
    let flags = binary::read_u16_le_at(data, 0).ok()?;
    let attributes = binary::read_u16_le_at(data, 2).ok()?;
    let name = parse_string_record(&data[10..], &Encoding::Utf16Le).ok()?;
    Some(QueryTable {
        name,
        titles: flags & QSI_TITLES != 0,
        row_numbers: flags & QSI_ROW_NUMS != 0,
        disable_refresh: flags & QSI_DISABLE_REFRESH != 0,
        async_refresh: flags & QSI_ASYNC != 0,
        first_refresh_pending: flags & QSI_NEW_ASYNC != 0,
        auto_refresh: flags & QSI_AUTO_REFRESH != 0,
        shrink: flags & QSI_SHRINK != 0,
        fill: flags & QSI_FILL != 0,
        auto_format: flags & QSI_AUTO_FORMAT != 0,
        save_data: flags & QSI_SAVE_DATA != 0,
        disable_edit: flags & QSI_DISABLE_EDIT != 0,
        overwrite: flags & QSI_OVERWRITE != 0,
        auto_format_index: binary::read_u16_le_at(data, 4).ok()?,
        auto_format_number: attributes & QSI_ATR_NUM != 0,
        auto_format_font: attributes & QSI_ATR_FNT != 0,
        auto_format_alignment: attributes & QSI_ATR_ALC != 0,
        auto_format_border: attributes & QSI_ATR_BDR != 0,
        auto_format_pattern: attributes & QSI_ATR_PAT != 0,
        auto_format_protection: attributes & QSI_ATR_PROT != 0,
        ..QueryTable::default()
    })
}

fn parse_db_query(build: &mut QueryTableBuild, data: &[u8]) -> Option<()> {
    if data.len() < 12 {
        return None;
    }
    let flags = binary::read_u16_le_at(data, 0).ok()?;
    build.table.source = QuerySource::from_dbt(flags & DBQUERY_DBT_MASK);
    build.table.save_password = flags & DBQUERY_SAVE_PWD != 0;
    build.table.tables_only_html = flags & DBQUERY_TABLES_ONLY_HTML != 0;
    build.remaining_query = string_count(data, 4, flags & (DBQUERY_SQL | DBQUERY_WEB) != 0)?;
    build.remaining_odbc_conn = string_count(data, 10, flags & DBQUERY_ODBC_CONN != 0)?;
    build.remaining_web_post = string_count(data, 6, flags & DBQUERY_WEB != 0)?;
    build.remaining_sql_sav = string_count(data, 8, flags & DBQUERY_SQL_SAV != 0)?;
    build.dbquery_seen = true;
    Some(())
}

/// A declared SXString chunk count; negative counts are clamped to zero.
fn string_count(data: &[u8], offset: usize, present: bool) -> Option<u16> {
    let count = binary::read_i16_le(data, offset).ok()?;
    Some(if present { count.max(0) as u16 } else { 0 })
}

fn parse_param_qry(build: &mut QueryTableBuild, data: &[u8]) -> Option<()> {
    if data.len() < 8 {
        return None;
    }
    let sql_type = binary::read_u16_le_at(data, 0).ok()?;
    let flags = binary::read_u16_le_at(data, 2).ok()?;
    let value_type = binary::read_u16_le_at(data, 4).ok()?;
    let pbt = flags & PARAMQRY_PBT_MASK;
    let parameter_type = match pbt {
        0 => QueryParameterType::Prompt,
        1 => QueryParameterType::Value,
        2 => QueryParameterType::CellReference,
        other => QueryParameterType::Unknown(other),
    };
    // For prompt parameters the trailing field is an SXString prompt
    // followed by an unused byte; other value encodings stay uninterpreted.
    let prompt = if pbt == 0 && data.len() > 8 {
        parse_string_record(&data[8..], &Encoding::Utf16Le).ok()
    } else {
        None
    };
    build.table.parameters.push(QueryParameter {
        name: build.pending_param_name.take().unwrap_or_default(),
        parameter_type,
        sql_type,
        non_default_name: flags & PARAMQRY_NON_DEFAULT_NAME != 0,
        value_type,
        prompt,
    });
    Some(())
}

fn parse_db_query_ext(build: &mut QueryTableBuild, data: &[u8]) -> Option<()> {
    if data.len() < 28 || binary::read_u16_le_at(data, 0).ok()? != DB_QUERY_EXT_RECORD_TYPE {
        return None;
    }
    let flags = binary::read_u16_le_at(data, 6).ok()?;
    let trailing = binary::read_u16_le_at(data, 10).ok()?;
    let parameter_count = usize::from(binary::read_u16_le_at(data, 26).ok()?);
    let parameter_end = 28usize.checked_add(parameter_count.checked_mul(2)?)?;
    if data.len() < parameter_end {
        return None;
    }
    let future_count = usize::from(binary::read_u16_le_at(data, 20).ok()?);
    let future_end = parameter_end.checked_add(future_count)?;
    // The DBQueryExt DataSourceType supersedes the 3-bit DbQuery dbt.
    build.table.source = QuerySource::from_dbt(binary::read_u16_le_at(data, 4).ok()?);
    build.table.maintain_connection = flags & DBEXT_MAINTAIN != 0;
    build.table.new_query = flags & DBEXT_NEW_QUERY != 0;
    build.table.import_xml_source = flags & DBEXT_IMPORT_XML_SOURCE != 0;
    build.table.sharepoint_list_source = flags & DBEXT_SP_LIST_SRC != 0;
    build.table.sharepoint_list_reinit = flags & DBEXT_SP_LIST_REINIT_CACHE != 0;
    build.table.source_is_xml = flags & DBEXT_SRC_IS_XML != 0;
    build.table.connection_flags = binary::read_u16_le_at(data, 8).ok()?;
    build.table.edited_version = data[12];
    build.table.refreshed_version = data[13];
    build.table.refreshable_min_version = data[14];
    build.table.refresh_interval = binary::read_u16_le_at(data, 22).ok()?;
    build.table.html_formatting = match binary::read_u16_le_at(data, 24).ok()? {
        0x0001 => HtmlFormatting::None,
        0x0002 => HtmlFormatting::RichText,
        0x0003 => HtmlFormatting::Full,
        other => HtmlFormatting::Unknown(other),
    };
    build.table.parameter_flags = data[28..parameter_end]
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
        .collect();
    build.table.future_bytes = data[parameter_end..future_end.min(data.len())].to_vec();
    build.ext_context = if trailing & DBEXT_TABLE_NAMES != 0 {
        ExtContext::TableNames
    } else {
        ExtContext::None
    };
    Some(())
}

fn parse_txt_qry(data: &[u8]) -> Option<TextQuery> {
    if data.len() < 22 || binary::read_u16_le_at(data, 0).ok()? != TXT_QRY_RECORD_TYPE {
        return None;
    }
    let flags = binary::read_u16_le_at(data, 4).ok()?;
    if flags & TXT_FILE == 0 {
        return None;
    }
    let delimiters = data[12];
    let field_count = binary::read_i32_le(data, 16).ok()?;
    if field_count <= 0 || field_count as usize > MAX_TXT_FIELDS {
        return None;
    }
    let fields_end = 22usize.checked_add((field_count as usize).checked_mul(8)?)?;
    if data.len() < fields_end {
        return None;
    }
    let fields = data[22..fields_end]
        .chunks_exact(8)
        .map(|chunk| TextField {
            format: match u32::from_le_bytes([chunk[0], chunk[1], chunk[2], chunk[3]]) {
                0 => TextFieldFormat::General,
                1 => TextFieldFormat::Text,
                2 => TextFieldFormat::DateMdy,
                3 => TextFieldFormat::DateDmy,
                4 => TextFieldFormat::DateYmd,
                5 => TextFieldFormat::DateMyd,
                6 => TextFieldFormat::DateDym,
                7 => TextFieldFormat::DateYdm,
                8 => TextFieldFormat::Skip,
                other => TextFieldFormat::Unknown(other),
            },
            start: i32::from_le_bytes([chunk[4], chunk[5], chunk[6], chunk[7]]),
        })
        .collect();
    let file = parse_string_record(&data[fields_end..], &Encoding::Utf16Le).ok()?;
    Some(TextQuery {
        delimited: flags & TXT_DELIMITED != 0,
        codepage: match (flags >> TXT_CPID_SHIFT) & TXT_CPID_MASK {
            0 => TextCodePage::Macintosh,
            1 => TextCodePage::WindowsAnsi,
            2 => TextCodePage::MsDos,
            other => TextCodePage::Unknown(other),
        },
        new_codepage: (flags >> 5) & 0x03FF,
        use_new_codepage: flags & TXT_USE_NEW_CPID != 0,
        prompt_for_file: flags & TXT_PROMPT_FOR_FILE != 0,
        row_start_at: binary::read_i32_le(data, 8).ok()?,
        tab: delimiters & TXT_DELIM_TAB != 0,
        space: delimiters & TXT_DELIM_SPACE != 0,
        comma: delimiters & TXT_DELIM_COMMA != 0,
        semicolon: delimiters & TXT_DELIM_SEMICOLON != 0,
        custom_delimiter: if delimiters & TXT_DELIM_CUSTOM != 0 {
            char::from_u32(u32::from(binary::read_u16_le_at(data, 13).ok()?))
        } else {
            None
        },
        consecutive: delimiters & TXT_DELIM_CONSECUTIVE != 0,
        text_delimiter: match (delimiters >> TXT_TEXT_DELIM_SHIFT) & TXT_TEXT_DELIM_MASK {
            0 => TextDelimiter::QuotationMark,
            1 => TextDelimiter::Apostrophe,
            2 => TextDelimiter::None,
            other => TextDelimiter::Unknown(other),
        },
        decimal_separator: char::from(data[20]),
        thousands_separator: char::from(data[21]),
        fields,
        file,
        connection_string: String::new(),
    })
}

/// Ordered worksheet `QUERYTABLE` sequence collector. Multiple query tables
/// per sheet are supported. See the module documentation for the inertness
/// contract and the `SORTDATA12` interaction.
#[derive(Debug, Default)]
pub(crate) struct QueryTableCollector {
    completed: Vec<QueryTable>,
    current: Option<QueryTableBuild>,
}

impl QueryTableCollector {
    pub(crate) fn new() -> Self {
        Self::default()
    }

    fn finalize_current(&mut self) {
        if let Some(build) = self.current.take() {
            self.completed.push(build.finish());
        }
    }

    /// Returns true when the record belongs to a `QUERYTABLE` sequence.
    ///
    /// Never fails: malformed core records drop the in-progress sequence and
    /// malformed optional records are ignored, so a broken query table can
    /// not abort worksheet parsing.
    pub(crate) fn feed_record(&mut self, record_type: u16, data: &[u8]) -> bool {
        if record_type == QSI_RECORD_TYPE {
            // A new Qsi starts a new sequence (one query table each).
            self.finalize_current();
            if let Some(table) = parse_qsi(data) {
                self.current = Some(QueryTableBuild::new(table));
            }
            return true;
        }

        let Some(build) = self.current.as_mut() else {
            return false;
        };

        // A pending SortData only accepts its declared ContinueFrt12 records.
        if let Some(remaining) = build.sort_data_remaining {
            if record_type == CONTINUE_FRT12_RECORD_TYPE && remaining > 0 {
                build.table.sort_data_bytes.extend_from_slice(data);
                build.sort_data_remaining = Some(remaining - 1);
                return true;
            }
            build.sort_data_remaining = None;
            self.finalize_current();
            return false;
        }

        match record_type {
            DB_OR_PARAM_QRY_RECORD_TYPE => {
                // MS-XLS 2.4.79: after an SXString or a ParamQry the record
                // is a ParamQry; after anything else it is a DbQuery.
                let was_param = build.last_string_or_param;
                let parsed = if was_param {
                    parse_param_qry(build, data)
                } else {
                    parse_db_query(build, data)
                };
                if parsed.is_none() {
                    // Malformed core record: drop the in-progress sequence.
                    self.current = None;
                    return true;
                }
                // A DbQuery restarts the disambiguation chain; a ParamQry
                // extends it.
                self.current
                    .as_mut()
                    .expect("build present")
                    .last_string_or_param = was_param;
                true
            },
            SX_STRING_RECORD_TYPE => {
                let Ok(text) = parse_string_record(data, &Encoding::Utf16Le) else {
                    // Malformed chunk: drop the in-progress sequence.
                    self.current = None;
                    return true;
                };
                let build = self.current.as_mut().expect("build present");
                if build.remaining_query > 0 {
                    build.remaining_query -= 1;
                    build.query_chunks.push(text);
                } else if build.remaining_odbc_conn > 0 {
                    build.remaining_odbc_conn -= 1;
                    build.odbc_conn_chunks.push(text);
                } else if build.remaining_web_post > 0 {
                    build.remaining_web_post -= 1;
                    build.web_post_chunks.push(text);
                } else if build.remaining_sql_sav > 0 {
                    build.remaining_sql_sav -= 1;
                    build.sql_sav_chunks.push(text);
                } else {
                    // A parameter name preceding its ParamQry record.
                    build.pending_param_name = Some(text);
                }
                build.last_string_or_param = true;
                true
            },
            QSI_SX_TAG_RECORD_TYPE => {
                let Ok(tag) = parse_qsi_sx_tag(data) else {
                    // Malformed tag: ignored, the sequence continues.
                    return true;
                };
                if tag.table_type != 0 {
                    // fSx=1: the tag and its collection belong to a
                    // PivotTable view; hand it back to the pivot collector.
                    self.finalize_current();
                    return false;
                }
                let build = self.current.as_mut().expect("build present");
                if tag.table_name == build.table.name {
                    build.table.enable_refresh = Some(tag.flags & 0x0001 != 0);
                    build.table.qsi_future = tag.options;
                }
                // Name mismatches are ignored per MS-XLS 2.4.211.
                true
            },
            DB_QUERY_EXT_RECORD_TYPE => {
                if parse_db_query_ext(build, data).is_none() {
                    // Malformed core record: drop the in-progress sequence.
                    self.current = None;
                }
                true
            },
            EXT_STRING_RECORD_TYPE => {
                let text = if data.len() >= 7 {
                    parse_string_record(&data[4..], &Encoding::Utf16Le).ok()
                } else {
                    None
                };
                let Some(text) = text else { return true };
                let build = self.current.as_mut().expect("build present");
                match build.ext_context {
                    ExtContext::TableNames => {
                        build.table.table_names = Some(text);
                        build.ext_context = ExtContext::None;
                    },
                    ExtContext::OleDb(index) => {
                        if let Some(connection) = build.table.ole_db_connections.get_mut(index) {
                            connection.connection_string.push_str(&text);
                        }
                        build.ole_db_remaining = build.ole_db_remaining.saturating_sub(1);
                        if build.ole_db_remaining == 0 {
                            build.ext_context = ExtContext::None;
                        }
                    },
                    ExtContext::TextQuery => {
                        if let Some(text_query) = build.table.text_query.as_mut() {
                            text_query.connection_string.push_str(&text);
                        }
                    },
                    ExtContext::None => {},
                }
                true
            },
            TXT_QRY_RECORD_TYPE => {
                if let Some(text_query) = parse_txt_qry(data) {
                    build.table.text_query = Some(Box::new(text_query));
                    build.ext_context = ExtContext::TextQuery;
                }
                true
            },
            OLE_DB_CONN_RECORD_TYPE => {
                if data.len() >= 8
                    && binary::read_u16_le_at(data, 0).ok() == Some(OLE_DB_CONN_RECORD_TYPE)
                {
                    let flags = binary::read_u16_le_at(data, 4).unwrap_or(0);
                    build.table.ole_db_connections.push(OleDbConnection {
                        password_stripped: flags & OLECONN_PASSWD_STRIPPED != 0,
                        local: flags & OLECONN_LOCAL != 0,
                        connection_string: String::new(),
                    });
                    build.ole_db_remaining = binary::read_u16_le_at(data, 6).unwrap_or(0);
                    build.ext_context = ExtContext::OleDb(build.table.ole_db_connections.len() - 1);
                }
                true
            },
            // QSIR formatting records are consumed but not interpreted.
            QSIR_RECORD_TYPE | QSIF_RECORD_TYPE => true,
            SXADDL_RECORD_TYPE => {
                let sxc_qsi = data.len() >= 6 && data[4] == SXC_QSI_CLASS;
                if sxc_qsi {
                    build.in_sxaddl_qsi = data[5] != SXD_END;
                    true
                } else {
                    // Another SXAddl class: not part of this sequence.
                    self.finalize_current();
                    false
                }
            },
            SORT_DATA_RECORD_TYPE => {
                build.table.sort_data_bytes.extend_from_slice(data);
                let conditions = if data.len() >= 34 {
                    binary::read_u32_le_at(data, 30).unwrap_or(0)
                } else {
                    0
                };
                build.sort_data_remaining = Some(conditions);
                true
            },
            _ => {
                self.finalize_current();
                false
            },
        }
    }

    pub(crate) fn finish(mut self) -> Vec<QueryTable> {
        self.finalize_current();
        self.completed
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Encode an `XLUnicodeString` (compressed, 16-bit character count).
    fn xl_string(text: &str) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&(text.len() as u16).to_le_bytes());
        out.push(0); // compressed (one byte per character)
        out.extend_from_slice(text.as_bytes());
        out
    }

    fn qsi(name: &str, flags: u16) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&flags.to_le_bytes());
        out.extend_from_slice(&QSI_ATR_PAT.to_le_bytes());
        out.extend_from_slice(&18u16.to_le_bytes()); // itblAutoFmt
        out.extend_from_slice(&0u32.to_le_bytes()); // reserved
        out.extend_from_slice(&xl_string(name));
        out.extend_from_slice(&0u16.to_le_bytes()); // unused4
        out
    }

    #[allow(clippy::too_many_arguments)]
    fn db_query(
        flags: u16,
        cparams: i16,
        cst_query: i16,
        cst_web_post: i16,
        cst_sql_sav: i16,
        cst_odbc_conn: i16,
    ) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&flags.to_le_bytes());
        out.extend_from_slice(&cparams.to_le_bytes());
        out.extend_from_slice(&cst_query.to_le_bytes());
        out.extend_from_slice(&cst_web_post.to_le_bytes());
        out.extend_from_slice(&cst_sql_sav.to_le_bytes());
        out.extend_from_slice(&cst_odbc_conn.to_le_bytes());
        out
    }

    fn qsi_sx_tag(name: &str, f_sx: u16, flags: u16, future: u32) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&QSI_SX_TAG_RECORD_TYPE.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
        out.extend_from_slice(&f_sx.to_le_bytes());
        out.extend_from_slice(&flags.to_le_bytes());
        out.extend_from_slice(&future.to_le_bytes());
        out.push(0); // verSxLastUpdated
        out.push(0); // verSxUpdatableMin
        out.push(16); // obCchName
        out.push(0); // reserved2
        out.extend_from_slice(&xl_string(name));
        out.extend_from_slice(&0u16.to_le_bytes()); // unused
        out
    }

    #[allow(clippy::too_many_arguments)]
    fn db_query_ext(
        dbt: u16,
        flags: u16,
        grbit_dbt: u16,
        trailing: u16,
        coledb: u16,
        future_bytes: &[u8],
        refresh_interval: u16,
        html_fmt: u16,
        parameter_flags: &[u16],
    ) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&DB_QUERY_EXT_RECORD_TYPE.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
        out.extend_from_slice(&dbt.to_le_bytes());
        out.extend_from_slice(&flags.to_le_bytes());
        out.extend_from_slice(&grbit_dbt.to_le_bytes());
        out.extend_from_slice(&trailing.to_le_bytes());
        out.push(3); // bVerDbqueryEdit
        out.push(2); // bVerDbqueryRefreshed
        out.push(1); // bVerDbqueryRefreshableMin
        out.push(0); // reserved4
        out.extend_from_slice(&0u16.to_le_bytes()); // reserved5
        out.extend_from_slice(&coledb.to_le_bytes());
        out.extend_from_slice(&(future_bytes.len() as u16).to_le_bytes());
        out.extend_from_slice(&refresh_interval.to_le_bytes());
        out.extend_from_slice(&html_fmt.to_le_bytes());
        out.extend_from_slice(&(parameter_flags.len() as u16).to_le_bytes());
        for flag in parameter_flags {
            out.extend_from_slice(&flag.to_le_bytes());
        }
        out.extend_from_slice(future_bytes);
        out
    }

    fn ext_string(text: &str) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&EXT_STRING_RECORD_TYPE.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
        out.extend_from_slice(&xl_string(text));
        out
    }

    fn txt_qry(
        flags: u16,
        delimiters: u8,
        custom: Option<char>,
        row_start: i32,
        fields: &[(u32, i32)],
        file: &str,
    ) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&TXT_QRY_RECORD_TYPE.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes()); // reserved
        out.extend_from_slice(&flags.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes()); // unused1
        out.extend_from_slice(&row_start.to_le_bytes());
        out.push(delimiters);
        out.extend_from_slice(&custom.map_or(0u16, |c| c as u16).to_le_bytes());
        out.push(0); // unused2
        out.extend_from_slice(&(fields.len() as i32).to_le_bytes());
        out.push(b'.'); // chDecimal
        out.push(b','); // chThousSep
        for (format, start) in fields {
            out.extend_from_slice(&format.to_le_bytes());
            out.extend_from_slice(&start.to_le_bytes());
        }
        out.extend_from_slice(&xl_string(file));
        out
    }

    fn ole_db_conn(flags: u16, cst: u16) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&OLE_DB_CONN_RECORD_TYPE.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
        out.extend_from_slice(&flags.to_le_bytes());
        out.extend_from_slice(&cst.to_le_bytes());
        out.extend_from_slice(&0u32.to_le_bytes()); // reserved2
        out
    }

    fn param_qry(sql_type: u16, pbt: u16, grbit: u16, prompt: Option<&str>) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(&sql_type.to_le_bytes());
        out.extend_from_slice(&pbt.to_le_bytes());
        out.extend_from_slice(&grbit.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes()); // fVal
        if let Some(prompt) = prompt {
            out.extend_from_slice(&xl_string(prompt));
            out.push(0); // unused byte
        }
        out
    }

    fn sort_data(conditions: u32) -> Vec<u8> {
        let mut out = vec![0xAA; 34];
        out[30..34].copy_from_slice(&conditions.to_le_bytes());
        out
    }

    /// A record type that never belongs to a QUERYTABLE sequence.
    const OTHER_RECORD: u16 = 0x000A;

    #[test]
    fn text_query_sequence() {
        let mut collector = QueryTableCollector::new();
        assert!(collector.feed_record(QSI_RECORD_TYPE, &qsi("Table1", 0x2209)));
        assert!(collector.feed_record(
            DB_OR_PARAM_QRY_RECORD_TYPE,
            &db_query(0x0006, 0, 0, 0, 0, 0)
        ));
        assert!(collector.feed_record(QSI_SX_TAG_RECORD_TYPE, &qsi_sx_tag("Table1", 0, 0x0001, 3)));
        assert!(collector.feed_record(
            DB_QUERY_EXT_RECORD_TYPE,
            &db_query_ext(0x0006, 0, 0, 0, 0, &[1, 2, 3], 15, 0, &[]),
        ));
        let txt = txt_qry(
            0x0001 | 0x0002 | (1 << 2), // fFile, fDelimited, iCpid = Windows (ANSI)
            TXT_DELIM_COMMA | TXT_DELIM_CUSTOM,
            Some('|'),
            2,
            &[(0, 0), (1, 5)],
            "D:\\data\\table1.csv",
        );
        assert!(collector.feed_record(TXT_QRY_RECORD_TYPE, &txt));
        assert!(!collector.feed_record(OTHER_RECORD, &[]));

        let tables = collector.finish();
        assert_eq!(tables.len(), 1);
        let table = &tables[0];
        assert_eq!(table.name, "Table1");
        assert!(table.titles);
        assert!(table.async_refresh);
        assert!(table.save_data);
        assert!(!table.disable_edit);
        assert!(table.overwrite);
        assert!(!table.shrink);
        assert_eq!(table.auto_format_index, 18);
        assert!(table.auto_format_pattern);
        assert_eq!(table.enable_refresh, Some(true));
        assert_eq!(table.qsi_future, 3);
        assert_eq!(table.source, QuerySource::Text);
        assert_eq!(table.refresh_interval, 15);
        assert_eq!(table.edited_version, 3);
        assert_eq!(table.future_bytes, vec![1, 2, 3]);
        let text_query = table.text_query.as_ref().expect("text query present");
        assert!(text_query.delimited);
        assert_eq!(text_query.codepage, TextCodePage::WindowsAnsi);
        assert_eq!(text_query.row_start_at, 2);
        assert!(text_query.comma);
        assert!(!text_query.tab);
        assert_eq!(text_query.custom_delimiter, Some('|'));
        assert_eq!(text_query.text_delimiter, TextDelimiter::QuotationMark);
        assert_eq!(text_query.decimal_separator, '.');
        assert_eq!(text_query.thousands_separator, ',');
        assert_eq!(
            text_query.fields,
            vec![
                TextField {
                    format: TextFieldFormat::General,
                    start: 0
                },
                TextField {
                    format: TextFieldFormat::Text,
                    start: 5
                },
            ]
        );
        assert_eq!(text_query.file, "D:\\data\\table1.csv");
    }

    #[test]
    fn odbc_query_with_parameters() {
        let flags = 0x0001 | DBQUERY_SQL | DBQUERY_ODBC_CONN | DBQUERY_SAVE_PWD;
        let mut collector = QueryTableCollector::new();
        collector.feed_record(QSI_RECORD_TYPE, &qsi("Odbc1", 0));
        collector.feed_record(DB_OR_PARAM_QRY_RECORD_TYPE, &db_query(flags, 1, 2, 0, 0, 1));
        collector.feed_record(
            SX_STRING_RECORD_TYPE,
            &xl_string("SELECT * FROM t WHERE id = "),
        );
        collector.feed_record(SX_STRING_RECORD_TYPE, &xl_string("?"));
        collector.feed_record(SX_STRING_RECORD_TYPE, &xl_string("DSN=warehouse;UID=scott"));
        collector.feed_record(SX_STRING_RECORD_TYPE, &xl_string("id"));
        collector.feed_record(
            DB_OR_PARAM_QRY_RECORD_TYPE,
            &param_qry(4, 0, 0, Some("Enter id")),
        );
        collector.feed_record(QSI_SX_TAG_RECORD_TYPE, &qsi_sx_tag("Odbc1", 0, 0, 0));
        collector.feed_record(
            DB_QUERY_EXT_RECORD_TYPE,
            &db_query_ext(0x0001, DBEXT_MAINTAIN, 0, 0, 0, &[], 0, 0, &[0x0004]),
        );
        collector.feed_record(OTHER_RECORD, &[]);

        let tables = collector.finish();
        assert_eq!(tables.len(), 1);
        let table = &tables[0];
        assert_eq!(table.source, QuerySource::Odbc);
        assert!(table.maintain_connection);
        assert!(table.save_password);
        assert_eq!(
            table.command_text.as_deref(),
            Some("SELECT * FROM t WHERE id = ?")
        );
        assert_eq!(
            table.connection_string.as_deref(),
            Some("DSN=warehouse;UID=scott")
        );
        assert_eq!(table.parameter_flags, vec![0x0004]);
        assert_eq!(table.parameters.len(), 1);
        let parameter = &table.parameters[0];
        assert_eq!(parameter.name, "id");
        assert_eq!(parameter.parameter_type, QueryParameterType::Prompt);
        assert_eq!(parameter.sql_type, 4);
        assert_eq!(parameter.prompt.as_deref(), Some("Enter id"));
    }

    #[test]
    fn web_query_with_post_data() {
        let flags = 0x0004 | DBQUERY_WEB | DBQUERY_TABLES_ONLY_HTML;
        let mut collector = QueryTableCollector::new();
        collector.feed_record(QSI_RECORD_TYPE, &qsi("Web1", 0));
        collector.feed_record(DB_OR_PARAM_QRY_RECORD_TYPE, &db_query(flags, 0, 2, 1, 0, 0));
        collector.feed_record(
            SX_STRING_RECORD_TYPE,
            &xl_string("https://example.invalid/report?a="),
        );
        collector.feed_record(SX_STRING_RECORD_TYPE, &xl_string("1"));
        collector.feed_record(SX_STRING_RECORD_TYPE, &xl_string("q=1&r=2"));
        collector.feed_record(QSI_SX_TAG_RECORD_TYPE, &qsi_sx_tag("Web1", 0, 0, 0));
        collector.feed_record(
            DB_QUERY_EXT_RECORD_TYPE,
            &db_query_ext(0x0004, 0, 0x0023, DBEXT_TABLE_NAMES, 0, &[], 0, 0x0002, &[]),
        );
        collector.feed_record(EXT_STRING_RECORD_TYPE, &ext_string("Table1,Table2"));
        collector.feed_record(OTHER_RECORD, &[]);

        let tables = collector.finish();
        assert_eq!(tables.len(), 1);
        let table = &tables[0];
        assert_eq!(table.source, QuerySource::Web);
        assert!(table.tables_only_html);
        assert_eq!(
            table.command_text.as_deref(),
            Some("https://example.invalid/report?a=1")
        );
        assert_eq!(table.web_post.as_deref(), Some("q=1&r=2"));
        assert_eq!(table.connection_flags, 0x0023);
        assert_eq!(table.html_formatting, HtmlFormatting::RichText);
        assert_eq!(table.table_names.as_deref(), Some("Table1,Table2"));
    }

    #[test]
    fn ole_db_query_with_chunked_connection() {
        let mut collector = QueryTableCollector::new();
        collector.feed_record(QSI_RECORD_TYPE, &qsi("Ole1", 0));
        collector.feed_record(
            DB_OR_PARAM_QRY_RECORD_TYPE,
            &db_query(0x0005 | DBQUERY_SQL, 0, 1, 0, 0, 0),
        );
        collector.feed_record(SX_STRING_RECORD_TYPE, &xl_string("SELECT 1"));
        collector.feed_record(QSI_SX_TAG_RECORD_TYPE, &qsi_sx_tag("Ole1", 0, 0, 0));
        collector.feed_record(
            DB_QUERY_EXT_RECORD_TYPE,
            &db_query_ext(0x0005, 0, 0, 0, 1, &[], 0, 0, &[]),
        );
        collector.feed_record(OLE_DB_CONN_RECORD_TYPE, &ole_db_conn(0x0001, 2));
        collector.feed_record(EXT_STRING_RECORD_TYPE, &ext_string("Provider=SQLOLEDB;"));
        collector.feed_record(EXT_STRING_RECORD_TYPE, &ext_string("Data Source=srv"));
        collector.feed_record(OTHER_RECORD, &[]);

        let tables = collector.finish();
        assert_eq!(tables.len(), 1);
        let table = &tables[0];
        assert_eq!(table.source, QuerySource::OleDb);
        assert_eq!(table.ole_db_connections.len(), 1);
        let connection = &table.ole_db_connections[0];
        assert!(connection.password_stripped);
        assert!(!connection.local);
        assert_eq!(
            connection.connection_string,
            "Provider=SQLOLEDB;Data Source=srv"
        );
    }

    #[test]
    fn multiple_query_tables_per_sheet() {
        let mut collector = QueryTableCollector::new();
        for name in ["First", "Second"] {
            collector.feed_record(QSI_RECORD_TYPE, &qsi(name, 0));
            collector.feed_record(
                DB_OR_PARAM_QRY_RECORD_TYPE,
                &db_query(0x0006, 0, 0, 0, 0, 0),
            );
            collector.feed_record(QSI_SX_TAG_RECORD_TYPE, &qsi_sx_tag(name, 0, 0, 0));
            collector.feed_record(
                DB_QUERY_EXT_RECORD_TYPE,
                &db_query_ext(0x0006, 0, 0, 0, 0, &[], 0, 0, &[]),
            );
        }
        let tables = collector.finish();
        assert_eq!(tables.len(), 2);
        assert_eq!(tables[0].name, "First");
        assert_eq!(tables[1].name, "Second");
    }

    #[test]
    fn pivot_tag_is_handed_back() {
        // fSx=1 while a sequence is open: the tag belongs to a PivotTable
        // view and must not be consumed; the pending table is completed
        // without a bound tag.
        let mut collector = QueryTableCollector::new();
        collector.feed_record(QSI_RECORD_TYPE, &qsi("Table1", 0));
        collector.feed_record(
            DB_OR_PARAM_QRY_RECORD_TYPE,
            &db_query(0x0006, 0, 0, 0, 0, 0),
        );
        assert!(
            !collector.feed_record(QSI_SX_TAG_RECORD_TYPE, &qsi_sx_tag("PivotTable1", 1, 0, 0))
        );
        let tables = collector.finish();
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].enable_refresh, None);
        assert_eq!(tables[0].qsi_future, 0);
    }

    #[test]
    fn pivot_tag_without_query_table_is_ignored() {
        let mut collector = QueryTableCollector::new();
        assert!(
            !collector.feed_record(QSI_SX_TAG_RECORD_TYPE, &qsi_sx_tag("PivotTable1", 1, 0, 0))
        );
        assert!(!collector.feed_record(
            DB_QUERY_EXT_RECORD_TYPE,
            &db_query_ext(0x0005, 0, 0, 0, 0, &[], 0, 0, &[]),
        ));
        assert!(collector.finish().is_empty());
    }

    #[test]
    fn mismatched_tag_name_is_ignored() {
        let mut collector = QueryTableCollector::new();
        collector.feed_record(QSI_RECORD_TYPE, &qsi("Table1", 0));
        collector.feed_record(
            DB_OR_PARAM_QRY_RECORD_TYPE,
            &db_query(0x0006, 0, 0, 0, 0, 0),
        );
        assert!(collector.feed_record(QSI_SX_TAG_RECORD_TYPE, &qsi_sx_tag("Other", 0, 0x0001, 7)));
        collector.feed_record(
            DB_QUERY_EXT_RECORD_TYPE,
            &db_query_ext(0x0006, 0, 0, 0, 0, &[], 0, 0, &[]),
        );
        let tables = collector.finish();
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].enable_refresh, None);
        assert_eq!(tables[0].source, QuerySource::Text);
    }

    #[test]
    fn sort_data_inside_sequence_is_consumed() {
        let mut collector = QueryTableCollector::new();
        collector.feed_record(QSI_RECORD_TYPE, &qsi("Table1", 0));
        collector.feed_record(
            DB_OR_PARAM_QRY_RECORD_TYPE,
            &db_query(0x0006, 0, 0, 0, 0, 0),
        );
        collector.feed_record(QSI_SX_TAG_RECORD_TYPE, &qsi_sx_tag("Table1", 0, 0, 0));
        collector.feed_record(
            DB_QUERY_EXT_RECORD_TYPE,
            &db_query_ext(0x0006, 0, 0, 0, 0, &[], 0, 0, &[]),
        );
        let base = sort_data(1);
        assert!(collector.feed_record(SORT_DATA_RECORD_TYPE, &base));
        assert!(collector.feed_record(CONTINUE_FRT12_RECORD_TYPE, &[0xBB; 12]));
        assert!(!collector.feed_record(OTHER_RECORD, &[]));

        let tables = collector.finish();
        assert_eq!(tables.len(), 1);
        let mut expected = base.clone();
        expected.extend_from_slice(&[0xBB; 12]);
        assert_eq!(tables[0].sort_data_bytes, expected);
    }

    #[test]
    fn sort_data_with_missing_continues_still_completes() {
        let mut collector = QueryTableCollector::new();
        collector.feed_record(QSI_RECORD_TYPE, &qsi("Table1", 0));
        collector.feed_record(
            DB_OR_PARAM_QRY_RECORD_TYPE,
            &db_query(0x0006, 0, 0, 0, 0, 0),
        );
        let base = sort_data(3);
        assert!(collector.feed_record(SORT_DATA_RECORD_TYPE, &base));
        // No ContinueFrt12 records arrive; the sequence just ends.
        assert!(!collector.feed_record(OTHER_RECORD, &[]));
        let tables = collector.finish();
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].sort_data_bytes, base);
    }

    #[test]
    fn sxaddl_qsi_class_is_consumed() {
        let sxaddl = |sxd: u8| {
            let mut out = Vec::new();
            out.extend_from_slice(&SXADDL_RECORD_TYPE.to_le_bytes());
            out.extend_from_slice(&0u16.to_le_bytes());
            out.push(SXC_QSI_CLASS);
            out.push(sxd);
            out.extend_from_slice(&[0; 6]);
            out
        };
        let mut collector = QueryTableCollector::new();
        collector.feed_record(QSI_RECORD_TYPE, &qsi("Table1", 0));
        collector.feed_record(
            DB_OR_PARAM_QRY_RECORD_TYPE,
            &db_query(0x0006, 0, 0, 0, 0, 0),
        );
        collector.feed_record(QSI_SX_TAG_RECORD_TYPE, &qsi_sx_tag("Table1", 0, 0, 0));
        collector.feed_record(
            DB_QUERY_EXT_RECORD_TYPE,
            &db_query_ext(0x0006, 0, 0, 0, 0, &[], 0, 0, &[]),
        );
        assert!(collector.feed_record(SXADDL_RECORD_TYPE, &sxaddl(0x00))); // SXDId
        assert!(collector.feed_record(SXADDL_RECORD_TYPE, &sxaddl(SXD_END)));
        assert!(!collector.feed_record(OTHER_RECORD, &[]));
        assert_eq!(collector.finish().len(), 1);
    }

    #[test]
    fn foreign_sxaddl_class_ends_sequence() {
        let other_class = {
            let mut out = Vec::new();
            out.extend_from_slice(&SXADDL_RECORD_TYPE.to_le_bytes());
            out.extend_from_slice(&0u16.to_le_bytes());
            out.push(0x01); // SxcView class
            out.push(0x00);
            out
        };
        let mut collector = QueryTableCollector::new();
        collector.feed_record(QSI_RECORD_TYPE, &qsi("Table1", 0));
        collector.feed_record(
            DB_OR_PARAM_QRY_RECORD_TYPE,
            &db_query(0x0006, 0, 0, 0, 0, 0),
        );
        assert!(!collector.feed_record(SXADDL_RECORD_TYPE, &other_class));
        assert_eq!(collector.finish().len(), 1);
    }

    #[test]
    fn malformed_records_never_panic() {
        // Truncated Qsi: consumed, no sequence started.
        let mut collector = QueryTableCollector::new();
        assert!(collector.feed_record(QSI_RECORD_TYPE, &[0x01, 0x02]));
        assert!(collector.finish().is_empty());

        // Truncated DbQuery drops the in-progress sequence.
        let mut collector = QueryTableCollector::new();
        collector.feed_record(QSI_RECORD_TYPE, &qsi("Table1", 0));
        assert!(collector.feed_record(DB_OR_PARAM_QRY_RECORD_TYPE, &[0x06, 0x00]));
        assert!(collector.finish().is_empty());

        // Truncated DBQueryExt drops the in-progress sequence.
        let mut collector = QueryTableCollector::new();
        collector.feed_record(QSI_RECORD_TYPE, &qsi("Table1", 0));
        collector.feed_record(
            DB_OR_PARAM_QRY_RECORD_TYPE,
            &db_query(0x0006, 0, 0, 0, 0, 0),
        );
        assert!(collector.feed_record(DB_QUERY_EXT_RECORD_TYPE, &[0x03, 0x08, 0x00]));
        assert!(collector.finish().is_empty());

        // Malformed SXString drops the in-progress sequence.
        let mut collector = QueryTableCollector::new();
        collector.feed_record(QSI_RECORD_TYPE, &qsi("Table1", 0));
        collector.feed_record(
            DB_OR_PARAM_QRY_RECORD_TYPE,
            &db_query(0x0001 | DBQUERY_SQL, 0, 1, 0, 0, 0),
        );
        assert!(collector.feed_record(SX_STRING_RECORD_TYPE, &[0xFF, 0xFF, 0x01]));
        assert!(collector.finish().is_empty());

        // Malformed TxtQry is ignored; the table still completes.
        let mut collector = QueryTableCollector::new();
        collector.feed_record(QSI_RECORD_TYPE, &qsi("Table1", 0));
        collector.feed_record(
            DB_OR_PARAM_QRY_RECORD_TYPE,
            &db_query(0x0006, 0, 0, 0, 0, 0),
        );
        collector.feed_record(QSI_SX_TAG_RECORD_TYPE, &qsi_sx_tag("Table1", 0, 0, 0));
        collector.feed_record(
            DB_QUERY_EXT_RECORD_TYPE,
            &db_query_ext(0x0006, 0, 0, 0, 0, &[], 0, 0, &[]),
        );
        assert!(collector.feed_record(TXT_QRY_RECORD_TYPE, &[0x05, 0x08]));
        collector.feed_record(OTHER_RECORD, &[]);
        let tables = collector.finish();
        assert_eq!(tables.len(), 1);
        assert!(tables[0].text_query.is_none());

        // Declared counts exceeding the payload are clamped/dropped.
        let mut collector = QueryTableCollector::new();
        collector.feed_record(QSI_RECORD_TYPE, &qsi("Table1", 0));
        collector.feed_record(
            DB_OR_PARAM_QRY_RECORD_TYPE,
            &db_query(0x0006, 0, 0, 0, 0, 0),
        );
        let mut ext = db_query_ext(0x0006, 0, 0, 0, 0, &[], 0, 0, &[]);
        ext[20] = 0xFF; // cstFuture claims 255 bytes that are absent
        ext[21] = 0x00;
        assert!(collector.feed_record(DB_QUERY_EXT_RECORD_TYPE, &ext));
        collector.feed_record(OTHER_RECORD, &[]);
        assert_eq!(collector.finish().len(), 1);
    }

    /// Feed every record of a fixture's Workbook stream through the
    /// collector, mirroring the worksheet walker.
    fn collect_from_fixture(path: &str) -> Vec<QueryTable> {
        let bytes = std::fs::read(path).expect("fixture readable");
        let mut ole = litchi_cfb::OleFile::open(std::io::Cursor::new(bytes))
            .expect("fixture is a CFB container");
        let stream = ole
            .open_stream(&["Workbook"])
            .or_else(|_| ole.open_stream(&["Book"]))
            .expect("workbook stream present");
        let mut collector = QueryTableCollector::new();
        let mut offset = 0;
        while offset + 4 <= stream.len() {
            let record_type = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
            let length = u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]) as usize;
            let body_end = (offset + 4 + length).min(stream.len());
            collector.feed_record(record_type, &stream[offset + 4..body_end]);
            offset = body_end;
        }
        collector.finish()
    }

    /// 57456.xls declares an SST whose total count underflows its unique
    /// count, which the (deliberately strict) workbook parser rejects before
    /// any worksheet is reached; walk the fixture's records directly so the
    /// real bytes still drive the collector.
    #[test]
    fn web_query_57456_fixture() {
        let tables = collect_from_fixture("../../test-data/poi/test-data/spreadsheet/57456.xls");
        assert_eq!(tables.len(), 1);
        let table = &tables[0];
        assert_eq!(table.name, "ExternalData_1");
        assert_eq!(table.source, QuerySource::Web);
        assert!(table.titles);
        assert!(table.async_refresh);
        assert!(table.auto_refresh);
        assert!(table.shrink);
        assert_eq!(table.enable_refresh, Some(false));
        assert_eq!(table.qsi_future, 3);
        assert_eq!(table.html_formatting, HtmlFormatting::None);
        assert_eq!(table.connection_flags, 0x0023);
        assert!(table.text_query.is_none());
        let url = table.command_text.as_deref().expect("web query URL");
        assert!(
            url.starts_with("http://bugstop.lenexa.ibm.com:8080/disp_bugs.php?flow=&comp=totals")
        );
        assert!(url.ends_with("txmgt:wlm&x=z"));
        assert_eq!(url.len(), 867);
        assert!(table.connection_string.is_none());
        assert!(table.web_post.is_none());
    }

    /// The text-query fixtures parse through the public workbook API; this
    /// cross-checks the raw record walk against the same bytes.
    #[test]
    fn text_query_45365_fixture() {
        let tables = collect_from_fixture("../../test-data/poi/test-data/spreadsheet/45365.xls");
        assert_eq!(tables.len(), 1);
        let table = &tables[0];
        assert_eq!(table.name, "Jac-Jackson-MSC_1");
        assert_eq!(table.source, QuerySource::Text);
        let text_query = table.text_query.as_ref().expect("text query present");
        assert_eq!(text_query.file, "D:\\Jac-Jackson-MSC_1.csv");
        assert_eq!(text_query.row_start_at, 1);
        assert!(text_query.tab);
        assert!(text_query.comma);
    }

    #[test]
    fn negative_chunk_counts_are_clamped() {
        let mut collector = QueryTableCollector::new();
        collector.feed_record(QSI_RECORD_TYPE, &qsi("Table1", 0));
        collector.feed_record(
            DB_OR_PARAM_QRY_RECORD_TYPE,
            &db_query(0x0001 | DBQUERY_SQL | DBQUERY_ODBC_CONN, -1, -2, -3, -4, -5),
        );
        collector.feed_record(OTHER_RECORD, &[]);
        let tables = collector.finish();
        assert_eq!(tables.len(), 1);
        assert_eq!(tables[0].command_text, None);
        assert_eq!(tables[0].connection_string, None);
    }
}
