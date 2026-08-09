/// Data source type of an external connection (MS-XLS 2.5.64
/// `DataSourceType`, also the 3-bit `dbt` field of `DbQuery`, MS-XLS 2.4.80).
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
    pub(crate) fn from_dbt(dbt: u16) -> Self {
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

/// A parameter of a parameterized query (`SXString` name followed by a
/// `ParamQry` record, MS-XLS 2.4.190).
#[derive(Debug, Clone, PartialEq)]
pub struct QueryParameter {
    /// Parameter name from the preceding `SXString` record.
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

/// An OLE DB connection of an external connection (`OleDbConn` followed by its
/// `ExtString` records, MS-XLS 2.4.186). The connection string is stored
/// verbatim and is never used.
#[derive(Debug, Clone, PartialEq)]
pub struct OleDbConnection {
    /// Whether the password was stripped from the connection string
    /// (`fPasswd`).
    pub password_stripped: bool,
    /// Whether this is the main (false) or an alternate (true) connection
    /// string (`fLocal`).
    pub local: bool,
    /// Concatenated `ExtString` chunks of the connection string.
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
    /// `AutoFormat` flag (`fAutoFormat`; unused per MS-XLS).
    pub auto_format: bool,
    /// `AutoFormat` table index (`itblAutoFmt`).
    pub auto_format_index: u16,
    /// `AutoFormat` applies to numeric cell data.
    pub auto_format_number: bool,
    /// `AutoFormat` applies to cell text.
    pub auto_format_font: bool,
    /// `AutoFormat` applies to cell text alignment.
    pub auto_format_alignment: bool,
    /// `AutoFormat` applies to borders.
    pub auto_format_border: bool,
    /// `AutoFormat` applies to patterns.
    pub auto_format_pattern: bool,
    /// `AutoFormat` applies to cell protection.
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
    /// from its `SXString` chunks. Stored verbatim, never executed or
    /// contacted.
    pub command_text: Option<String>,
    /// ODBC connection string, concatenated from its `SXString` chunks.
    /// Stored verbatim, never used.
    pub connection_string: Option<String>,
    /// Web query post statement, concatenated from its `SXString` chunks.
    pub web_post: Option<String>,
    /// SQL statement for server-based fields (`cstSQLSav` chunks).
    pub sql_server_fields: Option<String>,
    /// Query parameters with their prompts, in record order.
    pub parameters: Vec<QueryParameter>,
    /// Comma-delimited list of table names to import (`ExtString` after
    /// `DBQueryExt` when `fTableNames` is set).
    pub table_names: Option<String>,
    /// Raw `ConnGrbitDbt` flags of the `DBQueryExt` record (`grbitDbt`).
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
    /// `rgbFutureBytes` of the `DBQueryExt` record, preserved verbatim.
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
