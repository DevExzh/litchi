//! Typed, inert model of the XLSB External Data Connections part
//! (MS-XLSB 2.1.7.24).
//!
//! Connection strings, commands, URLs, file paths, and credential metadata
//! are stored exactly as declared and are never resolved, opened, contacted,
//! refreshed, or executed.

/// `DBType` (MS-XLSB 2.5.31): the data source type of a connection.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[repr(u32)]
pub enum XlsbConnectionSourceType {
    /// ODBC data source (`DBTODBC`).
    #[default]
    Odbc = 1,
    /// DAO data source (`DBTDAO`).
    Dao = 2,
    /// HTML data source (`DBTWEB`).
    Web = 4,
    /// OLE DB data source (`DBTOLEDB`).
    OleDb = 5,
    /// Text data source (`DBTTEXT`).
    Text = 6,
    /// ADO record set (`DBTADO`).
    Ado = 7,
    /// OLE DB data source created by the spreadsheet data model (`DBTOLEDBPP`).
    OleDbDataModel = 0x64,
    /// Data feed data source created by the spreadsheet data model (`DBTDATAFEED`).
    DataFeed = 0x65,
    /// Worksheet data source created by the spreadsheet data model (`DBTWORKSHEET`).
    Worksheet = 0x66,
    /// Text data source created by the spreadsheet data model (`DBTTEXTPP`).
    TextDataModel = 0x67,
}

impl TryFrom<u32> for XlsbConnectionSourceType {
    type Error = crate::xlsb::error::XlsbError;

    fn try_from(value: u32) -> Result<Self, crate::xlsb::error::XlsbError> {
        Ok(match value {
            1 => Self::Odbc,
            2 => Self::Dao,
            4 => Self::Web,
            5 => Self::OleDb,
            6 => Self::Text,
            7 => Self::Ado,
            0x64 => Self::OleDbDataModel,
            0x65 => Self::DataFeed,
            0x66 => Self::Worksheet,
            0x67 => Self::TextDataModel,
            other => {
                return Err(crate::xlsb::walker::malformed(
                    "BrtBeginExtConnection",
                    format!("unknown DBType {other:#x}"),
                ));
            },
        })
    }
}

/// `CmdType` (MS-XLSB 2.5.21): the meaning of the database command.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
#[repr(u32)]
pub enum XlsbCommandType {
    /// No command specified (`CMDNULL`).
    #[default]
    None = 0,
    /// Cube name within an OLAP database (`CMDCUBE`).
    Cube = 1,
    /// SQL statement (`CMDSQL`).
    Sql = 2,
    /// Database table name (`CMDTABLE`).
    Table = 3,
    /// Statement in the database's default language (`CMDDEFAULT`).
    Default = 4,
    /// List from a Web-based data provider (`CMDSPLIST`).
    SpList = 5,
}

impl TryFrom<u32> for XlsbCommandType {
    type Error = crate::xlsb::error::XlsbError;

    fn try_from(value: u32) -> Result<Self, crate::xlsb::error::XlsbError> {
        Ok(match value {
            0 => Self::None,
            1 => Self::Cube,
            2 => Self::Sql,
            3 => Self::Table,
            4 => Self::Default,
            5 => Self::SpList,
            other => {
                return Err(crate::xlsb::walker::malformed(
                    "BrtBeginECDbProps",
                    format!("unknown CmdType {other:#x}"),
                ));
            },
        })
    }
}

/// Whether the password is saved in the connection string (`pc` field).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum XlsbPasswordState {
    /// The password is saved in the connection string.
    Saved = 1,
    /// The password is not saved in the connection string.
    NotSaved = 2,
}

/// The authentication method for a database connection (`iCredMethod`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
pub enum XlsbCredentialMethod {
    /// Integrated authentication.
    Integrated = 0,
    /// No credentials.
    None = 1,
    /// Credentials stored in a single sign-on repository.
    SingleSignOn = 2,
}

/// When connection information is retrieved from the connection file
/// (`irecontype`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum XlsbReconnectionType {
    /// Retrieve updated information only after a refresh failure.
    AsRequired = 1,
    /// Always retrieve updated information from the connection file.
    Always = 2,
    /// Never retrieve updated information.
    Never = 3,
}

/// How HTML formatting is imported by a Web connection (`wHTMLFmt`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
#[derive(Default)]
pub enum XlsbHtmlFormat {
    /// No formatting is imported.
    #[default]
    None = 0,
    /// Rich-text formatting is imported.
    RichText = 1,
    /// All formatting is imported.
    All = 2,
    /// Any other value is preserved verbatim.
    Other(u8),
}

/// Database command properties (`BrtBeginECDbProps`, MS-XLSB 2.4.61).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct XlsbDbProperties {
    /// The database command type.
    pub command_type: XlsbCommandType,
    /// The connection string (inert).
    pub connection_string: String,
    /// The database command, when stored (inert).
    pub command: Option<String>,
    /// The server-based page-field command, when stored (inert).
    pub server_command: Option<String>,
}

/// OLAP connection properties (`BrtBeginECOlapProps`, MS-XLSB 2.4.62).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct XlsbOlapProperties {
    /// Whether data is retrieved from a local cube file.
    pub local_connection: bool,
    /// Whether the provider is requested not to rebuild the local cube file.
    pub no_refresh_cube: bool,
    /// Server-retrieved background color is applied to cells.
    pub server_format_back: bool,
    /// Server-retrieved font color is applied to cells.
    pub server_format_fore: bool,
    /// Server-retrieved font family is applied to cells.
    pub server_format_flags: bool,
    /// Server-retrieved format string is applied to cells.
    pub server_format_number: bool,
    /// Whether the office LCID is sent to the provider.
    pub use_office_lcid: bool,
    /// Maximum drillthrough rows.
    pub drillthrough_rows: u32,
    /// The local cube connection string, when stored (inert).
    pub local_connection_string: Option<String>,
}

/// Web connection properties (`BrtBeginECWebProps`, MS-XLSB 2.4.71).
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct XlsbWebProperties {
    /// How HTML formatting is imported.
    pub html_format: XlsbHtmlFormat,
    /// The source is XML rather than an HTML table.
    pub source_is_xml: bool,
    /// Data is imported from the URL rather than the table.
    pub import_source_data: bool,
    /// Pre-formatted blocks are parsed into columns.
    pub parse_pre_formatted: bool,
    /// Consecutive delimiters are treated as one.
    pub consecutive_delimiters: bool,
    /// Tables in a block share the first row's width settings.
    pub same_settings: bool,
    /// Created with the legacy application version.
    pub excel97_format: bool,
    /// Dates are imported as text.
    pub no_date_recognition: bool,
    /// Refreshed with a newer application version.
    pub refreshed_in_excel9: bool,
    /// Only works on HTML tables.
    pub tables_only_html: bool,
    /// The refresh URL (inert).
    pub url: Option<String>,
    /// The HTTP post string, when stored (inert).
    pub web_post: Option<String>,
    /// The user-facing edit page URL, when stored (inert).
    pub edit_web_page: Option<String>,
}

/// The source-specific property block of a connection.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub enum XlsbConnectionProperties {
    /// ODBC or OLE DB command properties.
    Database(XlsbDbProperties),
    /// OLAP connection properties.
    Olap(XlsbOlapProperties),
    /// Web connection properties.
    Web(XlsbWebProperties),
    /// No property block was stored (other or deleted connection kinds).
    #[default]
    None,
}

/// The type of a connection parameter (`pbt`).
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum XlsbParameterType {
    /// The user is prompted for the value.
    #[default]
    Prompt,
    /// A stored value (number, string, or Boolean).
    Value,
    /// A cell reference supplying the value.
    CellReference,
    /// Any other `pbt` value.
    Other(u8),
}

/// A stored parameter value.
#[derive(Debug, Clone, PartialEq)]
pub enum XlsbParameterValue {
    /// A numeric value (`xnumVal`).
    Number(f64),
    /// A string value (`stVal`).
    Text(String),
    /// A Boolean value (`boolVal`).
    Boolean(bool),
    /// Raw formula tokens of the cell reference (inert).
    CellFormula(Vec<u8>),
}

/// One connection parameter (`BrtBeginECParam`, MS-XLSB 2.4.63).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct XlsbConnectionParameter {
    /// The parameter type.
    pub parameter_type: XlsbParameterType,
    /// Whether data refreshes when the parameter cell changes.
    pub auto_refresh: bool,
    /// The SQL data type of the parameter (`TypeSql`, MS-XLSB 2.5.152).
    pub sql_type: u16,
    /// The parameter name.
    pub name: String,
    /// The prompt string, when stored.
    pub prompt: Option<String>,
    /// The stored value.
    pub value: Option<XlsbParameterValue>,
}

/// One Web query table reference (`BrtBeginEcWpTables` items).
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum XlsbWebTableItem {
    /// An HTML table by index (`BrtPCDIIndex`).
    Index(u32),
    /// An HTML table by `id` attribute (`BrtPCDIString`).
    Named(String),
    /// An invalid or missing reference (`BrtPCDIMissing`).
    Missing,
}

/// One external connection (`BrtBeginExtConnection`, MS-XLSB 2.4.80).
#[derive(Debug, Clone, Default, PartialEq)]
pub struct XlsbConnection {
    /// The unique connection identifier (`dwConnID`).
    pub connection_id: u32,
    /// The data source type (`idbtype`).
    pub source_type: XlsbConnectionSourceType,
    /// The connection name.
    pub name: String,
    /// The description, when stored.
    pub description: Option<String>,
    /// The external connection file path, when stored (inert).
    pub connection_file: Option<String>,
    /// The data file path, when stored (inert).
    pub data_file: Option<String>,
    /// The single sign-on application identifier, when stored.
    pub sso_id: Option<String>,
    /// The authentication method, when applicable.
    pub credential_method: Option<XlsbCredentialMethod>,
    /// Whether the password is saved in the connection string, when applicable.
    pub password_state: Option<XlsbPasswordState>,
    /// The data functionality level last refreshed with.
    pub refreshed_version: u8,
    /// The minimum data functionality level required to refresh.
    pub refreshable_min_version: u8,
    /// Automatic refresh interval in minutes (0 = never).
    pub refresh_interval_minutes: u16,
    /// Whether the connection is maintained after refresh.
    pub maintain: bool,
    /// Whether the connection has never been refreshed.
    pub new_query: bool,
    /// Whether the connection has been deleted.
    pub deleted: bool,
    /// Whether the connection file is always used on refresh.
    pub always_use_connection_file: bool,
    /// Whether refresh runs asynchronously in the background.
    pub background_query: bool,
    /// Whether the connection refreshes when the workbook opens.
    pub refresh_on_load: bool,
    /// Whether retrieved data is saved within the workbook.
    pub save_data: bool,
    /// When connection information is retrieved from the connection file.
    pub reconnection_type: Option<XlsbReconnectionType>,
    /// The source-specific property block.
    pub properties: XlsbConnectionProperties,
    /// Connection parameters.
    pub parameters: Vec<XlsbConnectionParameter>,
    /// Web query table references.
    pub web_tables: Vec<XlsbWebTableItem>,
}

/// The parsed External Data Connections part.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct XlsbConnections {
    /// The declared external connections, in part order.
    pub connections: Vec<XlsbConnection>,
}

impl XlsbConnections {
    /// Find a connection by its unique identifier (`dwConnID`).
    pub fn by_id(&self, connection_id: u32) -> Option<&XlsbConnection> {
        self.connections
            .iter()
            .find(|connection| connection.connection_id == connection_id)
    }

    /// Find a connection by its workbook-unique name.
    pub fn by_name(&self, name: &str) -> Option<&XlsbConnection> {
        self.connections
            .iter()
            .find(|connection| connection.name == name)
    }
}
