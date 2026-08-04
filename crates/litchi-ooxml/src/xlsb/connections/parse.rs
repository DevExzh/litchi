//! Record-walking parser for the XLSB External Data Connections part
//! (MS-XLSB 2.1.7.24).
//!
//! The parser is strict about record payloads it fully understands and
//! tolerant about everything else: unknown record types are ignored, and
//! known begin/end record pairs that carry no modelled data (text-import
//! wizards, Excel 2010/15 extensions, model data source extensions) are
//! skipped as balanced collections.

use crate::xlsb::connections::model::*;
use crate::xlsb::error::{Error, Result};
use crate::xlsb::walker::{RecordWalker, malformed};
use litchi_xlsb::raw::{Cursor, kind as rt};

/// Maximum number of connections in one part.
const MAX_CONNECTIONS: usize = super::write::MAX_CONNECTIONS;
/// Maximum number of parameters on one connection.
const MAX_PARAMETERS: usize = super::write::MAX_PARAMETERS;
/// Maximum number of Web query table items on one connection.
const MAX_WEB_TABLE_ITEMS: usize = super::write::MAX_WEB_TABLE_ITEMS;

// `BrtBeginExtConnection` flags word 1 (MS-XLSB 2.4.80).
const CONN_MAINTAIN: u16 = 1 << 0;
const CONN_NEW_QUERY: u16 = 1 << 1;
const CONN_DELETED: u16 = 1 << 2;
const CONN_ALWAYS_USE_CONNECTION_FILE: u16 = 1 << 3;
const CONN_BACKGROUND_QUERY: u16 = 1 << 4;
const CONN_REFRESH_ON_LOAD: u16 = 1 << 5;
const CONN_SAVE_DATA: u16 = 1 << 6;

// `BrtBeginExtConnection` flags word 2: optional-field presence.
const CONN_LOAD_DATA_FILE: u16 = 1 << 0;
const CONN_LOAD_CONNECTION_FILE: u16 = 1 << 1;
const CONN_LOAD_DESCRIPTION: u16 = 1 << 2;
const CONN_LOAD_SSO: u16 = 1 << 4;

// `BrtBeginECDbProps` flags byte (MS-XLSB 2.4.61).
const DB_LOAD_SERVER_COMMAND: u8 = 1 << 0;
const DB_LOAD_COMMAND: u8 = 1 << 1;

// `BrtBeginECOlapProps` flags bytes (MS-XLSB 2.4.62).
const OLAP_LOCAL_CONNECTION: u8 = 1 << 0;
const OLAP_NO_REFRESH_CUBE: u8 = 1 << 1;
const OLAP_SERVER_FORMAT_BACK: u8 = 1 << 2;
const OLAP_SERVER_FORMAT_FORE: u8 = 1 << 3;
const OLAP_SERVER_FORMAT_FLAGS: u8 = 1 << 4;
const OLAP_SERVER_FORMAT_NUMBER: u8 = 1 << 5;
const OLAP_USE_OFFICE_LCID: u8 = 1 << 6;
const OLAP_LOAD_LOCAL_CONNECTION: u8 = 1 << 0;

// `BrtBeginECWebProps` flags bytes (MS-XLSB 2.4.71).
const WEB_SOURCE_IS_XML: u8 = 1 << 0;
const WEB_IMPORT_SOURCE_DATA: u8 = 1 << 1;
const WEB_PARSE_PRE_FORMATTED: u8 = 1 << 2;
const WEB_CONSECUTIVE_DELIMITERS: u8 = 1 << 3;
const WEB_SAME_SETTINGS: u8 = 1 << 4;
const WEB_EXCEL97_FORMAT: u8 = 1 << 5;
const WEB_NO_DATE_RECOGNITION: u8 = 1 << 6;
const WEB_REFRESHED_IN_EXCEL9: u8 = 1 << 7;
const WEB_TABLES_ONLY_HTML: u16 = 1 << 0;
const WEB_LOAD_POST: u16 = 1 << 1;
const WEB_LOAD_EDIT_PAGE: u16 = 1 << 2;
const WEB_LOAD_URL: u16 = 1 << 3;

// `BrtBeginECParam` flags word (MS-XLSB 2.4.63).
const PARAM_TYPE_MASK: u16 = 0x7;
const PARAM_AUTO_REFRESH: u16 = 1 << 3;

// Parameter `pbt` values (MS-XLSB 2.4.63).
const PBT_PROMPT: u16 = 0x0;
const PBT_VALUE: u16 = 0x1;
const PBT_CELL_REFERENCE: u16 = 0x2;

// Parameter `dataType` values (MS-XLSB 2.4.63).
const PARAM_DATA_DOUBLE: u32 = 1;
const PARAM_DATA_STRING: u32 = 2;
const PARAM_DATA_BOOLEAN: u32 = 4;

/// Parse the complete External Data Connections part.
pub fn parse_connections_part(data: &[u8]) -> Result<Connections> {
    const CONTEXT: &str = "ExternalDataConnections";
    let mut walker = RecordWalker::new(data);
    let begin = walker.required_begin(rt::BEGIN_EXT_CONNECTIONS, CONTEXT)?;
    Cursor::new(begin.payload(), "BrtBeginExtConnections").finish()?;

    let mut connections = Connections::default();
    loop {
        let Some(record) = walker.next()? else {
            return Err(Error::UnexpectedEndOfStream(CONTEXT.to_string()));
        };
        match record.kind() {
            rt::BEGIN_EXT_CONNECTION => {
                if connections.connections.len() >= MAX_CONNECTIONS {
                    return Err(malformed(CONTEXT, "connection count limit exceeded"));
                }
                connections
                    .connections
                    .push(parse_connection(&mut walker, record.payload())?);
            },
            rt::END_EXT_CONNECTIONS => {
                Cursor::new(record.payload(), "BrtEndExtConnections").finish()?;
                super::write::validate_connections(&connections)?;
                return Ok(connections);
            },
            other => walker.skip_unhandled(other, CONTEXT)?,
        }
    }
}

/// Parse one `BrtBeginExtConnection` collection through its end record.
fn parse_connection(walker: &mut RecordWalker<'_>, data: &[u8]) -> Result<Connection> {
    const CONTEXT: &str = "BrtBeginExtConnection";
    let mut connection = parse_ext_connection(data)?;
    loop {
        let record = walker.required(CONTEXT)?;
        match record.kind() {
            rt::BEGIN_EC_DB_PROPS => {
                connection.properties = Properties::Database(parse_db_props(record.payload())?);
                walker.expect_end(rt::END_EC_DB_PROPS, "BrtBeginECDbProps")?;
            },
            rt::BEGIN_EC_OLAP_PROPS => {
                connection.properties = Properties::Olap(parse_olap_props(record.payload())?);
                walker.expect_end(rt::END_EC_OLAP_PROPS, "BrtBeginECOlapProps")?;
            },
            rt::BEGIN_EC_WEB_PROPS => {
                connection.properties = Properties::Web(parse_web_props(record.payload())?);
                walker.expect_end(rt::END_EC_WEB_PROPS, "BrtBeginECWebProps")?;
            },
            rt::BEGIN_EC_PARAMS => {
                parse_params(walker, &mut connection)?;
            },
            rt::BEGIN_EC_WP_TABLES => {
                parse_web_tables(walker, &mut connection)?;
            },
            rt::END_EXT_CONNECTION => return Ok(connection),
            other => walker.skip_unhandled(other, CONTEXT)?,
        }
    }
}

/// `BrtBeginExtConnection` payload (MS-XLSB 2.4.80).
fn parse_ext_connection(data: &[u8]) -> Result<Connection> {
    const CONTEXT: &str = "BrtBeginExtConnection";
    let mut cursor = Cursor::new(data, CONTEXT);
    let refreshed_version = cursor.read_u8()?;
    let refreshable_min_version = cursor.read_u8()?;
    let pc = cursor.read_u8()?;
    cursor.skip(1)?; // reserved1
    let refresh_interval_minutes = cursor.read_u16()?;
    let flags1 = cursor.read_u16()?;
    let flags2 = cursor.read_u16()?;
    let source_type = SourceType::try_from(cursor.read_u32()?)?;
    let reconnection_type = match cursor.read_u32()? {
        1 => Some(ReconnectionType::AsRequired),
        2 => Some(ReconnectionType::Always),
        3 => Some(ReconnectionType::Never),
        _ => None, // ignored when fAlwaysUseConnectionFile is set
    };
    let connection_id = cursor.read_u32()?;
    // `iCredMethod` and `pc` MUST be ignored when the source is not OLE DB or
    // ODBC (MS-XLSB 2.4.80), so only surface them for database sources.
    let database_source = matches!(source_type, SourceType::OleDb | SourceType::Odbc);
    let credential_method = match cursor.read_u8()? {
        0 if database_source => Some(CredentialMethod::Integrated),
        1 if database_source => Some(CredentialMethod::None),
        2 if database_source => Some(CredentialMethod::SingleSignOn),
        _ => None,
    };
    let password_state = match pc {
        1 if database_source => Some(PasswordState::Saved),
        2 if database_source => Some(PasswordState::NotSaved),
        _ => None,
    };
    let data_file = if flags2 & CONN_LOAD_DATA_FILE != 0 {
        Some(cursor.read_wide_string()?)
    } else {
        None
    };
    let connection_file = if flags2 & CONN_LOAD_CONNECTION_FILE != 0 {
        Some(cursor.read_wide_string()?)
    } else {
        None
    };
    let description = if flags2 & CONN_LOAD_DESCRIPTION != 0 {
        Some(cursor.read_wide_string()?)
    } else {
        None
    };
    let name = cursor.read_wide_string()?;
    let sso_id = if flags2 & CONN_LOAD_SSO != 0 {
        Some(cursor.read_wide_string()?)
    } else {
        None
    };
    cursor.finish()?;
    Ok(Connection {
        connection_id,
        source_type,
        name,
        description,
        connection_file,
        data_file,
        sso_id,
        credential_method,
        password_state,
        refreshed_version,
        refreshable_min_version,
        refresh_interval_minutes,
        maintain: flags1 & CONN_MAINTAIN != 0,
        new_query: flags1 & CONN_NEW_QUERY != 0,
        deleted: flags1 & CONN_DELETED != 0,
        always_use_connection_file: flags1 & CONN_ALWAYS_USE_CONNECTION_FILE != 0,
        background_query: flags1 & CONN_BACKGROUND_QUERY != 0,
        refresh_on_load: flags1 & CONN_REFRESH_ON_LOAD != 0,
        save_data: flags1 & CONN_SAVE_DATA != 0,
        reconnection_type,
        properties: Properties::None,
        parameters: Vec::new(),
        web_tables: Vec::new(),
    })
}

/// `BrtBeginECDbProps` payload (MS-XLSB 2.4.61).
fn parse_db_props(data: &[u8]) -> Result<DbProperties> {
    let mut cursor = Cursor::new(data, "BrtBeginECDbProps");
    let command_type = CommandType::try_from(cursor.read_u32()?)?;
    let flags = cursor.read_u8()?;
    let connection_string = cursor.read_wide_string()?;
    let command = if flags & DB_LOAD_COMMAND != 0 {
        Some(cursor.read_wide_string()?)
    } else {
        None
    };
    let server_command = if flags & DB_LOAD_SERVER_COMMAND != 0 {
        Some(cursor.read_wide_string()?)
    } else {
        None
    };
    cursor.finish()?;
    Ok(DbProperties {
        command_type,
        connection_string,
        command,
        server_command,
    })
}

/// `BrtBeginECOlapProps` payload (MS-XLSB 2.4.62).
fn parse_olap_props(data: &[u8]) -> Result<OlapProperties> {
    let mut cursor = Cursor::new(data, "BrtBeginECOlapProps");
    let flags = cursor.read_u8()?;
    let drillthrough_rows = cursor.read_u32()?;
    let load_flags = cursor.read_u8()?;
    let local_connection_string = if load_flags & OLAP_LOAD_LOCAL_CONNECTION != 0 {
        Some(cursor.read_wide_string()?)
    } else {
        None
    };
    cursor.finish()?;
    Ok(OlapProperties {
        local_connection: flags & OLAP_LOCAL_CONNECTION != 0,
        no_refresh_cube: flags & OLAP_NO_REFRESH_CUBE != 0,
        server_format_back: flags & OLAP_SERVER_FORMAT_BACK != 0,
        server_format_fore: flags & OLAP_SERVER_FORMAT_FORE != 0,
        server_format_flags: flags & OLAP_SERVER_FORMAT_FLAGS != 0,
        server_format_number: flags & OLAP_SERVER_FORMAT_NUMBER != 0,
        use_office_lcid: flags & OLAP_USE_OFFICE_LCID != 0,
        drillthrough_rows,
        local_connection_string,
    })
}

/// `BrtBeginECWebProps` payload (MS-XLSB 2.4.71).
fn parse_web_props(data: &[u8]) -> Result<WebProperties> {
    let mut cursor = Cursor::new(data, "BrtBeginECWebProps");
    let html_format = match cursor.read_u8()? {
        0 => HtmlFormat::None,
        1 => HtmlFormat::RichText,
        2 => HtmlFormat::All,
        other => HtmlFormat::Other(other),
    };
    let flags = cursor.read_u8()?;
    let load_flags = cursor.read_u16()?;
    let url = if load_flags & WEB_LOAD_URL != 0 {
        Some(cursor.read_wide_string()?)
    } else {
        None
    };
    let web_post = if load_flags & WEB_LOAD_POST != 0 {
        Some(cursor.read_wide_string()?)
    } else {
        None
    };
    let edit_web_page = if load_flags & WEB_LOAD_EDIT_PAGE != 0 {
        Some(cursor.read_wide_string()?)
    } else {
        None
    };
    cursor.finish()?;
    Ok(WebProperties {
        html_format,
        source_is_xml: flags & WEB_SOURCE_IS_XML != 0,
        import_source_data: flags & WEB_IMPORT_SOURCE_DATA != 0,
        parse_pre_formatted: flags & WEB_PARSE_PRE_FORMATTED != 0,
        consecutive_delimiters: flags & WEB_CONSECUTIVE_DELIMITERS != 0,
        same_settings: flags & WEB_SAME_SETTINGS != 0,
        excel97_format: flags & WEB_EXCEL97_FORMAT != 0,
        no_date_recognition: flags & WEB_NO_DATE_RECOGNITION != 0,
        refreshed_in_excel9: flags & WEB_REFRESHED_IN_EXCEL9 != 0,
        tables_only_html: load_flags & WEB_TABLES_ONLY_HTML != 0,
        url,
        web_post,
        edit_web_page,
    })
}

/// Parse a `BrtBeginECParams` collection into the connection.
fn parse_params(walker: &mut RecordWalker<'_>, connection: &mut Connection) -> Result<()> {
    const CONTEXT: &str = "BrtBeginECParams";
    loop {
        let record = walker.required(CONTEXT)?;
        match record.kind() {
            rt::BEGIN_EC_PARAM => {
                if connection.parameters.len() >= MAX_PARAMETERS {
                    return Err(malformed(CONTEXT, "parameter count limit exceeded"));
                }
                connection.parameters.push(parse_param(record.payload())?);
                walker.expect_end(rt::END_EC_PARAM, "BrtBeginECParam")?;
            },
            rt::END_EC_PARAMS => return Ok(()),
            other => walker.skip_unhandled(other, CONTEXT)?,
        }
    }
}

/// `BrtBeginECParam` payload (MS-XLSB 2.4.63).
fn parse_param(data: &[u8]) -> Result<Parameter> {
    const CONTEXT: &str = "BrtBeginECParam";
    let mut cursor = Cursor::new(data, CONTEXT);
    let flags = cursor.read_u16()?;
    let pbt = flags & PARAM_TYPE_MASK;
    let sql_type = cursor.read_u16()?;
    let data_type = if pbt != PBT_PROMPT {
        Some(cursor.read_u32()?)
    } else {
        None
    };
    let name = cursor.read_wide_string()?;
    let prompt = cursor.read_wide_string()?;
    let prompt = if prompt.is_empty() {
        None
    } else {
        Some(prompt)
    };
    let value = match pbt {
        PBT_VALUE => match data_type {
            Some(PARAM_DATA_DOUBLE) => Some(ParameterValue::Number(cursor.read_f64()?)),
            Some(PARAM_DATA_STRING) => Some(ParameterValue::Text(cursor.read_wide_string()?)),
            Some(PARAM_DATA_BOOLEAN) => {
                // `boolVal`: a 32-bit MS-XLSB Boolean when it fills the
                // record, a single byte for short payloads.
                let boolean = if cursor.remaining() == 1 {
                    u32::from(cursor.read_u8()?)
                } else {
                    u32::from(cursor.read_bool32()?)
                };
                Some(ParameterValue::Boolean(boolean != 0))
            },
            Some(other) => {
                return Err(malformed(
                    CONTEXT,
                    format!("unknown parameter dataType {other:#x}"),
                ));
            },
            None => return Err(malformed(CONTEXT, "value parameter lacks a dataType")),
        },
        PBT_CELL_REFERENCE => {
            // `fmla` tokens fill the remainder of the record (inert).
            let tokens = cursor.read_bytes(cursor.remaining())?.to_vec();
            Some(ParameterValue::CellFormula(tokens))
        },
        _ => None,
    };
    cursor.finish()?;
    let parameter_type = match pbt {
        PBT_PROMPT => ParameterType::Prompt,
        PBT_VALUE => ParameterType::Value,
        PBT_CELL_REFERENCE => ParameterType::CellReference,
        other => ParameterType::Other(other as u8),
    };
    Ok(Parameter {
        parameter_type,
        auto_refresh: flags & PARAM_AUTO_REFRESH != 0,
        sql_type,
        name,
        prompt,
        value,
    })
}

/// Parse a `BrtBeginEcWpTables` collection into the connection.
fn parse_web_tables(walker: &mut RecordWalker<'_>, connection: &mut Connection) -> Result<()> {
    const CONTEXT: &str = "BrtBeginEcWpTables";
    loop {
        let record = walker.required(CONTEXT)?;
        match record.kind() {
            rt::PCDI_MISSING => {
                if connection.web_tables.len() >= MAX_WEB_TABLE_ITEMS {
                    return Err(malformed(CONTEXT, "Web table item limit exceeded"));
                }
                connection.web_tables.push(WebTableItem::Missing);
            },
            rt::PCDI_STRING => {
                if connection.web_tables.len() >= MAX_WEB_TABLE_ITEMS {
                    return Err(malformed(CONTEXT, "Web table item limit exceeded"));
                }
                let mut cursor = Cursor::new(record.payload(), "BrtPCDIString");
                connection
                    .web_tables
                    .push(WebTableItem::Named(cursor.read_wide_string()?));
                cursor.finish()?;
            },
            rt::PCDI_INDEX => {
                if connection.web_tables.len() >= MAX_WEB_TABLE_ITEMS {
                    return Err(malformed(CONTEXT, "Web table item limit exceeded"));
                }
                let mut cursor = Cursor::new(record.payload(), "BrtPCDIIndex");
                connection
                    .web_tables
                    .push(WebTableItem::Index(cursor.read_u32()?));
                cursor.finish()?;
            },
            rt::END_EC_WP_TABLES => return Ok(()),
            other => walker.skip_unhandled(other, CONTEXT)?,
        }
    }
}
