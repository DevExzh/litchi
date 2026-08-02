//! Serializer for the XLSB External Data Connections part (MS-XLSB 2.1.7.24).
//!
//! This is the exact inverse of `parse.rs`: payload layouts, flag bits,
//! optional-field presence flags, and collection structure all mirror the
//! reader so authored connections round-trip through
//! `parse_connections_part` and `XlsbWorkbook::connections`.

use crate::xlsb::connections::model::*;
use crate::xlsb::error::{XlsbError, XlsbResult};
use litchi_xlsb::raw::Writer;
use litchi_xlsb::raw::kind as rt;
use std::collections::HashSet;

pub(crate) const MAX_CONNECTIONS: usize = 4_096;
pub(crate) const MAX_PARAMETERS: usize = 1_024;
pub(crate) const MAX_WEB_TABLE_ITEMS: usize = 1_024;
const MAX_SHORT_STRING_UTF16_UNITS: usize = 255;
const MAX_STRING_UTF16_UNITS: usize = 1_048_576;
const MAX_FORMULA_TOKEN_BYTES: usize = 16 * 1024 * 1024;
const MAX_CONNECTIONS_PART_BYTES: usize = 32 * 1024 * 1024;

/// `BrtPCDIMissing` / `BrtPCDIString` / `BrtPCDIIndex` record types.
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
const CONN_RESERVED3: u16 = 1 << 3;
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
const PARAM_AUTO_REFRESH: u16 = 1 << 3;
const PBT_PROMPT: u16 = 0x0;
const PBT_VALUE: u16 = 0x1;
const PBT_CELL_REFERENCE: u16 = 0x2;
const PARAM_DATA_DOUBLE: u32 = 1;
const PARAM_DATA_STRING: u32 = 2;
const PARAM_DATA_BOOLEAN: u32 = 4;

fn malformed(context: &str, detail: impl Into<String>) -> XlsbError {
    XlsbError::Unrecognized {
        typ: context.to_string(),
        val: detail.into(),
    }
}

fn validate_string(value: &str, context: &str, max_utf16_units: usize) -> XlsbResult<()> {
    if value.encode_utf16().count() > max_utf16_units {
        return Err(malformed(
            context,
            format!("string exceeds {max_utf16_units} UTF-16 code units"),
        ));
    }
    Ok(())
}

/// Validate model invariants shared by parsed-package and new-workbook authors.
pub(crate) fn validate_connections(connections: &XlsbConnections) -> XlsbResult<()> {
    if connections.connections.len() > MAX_CONNECTIONS {
        return Err(malformed(
            "ExternalDataConnections",
            "connection count limit exceeded",
        ));
    }
    let mut ids = HashSet::with_capacity(connections.connections.len());
    let mut names = HashSet::with_capacity(connections.connections.len());
    for connection in &connections.connections {
        if connection.connection_id == 0 {
            return Err(malformed(
                "BrtBeginExtConnection",
                "connection id must be greater than zero",
            ));
        }
        if !ids.insert(connection.connection_id) {
            return Err(malformed(
                "BrtBeginExtConnection",
                format!("duplicate connection id {}", connection.connection_id),
            ));
        }
        let name_len = connection.name.encode_utf16().count();
        if name_len == 0 || name_len > MAX_SHORT_STRING_UTF16_UNITS {
            return Err(malformed(
                "BrtBeginExtConnection",
                "connection name must contain 1 to 255 UTF-16 code units",
            ));
        }
        if !names.insert(connection.name.to_lowercase()) {
            return Err(malformed(
                "BrtBeginExtConnection",
                format!("duplicate connection name '{}'", connection.name),
            ));
        }
        for (context, value) in [
            ("data file", connection.data_file.as_deref()),
            ("connection file", connection.connection_file.as_deref()),
            ("description", connection.description.as_deref()),
            ("SSO identifier", connection.sso_id.as_deref()),
        ] {
            if let Some(value) = value {
                validate_string(value, context, MAX_SHORT_STRING_UTF16_UNITS)?;
            }
        }
        if connection.refresh_interval_minutes >= 32_768 {
            return Err(malformed(
                "BrtBeginExtConnection",
                "refresh interval must be less than 32768 minutes",
            ));
        }
        match &connection.properties {
            XlsbConnectionProperties::Database(properties) => {
                validate_string(
                    &properties.connection_string,
                    "database connection string",
                    MAX_STRING_UTF16_UNITS,
                )?;
                for (context, value) in [
                    ("database command", properties.command.as_deref()),
                    (
                        "database server command",
                        properties.server_command.as_deref(),
                    ),
                ] {
                    if let Some(value) = value {
                        validate_string(value, context, MAX_STRING_UTF16_UNITS)?;
                    }
                }
            },
            XlsbConnectionProperties::Olap(properties) => {
                if let Some(value) = &properties.local_connection_string {
                    validate_string(
                        value,
                        "OLAP local connection string",
                        MAX_STRING_UTF16_UNITS,
                    )?;
                }
            },
            XlsbConnectionProperties::Web(properties) => {
                for (context, value) in [
                    ("Web connection URL", properties.url.as_deref()),
                    ("Web POST body", properties.web_post.as_deref()),
                    ("Web edit-page URL", properties.edit_web_page.as_deref()),
                ] {
                    if let Some(value) = value {
                        validate_string(value, context, MAX_STRING_UTF16_UNITS)?;
                    }
                }
            },
            XlsbConnectionProperties::None => {},
        }
        if connection.parameters.len() > MAX_PARAMETERS {
            return Err(malformed(
                "BrtBeginECParams",
                "parameter count limit exceeded",
            ));
        }
        for parameter in &connection.parameters {
            validate_string(
                &parameter.name,
                "connection parameter name",
                MAX_STRING_UTF16_UNITS,
            )?;
            if let Some(prompt) = &parameter.prompt {
                validate_string(
                    prompt,
                    "connection parameter prompt",
                    MAX_STRING_UTF16_UNITS,
                )?;
            }
            match &parameter.value {
                Some(XlsbParameterValue::Text(value)) => {
                    validate_string(value, "connection parameter value", MAX_STRING_UTF16_UNITS)?
                },
                Some(XlsbParameterValue::CellFormula(tokens))
                    if tokens.len() > MAX_FORMULA_TOKEN_BYTES =>
                {
                    return Err(malformed(
                        "BrtBeginECParam",
                        "cell formula token limit exceeded",
                    ));
                },
                _ => {},
            }
        }
        if connection.web_tables.len() > MAX_WEB_TABLE_ITEMS {
            return Err(malformed(
                "BrtBeginEcWpTables",
                "Web table item count limit exceeded",
            ));
        }
        for item in &connection.web_tables {
            if let XlsbWebTableItem::Named(name) = item {
                validate_string(name, "Web table name", MAX_STRING_UTF16_UNITS)?;
            }
        }
    }
    Ok(())
}

fn write_wide_string(data: &mut Vec<u8>, value: &str) {
    data.extend_from_slice(&(value.encode_utf16().count() as u32).to_le_bytes());
    for unit in value.encode_utf16() {
        data.extend_from_slice(&unit.to_le_bytes());
    }
}

/// `BrtBeginExtConnection` payload (MS-XLSB 2.4.80).
fn ext_connection_payload(connection: &XlsbConnection) -> Vec<u8> {
    let mut data = Vec::with_capacity(64);
    data.push(connection.refreshed_version);
    data.push(connection.refreshable_min_version);
    data.push(match connection.password_state {
        Some(XlsbPasswordState::Saved) => 1,
        Some(XlsbPasswordState::NotSaved) => 2,
        None => 0,
    });
    data.push(0); // reserved1
    data.extend_from_slice(&connection.refresh_interval_minutes.to_le_bytes());
    let mut flags1 = 0u16;
    if connection.maintain {
        flags1 |= CONN_MAINTAIN;
    }
    if connection.new_query {
        flags1 |= CONN_NEW_QUERY;
    }
    if connection.deleted {
        flags1 |= CONN_DELETED;
    }
    if connection.always_use_connection_file {
        flags1 |= CONN_ALWAYS_USE_CONNECTION_FILE;
    }
    if connection.background_query {
        flags1 |= CONN_BACKGROUND_QUERY;
    }
    if connection.refresh_on_load {
        flags1 |= CONN_REFRESH_ON_LOAD;
    }
    if connection.save_data {
        flags1 |= CONN_SAVE_DATA;
    }
    data.extend_from_slice(&flags1.to_le_bytes());
    let mut flags2 = CONN_RESERVED3;
    if connection.data_file.is_some() {
        flags2 |= CONN_LOAD_DATA_FILE;
    }
    if connection.connection_file.is_some() {
        flags2 |= CONN_LOAD_CONNECTION_FILE;
    }
    if connection.description.is_some() {
        flags2 |= CONN_LOAD_DESCRIPTION;
    }
    if connection.sso_id.is_some() {
        flags2 |= CONN_LOAD_SSO;
    }
    data.extend_from_slice(&flags2.to_le_bytes());
    data.extend_from_slice(&(connection.source_type as u32).to_le_bytes());
    data.extend_from_slice(
        &match connection.reconnection_type {
            Some(XlsbReconnectionType::AsRequired) => 1u32,
            Some(XlsbReconnectionType::Always) => 2,
            Some(XlsbReconnectionType::Never) => 3,
            None => 0,
        }
        .to_le_bytes(),
    );
    data.extend_from_slice(&connection.connection_id.to_le_bytes());
    data.push(match connection.credential_method {
        Some(XlsbCredentialMethod::Integrated) | None => 0,
        Some(XlsbCredentialMethod::None) => 1,
        Some(XlsbCredentialMethod::SingleSignOn) => 2,
    });
    if let Some(data_file) = &connection.data_file {
        write_wide_string(&mut data, data_file);
    }
    if let Some(connection_file) = &connection.connection_file {
        write_wide_string(&mut data, connection_file);
    }
    if let Some(description) = &connection.description {
        write_wide_string(&mut data, description);
    }
    write_wide_string(&mut data, &connection.name);
    if let Some(sso_id) = &connection.sso_id {
        write_wide_string(&mut data, sso_id);
    }
    data
}

/// `BrtBeginECDbProps` payload (MS-XLSB 2.4.61).
fn db_props_payload(props: &XlsbDbProperties) -> Vec<u8> {
    let mut data = Vec::with_capacity(32);
    data.extend_from_slice(&(props.command_type as u32).to_le_bytes());
    let mut flags = 0u8;
    if props.server_command.is_some() {
        flags |= DB_LOAD_SERVER_COMMAND;
    }
    if props.command.is_some() {
        flags |= DB_LOAD_COMMAND;
    }
    data.push(flags);
    write_wide_string(&mut data, &props.connection_string);
    if let Some(command) = &props.command {
        write_wide_string(&mut data, command);
    }
    if let Some(server_command) = &props.server_command {
        write_wide_string(&mut data, server_command);
    }
    data
}

/// `BrtBeginECOlapProps` payload (MS-XLSB 2.4.62).
fn olap_props_payload(props: &XlsbOlapProperties) -> Vec<u8> {
    let mut flags = 0u8;
    if props.local_connection {
        flags |= OLAP_LOCAL_CONNECTION;
    }
    if props.no_refresh_cube {
        flags |= OLAP_NO_REFRESH_CUBE;
    }
    if props.server_format_back {
        flags |= OLAP_SERVER_FORMAT_BACK;
    }
    if props.server_format_fore {
        flags |= OLAP_SERVER_FORMAT_FORE;
    }
    if props.server_format_flags {
        flags |= OLAP_SERVER_FORMAT_FLAGS;
    }
    if props.server_format_number {
        flags |= OLAP_SERVER_FORMAT_NUMBER;
    }
    if props.use_office_lcid {
        flags |= OLAP_USE_OFFICE_LCID;
    }
    let mut data = Vec::with_capacity(16);
    data.push(flags);
    data.extend_from_slice(&props.drillthrough_rows.to_le_bytes());
    data.push(if props.local_connection_string.is_some() {
        OLAP_LOAD_LOCAL_CONNECTION
    } else {
        0
    });
    if let Some(local) = &props.local_connection_string {
        write_wide_string(&mut data, local);
    }
    data
}

/// `BrtBeginECWebProps` payload (MS-XLSB 2.4.71).
fn web_props_payload(props: &XlsbWebProperties) -> Vec<u8> {
    let mut flags = 0u8;
    if props.source_is_xml {
        flags |= WEB_SOURCE_IS_XML;
    }
    if props.import_source_data {
        flags |= WEB_IMPORT_SOURCE_DATA;
    }
    if props.parse_pre_formatted {
        flags |= WEB_PARSE_PRE_FORMATTED;
    }
    if props.consecutive_delimiters {
        flags |= WEB_CONSECUTIVE_DELIMITERS;
    }
    if props.same_settings {
        flags |= WEB_SAME_SETTINGS;
    }
    if props.excel97_format {
        flags |= WEB_EXCEL97_FORMAT;
    }
    if props.no_date_recognition {
        flags |= WEB_NO_DATE_RECOGNITION;
    }
    if props.refreshed_in_excel9 {
        flags |= WEB_REFRESHED_IN_EXCEL9;
    }
    let mut load_flags = 0u16;
    if props.tables_only_html {
        load_flags |= WEB_TABLES_ONLY_HTML;
    }
    if props.web_post.is_some() {
        load_flags |= WEB_LOAD_POST;
    }
    if props.edit_web_page.is_some() {
        load_flags |= WEB_LOAD_EDIT_PAGE;
    }
    if props.url.is_some() {
        load_flags |= WEB_LOAD_URL;
    }
    let mut data = Vec::with_capacity(16);
    data.push(match props.html_format {
        XlsbHtmlFormat::None => 0,
        XlsbHtmlFormat::RichText => 1,
        XlsbHtmlFormat::All => 2,
        XlsbHtmlFormat::Other(value) => value,
    });
    data.push(flags);
    data.extend_from_slice(&load_flags.to_le_bytes());
    if let Some(url) = &props.url {
        write_wide_string(&mut data, url);
    }
    if let Some(web_post) = &props.web_post {
        write_wide_string(&mut data, web_post);
    }
    if let Some(edit_web_page) = &props.edit_web_page {
        write_wide_string(&mut data, edit_web_page);
    }
    data
}

/// `BrtBeginECParam` payload (MS-XLSB 2.4.63).
fn param_payload(parameter: &XlsbConnectionParameter) -> XlsbResult<Vec<u8>> {
    const CONTEXT: &str = "BrtBeginECParam";
    let (pbt, data_type) = match parameter.parameter_type {
        XlsbParameterType::Prompt => (PBT_PROMPT, None),
        XlsbParameterType::Value => {
            let data_type = match &parameter.value {
                Some(XlsbParameterValue::Number(_)) => PARAM_DATA_DOUBLE,
                Some(XlsbParameterValue::Text(_)) => PARAM_DATA_STRING,
                Some(XlsbParameterValue::Boolean(_)) => PARAM_DATA_BOOLEAN,
                Some(XlsbParameterValue::CellFormula(_)) => {
                    return Err(malformed(
                        CONTEXT,
                        "cell formula value requires the cell-reference parameter type",
                    ));
                },
                None => {
                    return Err(malformed(CONTEXT, "value parameter lacks a value"));
                },
            };
            (PBT_VALUE, Some(data_type))
        },
        XlsbParameterType::CellReference => (PBT_CELL_REFERENCE, Some(0)),
        XlsbParameterType::Other(value) => (u16::from(value), None),
    };
    let mut flags = pbt;
    if parameter.auto_refresh {
        flags |= PARAM_AUTO_REFRESH;
    }
    let mut data = Vec::with_capacity(32);
    data.extend_from_slice(&flags.to_le_bytes());
    data.extend_from_slice(&parameter.sql_type.to_le_bytes());
    if pbt != PBT_PROMPT {
        data.extend_from_slice(&data_type.unwrap_or(0).to_le_bytes());
    }
    write_wide_string(&mut data, &parameter.name);
    write_wide_string(&mut data, parameter.prompt.as_deref().unwrap_or(""));
    match &parameter.value {
        Some(XlsbParameterValue::Number(value)) => data.extend_from_slice(&value.to_le_bytes()),
        Some(XlsbParameterValue::Text(value)) => write_wide_string(&mut data, value),
        Some(XlsbParameterValue::Boolean(value)) => {
            data.extend_from_slice(&u32::from(*value).to_le_bytes());
        },
        Some(XlsbParameterValue::CellFormula(tokens)) => data.extend_from_slice(tokens),
        None => {},
    }
    Ok(data)
}

/// Serialize the complete External Data Connections part.
pub(crate) fn write_connections_part(connections: &XlsbConnections) -> XlsbResult<Vec<u8>> {
    validate_connections(connections)?;
    let mut data = Vec::with_capacity(512);
    let mut writer = Writer::new(&mut data);
    writer.write_record(rt::BEGIN_EXT_CONNECTIONS, &[])?;
    for connection in &connections.connections {
        if connection.name.is_empty() {
            return Err(malformed(
                "BrtBeginExtConnection",
                "connection requires a name",
            ));
        }
        writer.write_record(
            rt::BEGIN_EXT_CONNECTION,
            &ext_connection_payload(connection),
        )?;
        match &connection.properties {
            XlsbConnectionProperties::Database(props) => {
                writer.write_record(rt::BEGIN_EC_DB_PROPS, &db_props_payload(props))?;
                writer.write_record(rt::END_EC_DB_PROPS, &[])?;
            },
            XlsbConnectionProperties::Olap(props) => {
                writer.write_record(rt::BEGIN_EC_OLAP_PROPS, &olap_props_payload(props))?;
                writer.write_record(rt::END_EC_OLAP_PROPS, &[])?;
            },
            XlsbConnectionProperties::Web(props) => {
                writer.write_record(rt::BEGIN_EC_WEB_PROPS, &web_props_payload(props))?;
                writer.write_record(rt::END_EC_WEB_PROPS, &[])?;
            },
            XlsbConnectionProperties::None => {},
        }
        if !connection.parameters.is_empty() {
            writer.write_record(rt::BEGIN_EC_PARAMS, &[])?;
            for parameter in &connection.parameters {
                writer.write_record(rt::BEGIN_EC_PARAM, &param_payload(parameter)?)?;
                writer.write_record(rt::END_EC_PARAM, &[])?;
            }
            writer.write_record(rt::END_EC_PARAMS, &[])?;
        }
        if !connection.web_tables.is_empty() {
            writer.write_record(rt::BEGIN_EC_WP_TABLES, &[])?;
            for item in &connection.web_tables {
                match item {
                    XlsbWebTableItem::Missing => writer.write_record(rt::PCDI_MISSING, &[])?,
                    XlsbWebTableItem::Named(name) => {
                        let mut payload = Vec::new();
                        write_wide_string(&mut payload, name);
                        writer.write_record(rt::PCDI_STRING, &payload)?;
                    },
                    XlsbWebTableItem::Index(index) => {
                        writer.write_record(rt::PCDI_INDEX, &index.to_le_bytes())?;
                    },
                }
            }
            writer.write_record(rt::END_EC_WP_TABLES, &[])?;
        }
        writer.write_record(rt::END_EXT_CONNECTION, &[])?;
    }
    writer.write_record(rt::END_EXT_CONNECTIONS, &[])?;
    if data.len() > MAX_CONNECTIONS_PART_BYTES {
        return Err(malformed(
            "ExternalDataConnections",
            "serialized part size limit exceeded",
        ));
    }
    Ok(data)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsb::connections::parse_connections_part;

    fn sample_connections() -> XlsbConnections {
        XlsbConnections {
            connections: vec![
                XlsbConnection {
                    connection_id: 42,
                    source_type: XlsbConnectionSourceType::Odbc,
                    name: "Warehouse".to_string(),
                    description: Some("Main warehouse".to_string()),
                    refreshed_version: 7,
                    refreshable_min_version: 5,
                    refresh_interval_minutes: 30,
                    maintain: true,
                    background_query: true,
                    save_data: true,
                    reconnection_type: Some(XlsbReconnectionType::Never),
                    credential_method: Some(XlsbCredentialMethod::Integrated),
                    password_state: Some(XlsbPasswordState::NotSaved),
                    properties: XlsbConnectionProperties::Database(XlsbDbProperties {
                        command_type: XlsbCommandType::Sql,
                        connection_string: "Driver={SQL Server};Server=db".to_string(),
                        command: Some("SELECT * FROM T".to_string()),
                        server_command: None,
                    }),
                    parameters: vec![
                        XlsbConnectionParameter {
                            parameter_type: XlsbParameterType::Value,
                            auto_refresh: true,
                            sql_type: 4,
                            name: "threshold".to_string(),
                            prompt: Some("Enter value".to_string()),
                            value: Some(XlsbParameterValue::Number(42.5)),
                        },
                        XlsbConnectionParameter {
                            parameter_type: XlsbParameterType::Prompt,
                            auto_refresh: false,
                            sql_type: 0,
                            name: "ask".to_string(),
                            prompt: Some("City?".to_string()),
                            value: None,
                        },
                    ],
                    web_tables: Vec::new(),
                    ..XlsbConnection::default()
                },
                XlsbConnection {
                    connection_id: 9,
                    source_type: XlsbConnectionSourceType::Web,
                    name: "Web Query".to_string(),
                    properties: XlsbConnectionProperties::Web(XlsbWebProperties {
                        html_format: XlsbHtmlFormat::All,
                        source_is_xml: true,
                        consecutive_delimiters: true,
                        url: Some("https://example.test/q".to_string()),
                        ..XlsbWebProperties::default()
                    }),
                    web_tables: vec![
                        XlsbWebTableItem::Missing,
                        XlsbWebTableItem::Named("results".to_string()),
                        XlsbWebTableItem::Index(3),
                    ],
                    ..XlsbConnection::default()
                },
                XlsbConnection {
                    connection_id: 7,
                    source_type: XlsbConnectionSourceType::OleDb,
                    name: "Cube".to_string(),
                    credential_method: Some(XlsbCredentialMethod::SingleSignOn),
                    sso_id: Some("sso-app".to_string()),
                    properties: XlsbConnectionProperties::Olap(XlsbOlapProperties {
                        local_connection: true,
                        server_format_back: true,
                        server_format_fore: true,
                        drillthrough_rows: 1024,
                        local_connection_string: Some(
                            "Provider=MSOLAP;Data Source=local.cub".to_string(),
                        ),
                        ..XlsbOlapProperties::default()
                    }),
                    ..XlsbConnection::default()
                },
            ],
        }
    }

    #[test]
    fn serialized_connections_round_trip_through_the_reader() {
        let connections = sample_connections();
        let bytes = write_connections_part(&connections).unwrap();
        let parsed = parse_connections_part(&bytes).unwrap();
        assert_eq!(parsed.connections.len(), 3);
        assert_eq!(parsed, connections);
    }

    #[test]
    fn rejects_unnamed_and_mistyped_parameters() {
        let mut connections = XlsbConnections::default();
        connections.connections.push(XlsbConnection::default());
        assert!(write_connections_part(&connections).is_err());

        let mut connections = XlsbConnections::default();
        connections.connections.push(XlsbConnection {
            name: "x".to_string(),
            parameters: vec![XlsbConnectionParameter {
                parameter_type: XlsbParameterType::Value,
                auto_refresh: false,
                sql_type: 0,
                name: "p".to_string(),
                prompt: None,
                value: None,
            }],
            ..XlsbConnection::default()
        });
        assert!(write_connections_part(&connections).is_err());
    }
}
