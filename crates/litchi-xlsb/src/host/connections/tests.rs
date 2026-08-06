//! Synthetic-stream tests for the External Data Connections parser.

use super::model::*;
use super::parse::parse_connections_part;
use crate::package::Workbook;
use crate::raw::{Kind, Writer, kind as rt};
use crate::writer::{MutableWorksheet, WorkbookWriter};
use litchi_opc::PackURI;
use std::io::Cursor;

fn wide(value: &str) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(4 + value.len() * 2);
    bytes.extend_from_slice(&(value.len() as u32).to_le_bytes());
    for unit in value.encode_utf16() {
        bytes.extend_from_slice(&unit.to_le_bytes());
    }
    bytes
}

fn record(record_type: Kind, payload: &[u8]) -> (Kind, Vec<u8>) {
    (record_type, payload.to_vec())
}

fn build(records: &[(Kind, Vec<u8>)]) -> Vec<u8> {
    let mut data = Vec::new();
    let mut writer = Writer::new(&mut data);
    for (record_type, payload) in records {
        writer.write_record(*record_type, payload).unwrap();
    }
    data
}

struct ExtConnectionBuilder {
    source_type: u32,
    connection_id: u32,
    name: String,
    flags1: u16,
    flags2: u16,
    description: Option<String>,
    connection_file: Option<String>,
    sso_id: Option<String>,
    cred_method: u8,
}

impl ExtConnectionBuilder {
    fn new(source_type: u32, connection_id: u32, name: &str) -> Self {
        Self {
            source_type,
            connection_id,
            name: name.to_string(),
            flags1: 0,
            flags2: 1 << 3, // reserved3 MUST be 1
            description: None,
            connection_file: None,
            sso_id: None,
            cred_method: 0,
        }
    }

    fn payload(self) -> Vec<u8> {
        let mut flags2 = self.flags2;
        if self.connection_file.is_some() {
            flags2 |= 1 << 1;
        }
        if self.description.is_some() {
            flags2 |= 1 << 2;
        }
        if self.sso_id.is_some() {
            flags2 |= 1 << 4;
        }
        let mut bytes = vec![7, 5, 2, 0];
        bytes.extend_from_slice(&30u16.to_le_bytes()); // wInterval
        bytes.extend_from_slice(&self.flags1.to_le_bytes());
        bytes.extend_from_slice(&flags2.to_le_bytes());
        bytes.extend_from_slice(&self.source_type.to_le_bytes());
        bytes.extend_from_slice(&3u32.to_le_bytes()); // irecontype
        bytes.extend_from_slice(&self.connection_id.to_le_bytes());
        bytes.push(self.cred_method);
        if let Some(file) = &self.connection_file {
            bytes.extend_from_slice(&wide(file));
        }
        if let Some(description) = &self.description {
            bytes.extend_from_slice(&wide(description));
        }
        bytes.extend_from_slice(&wide(&self.name));
        if let Some(sso) = &self.sso_id {
            bytes.extend_from_slice(&wide(sso));
        }
        bytes
    }
}

fn db_props_payload(command_type: u32, conn: &str, command: Option<&str>) -> Vec<u8> {
    let mut bytes = command_type.to_le_bytes().to_vec();
    bytes.push(if command.is_some() { 1 << 1 } else { 0 });
    bytes.extend_from_slice(&wide(conn));
    if let Some(command) = command {
        bytes.extend_from_slice(&wide(command));
    }
    bytes
}

fn olap_props_payload() -> Vec<u8> {
    let mut bytes = vec![0x2F]; // local|srvFmtBack|srvFmtFore|srvFmtFlags|srvFmtNum... bits 0-3,5
    bytes.extend_from_slice(&1024u32.to_le_bytes());
    bytes.push(1); // fLoadConnLocal
    bytes.extend_from_slice(&wide("Provider=MSOLAP;Data Source=local.cub"));
    bytes
}

fn web_props_payload(html_format: u8, url: Option<&str>) -> Vec<u8> {
    let mut bytes = vec![html_format];
    bytes.push(0x01 | 0x08); // fSrcIsXML | fConsecDelim
    bytes.extend_from_slice(&(if url.is_some() { 1u16 << 3 } else { 0 }).to_le_bytes());
    if let Some(url) = url {
        bytes.extend_from_slice(&wide(url));
    }
    bytes
}

fn param_payload(pbt: u16, data_type: Option<u32>, name: &str, tail: &[u8]) -> Vec<u8> {
    let mut bytes = (pbt | (1 << 3)).to_le_bytes().to_vec(); // fAutoRefresh
    bytes.extend_from_slice(&4u16.to_le_bytes()); // wTypeSql
    if let Some(data_type) = data_type {
        bytes.extend_from_slice(&data_type.to_le_bytes());
    }
    bytes.extend_from_slice(&wide(name));
    bytes.extend_from_slice(&wide("Enter value"));
    bytes.extend_from_slice(tail);
    bytes
}

fn full_part(extra_connection_records: &[(Kind, Vec<u8>)]) -> Vec<u8> {
    let mut records = vec![record(rt::BEGIN_EXT_CONNECTIONS, &[])];
    records.extend_from_slice(extra_connection_records);
    records.push(record(rt::END_EXT_CONNECTIONS, &[]));
    build(&records)
}

fn odbc_connection_records() -> Vec<(Kind, Vec<u8>)> {
    vec![
        record(
            rt::BEGIN_EXT_CONNECTION,
            &ExtConnectionBuilder::new(1, 42, "Warehouse").payload(),
        ),
        record(
            rt::BEGIN_EC_DB_PROPS,
            &db_props_payload(2, "Driver={SQL Server};Server=db", Some("SELECT * FROM T")),
        ),
        record(rt::END_EC_DB_PROPS, &[]),
        record(rt::BEGIN_EC_PARAMS, &[]),
        record(
            rt::BEGIN_EC_PARAM,
            &param_payload(1, Some(1), "threshold", &42.5f64.to_le_bytes()),
        ),
        record(rt::END_EC_PARAM, &[]),
        record(
            rt::BEGIN_EC_PARAM,
            &param_payload(1, Some(2), "city", &wide("Paris")),
        ),
        record(rt::END_EC_PARAM, &[]),
        record(
            rt::BEGIN_EC_PARAM,
            &param_payload(1, Some(4), "flag", &1u32.to_le_bytes()),
        ),
        record(rt::END_EC_PARAM, &[]),
        record(rt::BEGIN_EC_PARAM, &param_payload(0, None, "ask", &[])),
        record(rt::END_EC_PARAM, &[]),
        record(rt::END_EC_PARAMS, &[]),
        record(rt::END_EXT_CONNECTION, &[]),
    ]
}

#[test]
fn parses_odbc_connection_with_db_props_and_parameters() {
    let part = full_part(&odbc_connection_records());
    let connections = parse_connections_part(&part).unwrap();
    assert_eq!(connections.connections.len(), 1);
    let connection = &connections.connections[0];
    assert_eq!(connection.connection_id, 42);
    assert_eq!(connection.source_type, SourceType::Odbc);
    assert_eq!(connection.name, "Warehouse");
    assert_eq!(connection.refresh_interval_minutes, 30);
    assert_eq!(connection.refreshed_version, 7);
    assert_eq!(connection.refreshable_min_version, 5);
    assert_eq!(connection.reconnection_type, Some(ReconnectionType::Never));
    assert_eq!(connection.password_state, Some(PasswordState::NotSaved));
    assert_eq!(
        connection.credential_method,
        Some(CredentialMethod::Integrated)
    );
    match &connection.properties {
        Properties::Database(db) => {
            assert_eq!(db.command_type, CommandType::Sql);
            assert_eq!(db.connection_string, "Driver={SQL Server};Server=db");
            assert_eq!(db.command.as_deref(), Some("SELECT * FROM T"));
            assert_eq!(db.server_command, None);
        },
        other => panic!("expected database properties, got {other:?}"),
    }
    assert_eq!(connection.parameters.len(), 4);
    assert_eq!(
        connection.parameters[0].value,
        Some(ParameterValue::Number(42.5))
    );
    assert_eq!(
        connection.parameters[1].value,
        Some(ParameterValue::Text("Paris".to_string()))
    );
    assert_eq!(
        connection.parameters[2].value,
        Some(ParameterValue::Boolean(true))
    );
    assert_eq!(
        connection.parameters[3].parameter_type,
        ParameterType::Prompt
    );
    assert!(connection.parameters[0].auto_refresh);
    assert!(connections.by_id(42).is_some());
    assert!(connections.by_name("Warehouse").is_some());
}

#[test]
fn parses_olap_and_web_connections_with_lookup_helpers() {
    let records = vec![
        record(
            rt::BEGIN_EXT_CONNECTION,
            &ExtConnectionBuilder::new(5, 7, "Cube").payload(),
        ),
        record(rt::BEGIN_EC_OLAP_PROPS, &olap_props_payload()),
        record(rt::END_EC_OLAP_PROPS, &[]),
        record(rt::END_EXT_CONNECTION, &[]),
        record(
            rt::BEGIN_EXT_CONNECTION,
            &ExtConnectionBuilder::new(4, 9, "Web Query").payload(),
        ),
        record(
            rt::BEGIN_EC_WEB_PROPS,
            &web_props_payload(2, Some("https://example.test/q")),
        ),
        record(rt::END_EC_WEB_PROPS, &[]),
        record(rt::BEGIN_EC_WP_TABLES, &[]),
        record(rt::PCDI_MISSING, &[]),               // BrtPCDIMissing
        record(rt::PCDI_STRING, &wide("results")),   // BrtPCDIString
        record(rt::PCDI_INDEX, &3u32.to_le_bytes()), // BrtPCDIIndex
        record(rt::END_EC_WP_TABLES, &[]),
        record(rt::END_EXT_CONNECTION, &[]),
    ];
    let connections = parse_connections_part(&full_part(&records)).unwrap();
    assert_eq!(connections.connections.len(), 2);
    match &connections.connections[0].properties {
        Properties::Olap(olap) => {
            assert!(olap.local_connection);
            assert!(olap.server_format_back);
            assert!(!olap.use_office_lcid);
            assert_eq!(olap.drillthrough_rows, 1024);
            assert_eq!(
                olap.local_connection_string.as_deref(),
                Some("Provider=MSOLAP;Data Source=local.cub")
            );
        },
        other => panic!("expected OLAP properties, got {other:?}"),
    }
    match &connections.connections[1].properties {
        Properties::Web(web) => {
            assert_eq!(web.html_format, HtmlFormat::All);
            assert!(web.source_is_xml);
            assert!(web.consecutive_delimiters);
            assert_eq!(web.url.as_deref(), Some("https://example.test/q"));
        },
        other => panic!("expected Web properties, got {other:?}"),
    }
    assert_eq!(
        connections.connections[1].web_tables,
        vec![
            WebTableItem::Missing,
            WebTableItem::Named("results".to_string()),
            WebTableItem::Index(3),
        ]
    );
}

#[test]
fn skips_unknown_records_and_extension_collections() {
    let mut records = vec![
        record(
            rt::BEGIN_EXT_CONNECTION,
            &ExtConnectionBuilder::new(6, 1, "T").payload(),
        ),
        // Text-import wizard collection: skipped as a balanced collection.
        record(rt::BEGIN_EC_TXT_WIZ, &[0xAA, 0xBB]),
        record(rt::END_EC_TXT_WIZ, &[]),
        // Excel 2014 extension collection with unknown children inside.
        record(rt::BEGIN_EXT_CONN14, &[0x01]),
        record(Kind::new(0x3FFE).unwrap(), &[0x00, 0x00]),
        record(rt::END_EXT_CONN14, &[]),
        // Unknown standalone record.
        record(Kind::new(0x3FFD).unwrap(), &[0x10, 0x20, 0x30]),
        record(rt::END_EXT_CONNECTION, &[]),
    ];
    let mut part = vec![record(rt::BEGIN_EXT_CONNECTIONS, &[])];
    part.append(&mut records);
    part.push(record(rt::END_EXT_CONNECTIONS, &[]));
    let connections = parse_connections_part(&build(&part)).unwrap();
    assert_eq!(connections.connections.len(), 1);
    assert_eq!(connections.connections[0].source_type, SourceType::Text);
    assert_eq!(connections.connections[0].properties, Properties::None);
}

#[test]
fn rejects_malformed_parts() {
    // Not a connections part.
    assert!(parse_connections_part(&build(&[record(rt::BEGIN_LIST, &[])])).is_err());
    // Truncated stream.
    assert!(parse_connections_part(&build(&[record(rt::BEGIN_EXT_CONNECTIONS, &[])])).is_err());
    // Non-empty begin payload.
    assert!(
        parse_connections_part(&build(&[
            record(rt::BEGIN_EXT_CONNECTIONS, &[0x01]),
            record(rt::END_EXT_CONNECTIONS, &[]),
        ]))
        .is_err()
    );
    // Truncated connection payload.
    let mut records = vec![
        record(rt::BEGIN_EXT_CONNECTIONS, &[]),
        record(rt::BEGIN_EXT_CONNECTION, &[7, 5, 2]),
        record(rt::END_EXT_CONNECTION, &[]),
        record(rt::END_EXT_CONNECTIONS, &[]),
    ];
    assert!(parse_connections_part(&build(&records)).is_err());
    // Unknown DBType.
    records = vec![
        record(rt::BEGIN_EXT_CONNECTIONS, &[]),
        record(
            rt::BEGIN_EXT_CONNECTION,
            &ExtConnectionBuilder::new(99, 1, "x").payload(),
        ),
        record(rt::END_EXT_CONNECTION, &[]),
        record(rt::END_EXT_CONNECTIONS, &[]),
    ];
    assert!(parse_connections_part(&build(&records)).is_err());
    // Duplicate connection ids are rejected after the complete part is parsed.
    records = vec![
        record(rt::BEGIN_EXT_CONNECTIONS, &[]),
        record(
            rt::BEGIN_EXT_CONNECTION,
            &ExtConnectionBuilder::new(1, 7, "first").payload(),
        ),
        record(rt::END_EXT_CONNECTION, &[]),
        record(
            rt::BEGIN_EXT_CONNECTION,
            &ExtConnectionBuilder::new(1, 7, "second").payload(),
        ),
        record(rt::END_EXT_CONNECTION, &[]),
        record(rt::END_EXT_CONNECTIONS, &[]),
    ];
    assert!(parse_connections_part(&build(&records)).is_err());
    // Unterminated connection collection.
    records = vec![
        record(rt::BEGIN_EXT_CONNECTIONS, &[]),
        record(
            rt::BEGIN_EXT_CONNECTION,
            &ExtConnectionBuilder::new(1, 1, "x").payload(),
        ),
        record(rt::END_EXT_CONNECTIONS, &[]),
    ];
    assert!(parse_connections_part(&build(&records)).is_err());
}

fn generated_workbook() -> Workbook {
    let mut writer = WorkbookWriter::new();
    writer.add_worksheet(MutableWorksheet::new("Sheet1"));
    let mut bytes = Cursor::new(Vec::new());
    writer.save(&mut bytes).unwrap();
    Workbook::new(Cursor::new(bytes.into_inner())).unwrap()
}

fn sample_connection(name: &str) -> Connections {
    Connections {
        connections: vec![Connection {
            connection_id: 42,
            source_type: SourceType::Odbc,
            name: name.to_string(),
            properties: Properties::Database(DbProperties {
                command_type: CommandType::Sql,
                connection_string: "Driver={Generated};Server=example.invalid".to_string(),
                command: Some("SELECT 1".to_string()),
                server_command: None,
            }),
            ..Connection::default()
        }],
    }
}

#[test]
fn source_checked_transaction_preserves_opaque_records_and_inverse() {
    let mut workbook = generated_workbook();
    workbook
        .set_connections(sample_connection("Before"))
        .unwrap();
    let uri = PackURI::new(super::package::CONNECTIONS_PART_NAME).unwrap();
    let opaque = build(&[record(
        Kind::new(0x3ffd).unwrap(),
        &[0xDE, 0xAD, 0xBE, 0xEF],
    )]);
    let source = {
        let part = workbook.opc_package().get_part(&uri).unwrap();
        let mut bytes = part.blob().to_vec();
        bytes.extend_from_slice(&opaque);
        bytes
    };
    workbook
        .opc_package_mut()
        .get_part_mut(&uri)
        .unwrap()
        .set_blob(source.clone());

    let commit = {
        let package = workbook.opc_package_mut();
        let mut transaction = super::transaction::Transaction::new(package).unwrap();
        assert!(
            transaction
                .edit(42, |connection| {
                    connection.name = "After".to_string();
                    Ok(())
                })
                .unwrap()
        );
        transaction.commit().unwrap()
    };
    assert!(commit.changed());
    let package = workbook.opc_package();
    let edited = package.get_part(&uri).unwrap().blob();
    assert!(
        edited
            .windows(opaque.len())
            .any(|window| window == opaque.as_slice())
    );
    assert_eq!(
        parse_connections_part(edited).unwrap().connections[0].name,
        "After"
    );

    let mut restored = package.clone();
    commit.patch().inverse().apply(&mut restored).unwrap();
    assert_eq!(restored.get_part(&uri).unwrap().blob(), source);
}

#[test]
fn source_checked_transaction_rejects_stale_patch_and_supports_noop() {
    let mut workbook = generated_workbook();
    workbook
        .set_connections(sample_connection("Stable"))
        .unwrap();
    let noop = {
        let transaction = super::transaction::Transaction::new(workbook.opc_package_mut()).unwrap();
        transaction.commit().unwrap()
    };
    assert!(!noop.changed());
    assert!(noop.patch().is_empty());

    let mut changed = workbook.opc_package().clone();
    let uri = PackURI::new(super::package::CONNECTIONS_PART_NAME).unwrap();
    let mut bytes = changed.get_part(&uri).unwrap().blob().to_vec();
    bytes.push(0);
    changed.get_part_mut(&uri).unwrap().set_blob(bytes);
    assert!(noop.patch().apply(&mut changed).is_err());
}
