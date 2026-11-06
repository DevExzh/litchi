#![allow(
    clippy::cast_possible_truncation,
    clippy::wildcard_enum_match_arm,
    reason = "test fixture uses bounded literal casts, panic-on-failure extraction, exact floating sentinels, or explicit negative fallback solely to state its assertion"
)]

//! Synthetic-stream tests for the External Data Connections parser.

use super::model::*;
use super::parse::parse_connections_part;
use crate::builder::{self, Builder};
use crate::raw::{Encoder, Kind, kind as rt};
use litchi_opc::PackURI;

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
    let mut writer = Encoder::new(&mut data, crate::raw::Limits::DEFAULT);
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
    let connections = parse_connections_part(&part, crate::raw::Limits::DEFAULT).unwrap();
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
    let connections =
        parse_connections_part(&full_part(&records), crate::raw::Limits::DEFAULT).unwrap();
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
        // The model-data-source text extension remains an opaque collection.
        record(rt::BEGIN_EC_TXT_WIZ15, &[0xAA, 0xBB]),
        record(rt::END_EC_TXT_WIZ15, &[]),
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
    let connections = parse_connections_part(&build(&part), crate::raw::Limits::DEFAULT).unwrap();
    assert_eq!(connections.connections.len(), 1);
    assert_eq!(connections.connections[0].source_type, SourceType::Text);
    assert_eq!(connections.connections[0].properties, Properties::None);
}

#[test]
fn rejects_malformed_parts() {
    // Not a connections part.
    assert!(
        parse_connections_part(
            &build(&[record(rt::BEGIN_LIST, &[])]),
            crate::raw::Limits::DEFAULT
        )
        .is_err()
    );
    // Truncated stream.
    assert!(
        parse_connections_part(
            &build(&[record(rt::BEGIN_EXT_CONNECTIONS, &[])]),
            crate::raw::Limits::DEFAULT
        )
        .is_err()
    );
    // Non-empty begin payload.
    assert!(
        parse_connections_part(
            &build(&[
                record(rt::BEGIN_EXT_CONNECTIONS, &[0x01]),
                record(rt::END_EXT_CONNECTIONS, &[]),
            ]),
            crate::raw::Limits::DEFAULT
        )
        .is_err()
    );
    // Truncated connection payload.
    let mut records = vec![
        record(rt::BEGIN_EXT_CONNECTIONS, &[]),
        record(rt::BEGIN_EXT_CONNECTION, &[7, 5, 2]),
        record(rt::END_EXT_CONNECTION, &[]),
        record(rt::END_EXT_CONNECTIONS, &[]),
    ];
    assert!(parse_connections_part(&build(&records), crate::raw::Limits::DEFAULT).is_err());
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
    assert!(parse_connections_part(&build(&records), crate::raw::Limits::DEFAULT).is_err());
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
    assert!(parse_connections_part(&build(&records), crate::raw::Limits::DEFAULT).is_err());
    // Unterminated connection collection.
    records = vec![
        record(rt::BEGIN_EXT_CONNECTIONS, &[]),
        record(
            rt::BEGIN_EXT_CONNECTION,
            &ExtConnectionBuilder::new(1, 1, "x").payload(),
        ),
        record(rt::END_EXT_CONNECTIONS, &[]),
    ];
    assert!(parse_connections_part(&build(&records), crate::raw::Limits::DEFAULT).is_err());
}

fn generated_package() -> crate::Package {
    let mut writer = Builder::new();
    writer.add_worksheet(builder::Worksheet::new("Sheet1"));
    writer.build().unwrap()
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

fn with_connections(package: &crate::Package, value: Connections) -> crate::Package {
    let mut edit = package.edit_connections().unwrap();
    edit.replace(Some(value)).unwrap();
    edit.commit().unwrap().into_package()
}

#[test]
fn source_checked_transaction_preserves_opaque_records_and_inverse() {
    let package = with_connections(&generated_package(), sample_connection("Before"));
    let uri = PackURI::new(super::package::CONNECTIONS_PART_NAME).unwrap();
    let opaque = build(&[record(
        Kind::new(0x3ffd).unwrap(),
        &[0xDE, 0xAD, 0xBE, 0xEF],
    )]);
    let source = {
        let part = package.opc_package().get_part(&uri).unwrap();
        let mut bytes = part.blob().to_vec();
        bytes.extend_from_slice(&opaque);
        bytes
    };
    let mut raw = package.edit_opc().unwrap();
    raw.update(|opc| {
        opc.get_part_mut(&uri)?.set_blob(source.clone());
        Ok(())
    })
    .unwrap();
    let package = raw.commit().unwrap().into_package();

    let mut edit = package.edit_connections().unwrap();
    assert!(
        edit.edit(42, |connection| {
            connection.name = "After".to_string();
            Ok(())
        })
        .unwrap()
    );
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let edited = commit
        .package()
        .opc_package()
        .get_part(&uri)
        .unwrap()
        .blob();
    assert!(
        edited
            .windows(opaque.len())
            .any(|window| window == opaque.as_slice())
    );
    assert_eq!(
        parse_connections_part(edited, crate::raw::Limits::DEFAULT)
            .unwrap()
            .connections[0]
            .name,
        "After"
    );

    let restored = commit
        .package()
        .apply_connections(&commit.patch().inverse())
        .unwrap()
        .into_package();
    assert_eq!(
        restored.opc_package().get_part(&uri).unwrap().blob(),
        source
    );
}

#[test]
fn source_checked_transaction_rejects_stale_patch_and_supports_noop() {
    let package = with_connections(&generated_package(), sample_connection("Stable"));
    let noop = package.edit_connections().unwrap().commit().unwrap();
    assert!(!noop.diagnostics().changed());
    assert!(noop.patch().is_noop());

    let mut edit = package.edit_connections().unwrap();
    edit.replace(Some(sample_connection("Changed"))).unwrap();
    let changed = edit.commit().unwrap();
    let uri = PackURI::new(super::package::CONNECTIONS_PART_NAME).unwrap();
    let mut raw = package.edit_opc().unwrap();
    raw.update(|opc| {
        let mut bytes = opc.get_part(&uri)?.blob().to_vec();
        bytes.push(0);
        opc.get_part_mut(&uri)?.set_blob(bytes);
        Ok(())
    })
    .unwrap();
    let stale = raw.commit();
    if let Ok(stale) = stale {
        assert!(stale.package().apply_connections(changed.patch()).is_err());
    }
}

fn text_model(properties: TextImportProperties) -> Connections {
    let bytes = build(&[
        record(rt::BEGIN_EXT_CONNECTIONS, &[]),
        record(
            rt::BEGIN_EXT_CONNECTION,
            &ExtConnectionBuilder::new(6, 1, "Text").payload(),
        ),
        record(rt::END_EXT_CONNECTION, &[]),
        record(rt::END_EXT_CONNECTIONS, &[]),
    ]);
    let mut model = parse_connections_part(&bytes, crate::raw::Limits::DEFAULT).unwrap();
    model.connections[0].properties = Properties::Text(properties);
    model
}

#[test]
fn classic_text_flags_fields_and_ignored_bits_round_trip() {
    for code_page in [
        TextCodePage::Macintosh,
        TextCodePage::WindowsAnsi,
        TextCodePage::Dos,
    ] {
        for qualifier in [
            TextQualifier::DoubleQuote,
            TextQualifier::SingleQuote,
            TextQualifier::None,
            TextQualifier::NoneAlternate,
        ] {
            for fields in [
                None,
                Some(Vec::new()),
                Some(
                    [
                        TextFieldType::General,
                        TextFieldType::Text,
                        TextFieldType::Mdy,
                        TextFieldType::Dmy,
                        TextFieldType::Ymd,
                        TextFieldType::Myd,
                        TextFieldType::Dym,
                        TextFieldType::Ydm,
                        TextFieldType::Skip,
                        TextFieldType::Emd,
                    ]
                    .into_iter()
                    .enumerate()
                    .map(|(i, field_type)| TextField {
                        field_type,
                        start: i as u32 * 7,
                    })
                    .collect(),
                ),
            ] {
                let properties = TextImportProperties {
                    code_page,
                    qualifier,
                    application_code_page: 1023,
                    delimited: false,
                    tab: true,
                    space: true,
                    comma: true,
                    semicolon: true,
                    consecutive: true,
                    prompt_for_file: false,
                    custom_delimiter: true,
                    delimiter: 0x00a7,
                    first_row: 99_999_999,
                    decimal_separator: 0xb7,
                    thousands_separator: 0xa0,
                    source_file: "source-é-😀.csv".into(),
                    fields,
                    ignored_flags: 0xff90_0000,
                };
                let model = text_model(properties);
                let bytes =
                    super::write::write_connections_part(&model, crate::raw::Limits::DEFAULT)
                        .unwrap();
                assert_eq!(
                    parse_connections_part(&bytes, crate::raw::Limits::DEFAULT).unwrap(),
                    model
                );
            }
        }
    }
}

#[test]
fn malformed_text_fields_and_boundaries_do_not_become_empty_properties() {
    let base = TextImportProperties::default();
    let payload = super::text::payload(&base).unwrap();
    let stream = |body: Vec<(Kind, Vec<u8>)>| {
        let mut records = vec![
            record(rt::BEGIN_EXT_CONNECTIONS, &[]),
            record(
                rt::BEGIN_EXT_CONNECTION,
                &ExtConnectionBuilder::new(6, 1, "Text").payload(),
            ),
        ];
        records.extend(body);
        records.extend([
            record(rt::END_EXT_CONNECTION, &[]),
            record(rt::END_EXT_CONNECTIONS, &[]),
        ]);
        build(&records)
    };
    for body in [
        vec![
            record(rt::BEGIN_EC_TXT_WIZ, &[0xaa, 0xbb]),
            record(rt::END_EC_TXT_WIZ, &[]),
        ],
        vec![
            record(rt::BEGIN_EC_TXT_WIZ, &payload),
            record(rt::BEGIN_EC_TW_FLD_INFO_LIST, &1u32.to_le_bytes()),
            record(rt::END_EC_TW_FLD_INFO_LIST, &[]),
            record(rt::END_EC_TXT_WIZ, &[]),
        ],
        vec![
            record(rt::BEGIN_EC_TXT_WIZ, &payload),
            record(rt::BEGIN_EC_TW_FLD_INFO_LIST, &0u32.to_le_bytes()),
            record(rt::BEGIN_EC_TW_FLD_INFO, &[0; 8]),
            record(rt::END_EC_TW_FLD_INFO_LIST, &[]),
            record(rt::END_EC_TXT_WIZ, &[]),
        ],
        vec![
            record(rt::BEGIN_EC_TXT_WIZ, &payload),
            record(rt::BEGIN_EC_TW_FLD_INFO_LIST, &16_385u32.to_le_bytes()),
        ],
        vec![
            record(rt::BEGIN_EC_TXT_WIZ, &payload),
            record(rt::BEGIN_EC_TW_FLD_INFO, &[0; 8]),
            record(rt::END_EC_TXT_WIZ, &[]),
        ],
        vec![
            record(rt::BEGIN_EC_TXT_WIZ, &payload),
            record(rt::BEGIN_EC_TW_FLD_INFO_LIST, &1u32.to_le_bytes()),
            record(rt::BEGIN_EC_TW_FLD_INFO, &[10, 0, 0, 0, 0, 0, 0, 0]),
            record(rt::END_EC_TW_FLD_INFO_LIST, &[]),
            record(rt::END_EC_TXT_WIZ, &[]),
        ],
        vec![
            record(rt::BEGIN_EC_TXT_WIZ, &payload),
            record(rt::END_EC_TXT_WIZ, &[]),
            record(rt::BEGIN_EC_TXT_WIZ, &payload),
            record(rt::END_EC_TXT_WIZ, &[]),
        ],
    ] {
        assert!(parse_connections_part(&stream(body), crate::raw::Limits::DEFAULT).is_err());
    }
    for mutate in [0, 1, 2] {
        let mut invalid = payload.clone();
        match mutate {
            0 => invalid[0] |= 3,
            1 => invalid[6..10].fill(0),
            _ => invalid[2] &= !0x20,
        }
        assert!(
            parse_connections_part(
                &stream(vec![
                    record(rt::BEGIN_EC_TXT_WIZ, &invalid),
                    record(rt::END_EC_TXT_WIZ, &[])
                ]),
                crate::raw::Limits::DEFAULT
            )
            .is_err()
        );
    }
}

#[test]
fn text_authoring_checks_path_row_and_field_limits_before_writing() {
    for invalid in [
        TextImportProperties {
            source_file: "a".repeat(219),
            ..Default::default()
        },
        TextImportProperties {
            first_row: 0,
            ..Default::default()
        },
        TextImportProperties {
            first_row: 100_000_000,
            ..Default::default()
        },
        TextImportProperties {
            application_code_page: 1024,
            ..Default::default()
        },
        TextImportProperties {
            prompt_for_file: false,
            ..Default::default()
        },
        TextImportProperties {
            ignored_flags: 1,
            ..Default::default()
        },
        TextImportProperties {
            fields: Some(vec![
                TextField {
                    field_type: TextFieldType::General,
                    start: 0
                };
                16_385
            ]),
            ..Default::default()
        },
    ] {
        assert!(
            super::write::write_connections_part(&text_model(invalid), crate::raw::Limits::DEFAULT)
                .is_err()
        );
    }
    let valid = text_model(TextImportProperties {
        source_file: "a".repeat(218),
        ..Default::default()
    });
    assert!(super::write::write_connections_part(&valid, crate::raw::Limits::DEFAULT).is_ok());
}

#[test]
fn opaque_records_keep_root_local_and_trailing_positions() {
    let unknown = Kind::new(0x3ffd).unwrap();
    let bytes = build(&[
        record(rt::BEGIN_EXT_CONNECTIONS, &[]),
        record(unknown, b"root"),
        record(
            rt::BEGIN_EXT_CONNECTION,
            &ExtConnectionBuilder::new(1, 1, "Before").payload(),
        ),
        record(unknown, b"local"),
        record(rt::END_EXT_CONNECTION, &[]),
        record(rt::END_EXT_CONNECTIONS, &[]),
        record(unknown, b"tail"),
    ]);
    let source = super::codec::parse_source(&bytes, crate::raw::Limits::DEFAULT).unwrap();
    assert_eq!(
        source
            .unknown_records
            .iter()
            .map(|r| (r.connection_id(), r.after_known()))
            .collect::<Vec<_>>(),
        [(None, 1), (Some(1), 0), (None, 3)]
    );
    let mut model = parse_connections_part(&bytes, crate::raw::Limits::DEFAULT).unwrap();
    model.connections[0].name = "After".into();
    let changed = super::write::write_connections_part_with_unknown(
        &model,
        &source.unknown_records,
        crate::raw::Limits::DEFAULT,
    )
    .unwrap();
    let after = super::codec::parse_source(&changed, crate::raw::Limits::DEFAULT).unwrap();
    assert_eq!(after.unknown_records, source.unknown_records);
    assert_eq!(
        parse_connections_part(&changed, crate::raw::Limits::DEFAULT)
            .unwrap()
            .connections[0]
            .name,
        "After"
    );
}

#[test]
fn ignored_collections_keep_known_record_kinds_opaque_and_do_not_create_connections() {
    let bytes = build(&[
        record(rt::BEGIN_EXT_CONNECTIONS, &[]),
        record(
            rt::BEGIN_EXT_CONNECTION,
            &ExtConnectionBuilder::new(1, 1, "Before").payload(),
        ),
        record(rt::FRT_BEGIN, &[0; 4]),
        record(
            rt::BEGIN_EXT_CONNECTION,
            &ExtConnectionBuilder::new(1, 99, "Hidden").payload(),
        ),
        record(rt::END_EXT_CONNECTION, &[]),
        record(rt::FRT_END, &[]),
        record(rt::END_EXT_CONNECTION, &[]),
        record(rt::END_EXT_CONNECTIONS, &[]),
    ]);
    let source = super::codec::parse_source(&bytes, crate::raw::Limits::DEFAULT).unwrap();
    assert_eq!(source.unknown_records.len(), 4);
    assert!(
        source
            .unknown_records
            .iter()
            .all(|r| r.connection_id() == Some(1) && r.after_known() == 0)
    );
    super::validation::unknown_records(&source.unknown_records).unwrap();
    let mut model = parse_connections_part(&bytes, crate::raw::Limits::DEFAULT).unwrap();
    assert_eq!(model.connections.len(), 1);
    model.connections[0].name = "After".into();
    let changed = super::write::write_connections_part_with_unknown(
        &model,
        &source.unknown_records,
        crate::raw::Limits::DEFAULT,
    )
    .unwrap();
    assert_eq!(
        parse_connections_part(&changed, crate::raw::Limits::DEFAULT).unwrap(),
        model
    );
    assert_eq!(
        super::codec::parse_source(&changed, crate::raw::Limits::DEFAULT)
            .unwrap()
            .unknown_records,
        source.unknown_records
    );
}
