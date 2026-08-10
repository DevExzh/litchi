#![allow(
    clippy::expect_used,
    reason = "test fixture uses bounded literal casts, panic-on-failure extraction, exact floating sentinels, or explicit negative fallback solely to state its assertion"
)]

//! Workbook-writer serialization and round-trip tests.

use super::super::WorkbookWriter;
use crate::calc::{Delta, Mode, Opts, Threads};
use crate::comments::{Record, Run};
use crate::raw::kind;
use crate::writer::{MutableChartSheet, MutableWorksheet};
use litchi_core::sheet::WorkbookTrait;
use litchi_opc::constants::relationship_type as rel;
use litchi_opc::{OpcPackage, PackURI};
use std::io::Cursor;

#[test]
fn external_links_round_trip_with_inert_package_topology() {
    use crate::external_link::{
        CachedValue, CellLocation, CellReference, DdeItem, DefinedName, ErrorValue, Kind, Link,
        NameFormula, NameFormulaKind, OleItem, SheetRange, ValueMatrix,
    };

    let mut workbook = WorkbookWriter::new();
    workbook.add_worksheet(MutableWorksheet::new("Host"));
    let external_formula = NameFormula::cell_reference(CellReference::new(
        SheetRange::sheets(0, 0).unwrap(),
        CellLocation::new(3, 2),
    ));
    workbook
        .add_external_link(
            Link::workbook_with_defined_names(
                "file:///data/Budget.xlsx",
                vec!["Data".to_string(), "Rates".to_string()],
                vec![
                    DefinedName::new("ExchangeRate")
                        .unwrap()
                        .with_formula(external_formula)
                        .with_built_in(true)
                        .with_sheet_scope(1),
                ],
            )
            .unwrap(),
        )
        .unwrap();
    let dde_cache = ValueMatrix::new(
        1,
        5,
        vec![
            CachedValue::Empty,
            CachedValue::Number(42.5),
            CachedValue::Boolean(true),
            CachedValue::Error(ErrorValue::NotAvailable),
            CachedValue::String("Ready".to_string()),
        ],
    )
    .unwrap();
    workbook
        .add_external_link(
            Link::dde_with_items(
                "Excel",
                "System",
                vec![
                    DdeItem::new("StatusItem")
                        .unwrap()
                        .with_advise(true)
                        .with_picture(true)
                        .with_cached_values(dde_cache),
                ],
            )
            .unwrap(),
        )
        .unwrap();
    workbook
        .add_external_link(
            Link::ole_with_items(
                "file:///data/Model.xlsx",
                "Excel.Sheet.12",
                vec![
                    OleItem::new("ModelItem")
                        .unwrap()
                        .with_advise(true)
                        .with_icon(true)
                        .with_cached_values(
                            ValueMatrix::new(1, 1, vec![CachedValue::Number(7.0)]).unwrap(),
                        ),
                ],
            )
            .unwrap(),
        )
        .unwrap();

    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    let bytes = output.into_inner();
    let package = OpcPackage::from_bytes(&bytes).unwrap();
    let workbook_part = package
        .get_part(&PackURI::new("/xl/workbook.bin").unwrap())
        .unwrap();
    let external_relationships = workbook_part
        .rels()
        .iter()
        .filter(|relationship| relationship.reltype() == rel::EXTERNAL_LINK)
        .collect::<Vec<_>>();
    assert_eq!(external_relationships.len(), 3);
    let relationships = (1..=3)
        .map(|index| {
            let expected =
                PackURI::new(format!("/xl/externalLinks/externalLink{index}.bin")).unwrap();
            external_relationships
                .iter()
                .copied()
                .find(|relationship| relationship.target_partname().unwrap() == expected)
                .expect("external-link relationship missing")
        })
        .collect::<Vec<_>>();
    let mut support_relationship_ids = Vec::new();
    for record in crate::raw::Records::new(workbook_part.blob()) {
        let record = record.unwrap();
        if record.kind() == kind::SUP_BOOK_SRC {
            let (relationship_id, consumed) =
                crate::package::records::decode_string(record.payload()).unwrap();
            assert_eq!(consumed, record.payload().len());
            support_relationship_ids.push(relationship_id);
        }
    }
    assert_eq!(
        support_relationship_ids,
        relationships
            .iter()
            .map(|relationship| relationship.r_id().to_string())
            .collect::<Vec<_>>()
    );
    for (index, relationship) in relationships.iter().enumerate() {
        assert!(!relationship.is_external());
        assert_eq!(
            relationship.target_partname().unwrap(),
            PackURI::new(format!("/xl/externalLinks/externalLink{}.bin", index + 1)).unwrap()
        );
    }
    let workbook_link = package
        .get_part(&relationships[0].target_partname().unwrap())
        .unwrap();
    assert_eq!(workbook_link.rels().len(), 1);
    assert!(
        workbook_link
            .rels()
            .iter()
            .all(|relationship| relationship.is_external()
                && relationship.reltype() == rel::EXTERNAL_LINK_PATH)
    );
    let dde_link = package
        .get_part(&relationships[1].target_partname().unwrap())
        .unwrap();
    assert!(dde_link.rels().is_empty());
    let ole_link = package
        .get_part(&relationships[2].target_partname().unwrap())
        .unwrap();
    assert_eq!(ole_link.rels().len(), 1);
    assert!(ole_link.rels().iter().all(
        |relationship| relationship.is_external() && relationship.reltype() == rel::OLE_OBJECT
    ));

    let reader = crate::Workbook::new(Cursor::new(bytes)).unwrap();
    assert_eq!(reader.external_link_count(), 3);
    assert_eq!(reader.external_link_iter().len(), 3);
    assert_eq!(reader.external_link(1).unwrap().dde_topic(), Some("System"));
    let links = reader.external_links();
    assert_eq!(links.len(), 3);
    assert_eq!(links[0].kind(), Kind::Workbook);
    assert_eq!(links[0].source(), "file:///data/Budget.xlsx");
    assert_eq!(links[0].sheet_names(), ["Data", "Rates"]);
    let defined_name = &links[0].defined_names()[0];
    assert_eq!(defined_name.name(), "ExchangeRate");
    assert!(defined_name.is_built_in());
    assert_eq!(defined_name.scope_sheet_index(), Some(1));
    assert_eq!(
        defined_name.formula().unwrap().kind(),
        NameFormulaKind::CellReference
    );
    assert_eq!(
        defined_name.formula().unwrap().tokens(),
        [0x3A, 0, 0, 0, 0, 3, 0, 2, 0]
    );
    assert_eq!(links[1].kind(), Kind::Dde);
    assert_eq!(links[1].source(), "Excel");
    assert_eq!(links[1].dde_topic(), Some("System"));
    let dde_item = &links[1].dde_items()[0];
    assert_eq!(dde_item.name(), "StatusItem");
    assert!(dde_item.wants_advise());
    assert!(dde_item.wants_picture());
    assert_eq!(dde_item.cached_values().unwrap().rows(), 1);
    assert_eq!(dde_item.cached_values().unwrap().columns(), 5);
    assert_eq!(
        dde_item.cached_values().unwrap().values(),
        [
            CachedValue::Empty,
            CachedValue::Number(42.5),
            CachedValue::Boolean(true),
            CachedValue::Error(ErrorValue::NotAvailable),
            CachedValue::String("Ready".to_string()),
        ]
    );
    assert_eq!(links[2].kind(), Kind::Ole);
    assert_eq!(links[2].source(), "file:///data/Model.xlsx");
    assert_eq!(links[2].ole_program_id(), Some("Excel.Sheet.12"));
    let ole_item = &links[2].ole_items()[0];
    assert_eq!(ole_item.name(), "ModelItem");
    assert!(ole_item.wants_advise());
    assert!(ole_item.displays_as_icon());
    assert_eq!(
        ole_item.cached_values().unwrap().values(),
        [CachedValue::Number(7.0)]
    );
}

#[test]
fn external_link_constructors_refuse_malformed_metadata() {
    use crate::external_link::{CachedValue, DdeItem, DefinedName, Link, NameFormula, ValueMatrix};

    assert!(Link::workbook("", Vec::new(), Vec::new()).is_err());
    assert!(
        Link::workbook(
            "Book.xlsx",
            vec!["Data".to_string(), "data".to_string()],
            Vec::new(),
        )
        .is_err()
    );
    assert!(Link::dde("Excel", "", vec!["Item".to_string()]).is_err());
    assert!(Link::ole("Model.xlsx", "Excel.Sheet", vec!["A1".to_string()],).is_err());
    assert!(NameFormula::from_tokens(vec![0x3A, 0]).is_err());
    assert!(ValueMatrix::new(2, 2, vec![CachedValue::Number(1.0)]).is_err());
    assert!(ValueMatrix::new(1, 1, vec![CachedValue::Number(-0.0)]).is_err());
    assert!(ValueMatrix::new(1, 1, vec![CachedValue::Number(f64::from_bits(1))]).is_err());
    assert!(ValueMatrix::new(1, 1, vec![CachedValue::String(String::new())]).is_ok());
    let invalid_formula = NameFormula::from_tokens(vec![0x3A, 2, 0, 2, 0, 0, 0, 0, 0]).unwrap();
    assert!(
        Link::workbook_with_defined_names(
            "Book.xlsx",
            vec!["OnlySheet".to_string()],
            vec![
                DefinedName::new("BadScope")
                    .unwrap()
                    .with_formula(invalid_formula)
            ],
        )
        .is_err()
    );
    assert!(
        Link::dde_with_items(
            "Excel",
            "System",
            vec![
                DdeItem::new("NotStdDocumentName")
                    .unwrap()
                    .with_ole_support(true)
                    .with_cached_values(ValueMatrix::new(1, 1, vec![CachedValue::Empty]).unwrap())
            ],
        )
        .is_err()
    );
}

#[test]
fn chart_sheet_metadata_chart_and_printer_settings_round_trip_in_sheet_order() {
    use crate::chart::{Anchor, Chart};
    use crate::package::chartsheet::{Color, ColorType, PageSetup, Protection, State, View};
    use crate::sheet::StrongProtection;

    let chart = Chart::bar_chart_with_cache(
        "Sales",
        "Data!$A$2:$A$3",
        &["North", "South"],
        "Data!$B$2:$B$3",
        &[42.0, 55.0],
        Anchor::new(0, 0, 10, 20),
    )
    .unwrap();
    let mut chart_sheet = MutableChartSheet::new("Sales Chart", chart);
    {
        let metadata = chart_sheet.metadata_mut();
        metadata.state = State::Hidden;
        metadata.code_name = "ChartCode".to_string();
        metadata.published = true;
        metadata.tab_color = Color {
            valid_rgb: true,
            color_type: ColorType::Rgb,
            index: 0,
            tint: -100,
            rgba: [0x44, 0x72, 0xc4, 0xff],
        };
        metadata.views = vec![View {
            selected: true,
            scale: 125,
            workbook_view_index: 0,
        }];
        metadata.protection = Some(Protection {
            password_verifier: 0x1234,
            locked: true,
            objects: false,
        });
        metadata.strong_protection = Some(StrongProtection {
            spin_count: 100_000,
            hash: vec![7; 64],
            salt: vec![3; 16],
            algorithm: "SHA-512".to_string(),
        });
    }
    chart_sheet
        .set_page_setup(
            PageSetup {
                paper_size: 9,
                horizontal_resolution: 600,
                vertical_resolution: 600,
                copies: 2,
                page_start: 4,
                landscape: true,
                black_and_white: true,
                use_default_orientation: false,
                use_page_start: true,
                draft: false,
                printer_settings_rel_id: "caller-id-is-replaced".to_string(),
            },
            vec![1, 2, 3, 4],
        )
        .unwrap();

    let mut workbook = WorkbookWriter::new();
    workbook.add_worksheet(MutableWorksheet::new("Data"));
    workbook.add_chart_sheet(chart_sheet).unwrap();
    workbook.add_worksheet(MutableWorksheet::new("Tail"));
    assert_eq!(workbook.chart_sheet_count(), 1);

    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    let bytes = output.into_inner();
    let package = OpcPackage::from_bytes(&bytes).unwrap();
    let reader = crate::Workbook::new(Cursor::new(bytes)).unwrap();
    assert_eq!(
        reader.worksheet_names(),
        &[
            "Data".to_string(),
            "Sales Chart".to_string(),
            "Tail".to_string()
        ]
    );
    let parsed = reader.chart_sheet(1).expect("chart sheet missing");
    assert_eq!(parsed.state, State::Hidden);
    assert_eq!(parsed.code_name, "ChartCode");
    assert!(parsed.published);
    assert_eq!(parsed.tab_color.rgba, [0x44, 0x72, 0xc4, 0xff]);
    assert_eq!(parsed.views[0].scale, 125);
    assert_eq!(parsed.protection.unwrap().password_verifier, 0);
    assert_eq!(
        parsed.strong_protection.as_ref().unwrap().algorithm,
        "SHA-512"
    );
    let page_setup = parsed.page_setup.as_ref().unwrap();
    assert_eq!(page_setup.copies, 2);
    assert_eq!(page_setup.printer_settings_rel_id, "rId2");

    let drawing = reader.sheet_drawing(1).expect("chart drawing missing");
    assert_eq!(drawing.drawing.anchors.len(), 1);
    assert_eq!(drawing.charts.len(), 1);
    let printer = package
        .get_part(&PackURI::new("/xl/printerSettings/printerSettings1.bin").unwrap())
        .unwrap();
    assert_eq!(printer.blob(), &[1, 2, 3, 4]);
}

#[test]
fn chart_sheet_validation_is_lossless_or_refuse() {
    use crate::chart::{Anchor, Chart};
    use crate::package::chartsheet::{Color, ColorType};

    let chart = Chart::bar_chart(
        "T",
        "Data!$A$1:$A$2",
        "Data!$B$1:$B$2",
        Anchor::new(0, 0, 5, 5),
    )
    .unwrap();
    let mut workbook = WorkbookWriter::new();
    workbook.add_worksheet(MutableWorksheet::new("Data"));
    workbook
        .add_chart_sheet(MutableChartSheet::new("Chart", chart.clone()))
        .unwrap();
    assert!(
        workbook
            .add_chart_sheet(MutableChartSheet::new("chart", chart.clone()))
            .is_err()
    );

    let mut invalid = MutableChartSheet::new("Invalid", chart.clone());
    invalid.metadata_mut().views[0].scale = 401;
    assert!(workbook.add_chart_sheet(invalid).is_err());

    let mut invalid_color = MutableChartSheet::new("Invalid Color", chart);
    invalid_color.metadata_mut().tab_color = Color {
        valid_rgb: false,
        color_type: ColorType::Indexed,
        index: 0x52,
        tint: 0,
        rgba: [0; 4],
    };
    assert!(workbook.add_chart_sheet(invalid_color).is_err());
}

#[test]
fn test_set_date_system() {
    let mut workbook = WorkbookWriter::new();
    workbook.set_date_system(true);
    assert!(workbook.is_1904);
}

#[test]
fn calc_survives_package_roundtrip() {
    let mut workbook = WorkbookWriter::new();
    let properties = workbook.calc_mut();
    properties
        .set_mode(Mode::Manual)
        .set_iters(25)
        .set_delta(Delta::new(0.000_01).unwrap())
        .set_threads(Threads::new(4).unwrap());
    properties
        .set_opt(
            Opts::ITERATE | Opts::USER_THREADS | Opts::FULL_ON_LOAD,
            true,
        )
        .unwrap();
    workbook.add_worksheet(MutableWorksheet::new("Sheet1"));

    let expected = workbook.calc().clone();
    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    let reader = crate::Workbook::new(Cursor::new(output.into_inner())).unwrap();
    assert_eq!(reader.calc(), &expected);
}

#[test]
fn connections_round_trip_through_save_and_read() {
    use crate::package::connections::*;

    let connections = Connections {
        connections: vec![
            Connection {
                connection_id: 42,
                source_type: SourceType::Odbc,
                name: "Warehouse".to_string(),
                refresh_interval_minutes: 30,
                background_query: true,
                credential_method: Some(CredentialMethod::Integrated),
                properties: Properties::Database(DbProperties {
                    command_type: CommandType::Sql,
                    connection_string: "Driver={SQL Server};Server=db".to_string(),
                    command: Some("SELECT * FROM T".to_string()),
                    server_command: None,
                }),
                ..Connection::default()
            },
            Connection {
                connection_id: 9,
                source_type: SourceType::Web,
                name: "Web Query".to_string(),
                properties: Properties::Web(WebProperties {
                    html_format: HtmlFormat::All,
                    url: Some("https://example.test/q".to_string()),
                    ..WebProperties::default()
                }),
                web_tables: vec![WebTableItem::Index(1)],
                ..Connection::default()
            },
        ],
    };

    let mut workbook = WorkbookWriter::new();
    workbook.add_worksheet(MutableWorksheet::new("Sheet1"));
    workbook.set_connections(connections.clone()).unwrap();
    // Validation: zero id, duplicate id, duplicate name (case-insensitive).
    assert!(
        workbook
            .set_connections(Connections {
                connections: vec![Connection {
                    connection_id: 0,
                    name: "bad".to_string(),
                    ..Connection::default()
                }],
            })
            .is_err()
    );
    assert!(
        workbook
            .set_connections(Connections {
                connections: vec![
                    Connection {
                        connection_id: 5,
                        name: "a".to_string(),
                        ..Connection::default()
                    },
                    Connection {
                        connection_id: 5,
                        name: "b".to_string(),
                        ..Connection::default()
                    },
                ],
            })
            .is_err()
    );
    assert!(
        workbook
            .set_connections(Connections {
                connections: vec![
                    Connection {
                        connection_id: 5,
                        name: "Dup".to_string(),
                        ..Connection::default()
                    },
                    Connection {
                        connection_id: 6,
                        name: "dup".to_string(),
                        ..Connection::default()
                    },
                ],
            })
            .is_err()
    );

    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    let reader = crate::Workbook::new(Cursor::new(output.into_inner())).unwrap();
    let parsed = reader.connections().expect("connections part missing");
    assert_eq!(parsed, &connections);
    assert_eq!(parsed.by_id(42).unwrap().name, "Warehouse");
    assert!(parsed.by_name("Web Query").is_some());
}

#[test]
fn structured_tables_round_trip_through_save_and_read() {
    use crate::package::table::{
        Column, Formula, Range, StyleInfo, Table, TotalsRowFunction, Type,
    };

    let table = Table {
        id: 3,
        name: Some("SalesTable".to_string()),
        display_name: Some("SalesTable".to_string()),
        range: Range {
            first_row: 0,
            last_row: 2,
            first_column: 0,
            last_column: 1,
        },
        table_type: Type::Range,
        header_row_count: 1,
        columns: vec![
            Column {
                id: 1,
                name: Some("Region".to_string()),
                ..Column::default()
            },
            Column {
                id: 2,
                name: Some("Amount".to_string()),
                totals_row_function: TotalsRowFunction::Sum,
                calculated_column_formula: Some(Formula {
                    array: false,
                    tokens: vec![0x1E, 0x02],
                    extra: Vec::new(),
                }),
                ..Column::default()
            },
        ],
        style_info: Some(StyleInfo {
            name: Some("TableStyleMedium2".to_string()),
            show_first_column: false,
            show_last_column: false,
            show_row_stripes: true,
            show_column_stripes: false,
        }),
        ..Table::default()
    };

    let mut workbook = WorkbookWriter::new();
    let mut sheet = MutableWorksheet::new("Sheet1");
    sheet.set_cell(0, 0, "Region");
    sheet.set_cell(0, 1, "Amount");
    sheet.set_cell(1, 0, "North");
    sheet.set_cell(1, 1, 42.5);
    sheet.add_table(table.clone()).unwrap();
    // Validation: missing display name, inverted range, width mismatch,
    // duplicate id.
    assert!(
        sheet
            .add_table(Table {
                id: 9,
                range: table.range,
                ..Table::default()
            })
            .is_err()
    );
    assert!(
        sheet
            .add_table(Table {
                id: 9,
                display_name: Some("Bad".to_string()),
                range: Range {
                    first_row: 5,
                    last_row: 2,
                    first_column: 0,
                    last_column: 0,
                },
                ..Table::default()
            })
            .is_err()
    );
    assert!(
        sheet
            .add_table(Table {
                id: 9,
                display_name: Some("Bad".to_string()),
                range: table.range,
                columns: vec![Column::default()],
                ..Table::default()
            })
            .is_err()
    );
    let mut duplicate = table.clone();
    duplicate.display_name = Some("Other".to_string());
    assert!(sheet.add_table(duplicate).is_err());
    workbook.add_worksheet(sheet);

    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    let reader = crate::Workbook::new(Cursor::new(output.into_inner())).unwrap();
    let tables = reader.structured_tables();
    assert_eq!(tables.len(), 1);
    assert_eq!(tables[0].0, 0);
    assert_eq!(tables[0].1, table);
    assert_eq!(reader.tables_on_sheet(0).len(), 1);
    assert!(reader.tables_on_sheet(1).is_empty());
}

#[test]
fn comments_survive_package_roundtrip() {
    let mut workbook = WorkbookWriter::new();
    let mut sheet = MutableWorksheet::new("Notes");
    let mut first = Record::new(2, 3, "Author".to_string(), "formatted note".to_string());
    first.runs = vec![Run {
        character_index: 0,
        font_id: 0,
    }];
    first.guid = [7; 16];
    sheet.add_comment(first);
    sheet.add_comment(Record::new(
        4,
        1,
        "Author".to_string(),
        "second note".to_string(),
    ));
    workbook.add_worksheet(sheet);

    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    let reader = crate::Workbook::new(Cursor::new(output.into_inner())).unwrap();
    let worksheet = reader.worksheet(0).unwrap();
    assert_eq!(worksheet.comments().len(), 2);
    assert_eq!(worksheet.comments()[0].text, "formatted note");
    assert_eq!(worksheet.comments()[0].runs.len(), 1);
    assert_eq!(worksheet.comments()[0].guid, [7; 16]);
    assert_eq!(worksheet.comments()[1].author, "Author");
}
