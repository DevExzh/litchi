use std::io::Cursor;

use litchi_ooxml_common::spreadsheet_xml_maps::{
    DataBinding, XmlMap, XmlMapConformance, XmlMapInfo, XmlSchema,
};
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part};
use litchi_xlsb::Workbook;
use litchi_xlsb::package::table::{Column, Range, Table, Type};
use litchi_xlsb::writer::{MutableWorksheet, WorkbookWriter};
use litchi_xlsb::xml_maps::{
    CellReference, ColumnBinding, MappedTable, ReadLimits, SingleCellBinding, XPath, XmlDataType,
};

const WORKBOOK_PART: &str = "/xl/workbook.bin";
const WORKSHEET_PART: &str = "/xl/worksheets/sheet1.bin";
const MAP_INFO_PART: &str = "/xl/xmlMaps.xml";
const SINGLE_CELLS_PART: &str = "/xl/tables/singleCells1.bin";
const TABLE_PART: &str = "/xl/tables/table1.bin";
const SIGNATURE_ORIGIN_PART: &str = "/_xmlsignatures/origin.sigs";

const OFFICE_DOCUMENT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const WORKSHEET_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const XML_MAPS_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/xmlMaps";
const TABLE_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
const SINGLE_CELLS_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableSingleCells";
const SIGNATURE_ORIGIN_REL: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";

const WORKBOOK_CONTENT_TYPE: &str = "application/vnd.ms-excel.sheet.binary.macroEnabled.main";
const WORKSHEET_CONTENT_TYPE: &str = "application/vnd.ms-excel.worksheet";
const TABLE_CONTENT_TYPE: &str = "application/vnd.ms-excel.table";
const SINGLE_CELLS_CONTENT_TYPE: &str = "application/vnd.ms-excel.tableSingleCells";
const SIGNATURE_ORIGIN_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-package.digital-signature-origin";

const BRT_BEGIN_SHEET: u16 = 129;
const BRT_END_SHEET: u16 = 130;
const BRT_BEGIN_SHEET_DATA: u16 = 145;
const BRT_END_SHEET_DATA: u16 = 146;
const BRT_WS_DIM: u16 = 148;
const BRT_BUNDLE_SH: u16 = 156;
const BRT_BEGIN_SINGLE_CELLS: u16 = 341;
const BRT_END_SINGLE_CELLS: u16 = 342;
const BRT_BEGIN_LIST: u16 = 343;
const BRT_END_LIST: u16 = 344;
const BRT_BEGIN_LIST_COLS: u16 = 345;
const BRT_END_LIST_COLS: u16 = 346;
const BRT_BEGIN_LIST_COL: u16 = 347;
const BRT_END_LIST_COL: u16 = 348;
const BRT_BEGIN_LIST_XML_CPR: u16 = 349;
const BRT_END_LIST_XML_CPR: u16 = 350;
const BRT_BEGIN_LIST_PARTS: u16 = 660;
const BRT_LIST_PART: u16 = 661;
const BRT_END_LIST_PARTS: u16 = 662;
const BRT_FRT_BEGIN: u16 = 35;
const BRT_FRT_END: u16 = 36;
const BRT_AC_BEGIN: u16 = 37;
const BRT_AC_END: u16 = 38;
const BRT_FIXTURE_UNKNOWN: u16 = 777;

const MAP_INFO_XML: &[u8] = br#"<?xml version='1.0' encoding='UTF-8'?>
<MapInfo xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" SelectionNamespaces="xmlns:e='urn:litchi:fixture'">
  <!-- this spelling and whitespace must survive an unrelated binary no-op -->
  <Schema ID="schema-7" SchemaRef="urn:litchi:fixture" Namespace="urn:litchi:fixture"><e:schema xmlns:e="urn:litchi:fixture"/></Schema>
  <Map ID="7" Name="Orders" RootElement="root" SchemaID="schema-7" ShowImportExportValidationErrors="true" AutoFit="true" Append="false" PreserveSortAFLayout="true" PreserveFormat="true"><DataBinding DataBindingName="inert" FileBinding="false" DataBindingLoadMode="1"/></Map>
</MapInfo>"#;

const SINGLE_208_HEX: &str = "D50200D70258000000000000000000000000000000000200000001000000000000000000000002000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF00000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFD9020401000000DB02300100000000000000FFFFFFFFFFFFFFFFFFFFFFFF00000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFDD02260700000002000000010000000B0000002F0072006F006F0074002F00760061006C0075006500DE0200DC0200DA0200D80200D60200";
const TABLE_248_HEX: &str = "D70270010000000300000001000000010000000200000002000000010000000000000000000000FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF00000000FFFFFFFF0C0000004D00610070007000650064005400610062006C0065003100FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFD9020401000000DB02460100000000000000FFFFFFFFFFFFFFFFFFFFFFFF00000000FFFFFFFF0B0000004D0061007000700065006400560061006C0075006500FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFDD022607000000000000000D0000000B0000002F0072006F006F0074002F00760061006C0075006500DE0200DC0200DA0200D80200";

#[derive(Clone, Copy)]
enum RecordOrder {
    Valid,
    EndColumnBeforeXml,
    MissingXmlEnd,
}

#[derive(Clone)]
struct BindingWire {
    map_id: u32,
    flags: u32,
    data_type: u32,
    xpath: String,
    order: RecordOrder,
}

impl Default for BindingWire {
    fn default() -> Self {
        Self {
            map_id: 7,
            flags: 2,
            data_type: 1,
            xpath: "/root/value".to_owned(),
            order: RecordOrder::Valid,
        }
    }
}

#[derive(Clone, Copy)]
enum OutboundOwner {
    Maps,
    SingleCells,
}

#[derive(Clone, Copy)]
struct Shape {
    maps_content_type: &'static str,
    single_content_type: &'static str,
    table_content_type: &'static str,
    omit_maps_relationship: bool,
    maps_external: bool,
    single_external: bool,
    table_external: bool,
    duplicate_single_relationship: bool,
    omit_table_relationship: bool,
    table_relationship_type: &'static str,
    outbound: Option<OutboundOwner>,
    signed: bool,
}

impl Default for Shape {
    fn default() -> Self {
        Self {
            maps_content_type: "application/xml",
            single_content_type: SINGLE_CELLS_CONTENT_TYPE,
            table_content_type: TABLE_CONTENT_TYPE,
            omit_maps_relationship: false,
            maps_external: false,
            single_external: false,
            table_external: false,
            duplicate_single_relationship: false,
            omit_table_relationship: false,
            table_relationship_type: TABLE_REL,
            outbound: None,
            signed: false,
        }
    }
}

#[derive(Debug, PartialEq, Eq)]
struct PhysicalState {
    root_relationships: Vec<String>,
    parts: Vec<(String, String, Vec<u8>, Vec<String>)>,
}

#[test]
fn independent_vectors_have_the_normative_208_and_248_byte_shapes() {
    let single = single_cell_vector(&BindingWire::default(), 2);
    let table = normal_table_vector(
        &BindingWire {
            flags: 0,
            data_type: 13,
            ..BindingWire::default()
        },
        0,
    );

    assert_eq!(single.len(), 208);
    assert_eq!(table.len(), 248);
    assert_eq!(single, decode_hex(SINGLE_208_HEX));
    assert_eq!(table, decode_hex(TABLE_248_HEX));
    assert_eq!(&single[..3], &[0xd5, 0x02, 0x00]);
    assert_eq!(&single[3..6], &[0xd7, 0x02, 0x58]);
    assert_eq!(&single[152..155], &[0xdd, 0x02, 0x26]);
    assert_eq!(&single[205..], &[0xd6, 0x02, 0x00]);
    assert_eq!(&table[..3], &[0xd7, 0x02, 0x70]);
    assert_eq!(&table[195..198], &[0xdd, 0x02, 0x26]);
    assert_eq!(&table[245..], &[0xd8, 0x02, 0x00]);
}

#[test]
fn public_reader_projects_map_info_and_both_binary_binding_families() {
    let workbook = open(valid_package());
    let snapshot = workbook.xml_maps().expect("read public XML Maps snapshot");

    assert_eq!(snapshot.source_xml(), Some(MAP_INFO_XML));
    let info = snapshot.map_info().expect("fixture MapInfo");
    assert_eq!(info.selection_namespaces, "xmlns:e='urn:litchi:fixture'");
    assert_eq!(info.schemas.len(), 1);
    assert_eq!(info.schemas[0].id, "schema-7");
    assert_eq!(snapshot.maps().len(), 1);
    assert_eq!(snapshot.maps()[0].id, 7);
    assert_eq!(snapshot.maps()[0].name, "Orders");

    let mapped = snapshot.mapped_tables();
    assert_eq!(mapped.len(), 1);
    assert_eq!(mapped[0].table_id(), 2);
    let column = &mapped[0].columns()[0];
    assert_eq!(column.column_id(), 1);
    assert_eq!(column.map_id(), 7);
    assert_eq!(column.data_type().get(), 13);
    assert_eq!(column.xpath().as_str(), "/root/value");
    assert!(!column.can_be_single());

    assert_eq!(snapshot.single_cell_tables().len(), 1);
    let single = snapshot
        .single_cell_bindings(0)
        .expect("sheet has a Single Cell Tables part");
    assert_eq!(single.len(), 1);
    assert_eq!(single[0].table_id(), 1);
    assert_eq!(single[0].column_id(), 1);
    assert_eq!(single[0].cell().row(), 0);
    assert_eq!(single[0].cell().column(), 0);
    assert_eq!(single[0].map_id(), 7);
    assert_eq!(single[0].data_type().get(), 1);
    assert_eq!(single[0].xpath().as_str(), "/root/value");
    assert!(single[0].can_be_single());
    assert!(snapshot.single_cell_bindings(1).is_none());
}

#[test]
fn public_reader_rejects_malformed_ids_datatypes_xpaths_flags_and_record_order() {
    let relative = BindingWire {
        xpath: "root/value".to_owned(),
        ..BindingWire::default()
    };
    let axis = BindingWire {
        xpath: "/root/child::value".to_owned(),
        ..BindingWire::default()
    };
    let compound_predicate = BindingWire {
        xpath: "/root/item[@a='x' and @b='y']/value".to_owned(),
        ..BindingWire::default()
    };
    let empty_xpath = BindingWire {
        xpath: String::new(),
        ..BindingWire::default()
    };
    let oversized_xpath = BindingWire {
        xpath: format!("/{}", "a".repeat(31_999)),
        ..BindingWire::default()
    };
    let cases = vec![
        (
            "unknown map id",
            BindingWire {
                map_id: 999,
                ..BindingWire::default()
            },
            "unknown map ID",
        ),
        (
            "zero datatype",
            BindingWire {
                data_type: 0,
                ..BindingWire::default()
            },
            "XmlDataType",
        ),
        (
            "datatype above enum",
            BindingWire {
                data_type: 46,
                ..BindingWire::default()
            },
            "XmlDataType",
        ),
        ("relative XPath", relative, "XmlMappedXpath"),
        ("explicit XPath axis", axis, "XmlMappedXpath"),
        (
            "compound XPath predicate",
            compound_predicate,
            "XmlMappedXpath",
        ),
        ("empty XPath", empty_xpath, "XmlMappedXpath"),
        ("32000-unit XPath", oversized_xpath, "UTF-16 string length"),
        (
            "missing singleton flag",
            BindingWire {
                flags: 0,
                ..BindingWire::default()
            },
            "fCanBeSingle",
        ),
        (
            "misordered XML property delimiters",
            BindingWire {
                order: RecordOrder::EndColumnBeforeXml,
                ..BindingWire::default()
            },
            "record",
        ),
        (
            "missing XML property end",
            BindingWire {
                order: RecordOrder::MissingXmlEnd,
                ..BindingWire::default()
            },
            "record",
        ),
    ];

    for (label, wire, context) in cases {
        let package = package_with(
            MAP_INFO_XML.to_vec(),
            single_cell_vector(&wire, 2),
            normal_table_vector(
                &BindingWire {
                    flags: 0,
                    data_type: 13,
                    ..BindingWire::default()
                },
                0,
            ),
            Shape::default(),
        );
        assert_reader_error(label, package, context);
    }

    let duplicate_ids = String::from_utf8(MAP_INFO_XML.to_vec())
        .unwrap()
        .replace(
            "</MapInfo>",
            "<Map ID=\"7\" Name=\"Duplicate\" RootElement=\"root\" SchemaID=\"schema-7\" ShowImportExportValidationErrors=\"false\" AutoFit=\"false\" Append=\"false\" PreserveSortAFLayout=\"false\" PreserveFormat=\"false\"/></MapInfo>",
        )
        .into_bytes();
    assert_reader_error(
        "duplicate Map IDs",
        package_with(
            duplicate_ids,
            single_cell_vector(&BindingWire::default(), 2),
            normal_table_vector(
                &BindingWire {
                    flags: 0,
                    data_type: 13,
                    ..BindingWire::default()
                },
                0,
            ),
            Shape::default(),
        ),
        "duplicate Map ID",
    );

    let non_numeric_id = String::from_utf8(MAP_INFO_XML.to_vec())
        .unwrap()
        .replace("ID=\"7\" Name=\"Orders\"", "ID=\"NaN\" Name=\"Orders\"")
        .into_bytes();
    assert_reader_error(
        "non-numeric Map ID",
        package_with(
            non_numeric_id,
            single_cell_vector(&BindingWire::default(), 2),
            normal_table_vector(
                &BindingWire {
                    flags: 0,
                    data_type: 13,
                    ..BindingWire::default()
                },
                0,
            ),
            Shape::default(),
        ),
        "unsigned integer attribute",
    );
}

#[test]
fn public_reader_rejects_malformed_relationships_content_types_and_outbound_rels() {
    let wrong = "application/octet-stream";
    let mut cases = Vec::new();
    let mut base_cases = Vec::new();

    let mut shape = Shape {
        maps_content_type: wrong,
        ..Shape::default()
    };
    cases.push(("MapInfo content type", shape, "Invalid content type"));
    shape = Shape {
        single_content_type: wrong,
        ..Shape::default()
    };
    cases.push(("single-cell content type", shape, "Invalid content type"));
    shape = Shape {
        table_content_type: wrong,
        ..Shape::default()
    };
    cases.push(("table content type", shape, "Invalid content type"));
    shape = Shape {
        omit_maps_relationship: true,
        ..Shape::default()
    };
    cases.push((
        "missing MapInfo relationship",
        shape,
        "require a Custom XML Maps part",
    ));
    shape = Shape {
        maps_external: true,
        ..Shape::default()
    };
    cases.push((
        "external MapInfo relationship",
        shape,
        "Custom XML Maps relationship",
    ));
    shape = Shape {
        single_external: true,
        ..Shape::default()
    };
    cases.push((
        "external singleton relationship",
        shape,
        "tableSingleCells relationship",
    ));
    shape = Shape {
        table_external: true,
        ..Shape::default()
    };
    base_cases.push((
        "external table relationship",
        shape,
        "BrtListPart relationship",
    ));
    shape = Shape {
        duplicate_single_relationship: true,
        ..Shape::default()
    };
    cases.push((
        "two SCT relationships from one worksheet",
        shape,
        "tableSingleCells",
    ));
    shape = Shape {
        omit_table_relationship: true,
        ..Shape::default()
    };
    base_cases.push((
        "BrtListPart without relationship",
        shape,
        "BrtListPart relationship",
    ));
    shape = Shape {
        table_relationship_type: SINGLE_CELLS_REL,
        ..Shape::default()
    };
    cases.push((
        "BrtListPart with wrong relationship family",
        shape,
        "tableSingleCells",
    ));
    shape = Shape {
        outbound: Some(OutboundOwner::Maps),
        ..Shape::default()
    };
    cases.push((
        "MapInfo outbound relationship",
        shape,
        "Custom XML Maps part",
    ));
    shape = Shape {
        outbound: Some(OutboundOwner::SingleCells),
        ..Shape::default()
    };
    cases.push((
        "singleton outbound relationship",
        shape,
        "tableSingleCells part",
    ));
    for (label, shape, context) in cases {
        assert_reader_error(
            label,
            package_with(
                MAP_INFO_XML.to_vec(),
                single_cell_vector(
                    &BindingWire {
                        flags: u32::MAX,
                        ..BindingWire::default()
                    },
                    2,
                ),
                normal_table_vector(
                    &BindingWire {
                        flags: u32::MAX & !2,
                        data_type: 13,
                        ..BindingWire::default()
                    },
                    0,
                ),
                shape,
            ),
            context,
        );
    }
    for (label, shape, context) in base_cases {
        assert_base_open_error(
            label,
            package_with(
                MAP_INFO_XML.to_vec(),
                single_cell_vector(
                    &BindingWire {
                        flags: u32::MAX,
                        ..BindingWire::default()
                    },
                    2,
                ),
                normal_table_vector(
                    &BindingWire {
                        flags: u32::MAX & !2,
                        data_type: 13,
                        ..BindingWire::default()
                    },
                    0,
                ),
                shape,
            ),
            context,
        );
    }
}

#[test]
fn public_reader_enforces_exact_xml_binary_and_graph_limits() {
    let workbook = open(valid_package());

    let mut exact = ReadLimits::default();
    exact.xml_maps.max_part_bytes = MAP_INFO_XML.len();
    exact.bindings.max_part_bytes = 248;
    workbook
        .xml_maps_with_limits(exact)
        .expect("exact source limits are inclusive");

    let mut cases = Vec::new();
    let mut limits = ReadLimits::default();
    limits.xml_maps.max_part_bytes = MAP_INFO_XML.len() - 1;
    cases.push(("XML part bytes", limits));
    limits = ReadLimits::default();
    limits.bindings.max_part_bytes = 247;
    cases.push(("binary part bytes", limits));
    limits = ReadLimits::default();
    limits.bindings.max_xpath_units = 10;
    cases.push(("XPath units", limits));
    limits = ReadLimits::default();
    limits.bindings.max_bindings = 0;
    cases.push(("binding count", limits));
    limits = ReadLimits::default();
    limits.bindings.max_records = 7;
    cases.push(("record count", limits));
    limits = ReadLimits::default();
    limits.max_parts = 2;
    cases.push(("part count", limits));
    limits = ReadLimits::default();
    limits.max_relationships = 2;
    cases.push(("relationship count", limits));
    limits = ReadLimits::default();
    limits.max_total_bytes = MAP_INFO_XML.len();
    cases.push(("aggregate bytes", limits));

    for (label, limits) in cases {
        let error = workbook.xml_maps_with_limits(limits).expect_err(label);
        assert!(
            !error.to_string().is_empty(),
            "{label} returned an empty error"
        );
    }
}

#[test]
fn caller_xml_limits_are_inclusive_at_every_public_boundary() {
    let workbook = open(valid_package());
    assert_xml_limit_boundary(
        &workbook,
        MAP_INFO_XML.len(),
        |limits, value| {
            limits.xml_maps.max_part_bytes = value;
        },
        "part bytes",
    );
    assert_xml_limit_boundary(
        &workbook,
        2,
        |limits, value| {
            limits.xml_maps.max_schemas = value;
        },
        "schemas",
    );
    assert_xml_limit_boundary(
        &workbook,
        2,
        |limits, value| {
            limits.xml_maps.max_maps = value;
        },
        "maps",
    );
    assert_xml_limit_boundary(
        &workbook,
        MAP_INFO_XML.len(),
        |limits, value| {
            limits.xml_maps.max_string_bytes = value;
        },
        "string bytes",
    );
    assert_xml_limit_boundary(
        &workbook,
        MAP_INFO_XML.len(),
        |limits, value| {
            limits.xml_maps.max_opaque_bytes = value;
        },
        "opaque bytes",
    );
    assert_xml_limit_boundary(
        &workbook,
        32,
        |limits, value| {
            limits.xml_maps.max_depth = value;
        },
        "XML depth",
    );
    assert_xml_limit_boundary(
        &workbook,
        256,
        |limits, value| {
            limits.xml_maps.max_events = value;
        },
        "XML events",
    );

    let mut exact = ReadLimits::default();
    exact.max_total_bindings = 2;
    exact.max_total_xpath_units = 22;
    workbook.xml_maps_with_limits(exact).unwrap();
    let mut below = exact;
    below.max_total_bindings = 1;
    assert_error_contains(
        workbook.xml_maps_with_limits(below).unwrap_err(),
        "binding",
        "aggregate binding boundary",
    );
    below = exact;
    below.max_total_xpath_units = 21;
    assert_error_contains(
        workbook.xml_maps_with_limits(below).unwrap_err(),
        "XPath",
        "aggregate XPath boundary",
    );
}

#[test]
fn no_op_is_byte_exact_and_preserves_a_signature_origin() {
    let mut shape = Shape::default();
    shape.signed = true;
    let mut workbook = open(package_with(
        MAP_INFO_XML.to_vec(),
        single_cell_vector(
            &BindingWire {
                flags: u32::MAX,
                ..BindingWire::default()
            },
            2,
        ),
        normal_table_vector(
            &BindingWire {
                flags: u32::MAX & !2,
                data_type: 13,
                ..BindingWire::default()
            },
            0,
        ),
        shape,
    ));
    assert!(workbook.is_signed());
    let before = physical_state(&workbook);

    let commit = workbook
        .xml_maps()
        .unwrap()
        .edit()
        .commit()
        .expect("commit exact no-op");
    assert!(commit.patch().is_empty());
    workbook
        .apply_xml_maps(&commit)
        .expect("publish exact no-op");

    assert!(workbook.is_signed());
    assert_eq!(physical_state(&workbook), before);
    assert_eq!(
        workbook.xml_maps().unwrap().source_xml(),
        Some(MAP_INFO_XML)
    );
}

#[test]
fn patch_is_stale_checked_retryable_apply_once_and_exactly_invertible() {
    let mut workbook = open(valid_package());
    let before = physical_state(&workbook);
    let original_single = part_bytes(&workbook, SINGLE_CELLS_PART);
    let original_table = part_bytes(&workbook, TABLE_PART);

    let mut transaction = workbook.xml_maps().unwrap().edit();
    assert!(transaction.set_map_name(7, "Renamed").unwrap());
    let commit = transaction.commit().expect("build source-bound commit");
    assert!(!commit.patch().is_empty());
    let inverse = commit.patch().inverse();

    let mut stale = open(valid_package());
    stale
        .edit_opc(|package| {
            let uri = PackURI::new(MAP_INFO_PART)?;
            let mut changed = package.get_part(&uri)?.blob().to_vec();
            changed.extend_from_slice(b"\n");
            package.get_part_mut(&uri)?.set_blob(changed);
            Ok(())
        })
        .expect("make a still-valid but physically stale source");
    let stale_before = physical_state(&stale);
    stale
        .apply_xml_maps(&commit)
        .expect_err("stale source must reject publication");
    assert_eq!(physical_state(&stale), stale_before);

    workbook
        .apply_xml_maps(&commit)
        .expect("the same commit remains retryable on its real source");
    assert_eq!(workbook.xml_maps().unwrap().maps()[0].name, "Renamed");
    assert_eq!(part_bytes(&workbook, SINGLE_CELLS_PART), original_single);
    assert_eq!(part_bytes(&workbook, TABLE_PART), original_table);

    let changed = physical_state(&workbook);
    workbook
        .apply_xml_maps(&commit)
        .expect_err("a source-bound patch applies only once");
    assert_eq!(physical_state(&workbook), changed);

    workbook
        .apply_xml_maps_patch(&inverse)
        .expect("publish inverse patch");
    assert_eq!(physical_state(&workbook), before);
}

#[test]
fn a_real_change_unsigns_but_preserves_untouched_binary_parts() {
    let mut shape = Shape::default();
    shape.signed = true;
    let mut workbook = open(package_with(
        MAP_INFO_XML.to_vec(),
        single_cell_vector(
            &BindingWire {
                flags: u32::MAX,
                ..BindingWire::default()
            },
            2,
        ),
        normal_table_vector(
            &BindingWire {
                flags: u32::MAX & !2,
                data_type: 13,
                ..BindingWire::default()
            },
            0,
        ),
        shape,
    ));
    let single = part_bytes(&workbook, SINGLE_CELLS_PART);
    let table = part_bytes(&workbook, TABLE_PART);
    let mut transaction = workbook.xml_maps().unwrap().edit();
    transaction.set_map_name(7, "Changed").unwrap();
    let commit = transaction.commit().unwrap();

    workbook.apply_xml_maps(&commit).unwrap();

    assert!(!workbook.is_signed());
    assert!(
        !workbook
            .opc_package()
            .contains_part(&PackURI::new(SIGNATURE_ORIGIN_PART).unwrap())
    );
    assert_eq!(part_bytes(&workbook, SINGLE_CELLS_PART), single);
    assert_eq!(part_bytes(&workbook, TABLE_PART), table);
}

#[test]
fn public_writer_authors_both_families_and_reopens_through_the_reader() {
    let mut writer = WorkbookWriter::new();
    writer.set_xml_maps(semantic_map_info()).unwrap();

    let mut sheet = MutableWorksheet::new("Sheet1");
    sheet
        .add_table(Table {
            id: 2,
            display_name: Some("Mapped".to_owned()),
            range: Range {
                first_row: 0,
                last_row: 2,
                first_column: 0,
                last_column: 0,
            },
            table_type: Type::Xml,
            header_row_count: 1,
            columns: vec![Column {
                id: 1,
                name: Some("Value".to_owned()),
                ..Column::default()
            }],
            ..Table::default()
        })
        .unwrap();
    sheet
        .set_mapped_table(
            MappedTable::new(
                2,
                vec![
                    ColumnBinding::new(
                        1,
                        7,
                        XmlDataType::new(13).unwrap(),
                        XPath::new("/root/value").unwrap(),
                        false,
                    )
                    .unwrap(),
                ],
            )
            .unwrap(),
        )
        .unwrap();
    sheet
        .set_single_cell_mapping(
            SingleCellBinding::new(
                3,
                1,
                CellReference::new(0, 1).unwrap(),
                7,
                XmlDataType::new(1).unwrap(),
                XPath::new("/root/value").unwrap(),
            )
            .unwrap(),
        )
        .unwrap();
    writer.add_worksheet(sheet);

    let mut output = Cursor::new(Vec::new());
    writer.save(&mut output).expect("save authored XLSB");
    let reopened = Workbook::new(Cursor::new(output.into_inner())).expect("reopen authored XLSB");
    let snapshot = reopened.xml_maps().expect("read authored XML Maps");

    assert_eq!(snapshot.maps()[0].id, 7);
    assert_eq!(snapshot.mapped_tables().len(), 1);
    assert_eq!(snapshot.mapped_tables()[0].table_id(), 2);
    assert_eq!(snapshot.mapped_tables()[0].columns()[0].map_id(), 7);
    let single = snapshot
        .single_cell_bindings(0)
        .expect("writer emitted a Single Cell Tables part");
    assert_eq!(single.len(), 1);
    assert_eq!(single[0].table_id(), 3);
    assert_eq!(single[0].cell().column(), 1);
    assert_eq!(mapping_flags(&part_bytes(&reopened, TABLE_PART)), 0);
    assert_eq!(
        mapping_flags(&part_bytes(&reopened, "/xl/tables/tableSingleCells1.bin")),
        2
    );
}

#[test]
fn opaque_records_and_ignored_flags_are_preserved_and_block_owned_singleton_edits() {
    let mut workbook = open(opaque_package());
    let before = physical_state(&workbook);
    assert_eq!(
        mapping_flags(&part_bytes(&workbook, SINGLE_CELLS_PART)),
        u32::MAX
    );
    assert_eq!(
        mapping_flags(&part_bytes(&workbook, TABLE_PART)),
        u32::MAX & !2
    );

    let no_op = workbook.xml_maps().unwrap().edit().commit().unwrap();
    assert!(no_op.patch().is_empty());
    workbook.apply_xml_maps(&no_op).unwrap();
    assert_eq!(physical_state(&workbook), before);

    let single_before = part_bytes(&workbook, SINGLE_CELLS_PART);
    let table_before = part_bytes(&workbook, TABLE_PART);
    let mut rename = workbook.xml_maps().unwrap().edit();
    rename.set_map_name(7, "Renamed").unwrap();
    let renamed = rename.commit().unwrap();
    workbook.apply_xml_maps(&renamed).unwrap();
    assert_eq!(part_bytes(&workbook, SINGLE_CELLS_PART), single_before);
    assert_eq!(part_bytes(&workbook, TABLE_PART), table_before);
    let expected_xml = String::from_utf8(MAP_INFO_XML.to_vec())
        .unwrap()
        .replace("Name=\"Orders\"", "Name=\"Renamed\"")
        .into_bytes();
    assert_eq!(
        workbook.xml_maps().unwrap().source_xml(),
        Some(expected_xml.as_slice())
    );

    let mut owned = workbook.xml_maps().unwrap().edit();
    assert!(
        owned
            .put_single_cell_binding(
                0,
                SingleCellBinding::new(
                    1,
                    1,
                    CellReference::new(0, 0).unwrap(),
                    7,
                    XmlDataType::new(1).unwrap(),
                    XPath::new("/root/@code").unwrap(),
                )
                .unwrap(),
            )
            .unwrap()
    );
    assert_error_contains(
        owned.commit().unwrap_err(),
        "losslessly edit Single Cell Tables",
        "owned opaque singleton edit",
    );
}

#[test]
fn catalog_create_remove_dependency_edit_and_conformance_are_reversible() {
    let mut workbook = open(empty_package());
    let empty = physical_state(&workbook);
    assert!(workbook.xml_maps().unwrap().map_info().is_none());

    let mut create = workbook.xml_maps().unwrap().edit();
    assert!(create.set_catalog(semantic_map_info()).unwrap());
    assert!(create.set_conformance(XmlMapConformance::Strict).unwrap());
    let create = create.commit().unwrap();
    let uncreate = create.patch().inverse();
    workbook.apply_xml_maps(&create).unwrap();
    assert!(workbook.xml_maps().unwrap().conformance().is_strict());
    assert!(
        workbook
            .opc_package()
            .contains_part(&PackURI::new(MAP_INFO_PART).unwrap())
    );
    let created = physical_state(&workbook);
    workbook.apply_xml_maps_patch(&uncreate).unwrap();
    assert_eq!(physical_state(&workbook), empty);

    workbook.apply_xml_maps(&create).unwrap();
    let mut edit = workbook.xml_maps().unwrap().edit();
    assert!(
        edit.edit_map(7, |map| {
            map.name = "Edited".to_owned();
            Ok(())
        })
        .unwrap()
    );
    let edit = edit.commit().unwrap();
    let unedit = edit.patch().inverse();
    workbook.apply_xml_maps(&edit).unwrap();
    assert_eq!(workbook.xml_maps().unwrap().maps()[0].name, "Edited");
    workbook.apply_xml_maps_patch(&unedit).unwrap();
    assert_eq!(physical_state(&workbook), created);

    let mut remove = workbook.xml_maps().unwrap().edit();
    assert!(remove.remove_catalog().unwrap().is_some());
    let remove = remove.commit().unwrap();
    let restore = remove.patch().inverse();
    workbook.apply_xml_maps(&remove).unwrap();
    assert!(workbook.xml_maps().unwrap().map_info().is_none());
    assert!(
        !workbook
            .opc_package()
            .contains_part(&PackURI::new(MAP_INFO_PART).unwrap())
    );
    workbook.apply_xml_maps_patch(&restore).unwrap();
    assert_eq!(physical_state(&workbook), created);

    let bound = open(valid_package());
    let mut refused = bound.xml_maps().unwrap().edit();
    assert_error_contains(
        refused.remove_catalog().unwrap_err(),
        "binary bindings depend",
        "bound catalog removal",
    );
}

#[test]
fn public_binding_crud_missing_part_refusal_and_inverse_are_exact() {
    let mut workbook = open(valid_package());
    let before = physical_state(&workbook);
    let mut replace = workbook.xml_maps().unwrap().edit();
    assert!(
        replace
            .put_table_column_binding(
                2,
                ColumnBinding::new(
                    1,
                    7,
                    XmlDataType::new(16).unwrap(),
                    XPath::new("/root/@code").unwrap(),
                    true,
                )
                .unwrap(),
            )
            .unwrap()
    );
    assert!(
        replace
            .put_single_cell_binding(
                0,
                SingleCellBinding::new(
                    1,
                    1,
                    CellReference::new(0, 0).unwrap(),
                    7,
                    XmlDataType::new(16).unwrap(),
                    XPath::new("/root/@code").unwrap(),
                )
                .unwrap(),
            )
            .unwrap()
    );
    let replace = replace.commit().unwrap();
    let undo = replace.patch().inverse();
    workbook.apply_xml_maps(&replace).unwrap();
    assert_eq!(
        workbook.xml_maps().unwrap().mapped_tables()[0].columns()[0]
            .data_type()
            .get(),
        16
    );
    assert_eq!(
        workbook
            .xml_maps()
            .unwrap()
            .single_cell_bindings(0)
            .unwrap()[0]
            .xpath()
            .as_str(),
        "/root/@code"
    );
    workbook.apply_xml_maps_patch(&undo).unwrap();
    assert_eq!(physical_state(&workbook), before);

    let original_table = part_bytes(&workbook, TABLE_PART);
    let mut remove = workbook.xml_maps().unwrap().edit();
    assert!(remove.remove_table_column_binding(2, 1).unwrap().is_some());
    assert!(remove.remove_single_cell_binding(0, 1).unwrap().is_some());
    let remove = remove.commit().unwrap();
    let restore = remove.patch().inverse();
    workbook.apply_xml_maps(&remove).unwrap();
    assert!(workbook.xml_maps().unwrap().mapped_tables().is_empty());
    assert_eq!(
        workbook.xml_maps().unwrap().single_cell_bindings(0),
        Some(&[][..])
    );
    assert_eq!(
        part_bytes(&workbook, TABLE_PART),
        remove_first_record_pair(
            &original_table,
            BRT_BEGIN_LIST_XML_CPR,
            BRT_END_LIST_XML_CPR
        )
    );
    assert!(
        workbook
            .opc_package()
            .contains_part(&PackURI::new(TABLE_PART).unwrap())
    );
    assert!(
        workbook
            .opc_package()
            .contains_part(&PackURI::new(SINGLE_CELLS_PART).unwrap())
    );
    workbook.apply_xml_maps_patch(&restore).unwrap();
    assert_eq!(physical_state(&workbook), before);

    let empty = open(empty_package());
    let mut missing = empty.xml_maps().unwrap().edit();
    assert_error_contains(
        missing
            .put_single_cell_binding(
                0,
                SingleCellBinding::new(
                    1,
                    1,
                    CellReference::new(0, 0).unwrap(),
                    7,
                    XmlDataType::new(1).unwrap(),
                    XPath::new("/root/value").unwrap(),
                )
                .unwrap(),
            )
            .unwrap_err(),
        "creating a new tableSingleCells part",
        "missing singleton part",
    );
    assert_error_contains(
        missing
            .put_table_column_binding(
                99,
                ColumnBinding::new(
                    1,
                    7,
                    XmlDataType::new(1).unwrap(),
                    XPath::new("/root/value").unwrap(),
                    false,
                )
                .unwrap(),
            )
            .unwrap_err(),
        "mapped table ID 99 was not found",
        "missing mapped table",
    );

    let mut absent_column = open(valid_package()).xml_maps().unwrap().edit();
    let error = absent_column
        .put_table_column_binding(
            2,
            ColumnBinding::new(
                99,
                7,
                XmlDataType::new(1).unwrap(),
                XPath::new("/root/value").unwrap(),
                false,
            )
            .unwrap(),
        )
        .unwrap_err();
    assert_error_contains(
        error,
        "column 99 is absent from physical table ID 2",
        "absent physical column",
    );
}

#[test]
fn isolated_shared_sct_and_singleton_geometry_fail_at_the_intended_reader_invariant() {
    assert_reader_error(
        "two valid worksheets share one SCT target",
        package_with_shared_sct_across_two_worksheets(),
        "tableSingleCells part is shared by multiple worksheets",
    );

    assert_reader_error(
        "singleton cell overlaps an ordinary table",
        package_with(
            MAP_INFO_XML.to_vec(),
            set_begin_list_range(single_cell_vector(&BindingWire::default(), 2), [1, 1, 1, 1]),
            normal_table_vector(
                &BindingWire {
                    flags: 0,
                    data_type: 13,
                    ..BindingWire::default()
                },
                0,
            ),
            Shape::default(),
        ),
        "single-cell XML mapping 1 overlaps an ordinary table",
    );

    let mut auto_filter_overlap = package_with(
        MAP_INFO_XML.to_vec(),
        set_begin_list_range(single_cell_vector(&BindingWire::default(), 2), [1, 1, 0, 0]),
        normal_table_vector(
            &BindingWire {
                flags: 0,
                data_type: 13,
                ..BindingWire::default()
            },
            0,
        ),
        Shape::default(),
    );
    insert_auto_filter_before_table_refs(&mut auto_filter_overlap, [1, 1, 0, 0]);
    assert_reader_error(
        "singleton cell overlaps worksheet AutoFilter",
        auto_filter_overlap,
        "single-cell XML mapping 1 overlaps the worksheet AutoFilter",
    );
}

#[test]
fn hostile_graph_ownership_orphans_global_ids_geometry_and_conformance_are_rejected() {
    let mut non_worksheet = valid_package();
    let mut owner = BlobPart::new(
        PackURI::new("/custom/sct-owner.bin").unwrap(),
        "application/octet-stream".to_owned(),
        Vec::new(),
    );
    owner.rels_mut().add_relationship(
        SINGLE_CELLS_REL.to_owned(),
        "../xl/tables/singleCells1.bin".to_owned(),
        "rIdIllegalSct".to_owned(),
        false,
    );
    non_worksheet.add_part(Box::new(owner));
    assert_reader_error(
        "non-worksheet SCT owner",
        non_worksheet,
        "must originate from a worksheet",
    );

    let mut orphan_sct = valid_package();
    remove_relationship_of_type(&mut orphan_sct, WORKSHEET_PART, SINGLE_CELLS_REL);
    assert_reader_error(
        "orphan SCT part",
        orphan_sct,
        "tableSingleCells part is orphaned",
    );

    let mut missing_sct = valid_package();
    let _ = missing_sct.remove_part(&PackURI::new(SINGLE_CELLS_PART).unwrap());
    assert_reader_error("missing SCT target", missing_sct, "Part not found");

    let mut orphan_table_relationship = valid_package();
    let orphan_table = replace_utf16(
        set_begin_list_id(
            orphan_table_relationship
                .get_part(&PackURI::new(TABLE_PART).unwrap())
                .unwrap()
                .blob()
                .to_vec(),
            3,
        ),
        "MappedTable1",
        "OrphanTable1",
    );
    orphan_table_relationship.add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/tables/orphan.bin").unwrap(),
        "application/vnd.ms-excel.table".to_owned(),
        orphan_table,
    )));
    orphan_table_relationship
        .get_part_mut(&PackURI::new(WORKSHEET_PART).unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            TABLE_REL.to_owned(),
            "../tables/orphan.bin".to_owned(),
            "rIdOrphanTable".to_owned(),
            false,
        );
    assert_reader_error(
        "orphan ordinary-table relationship",
        orphan_table_relationship,
        "orphan table relationship",
    );

    let mut missing_table_target = valid_package();
    add_worksheet_relationship_record(&mut missing_table_target, "rIdTable1", "rIdTable2");
    missing_table_target
        .get_part_mut(&PackURI::new(WORKSHEET_PART).unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            TABLE_REL.to_owned(),
            "../tables/missing.bin".to_owned(),
            "rIdTable2".to_owned(),
            false,
        );
    assert_base_open_error(
        "missing ordinary-table target",
        missing_table_target,
        "Part not found",
    );

    let duplicate_id = package_with(
        MAP_INFO_XML.to_vec(),
        single_cell_vector(&BindingWire::default(), 2),
        set_begin_list_id(
            normal_table_vector(
                &BindingWire {
                    flags: 0,
                    data_type: 13,
                    ..BindingWire::default()
                },
                0,
            ),
            1,
        ),
        Shape::default(),
    );
    assert_reader_error(
        "global duplicate list ID",
        duplicate_id,
        "duplicate XML mapping table ID 1",
    );

    let mut outside_dimension = valid_package();
    replace_record_payload_in_part(
        &mut outside_dimension,
        WORKSHEET_PART,
        148,
        range_payload([0, 2, 0, 1]),
    );
    assert_reader_error(
        "table outside BrtWsDim",
        outside_dimension,
        "exceeds BrtWsDim",
    );

    let no_data_rows = package_with(
        MAP_INFO_XML.to_vec(),
        single_cell_vector(&BindingWire::default(), 2),
        set_begin_list_range(
            normal_table_vector(
                &BindingWire {
                    flags: 0,
                    data_type: 13,
                    ..BindingWire::default()
                },
                0,
            ),
            [1, 1, 1, 1],
        ),
        Shape::default(),
    );
    assert_reader_error("ordinary table without data rows", no_data_rows, "data row");

    assert_reader_error(
        "overlapping ordinary tables",
        package_with_second_table(3, [1, 3, 1, 1]),
        "overlap",
    );

    let mut auto_filter_overlap = valid_package();
    insert_auto_filter_before_table_refs(&mut auto_filter_overlap, [1, 3, 1, 1]);
    assert_reader_error(
        "worksheet AutoFilter overlaps table",
        auto_filter_overlap,
        "overlap",
    );

    let strict_root = String::from_utf8(MAP_INFO_XML.to_vec())
        .unwrap()
        .replace(
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            "http://purl.oclc.org/ooxml/spreadsheetml/main",
        )
        .into_bytes();
    assert_reader_error(
        "XML root and relationship conformance mismatch",
        package_with(
            strict_root,
            single_cell_vector(&BindingWire::default(), 2),
            normal_table_vector(
                &BindingWire {
                    flags: 0,
                    data_type: 13,
                    ..BindingWire::default()
                },
                0,
            ),
            Shape::default(),
        ),
        "conformance",
    );

    let dangling_connection = String::from_utf8(MAP_INFO_XML.to_vec())
        .unwrap()
        .replace(
            "FileBinding=\"false\" DataBindingLoadMode=\"1\"",
            "FileBinding=\"true\" ConnectionID=\"9\" DataBindingLoadMode=\"1\"",
        )
        .into_bytes();
    assert_reader_error(
        "dangling DataBinding ConnectionID",
        package_with(
            dangling_connection,
            single_cell_vector(&BindingWire::default(), 2),
            normal_table_vector(
                &BindingWire {
                    flags: 0,
                    data_type: 13,
                    ..BindingWire::default()
                },
                0,
            ),
            Shape::default(),
        ),
        "missing connection ID 9",
    );
}

#[test]
fn malformed_ordinary_table_vectors_fail_in_the_base_workbook_layer() {
    for (label, order, context) in [
        (
            "ordinary table closes its column before XmlCPr",
            RecordOrder::EndColumnBeforeXml,
            "BrtEndListCol",
        ),
        (
            "ordinary table omits BrtEndListXmlCPr",
            RecordOrder::MissingXmlEnd,
            "BrtBeginListCol collection",
        ),
    ] {
        assert_base_open_error(
            label,
            package_with(
                MAP_INFO_XML.to_vec(),
                single_cell_vector(&BindingWire::default(), 2),
                normal_table_vector(
                    &BindingWire {
                        flags: 0,
                        data_type: 13,
                        order,
                        ..BindingWire::default()
                    },
                    0,
                ),
                Shape::default(),
            ),
            context,
        );
    }
}

#[test]
fn ordinary_table_edits_preserve_unrelated_opacity_and_refuse_owned_block_opacity() {
    let unrelated = with_opaque_records(
        normal_table_vector(
            &BindingWire {
                flags: 0,
                data_type: 13,
                ..BindingWire::default()
            },
            0,
        ),
        BRT_END_LIST,
    );
    let mut workbook = open(package_with(
        MAP_INFO_XML.to_vec(),
        single_cell_vector(&BindingWire::default(), 2),
        unrelated,
        Shape::default(),
    ));
    let before = part_bytes(&workbook, TABLE_PART);
    let mut edit = workbook.xml_maps().unwrap().edit();
    edit.put_table_column_binding(
        2,
        ColumnBinding::new(
            1,
            7,
            XmlDataType::new(16).unwrap(),
            XPath::new("/root/@code").unwrap(),
            true,
        )
        .unwrap(),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    workbook.apply_xml_maps(&commit).unwrap();
    let after = part_bytes(&workbook, TABLE_PART);
    assert_eq!(
        remove_first_record_pair(&before, BRT_BEGIN_LIST_XML_CPR, BRT_END_LIST_XML_CPR),
        remove_first_record_pair(&after, BRT_BEGIN_LIST_XML_CPR, BRT_END_LIST_XML_CPR),
    );

    let owned_opaque = with_opaque_records_before(
        normal_table_vector(
            &BindingWire {
                flags: 0,
                data_type: 13,
                ..BindingWire::default()
            },
            0,
        ),
        BRT_END_LIST_XML_CPR,
    );
    let workbook = open(package_with(
        MAP_INFO_XML.to_vec(),
        single_cell_vector(&BindingWire::default(), 2),
        owned_opaque,
        Shape::default(),
    ));
    let mut edit = workbook.xml_maps().unwrap().edit();
    let error = edit
        .put_table_column_binding(
            2,
            ColumnBinding::new(
                1,
                7,
                XmlDataType::new(16).unwrap(),
                XPath::new("/root/@code").unwrap(),
                true,
            )
            .unwrap(),
        )
        .unwrap_err();
    assert_error_contains(
        error,
        "losslessly edit",
        "owned opaque ordinary-table XML property block",
    );
}

#[test]
fn additive_crud_aggregate_limits_and_signed_inverse_are_publicly_enforced() {
    let mut workbook = open(valid_package());
    let mut remove = workbook.xml_maps().unwrap().edit();
    remove.remove_table_column_binding(2, 1).unwrap();
    workbook.apply_xml_maps(&remove.commit().unwrap()).unwrap();
    assert!(workbook.xml_maps().unwrap().mapped_tables().is_empty());
    let mut add_table = workbook.xml_maps().unwrap().edit();
    add_table
        .put_table_column_binding(
            2,
            ColumnBinding::new(
                1,
                7,
                XmlDataType::new(13).unwrap(),
                XPath::new("/root/value").unwrap(),
                false,
            )
            .unwrap(),
        )
        .unwrap();
    workbook
        .apply_xml_maps(&add_table.commit().unwrap())
        .unwrap();
    assert_eq!(workbook.xml_maps().unwrap().mapped_tables().len(), 1);

    let mut add_single = workbook.xml_maps().unwrap().edit();
    add_single
        .put_single_cell_binding(
            0,
            SingleCellBinding::new(
                3,
                1,
                CellReference::new(1, 0).unwrap(),
                7,
                XmlDataType::new(1).unwrap(),
                XPath::new("/root/@code").unwrap(),
            )
            .unwrap(),
        )
        .unwrap();
    workbook
        .apply_xml_maps(&add_single.commit().unwrap())
        .unwrap();
    assert_eq!(
        workbook
            .xml_maps()
            .unwrap()
            .single_cell_bindings(0)
            .unwrap()
            .len(),
        2
    );

    let source = open(valid_package());
    let mut exact = ReadLimits::default();
    exact.max_total_bindings = 2;
    exact.max_total_xpath_units = 22;
    let snapshot = source.xml_maps_with_limits(exact).unwrap();
    assert!(snapshot.edit().commit().unwrap().patch().is_empty());
    let mut within = snapshot.edit();
    within.set_map_name(7, "WithinLimits").unwrap();
    within.commit().unwrap();

    let mut over_bindings = snapshot.edit();
    let error = over_bindings
        .put_single_cell_binding(
            0,
            SingleCellBinding::new(
                3,
                1,
                CellReference::new(1, 0).unwrap(),
                7,
                XmlDataType::new(1).unwrap(),
                XPath::new("/x").unwrap(),
            )
            .unwrap(),
        )
        .unwrap_err();
    assert_error_contains(error, "aggregate XML binding", "one-over binding mutation");
    over_bindings.set_map_name(7, "RetryBinding").unwrap();
    over_bindings.commit().unwrap();

    let mut xpath_limits = exact;
    xpath_limits.max_total_bindings = 3;
    let xpath_snapshot = source.xml_maps_with_limits(xpath_limits).unwrap();
    let mut over_xpath = xpath_snapshot.edit();
    let error = over_xpath
        .put_single_cell_binding(
            0,
            SingleCellBinding::new(
                3,
                1,
                CellReference::new(1, 0).unwrap(),
                7,
                XmlDataType::new(1).unwrap(),
                XPath::new("/x").unwrap(),
            )
            .unwrap(),
        )
        .unwrap_err();
    assert_error_contains(
        error,
        "aggregate XML binding XPath units",
        "one-over XPath mutation",
    );
    over_xpath.set_map_name(7, "RetryXPath").unwrap();
    over_xpath.commit().unwrap();

    let mut shape = Shape::default();
    shape.signed = true;
    let mut signed = open(package_with(
        MAP_INFO_XML.to_vec(),
        single_cell_vector(&BindingWire::default(), 2),
        normal_table_vector(
            &BindingWire {
                flags: 0,
                data_type: 13,
                ..BindingWire::default()
            },
            0,
        ),
        shape,
    ));
    let mut rename = signed.xml_maps().unwrap().edit();
    rename.set_map_name(7, "Unsigned").unwrap();
    let rename = rename.commit().unwrap();
    let inverse = rename.patch().inverse();
    signed.apply_xml_maps(&rename).unwrap();
    assert!(!signed.is_signed());
    signed.apply_xml_maps_patch(&inverse).unwrap();
    assert_eq!(signed.xml_maps().unwrap().maps()[0].name, "Orders");
    assert!(
        !signed.is_signed(),
        "inverse must not resurrect invalid signatures"
    );
}

#[test]
fn stale_content_type_relationship_and_topology_are_failure_atomic() {
    let source = open(valid_package());
    let mut transaction = source.xml_maps().unwrap().edit();
    transaction.set_map_name(7, "Staged").unwrap();
    let commit = transaction.commit().unwrap();

    let mut content_type = open(valid_package());
    content_type
        .edit_opc(|package| {
            let uri = PackURI::new(MAP_INFO_PART)?;
            let bytes = package.get_part(&uri)?.blob().to_vec();
            package.remove_part(&uri);
            package.add_part(Box::new(BlobPart::new(
                uri,
                "application/octet-stream".to_owned(),
                bytes,
            )));
            Ok(())
        })
        .unwrap();
    assert_failed_apply(&mut content_type, &commit, "content type");

    let mut relationship = open(valid_package());
    relationship
        .edit_opc(|package| {
            let uri = PackURI::new(WORKBOOK_PART)?;
            package.get_part_mut(&uri)?.rels_mut().remove("rIdXmlMaps");
            package.get_part_mut(&uri)?.rels_mut().add_relationship(
                XML_MAPS_REL.to_owned(),
                "xmlMaps.xml".to_owned(),
                "rIdXmlMapsChanged".to_owned(),
                false,
            );
            Ok(())
        })
        .unwrap();
    assert_failed_apply(&mut relationship, &commit, "stale");

    let mut topology = open(valid_package());
    topology
        .edit_opc(|package| {
            package.add_part(Box::new(BlobPart::new(
                PackURI::new("/custom/extra.bin")?,
                "application/octet-stream".to_owned(),
                vec![1, 2, 3],
            )));
            Ok(())
        })
        .unwrap();
    assert_failed_apply(&mut topology, &commit, "stale");
}

#[test]
fn catalog_removal_refuses_an_unowned_inbound_dependency() {
    let mut workbook = open(empty_package());
    let mut create = workbook.xml_maps().unwrap().edit();
    create.set_catalog(semantic_map_info()).unwrap();
    workbook.apply_xml_maps(&create.commit().unwrap()).unwrap();
    workbook
        .edit_opc(|package| {
            let mut owner = BlobPart::new(
                PackURI::new("/custom/owner.bin")?,
                "application/octet-stream".to_owned(),
                Vec::new(),
            );
            owner.rels_mut().add_relationship(
                "urn:litchi:fixture:dependency".to_owned(),
                "../xl/xmlMaps.xml".to_owned(),
                "rIdMapsDependency".to_owned(),
                false,
            );
            package.add_part(Box::new(owner));
            Ok(())
        })
        .unwrap();
    let mut remove = workbook.xml_maps().unwrap().edit();
    remove.remove_catalog().unwrap();
    let remove = remove.commit().unwrap();
    let before = physical_state(&workbook);
    let error = workbook.apply_xml_maps(&remove).unwrap_err();
    assert_error_contains(
        error,
        "another part references",
        "inbound catalog dependency",
    );
    assert_eq!(physical_state(&workbook), before);
}

fn semantic_map_info() -> XmlMapInfo {
    XmlMapInfo {
        selection_namespaces: "xmlns:e='urn:litchi:fixture'".to_owned(),
        schemas: vec![XmlSchema {
            id: "schema-7".to_owned(),
            schema_reference: Some("urn:litchi:fixture".to_owned()),
            namespace: Some("urn:litchi:fixture".to_owned()),
            payload_xml: Some(br#"<e:schema xmlns:e="urn:litchi:fixture"/>"#.to_vec()),
        }],
        maps: vec![XmlMap {
            id: 7,
            name: "Orders".to_owned(),
            root_element: "root".to_owned(),
            schema_id: "schema-7".to_owned(),
            show_import_export_validation_errors: true,
            auto_fit: true,
            append: false,
            preserve_sort_auto_filter_layout: true,
            preserve_format: true,
            data_binding: Some(DataBinding {
                data_binding_name: Some("inert".to_owned()),
                file_binding: Some(false),
                connection_id: None,
                file_binding_name: None,
                load_mode: 1,
                payload_xml: None,
            }),
        }],
    }
}

fn valid_package() -> OpcPackage {
    package_with(
        MAP_INFO_XML.to_vec(),
        single_cell_vector(&BindingWire::default(), 2),
        normal_table_vector(
            &BindingWire {
                flags: 0,
                data_type: 13,
                ..BindingWire::default()
            },
            0,
        ),
        Shape::default(),
    )
}

fn opaque_package() -> OpcPackage {
    package_with(
        MAP_INFO_XML.to_vec(),
        with_opaque_records(
            single_cell_vector(
                &BindingWire {
                    flags: u32::MAX,
                    ..BindingWire::default()
                },
                2,
            ),
            BRT_END_SINGLE_CELLS,
        ),
        with_opaque_records(
            normal_table_vector(
                &BindingWire {
                    flags: u32::MAX & !2,
                    data_type: 13,
                    ..BindingWire::default()
                },
                0,
            ),
            BRT_END_LIST,
        ),
        Shape::default(),
    )
}

fn empty_package() -> OpcPackage {
    let mut bundle_sheet = Vec::new();
    bundle_sheet.extend_from_slice(&0u32.to_le_bytes());
    bundle_sheet.extend_from_slice(&1u32.to_le_bytes());
    bundle_sheet.extend_from_slice(&wide_string("rIdSheet1"));
    bundle_sheet.extend_from_slice(&wide_string("Sheet1"));
    let mut workbook_bytes = Vec::new();
    push_record(&mut workbook_bytes, BRT_BUNDLE_SH, &bundle_sheet);
    let mut workbook = BlobPart::new(
        PackURI::new(WORKBOOK_PART).unwrap(),
        WORKBOOK_CONTENT_TYPE.to_owned(),
        workbook_bytes,
    );
    workbook.rels_mut().add_relationship(
        WORKSHEET_REL.to_owned(),
        "worksheets/sheet1.bin".to_owned(),
        "rIdSheet1".to_owned(),
        false,
    );
    let mut sheet_bytes = Vec::new();
    push_record(&mut sheet_bytes, BRT_BEGIN_SHEET, &[]);
    push_record(&mut sheet_bytes, BRT_BEGIN_SHEET_DATA, &[]);
    push_record(&mut sheet_bytes, BRT_END_SHEET_DATA, &[]);
    push_record(&mut sheet_bytes, BRT_END_SHEET, &[]);
    let sheet = BlobPart::new(
        PackURI::new(WORKSHEET_PART).unwrap(),
        WORKSHEET_CONTENT_TYPE.to_owned(),
        sheet_bytes,
    );
    let mut package = OpcPackage::new();
    package.rels_mut().add_relationship(
        OFFICE_DOCUMENT_REL.to_owned(),
        "xl/workbook.bin".to_owned(),
        "rIdOfficeDocument".to_owned(),
        false,
    );
    package.add_part(Box::new(workbook));
    package.add_part(Box::new(sheet));
    package
}

fn package_with(
    map_info: Vec<u8>,
    single_cells: Vec<u8>,
    table: Vec<u8>,
    shape: Shape,
) -> OpcPackage {
    let mut workbook_payload = Vec::new();
    let mut bundle_sheet = Vec::new();
    bundle_sheet.extend_from_slice(&0u32.to_le_bytes());
    bundle_sheet.extend_from_slice(&1u32.to_le_bytes());
    bundle_sheet.extend_from_slice(&wide_string("rIdSheet1"));
    bundle_sheet.extend_from_slice(&wide_string("Sheet1"));
    push_record(&mut workbook_payload, BRT_BUNDLE_SH, &bundle_sheet);

    let mut workbook = BlobPart::new(
        PackURI::new(WORKBOOK_PART).unwrap(),
        WORKBOOK_CONTENT_TYPE.to_owned(),
        workbook_payload,
    );
    workbook.rels_mut().add_relationship(
        WORKSHEET_REL.to_owned(),
        "worksheets/sheet1.bin".to_owned(),
        "rIdSheet1".to_owned(),
        false,
    );
    if !shape.omit_maps_relationship {
        workbook.rels_mut().add_relationship(
            XML_MAPS_REL.to_owned(),
            if shape.maps_external {
                "https://example.invalid/xmlMaps.xml".to_owned()
            } else {
                "xmlMaps.xml".to_owned()
            },
            "rIdXmlMaps".to_owned(),
            shape.maps_external,
        );
    }

    let mut worksheet = BlobPart::new(
        PackURI::new(WORKSHEET_PART).unwrap(),
        WORKSHEET_CONTENT_TYPE.to_owned(),
        worksheet_stream(),
    );
    worksheet.rels_mut().add_relationship(
        SINGLE_CELLS_REL.to_owned(),
        if shape.single_external {
            "https://example.invalid/singleCells.bin".to_owned()
        } else {
            "../tables/singleCells1.bin".to_owned()
        },
        "rIdSingleCells".to_owned(),
        shape.single_external,
    );
    if shape.duplicate_single_relationship {
        worksheet.rels_mut().add_relationship(
            SINGLE_CELLS_REL.to_owned(),
            "../tables/singleCells1.bin".to_owned(),
            "rIdSingleCellsDuplicate".to_owned(),
            false,
        );
    }
    if !shape.omit_table_relationship {
        worksheet.rels_mut().add_relationship(
            shape.table_relationship_type.to_owned(),
            if shape.table_external {
                "https://example.invalid/table1.bin".to_owned()
            } else {
                "../tables/table1.bin".to_owned()
            },
            "rIdTable1".to_owned(),
            shape.table_external,
        );
    }

    let mut maps = BlobPart::new(
        PackURI::new(MAP_INFO_PART).unwrap(),
        shape.maps_content_type.to_owned(),
        map_info,
    );
    let mut single = BlobPart::new(
        PackURI::new(SINGLE_CELLS_PART).unwrap(),
        shape.single_content_type.to_owned(),
        single_cells,
    );
    let table_part = BlobPart::new(
        PackURI::new(TABLE_PART).unwrap(),
        shape.table_content_type.to_owned(),
        table,
    );
    match shape.outbound {
        Some(OutboundOwner::Maps) => {
            maps.rels_mut().add_relationship(
                WORKSHEET_REL.to_owned(),
                "worksheets/sheet1.bin".to_owned(),
                "rIdOutbound".to_owned(),
                false,
            );
        },
        Some(OutboundOwner::SingleCells) => {
            single.rels_mut().add_relationship(
                WORKSHEET_REL.to_owned(),
                "../worksheets/sheet1.bin".to_owned(),
                "rIdOutbound".to_owned(),
                false,
            );
        },
        None => {},
    }

    let mut package = OpcPackage::new();
    package.rels_mut().add_relationship(
        OFFICE_DOCUMENT_REL.to_owned(),
        "xl/workbook.bin".to_owned(),
        "rIdOfficeDocument".to_owned(),
        false,
    );
    package.add_part(Box::new(workbook));
    package.add_part(Box::new(worksheet));
    package.add_part(Box::new(maps));
    package.add_part(Box::new(single));
    package.add_part(Box::new(table_part));
    if shape.signed {
        package.add_part(Box::new(BlobPart::new(
            PackURI::new(SIGNATURE_ORIGIN_PART).unwrap(),
            SIGNATURE_ORIGIN_CONTENT_TYPE.to_owned(),
            Vec::new(),
        )));
        package.rels_mut().add_relationship(
            SIGNATURE_ORIGIN_REL.to_owned(),
            "_xmlsignatures/origin.sigs".to_owned(),
            "rIdSignatureOrigin".to_owned(),
            false,
        );
    }
    package
}

fn worksheet_stream() -> Vec<u8> {
    let mut bytes = Vec::new();
    push_record(&mut bytes, BRT_BEGIN_SHEET, &[]);
    let mut dimension = Vec::new();
    for value in [0u32, 3, 0, 1] {
        dimension.extend_from_slice(&value.to_le_bytes());
    }
    push_record(&mut bytes, BRT_WS_DIM, &dimension);
    push_record(&mut bytes, BRT_BEGIN_SHEET_DATA, &[]);
    push_record(&mut bytes, BRT_END_SHEET_DATA, &[]);
    push_record(&mut bytes, BRT_BEGIN_LIST_PARTS, &1u32.to_le_bytes());
    push_record(&mut bytes, BRT_LIST_PART, &wide_string("rIdTable1"));
    push_record(&mut bytes, BRT_END_LIST_PARTS, &[]);
    push_record(&mut bytes, BRT_END_SHEET, &[]);
    bytes
}

fn single_cell_vector(binding: &BindingWire, list_flags: u32) -> Vec<u8> {
    let mut bytes = Vec::new();
    push_record(&mut bytes, BRT_BEGIN_SINGLE_CELLS, &[]);
    push_binding_list(&mut bytes, true, binding, list_flags);
    push_record(&mut bytes, BRT_END_SINGLE_CELLS, &[]);
    bytes
}

fn normal_table_vector(binding: &BindingWire, list_flags: u32) -> Vec<u8> {
    let mut bytes = Vec::new();
    push_binding_list(&mut bytes, false, binding, list_flags);
    bytes
}

fn push_binding_list(bytes: &mut Vec<u8>, single: bool, binding: &BindingWire, list_flags: u32) {
    push_record(bytes, BRT_BEGIN_LIST, &list_payload(single, list_flags));
    push_record(bytes, BRT_BEGIN_LIST_COLS, &1u32.to_le_bytes());
    push_record(bytes, BRT_BEGIN_LIST_COL, &column_payload(single));
    push_record(bytes, BRT_BEGIN_LIST_XML_CPR, &mapping_payload(binding));
    match binding.order {
        RecordOrder::Valid => {
            push_record(bytes, BRT_END_LIST_XML_CPR, &[]);
            push_record(bytes, BRT_END_LIST_COL, &[]);
        },
        RecordOrder::EndColumnBeforeXml => {
            push_record(bytes, BRT_END_LIST_COL, &[]);
            push_record(bytes, BRT_END_LIST_XML_CPR, &[]);
        },
        RecordOrder::MissingXmlEnd => push_record(bytes, BRT_END_LIST_COL, &[]),
    }
    push_record(bytes, BRT_END_LIST_COLS, &[]);
    push_record(bytes, BRT_END_LIST, &[]);
}

fn list_payload(single: bool, flags: u32) -> Vec<u8> {
    let mut payload = Vec::new();
    let range = if single {
        [0u32, 0, 0, 0]
    } else {
        [1u32, 3, 1, 1]
    };
    for value in range {
        payload.extend_from_slice(&value.to_le_bytes());
    }
    payload.extend_from_slice(&2u32.to_le_bytes());
    payload.extend_from_slice(&(if single { 1u32 } else { 2 }).to_le_bytes());
    payload.extend_from_slice(&(if single { 0u32 } else { 1 }).to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&flags.to_le_bytes());
    for _ in 0..6 {
        payload.extend_from_slice(&u32::MAX.to_le_bytes());
    }
    payload.extend_from_slice(&0u32.to_le_bytes());
    push_nullable(&mut payload, None);
    push_nullable(
        &mut payload,
        if single { None } else { Some("MappedTable1") },
    );
    for _ in 0..4 {
        push_nullable(&mut payload, None);
    }
    assert_eq!(payload.len(), if single { 88 } else { 112 });
    payload
}

fn column_payload(single: bool) -> Vec<u8> {
    let mut payload = Vec::new();
    payload.extend_from_slice(&1u32.to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
    for _ in 0..3 {
        payload.extend_from_slice(&u32::MAX.to_le_bytes());
    }
    payload.extend_from_slice(&0u32.to_le_bytes());
    push_nullable(&mut payload, None);
    push_nullable(
        &mut payload,
        if single { None } else { Some("MappedValue") },
    );
    for _ in 0..4 {
        push_nullable(&mut payload, None);
    }
    assert_eq!(payload.len(), if single { 48 } else { 70 });
    payload
}

fn mapping_payload(binding: &BindingWire) -> Vec<u8> {
    let mut payload = Vec::new();
    payload.extend_from_slice(&binding.map_id.to_le_bytes());
    payload.extend_from_slice(&binding.flags.to_le_bytes());
    payload.extend_from_slice(&binding.data_type.to_le_bytes());
    payload.extend_from_slice(&wide_string(&binding.xpath));
    payload
}

fn push_nullable(output: &mut Vec<u8>, value: Option<&str>) {
    match value {
        Some(value) => output.extend_from_slice(&wide_string(value)),
        None => output.extend_from_slice(&u32::MAX.to_le_bytes()),
    }
}

fn wide_string(value: &str) -> Vec<u8> {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(&(value.encode_utf16().count() as u32).to_le_bytes());
    for unit in value.encode_utf16() {
        bytes.extend_from_slice(&unit.to_le_bytes());
    }
    bytes
}

fn push_record(output: &mut Vec<u8>, kind: u16, payload: &[u8]) {
    if kind < 0x80 {
        output.push(kind as u8);
    } else {
        output.push(((kind & 0x7f) as u8) | 0x80);
        output.push((kind >> 7) as u8);
    }
    push_varint(output, payload.len());
    output.extend_from_slice(payload);
}

fn push_varint(output: &mut Vec<u8>, mut value: usize) {
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        output.push(byte);
        if value == 0 {
            return;
        }
    }
}

fn decode_hex(value: &str) -> Vec<u8> {
    assert_eq!(value.len() % 2, 0);
    value
        .as_bytes()
        .chunks_exact(2)
        .map(|pair| u8::from_str_radix(std::str::from_utf8(pair).unwrap(), 16).unwrap())
        .collect()
}

fn with_opaque_records(mut bytes: Vec<u8>, terminal: u16) -> Vec<u8> {
    bytes.truncate(bytes.len() - 3);
    push_record(&mut bytes, BRT_FIXTURE_UNKNOWN, &[0xde, 0xad]);
    push_record(&mut bytes, BRT_FRT_BEGIN, &[]);
    push_record(&mut bytes, BRT_FIXTURE_UNKNOWN + 1, &[0xbe, 0xef]);
    push_record(&mut bytes, BRT_FRT_END, &[]);
    push_record(&mut bytes, BRT_AC_BEGIN, &[]);
    push_record(&mut bytes, BRT_FIXTURE_UNKNOWN + 2, &[0xca, 0xfe]);
    push_record(&mut bytes, BRT_AC_END, &[]);
    push_record(&mut bytes, terminal, &[]);
    bytes
}

fn with_opaque_records_before(mut bytes: Vec<u8>, terminal: u16) -> Vec<u8> {
    let (_, start, _, _) = record_spans(&bytes)
        .into_iter()
        .find(|(kind, _, _, _)| *kind == terminal)
        .expect("fixture terminal record");
    let mut opaque = Vec::new();
    push_record(&mut opaque, BRT_FIXTURE_UNKNOWN, &[0xde, 0xad]);
    push_record(&mut opaque, BRT_FRT_BEGIN, &[]);
    push_record(&mut opaque, BRT_FIXTURE_UNKNOWN + 1, &[0xbe, 0xef]);
    push_record(&mut opaque, BRT_FRT_END, &[]);
    push_record(&mut opaque, BRT_AC_BEGIN, &[]);
    push_record(&mut opaque, BRT_FIXTURE_UNKNOWN + 2, &[0xca, 0xfe]);
    push_record(&mut opaque, BRT_AC_END, &[]);
    bytes.splice(start..start, opaque);
    bytes
}

fn record_spans(data: &[u8]) -> Vec<(u16, usize, usize, usize)> {
    let mut spans = Vec::new();
    let mut offset = 0usize;
    while offset < data.len() {
        let start = offset;
        let first = data[offset];
        offset += 1;
        let kind = if first & 0x80 == 0 {
            u16::from(first)
        } else {
            let second = data[offset];
            offset += 1;
            u16::from(first & 0x7f) | (u16::from(second) << 7)
        };
        let mut length = 0usize;
        let mut shift = 0usize;
        loop {
            let byte = data[offset];
            offset += 1;
            length |= usize::from(byte & 0x7f) << shift;
            if byte & 0x80 == 0 {
                break;
            }
            shift += 7;
        }
        let payload = offset;
        offset += length;
        spans.push((kind, start, payload, offset));
    }
    spans
}

fn mapping_flags(data: &[u8]) -> u32 {
    let (_, _, payload, end) = record_spans(data)
        .into_iter()
        .find(|(kind, _, _, _)| *kind == BRT_BEGIN_LIST_XML_CPR)
        .expect("fixture has BrtBeginListXmlCPr");
    assert!(end >= payload + 8);
    u32::from_le_bytes(data[payload + 4..payload + 8].try_into().unwrap())
}

fn remove_first_record_pair(data: &[u8], begin: u16, end: u16) -> Vec<u8> {
    let spans = record_spans(data);
    let first = spans.iter().position(|value| value.0 == begin).unwrap();
    let last = spans[first..]
        .iter()
        .position(|value| value.0 == end)
        .unwrap()
        + first;
    let mut result = data.to_vec();
    result.drain(spans[first].1..spans[last].3);
    result
}

fn package_with_shared_sct_across_two_worksheets() -> OpcPackage {
    let mut package = valid_package();

    let sheet1 = package
        .get_part(&PackURI::new(WORKSHEET_PART).unwrap())
        .unwrap()
        .blob();
    let mut sheet2 = Vec::new();
    for (kind, start, _, end) in record_spans(sheet1) {
        if !matches!(kind, 660..=662) {
            sheet2.extend_from_slice(&sheet1[start..end]);
        }
    }
    let mut sheet2_part = BlobPart::new(
        PackURI::new("/xl/worksheets/sheet2.bin").unwrap(),
        "application/vnd.ms-excel.worksheet".to_owned(),
        sheet2,
    );
    sheet2_part.rels_mut().add_relationship(
        SINGLE_CELLS_REL.to_owned(),
        "../tables/singleCells1.bin".to_owned(),
        "rIdSingleCells2".to_owned(),
        false,
    );
    package.add_part(Box::new(sheet2_part));

    let workbook_uri = PackURI::new(WORKBOOK_PART).unwrap();
    let workbook = package.get_part_mut(&workbook_uri).unwrap();
    let data = workbook.blob();
    let relationship = utf16_bytes("rIdSheet1");
    let offset = data
        .windows(relationship.len())
        .position(|window| window == relationship)
        .expect("workbook BundleSh relationship ID");
    let (_, start, payload, end) = record_spans(data)
        .into_iter()
        .find(|(_, start, _, end)| *start <= offset && offset < *end)
        .expect("BundleSh record span");
    let mut bundle_sheet = data[start..end].to_vec();
    bundle_sheet[payload - start + 4..payload - start + 8].copy_from_slice(&2u32.to_le_bytes());
    bundle_sheet = replace_utf16(bundle_sheet, "rIdSheet1", "rIdSheet2");
    bundle_sheet = replace_utf16(bundle_sheet, "Sheet1", "Sheet2");
    let mut workbook_data = data.to_vec();
    workbook_data.splice(end..end, bundle_sheet);
    workbook.set_blob(workbook_data);
    workbook.rels_mut().add_relationship(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet".to_owned(),
        "worksheets/sheet2.bin".to_owned(),
        "rIdSheet2".to_owned(),
        false,
    );
    package
}

fn remove_relationship_of_type(package: &mut OpcPackage, owner: &str, reltype: &str) {
    let owner = PackURI::new(owner).unwrap();
    let relationship_id = package
        .get_part(&owner)
        .unwrap()
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == reltype)
        .expect("fixture relationship")
        .r_id()
        .to_owned();
    package
        .get_part_mut(&owner)
        .unwrap()
        .rels_mut()
        .remove(&relationship_id);
}

fn range_payload(range: [u32; 4]) -> Vec<u8> {
    range.into_iter().flat_map(u32::to_le_bytes).collect()
}

fn set_begin_list_id(mut data: Vec<u8>, id: u32) -> Vec<u8> {
    let (_, _, payload, end) = record_spans(&data)
        .into_iter()
        .find(|(kind, _, _, _)| *kind == BRT_BEGIN_LIST)
        .expect("fixture BrtBeginList");
    assert!(end >= payload + 24);
    data[payload + 20..payload + 24].copy_from_slice(&id.to_le_bytes());
    data
}

fn set_begin_list_range(mut data: Vec<u8>, range: [u32; 4]) -> Vec<u8> {
    let (_, _, payload, end) = record_spans(&data)
        .into_iter()
        .find(|(kind, _, _, _)| *kind == BRT_BEGIN_LIST)
        .expect("fixture BrtBeginList");
    assert!(end >= payload + 16);
    data[payload..payload + 16].copy_from_slice(&range_payload(range));
    data
}

fn replace_record_payload_in_part(
    package: &mut OpcPackage,
    partname: &str,
    kind: u16,
    replacement: Vec<u8>,
) {
    let uri = PackURI::new(partname).unwrap();
    let part = package.get_part_mut(&uri).unwrap();
    let mut data = part.blob().to_vec();
    let (_, _, payload, end) = record_spans(&data)
        .into_iter()
        .find(|(record_kind, _, _, _)| *record_kind == kind)
        .expect("fixture record");
    assert_eq!(end - payload, replacement.len());
    data[payload..end].copy_from_slice(&replacement);
    part.set_blob(data);
}

fn utf16_bytes(value: &str) -> Vec<u8> {
    value.encode_utf16().flat_map(u16::to_le_bytes).collect()
}

fn replace_utf16(mut data: Vec<u8>, old: &str, new: &str) -> Vec<u8> {
    let old = utf16_bytes(old);
    let new = utf16_bytes(new);
    assert_eq!(old.len(), new.len());
    let offset = data
        .windows(old.len())
        .position(|window| window == old)
        .expect("fixture UTF-16 string");
    data[offset..offset + old.len()].copy_from_slice(&new);
    data
}

fn add_worksheet_relationship_record(package: &mut OpcPackage, old_id: &str, new_id: &str) {
    assert_eq!(old_id.encode_utf16().count(), new_id.encode_utf16().count());
    let uri = PackURI::new(WORKSHEET_PART).unwrap();
    let part = package.get_part_mut(&uri).unwrap();
    let data = part.blob();
    let old = utf16_bytes(old_id);
    let new = utf16_bytes(new_id);
    let offset = data
        .windows(old.len())
        .position(|window| window == old)
        .expect("worksheet relationship ID");
    let (_, start, _, end) = record_spans(data)
        .into_iter()
        .find(|(_, start, _, end)| *start <= offset && offset < *end)
        .expect("relationship record span");
    let mut duplicate = data[start..end].to_vec();
    let within = duplicate
        .windows(old.len())
        .position(|window| window == old)
        .unwrap();
    duplicate[within..within + old.len()].copy_from_slice(&new);
    let mut result = data.to_vec();
    result.splice(end..end, duplicate);
    let (_, _, count_payload, count_end) = record_spans(&result)
        .into_iter()
        .find(|(kind, _, _, _)| *kind == 660)
        .expect("BrtBeginListParts");
    assert_eq!(count_end - count_payload, 4);
    result[count_payload..count_end].copy_from_slice(&2u32.to_le_bytes());
    part.set_blob(result);
}

fn package_with_second_table(id: u32, range: [u32; 4]) -> OpcPackage {
    let mut package = valid_package();
    let source = package
        .get_part(&PackURI::new(TABLE_PART).unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let table = replace_utf16(
        set_begin_list_range(set_begin_list_id(source, id), range),
        "MappedTable1",
        "SecondTable1",
    );
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/xl/tables/table2.bin").unwrap(),
        "application/vnd.ms-excel.table".to_owned(),
        table,
    )));
    add_worksheet_relationship_record(&mut package, "rIdTable1", "rIdTable2");
    package
        .get_part_mut(&PackURI::new(WORKSHEET_PART).unwrap())
        .unwrap()
        .rels_mut()
        .add_relationship(
            TABLE_REL.to_owned(),
            "../tables/table2.bin".to_owned(),
            "rIdTable2".to_owned(),
            false,
        );
    package
}

fn insert_auto_filter_before_table_refs(package: &mut OpcPackage, range: [u32; 4]) {
    let uri = PackURI::new(WORKSHEET_PART).unwrap();
    let part = package.get_part_mut(&uri).unwrap();
    let data = part.blob();
    let needle = utf16_bytes("rIdTable1");
    let offset = data
        .windows(needle.len())
        .position(|window| window == needle)
        .expect("worksheet table relationship ID");
    let (_, start, _, _) = record_spans(data)
        .into_iter()
        .find(|(_, start, _, end)| *start <= offset && offset < *end)
        .expect("BrtListPart span");
    let mut records = Vec::new();
    push_record(&mut records, 161, &range_payload(range));
    push_record(&mut records, 162, &[]);
    let mut result = data.to_vec();
    result.splice(start..start, records);
    part.set_blob(result);
}

fn open(package: OpcPackage) -> Workbook {
    Workbook::from_opc_package(package).expect("open independent fixture")
}

fn assert_reader_error(label: &str, package: OpcPackage, context: &str) {
    let workbook = Workbook::from_opc_package(package)
        .unwrap_or_else(|error| panic!("{label} must pass base Workbook open: {error}"));
    let error = workbook
        .xml_maps()
        .expect_err("XML Maps reader must reject its malformed graph");
    assert_error_contains(error, context, label);
}

fn assert_base_open_error(label: &str, package: OpcPackage, context: &str) {
    let error = match Workbook::from_opc_package(package) {
        Ok(_) => panic!("{label}: ordinary Table graph must fail base Workbook open"),
        Err(error) => error,
    };
    assert_error_contains(error, context, label);
}

fn assert_error_contains(error: impl std::fmt::Display, context: &str, label: &str) {
    let error = error.to_string();
    assert!(
        error
            .to_ascii_lowercase()
            .contains(&context.to_ascii_lowercase()),
        "{label}: expected {context:?} context, found {error:?}"
    );
}

fn assert_xml_limit_boundary(
    workbook: &Workbook,
    upper: usize,
    set: impl Fn(&mut ReadLimits, usize),
    label: &str,
) {
    let mut low = 0usize;
    let mut high = upper;
    while low < high {
        let mid = low + (high - low) / 2;
        let mut limits = ReadLimits::default();
        set(&mut limits, mid);
        if workbook.xml_maps_with_limits(limits).is_ok() {
            high = mid;
        } else {
            low = mid + 1;
        }
    }
    assert!(low > 0, "{label} limit was not enforced");
    let mut exact = ReadLimits::default();
    set(&mut exact, low);
    workbook
        .xml_maps_with_limits(exact)
        .unwrap_or_else(|error| panic!("exact {label} boundary {low} failed: {error}"));
    let mut below = ReadLimits::default();
    set(&mut below, low - 1);
    assert!(
        workbook.xml_maps_with_limits(below).is_err(),
        "one below {label} boundary unexpectedly passed"
    );
}

fn assert_failed_apply(
    workbook: &mut Workbook,
    commit: &litchi_xlsb::xml_maps::Commit,
    context: &str,
) {
    let before = physical_state(workbook);
    let error = workbook.apply_xml_maps(commit).unwrap_err();
    assert_error_contains(error, context, "stale patch");
    assert_eq!(physical_state(workbook), before);
}

fn part_bytes(workbook: &Workbook, name: &str) -> Vec<u8> {
    workbook
        .opc_package()
        .get_part(&PackURI::new(name).unwrap())
        .unwrap()
        .blob()
        .to_vec()
}

fn physical_state(workbook: &Workbook) -> PhysicalState {
    let package = workbook.opc_package();
    let mut root_relationships = package
        .rels()
        .iter()
        .map(relationship_state)
        .collect::<Vec<_>>();
    root_relationships.sort();
    let mut parts = package
        .iter_parts()
        .map(|part| {
            let mut relationships = part
                .rels()
                .iter()
                .map(relationship_state)
                .collect::<Vec<_>>();
            relationships.sort();
            (
                part.partname().to_string(),
                part.content_type().to_owned(),
                part.blob().to_vec(),
                relationships,
            )
        })
        .collect::<Vec<_>>();
    parts.sort_by(|left, right| left.0.cmp(&right.0));
    PhysicalState {
        root_relationships,
        parts,
    }
}

fn relationship_state(relationship: &litchi_opc::Relationship) -> String {
    format!(
        "{}\0{}\0{}\0{}",
        relationship.r_id(),
        relationship.reltype(),
        relationship.target_ref(),
        relationship.is_external()
    )
}
