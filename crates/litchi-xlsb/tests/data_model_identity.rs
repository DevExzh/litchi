//! Source-bound Data Model table identity lifecycle coverage.
///
/// Positive cases use a complete synthetic MS-XLDM 140 storage stream built
/// from the same physical closure shape as the litchi-xldm identity tests.
/// The native date.xlsb fixture is intentionally used only for model-free
/// package/source-preservation checks; it is not presented as native model
/// evidence.
use std::collections::BTreeMap;
use std::io::Cursor;

use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, TargetMode};
use litchi_xlsb::Package;
use litchi_xlsb::data_model::{Definition, Model, ReadLimits, Relationship, Table};
use litchi_xlsb::package::connections::{Connection, Connections, SourceType};
use litchi_xlsb::writer::{MutableWorksheet, WorkbookWriter};

const XLDM_PAGE_SIZE: usize = 4096;
const CRC_SIZE: usize = 4;
const BOM: [u8; 2] = [0xFF, 0xFE];
const XLDM_STREAM_SIGNATURE: &str = "STREAM_STORAGE_SIGNATURE_)!@#$%^&*(";
const MODEL_URI: &str = "/xl/model/item.data";
const WORKBOOK_URI: &str = "/xl/workbook.bin";
const CONNECTIONS_URI: &str = "/xl/connections.bin";
const RELATION_METADATA_PATH: &str = "Model.1.db/T1.0.dim/R$T1$RelA.1.tbl.xml";
const CONNECTION_NAME: &str = "Model Connection";
const TABLE_ID: &str = "T1";
const OLD_TABLE_NAME: &str = "OldName";
const NEW_TABLE_NAME: &str = "Renamed";

fn utf16le(value: &str) -> Vec<u8> {
    value.encode_utf16().flat_map(u16::to_le_bytes).collect()
}

fn crc32(bytes: &[u8]) -> u32 {
    let mut crc = 0xFFFF_FFFFu32;
    for byte in bytes {
        let mut table = (((crc >> 24) ^ u32::from(*byte)) & 0xFF) << 24;
        for _ in 0..8 {
            table = if table & 0x8000_0000 != 0 {
                (table << 1) ^ 0x04C1_1DB7
            } else {
                table << 1
            };
        }
        crc = (crc << 8) ^ table;
    }
    crc
}

fn model_definition(name: &str, self_relationship: bool) -> Definition {
    Definition {
        min_version_load: 5,
        tables: vec![Table {
            id: TABLE_ID.to_owned(),
            name: name.to_owned(),
            connection: CONNECTION_NAME.to_owned(),
        }],
        relationships: self_relationship
            .then_some(vec![Relationship {
                from_table: name.to_owned(),
                from_column: "Key".to_owned(),
                to_table: name.to_owned(),
                to_column: "Key".to_owned(),
            }])
            .unwrap_or_default(),
        time_groupings: Vec::new(),
    }
}

fn connections() -> Connections {
    Connections {
        connections: vec![Connection {
            connection_id: 1,
            source_type: SourceType::Odbc,
            name: CONNECTION_NAME.to_owned(),
            ..Connection::default()
        }],
    }
}

fn package_with_model(definition: Definition, payload: Vec<u8>) -> Package {
    let mut writer = WorkbookWriter::new();
    writer.add_worksheet(MutableWorksheet::new("Sheet1"));
    writer.set_connections(connections()).expect("connections");
    writer
        .set_data_model(Model::from_bytes(definition, payload).expect("model"))
        .expect("set Data Model");
    let mut output = Cursor::new(Vec::new());
    writer.save(&mut output).expect("save");
    Package::from_bytes(output.into_inner()).expect("package")
}

fn complete_package(self_relationship: bool) -> Package {
    let payload = if self_relationship {
        complete_payload_named_with_relationship(OLD_TABLE_NAME)
    } else {
        complete_payload_named(OLD_TABLE_NAME)
    };
    package_with_model(model_definition(OLD_TABLE_NAME, self_relationship), payload)
}

fn complete_payload_named(name: &str) -> Vec<u8> {
    complete_payload_named_inner(name, false)
}

fn complete_payload_named_with_relationship(name: &str) -> Vec<u8> {
    complete_payload_named_inner(name, true)
}

fn complete_payload_named_inner(name: &str, with_relationship: bool) -> Vec<u8> {
    let source = if with_relationship {
        complete_distinct_id_storage_entries_with_relationship()
    } else {
        complete_distinct_id_storage_entries()
    };
    if name == OLD_TABLE_NAME {
        return build_identity_storage(&source);
    }
    let storage_bytes = build_identity_storage(&source);
    let storage = litchi_xldm::inspect_shared(&storage_bytes).expect("XLDM storage");
    let metadata = litchi_xldm::metadata::inspect(&storage).expect("metadata closure");
    let native = litchi_xldm::native::inspect(&storage, &metadata.native_parse_options())
        .expect("native closure");
    let generated =
        litchi_xldm::generated::inspect_system_generated(&storage).expect("generated closure");
    let olap = litchi_xldm::olap::inspect(&storage, &metadata).expect("OLAP closure");
    let closure =
        litchi_xldm::prove_xldm140_closure(&storage, &metadata, &olap, &native, &generated)
            .expect("complete identity closure");
    closure
        .rename_table_name_with_relationships(TABLE_ID, name)
        .expect("source fixture table rename")
        .after()
        .to_vec()
}

fn case_variant_endpoint_package() -> Package {
    package_with_model(
        Definition {
            min_version_load: 5,
            tables: vec![Table {
                id: TABLE_ID.to_owned(),
                name: "Sales".to_owned(),
                connection: CONNECTION_NAME.to_owned(),
            }],
            relationships: vec![Relationship {
                from_table: "sales".to_owned(),
                from_column: "Key".to_owned(),
                to_table: "SALES".to_owned(),
                to_column: "Key".to_owned(),
            }],
            time_groupings: Vec::new(),
        },
        complete_payload_named_with_relationship("Sales"),
    )
}

fn time_grouping(table_name: &str) -> litchi_xlsb::data_model::TimeGrouping {
    litchi_xlsb::data_model::TimeGrouping {
        table_name: table_name.to_owned(),
        column_name: "Key".to_owned(),
        column_id: "Key".to_owned(),
        columns: vec![litchi_xlsb::data_model::TimeGroupingColumn {
            is_selected: true,
            content_type: litchi_xlsb::data_model::ContentType::Years,
            column_name: "Year".to_owned(),
            column_id: "Year".to_owned(),
        }],
    }
}

fn duplicate_name_package() -> Package {
    package_with_model(
        Definition {
            min_version_load: 5,
            tables: vec![
                Table {
                    id: TABLE_ID.to_owned(),
                    name: OLD_TABLE_NAME.to_owned(),
                    connection: CONNECTION_NAME.to_owned(),
                },
                Table {
                    id: "T2".to_owned(),
                    name: "Other".to_owned(),
                    connection: CONNECTION_NAME.to_owned(),
                },
            ],
            relationships: Vec::new(),
            time_groupings: Vec::new(),
        },
        build_identity_storage(&complete_distinct_id_storage_entries()),
    )
}

fn opaque_package() -> Package {
    package_with_model(
        model_definition(OLD_TABLE_NAME, false),
        vec![0x01, 0x23, 0x45, 0x67],
    )
}

fn package_part_bytes(package: &Package, name: &str) -> Vec<u8> {
    package
        .opc_package()
        .get_part(&PackURI::new(name).expect("part URI"))
        .expect("part")
        .blob()
        .to_vec()
}

fn preserved_members(package: &Package) -> BTreeMap<String, Vec<u8>> {
    package
        .opc_package()
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
        .filter(|part| {
            !matches!(
                part.partname().as_str(),
                WORKBOOK_URI | CONNECTIONS_URI | MODEL_URI
            )
        })
        .map(|part| (part.partname().as_str().to_owned(), part.blob().to_vec()))
        .collect()
}

fn replace_utf16_once(bytes: &mut [u8], from: &str, to: &str) {
    let from = utf16le(from);
    let to = utf16le(to);
    assert_eq!(from.len(), to.len());
    let matches = bytes
        .windows(from.len())
        .enumerate()
        .filter_map(|(index, value)| (value == from.as_slice()).then_some(index))
        .collect::<Vec<_>>();
    assert_eq!(matches.len(), 1, "expected one UTF-16 marker");
    bytes[matches[0]..matches[0] + to.len()].copy_from_slice(&to);
}

fn replace_bytes_once(source: &[u8], from: &[u8], to: &[u8]) -> Vec<u8> {
    let matches = source
        .windows(from.len())
        .enumerate()
        .filter_map(|(index, value)| (value == from).then_some(index))
        .collect::<Vec<_>>();
    assert_eq!(matches.len(), 1, "expected one byte marker");
    let start = matches[0];
    let mut result = Vec::with_capacity(source.len() - from.len() + to.len());
    result.extend_from_slice(&source[..start]);
    result.extend_from_slice(to);
    result.extend_from_slice(&source[start + from.len()..]);
    result
}

fn package_with_corrupted_model_source(package: Package) -> Package {
    let mut opc = package.into_opc();
    let uri = PackURI::new(MODEL_URI).expect("model URI");
    let part = opc.get_part_mut(&uri).expect("model part");
    let mut bytes = part.blob().to_vec();
    bytes[0] ^= 0x80;
    part.set_blob(bytes);
    Package::from_opc(opc).expect("corrupted source remains an OPC-valid package")
}

fn authored_opc(source: OpcPackage) -> OpcPackage {
    let mut authored = OpcPackage::new();
    for relationship in source.rels().iter() {
        authored
            .rels_mut()
            .try_add_relationship(
                relationship.reltype().to_owned(),
                relationship.target_ref().to_owned(),
                relationship.r_id().to_owned(),
                relationship.target_mode(),
            )
            .expect("package relationship");
    }
    for part in source
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
    {
        let mut copy = BlobPart::new(
            part.partname().clone(),
            part.content_type().to_owned(),
            part.blob().to_vec(),
        );
        for relationship in part.rels().iter() {
            copy.rels_mut()
                .try_add_relationship(
                    relationship.reltype().to_owned(),
                    relationship.target_ref().to_owned(),
                    relationship.r_id().to_owned(),
                    relationship.target_mode(),
                )
                .expect("part relationship");
        }
        authored
            .try_add_part(Box::new(copy))
            .expect("authored part");
    }
    authored
}

fn add_signature_marker(package: &mut OpcPackage) {
    let origin = PackURI::new("/_xmlsignatures/origin.sigs").expect("signature origin");
    let signature = PackURI::new("/_xmlsignatures/sig1.xml").expect("signature");
    let mut origin_part = BlobPart::new(
        origin.clone(),
        litchi_opc::constants::content_type::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
        Vec::new(),
    );
    origin_part
        .rels_mut()
        .try_add_relationship(
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"
                .to_owned(),
            signature.relative_ref(origin.base_uri()),
            "rIdSignature".to_owned(),
            TargetMode::Internal,
        )
        .expect("signature relationship");
    package.add_part(Box::new(origin_part));
    package.add_part(Box::new(BlobPart::new(
        signature,
        litchi_opc::constants::content_type::OPC_DIGITAL_SIGNATURE_XMLSIGNATURE.to_owned(),
        b"<Signature/>".to_vec(),
    )));
    package
        .rels_mut()
        .try_add_relationship(
            litchi_opc::constants::relationship_type::DIGITAL_SIGNATURE_ORIGIN.to_owned(),
            origin.as_str().to_owned(),
            "rIdSignatureOrigin".to_owned(),
            TargetMode::Internal,
        )
        .expect("signature origin relationship");
}

fn signed_complete_package() -> Package {
    let mut opc = authored_opc(complete_package(false).into_opc());
    add_signature_marker(&mut opc);
    Package::from_opc(opc).expect("signed complete package")
}

fn assert_inner_table_name(package: &Package, expected: &str) {
    let part = package_part_bytes(package, MODEL_URI);
    let storage = litchi_xldm::inspect(&part).expect("XLDM storage");
    let metadata = litchi_xldm::metadata::inspect(&storage).expect("metadata closure");
    let table = metadata
        .files
        .iter()
        .find(|file| file.storage_path.ends_with("T1.1.tbl.xml"))
        .expect("T1 metadata file");
    assert_eq!(table.table.name.as_deref(), Some(expected));
}

fn inner_member_bytes(package: &Package) -> BTreeMap<String, Vec<u8>> {
    let part = package_part_bytes(package, MODEL_URI);
    let storage = litchi_xldm::inspect(&part).expect("XLDM storage");
    storage
        .files
        .iter()
        .enumerate()
        .map(|(index, file)| {
            (
                file.path.clone(),
                storage
                    .file_payload(index)
                    .expect("XLDM member payload")
                    .to_vec(),
            )
        })
        .collect()
}

fn column_data_fixture() -> Vec<u8> {
    let mut bytes = Vec::with_capacity(32);
    for value in [0x0102_0304_0506_0708_u64, 0_u64] {
        bytes.extend_from_slice(&1_u64.to_le_bytes());
        bytes.extend_from_slice(&value.to_le_bytes());
    }
    bytes
}

fn generated_mapping_fixture() -> Vec<u8> {
    let mut bytes = Vec::with_capacity(16);
    bytes.extend_from_slice(&1_u64.to_le_bytes());
    bytes.extend_from_slice(&0_u64.to_le_bytes());
    bytes
}

fn column_metadata_fixture() -> String {
    let mut metadata = r#"<XMObject class="XMSimpleTable" name="OldName"><Properties><Version>1</Version><Settings>0</Settings><RIViolationCount>0</RIViolationCount></Properties><Members><Member><Name>SegmentMap</Name><XMObject class="XMSegment1Map"><Properties><Records>1</Records></Properties></XMObject></Member><Member><Name>TableStats</Name><XMObject class="XMTableStats"><Properties><SegmentSize>1</SegmentSize><Usage>0</Usage></Properties></XMObject></Member></Members><Collections><Collection><Name>Partitions</Name></Collection><Collection><Name>Columns</Name><XMObject class="XMRawColumn" name="Key"><Properties><Settings>0</Settings><ColumnFlags>8</ColumnFlags><Collation></Collation><OrderByColumn></OrderByColumn><Locale>0</Locale><BinaryCharacters>0</BinaryCharacters></Properties><Members><Member><Name>IntrinsicHierarchy</Name><XMObject class="XMHierarchy"><Properties><SortOrder>0</SortOrder><IsProcessed>false</IsProcessed><TypeMaterialization>0</TypeMaterialization><ColumnPosition2DataID>-1</ColumnPosition2DataID><ColumnDataID2Position>-1</ColumnDataID2Position><DistinctDataIDs>1</DistinctDataIDs><TableStore>Key</TableStore></Properties></XMObject></Member><Member><Name>ColumnStats</Name><XMObject class="XMColumnStats"><Properties><DistinctStates>1</DistinctStates><MinDataID>0</MinDataID><MaxDataID>0</MaxDataID><OriginalMinSegmentDataID>0</OriginalMinSegmentDataID><RLESortOrder>-1</RLESortOrder><RowCount>1</RowCount><HasNulls>false</HasNulls><RLERuns>0</RLERuns><OthersRLERuns>0</OthersRLERuns><Usage>0</Usage><DBType>7</DBType><XMType>0</XMType><CompressionType>0</CompressionType><CompressionParam>0</CompressionParam><EncodingHint>0</EncodingHint><AggCounter>0</AggCounter><WhereCounter>0</WhereCounter><OrderByCounter>0</OrderByCounter></Properties></XMObject></Member></Members><Collections><Collection><Name>Segments</Name><XMObject class="XMColumnSegment"><Properties><Records>1</Records><Mask>0</Mask></Properties><Members><Member><Name>SubSegment</Name><XMObject class="XMColumnSegment"><Properties><Records>1</Records><Mask>0</Mask></Properties><Members><Member><Name>CompressionInfo</Name><XMObject class="XM123CompressionInfo"><Properties><Min>0</Min></Properties></XMObject></Member><Member><Name>ColumnSegmentStats</Name><XMObject class="XMColumnSegmentStats"><Properties><DistinctStates>1</DistinctStates><MinDataID>0</MinDataID><MaxDataID>0</MaxDataID><OriginalMinSegmentDataID>0</OriginalMinSegmentDataID><RLESortOrder>-1</RLESortOrder><RowCount>1</RowCount><HasNulls>false</HasNulls><RLERuns>0</RLERuns><OthersRLERuns>0</OthersRLERuns></Properties></XMObject></Member></Members></XMObject></Member><Member><Name>CompressionInfo</Name><XMObject class="XMHybridRLECompressionInfo&lt;class XM123CompressionInfo&gt;"><Members><Member><Name>RLECompression</Name><XMObject class="XMRLECompressionInfo"><Properties><BookmarkBits>0</BookmarkBits><StorageAllocSize>0</StorageAllocSize><StorageUsedSize>0</StorageUsedSize><SegmentNeedsResizing>false</SegmentNeedsResizing></Properties></XMObject></Member><Member><Name>SubCompression</Name><XMObject class="XM123CompressionInfo"><Properties><Min>0</Min></Properties></XMObject></Member></Members></XMObject></Member><Member><Name>ColumnSegmentStats</Name><XMObject class="XMColumnSegmentStats"><Properties><DistinctStates>1</DistinctStates><MinDataID>0</MinDataID><MaxDataID>0</MaxDataID><OriginalMinSegmentDataID>0</OriginalMinSegmentDataID><RLESortOrder>-1</RLESortOrder><RowCount>1</RowCount><HasNulls>false</HasNulls><RLERuns>0</RLERuns><OthersRLERuns>0</OthersRLERuns></Properties></XMObject></Member></Members></XMObject></Collection></Collections><DataObjects><DataObject><XMObject class="XMRawColumnPartitionDataObject" name="1.T1.Key.0.idf"><Properties><DataVersion>0</DataVersion><Partition>0</Partition><SegmentCount>1</SegmentCount></Properties></XMObject></DataObject><DataObject><XMObject class="XMValueDataDictionary&lt;XM_Long&gt;" name="1.T1.Key.dictionary"><Properties><DataVersion>0</DataVersion><BaseId>0</BaseId><Magnitude>0</Magnitude></Properties></XMObject></DataObject></DataObjects></XMObject></Collection><Collection><Name>Relationships</Name></Collection><Collection><Name>UserHierarchies</Name></Collection></Collections></XMObject>"#.to_owned();
    metadata = metadata
        .replace(
            "<IsProcessed>false</IsProcessed><TypeMaterialization>0</TypeMaterialization><ColumnPosition2DataID>-1</ColumnPosition2DataID>",
            "<IsProcessed>true</IsProcessed><TypeMaterialization>0</TypeMaterialization><ColumnPosition2DataID>0</ColumnPosition2DataID>",
        );
    let column_start = metadata.find("<XMObject class=\"XMRawColumn\"").unwrap();
    let column_end = metadata
        .find("</XMObject></Collection><Collection><Name>Relationships")
        .unwrap();
    let calculated = metadata[column_start..column_end + "</XMObject>".len()]
        .replace("name=\"Key\"", "name=\"Year\"")
        .replace("1.T1.Key.0.idf", "1.T1.Year.0.idf")
        .replace("1.T1.Key.dictionary", "1.T1.Year.dictionary")
        .replace("<Settings>0</Settings>", "<Settings>2</Settings>")
        .replace("<DBType>7</DBType>", "<DBType>20</DBType>")
        .replace(
            "<TableStore>Key</TableStore>",
            "<TableStore>Year</TableStore>",
        );
    metadata.insert_str(column_end + "</XMObject>".len(), &calculated);
    metadata
}

fn complete_distinct_id_storage_entries() -> Vec<(&'static str, Vec<u8>)> {
    complete_distinct_id_storage_entries_inner(false)
}

fn complete_distinct_id_storage_entries_with_relationship() -> Vec<(&'static str, Vec<u8>)> {
    complete_distinct_id_storage_entries_inner(true)
}

fn complete_distinct_id_storage_entries_inner(
    with_relationship: bool,
) -> Vec<(&'static str, Vec<u8>)> {
    let metadata = column_metadata_fixture().into_bytes();
    let mut entries = vec![
        ("Partitions", partitions_fixture().into_bytes()),
        ("Model.1.db.xml", database_definition_fixture().into_bytes()),
        (
            "Model.1.db/Source.1.ds.xml",
            datasource_definition_fixture().into_bytes(),
        ),
        (
            "Model.1.db/View.1.dsv.xml",
            datasource_view_definition_fixture().into_bytes(),
        ),
        (
            "Model.1.db/C.1.cub.xml",
            cube_definition_fixture().into_bytes(),
        ),
        (
            "Model.1.db/T1.1.dim.xml",
            dimension_definition_fixture(with_relationship).into_bytes(),
        ),
        (
            "Model.1.db/C.0.cub/MdxScript.0.scr.xml",
            mdx_script_definition_fixture().into_bytes(),
        ),
        (
            "Model.1.db/C.0.cub/T1.1.det.xml",
            measure_group_definition_fixture(with_relationship).into_bytes(),
        ),
        (
            "Model.1.db/C.0.cub/T1.0.det/T1.1.prt.xml",
            partition_definition_fixture().into_bytes(),
        ),
        ("Model.1.db/T1.0.dim/T1.1.tbl.xml", metadata),
        ("Model.1.db/T1.0.dim/1.T1.Key.0.idf", column_data_fixture()),
        ("Model.1.db/T1.0.dim/1.T1.Year.0.idf", column_data_fixture()),
        (
            "Model.1.db/T1.0.dim/1.H$T1$Key.POS_TO_ID.0.idf",
            generated_mapping_fixture(),
        ),
        (
            "Model.1.db/T1.0.dim/1.H$T1$Year.POS_TO_ID.0.idf",
            generated_mapping_fixture(),
        ),
    ];
    if with_relationship {
        entries.insert(
            10,
            (
                "Model.1.db/T1.0.dim/R$T1$RelA.1.tbl.xml",
                relationship_definition_fixture().into_bytes(),
            ),
        );
        entries.push((
            "Model.1.db/T1.0.dim/1.R$T1$RelA.INDEX.0.idf",
            column_data_fixture(),
        ));
    }
    let log = backup_log_fixture(&entries);
    entries.push(("BackupLog", log.into_bytes()));
    entries
}

fn build_identity_storage(entries: &[(&'static str, Vec<u8>)]) -> Vec<u8> {
    let mut bytes = vec![0; XLDM_PAGE_SIZE];
    bytes.extend_from_slice(&BOM);
    let mut allocations = Vec::new();
    for (index, (_, payload)) in entries.iter().enumerate() {
        if index + 1 == entries.len() {
            bytes.extend_from_slice(&BOM);
        }
        let offset = bytes.len();
        bytes.extend_from_slice(payload);
        bytes.extend_from_slice(&crc32(payload).to_le_bytes());
        allocations.push((offset, payload.len() + CRC_SIZE));
    }
    let directory_offset = bytes.len().div_ceil(XLDM_PAGE_SIZE) * XLDM_PAGE_SIZE;
    bytes.resize(directory_offset, 0);
    let mut directory = String::from("<VirtualDirectory>");
    for ((path, _), (offset, size)) in entries.iter().zip(&allocations) {
        directory.push_str(&format!(
                "<BackupFile><Path>{path}</Path><Size>{size}</Size><m_cbOffsetHeader>{offset}</m_cbOffsetHeader><Delete>false</Delete><CreatedTimestamp>0</CreatedTimestamp><Access>0</Access><LastWriteTime>0</LastWriteTime></BackupFile>"
            ));
    }
    directory.push_str("</VirtualDirectory>");
    let directory_bytes = utf16le(&directory);
    bytes.extend_from_slice(&directory_bytes);
    bytes.resize(bytes.len().div_ceil(XLDM_PAGE_SIZE) * XLDM_PAGE_SIZE, 0);
    let header = format!(
        "<BackupLog><BackupRestoreSyncVersion>140</BackupRestoreSyncVersion><Fault>false</Fault><faultcode>0</faultcode><ErrorCode>true</ErrorCode><EncryptionFlag>false</EncryptionFlag><EncryptionKey>0</EncryptionKey><ApplyCompression>true</ApplyCompression><m_cbOffsetHeader>{directory_offset}</m_cbOffsetHeader><DataSize>{}</DataSize><Files>{}</Files><ObjectID>11111111-2222-3333-4444-555555555500</ObjectID><m_cbOffsetData>4096</m_cbOffsetData></BackupLog>",
        directory_bytes.len(),
        entries.len()
    );
    let mut page = Vec::new();
    page.extend_from_slice(&BOM);
    page.extend_from_slice(&utf16le(XLDM_STREAM_SIGNATURE));
    page.extend_from_slice(&utf16le(&header));
    page.resize(XLDM_PAGE_SIZE, 0);
    bytes[..XLDM_PAGE_SIZE].copy_from_slice(&page);
    bytes
}

fn backup_log_fixture(entries: &[(&'static str, Vec<u8>)]) -> String {
    let with_relationship = entries
        .iter()
        .any(|(path, _)| *path == "Model.1.db/T1.0.dim/R$T1$RelA.1.tbl.xml");
    let dimension_paths: &[&str] = if with_relationship {
        &[
            "Model.1.db/T1.1.dim.xml",
            "Model.1.db/T1.0.dim/T1.1.tbl.xml",
            "Model.1.db/T1.0.dim/R$T1$RelA.1.tbl.xml",
            "Model.1.db/T1.0.dim/1.T1.Key.0.idf",
            "Model.1.db/T1.0.dim/1.T1.Year.0.idf",
            "Model.1.db/T1.0.dim/1.H$T1$Key.POS_TO_ID.0.idf",
            "Model.1.db/T1.0.dim/1.H$T1$Year.POS_TO_ID.0.idf",
            "Model.1.db/T1.0.dim/1.R$T1$RelA.INDEX.0.idf",
        ]
    } else {
        &[
            "Model.1.db/T1.1.dim.xml",
            "Model.1.db/T1.0.dim/T1.1.tbl.xml",
            "Model.1.db/T1.0.dim/1.T1.Key.0.idf",
            "Model.1.db/T1.0.dim/1.T1.Year.0.idf",
            "Model.1.db/T1.0.dim/1.H$T1$Key.POS_TO_ID.0.idf",
            "Model.1.db/T1.0.dim/1.H$T1$Year.POS_TO_ID.0.idf",
        ]
    };
    let groups: [(&str, i32, &str, &str, &[&str]); 8] = [
        (
            "100002",
            100_002,
            "DB",
            "11111111-2222-3333-4444-555555555551",
            &["Model.1.db.xml"],
        ),
        (
            "100003",
            100_003,
            "DS",
            "11111111-2222-3333-4444-555555555552",
            &["Model.1.db/Source.1.ds.xml"],
        ),
        (
            "100053",
            100_053,
            "DSV",
            "11111111-2222-3333-4444-555555555553",
            &["Model.1.db/View.1.dsv.xml"],
        ),
        (
            "100010",
            100_010,
            "CUBE",
            "11111111-2222-3333-4444-555555555554",
            &["Model.1.db/C.1.cub.xml"],
        ),
        (
            "100006",
            100_006,
            "DIM",
            "11111111-2222-3333-4444-555555555555",
            dimension_paths,
        ),
        (
            "100060",
            100_060,
            "SCRIPT",
            "11111111-2222-3333-4444-555555555556",
            &["Model.1.db/C.0.cub/MdxScript.0.scr.xml"],
        ),
        (
            "100016",
            100_016,
            "MG",
            "11111111-2222-3333-4444-555555555557",
            &["Model.1.db/C.0.cub/T1.1.det.xml"],
        ),
        (
            "100021",
            100_021,
            "PART",
            "11111111-2222-3333-4444-555555555558",
            &["Model.1.db/C.0.cub/T1.0.det/T1.1.prt.xml"],
        ),
    ];
    let mut result = String::from(
        "<BackupLog><BackupRestoreSyncVersion>1153</BackupRestoreSyncVersion><ServerRoot>C:\\inert</ServerRoot><SvrEncryptPwdFlag>true</SvrEncryptPwdFlag><ServerEnableBinaryXML>false</ServerEnableBinaryXML><ServerEnableCompression>false</ServerEnableCompression><CompressionFlag>false</CompressionFlag><EncryptionFlag>false</EncryptionFlag><ObjectName>Model</ObjectName><ObjectId>11111111-2222-3333-4444-555555555500</ObjectId><Write>ReadWrite</Write><OlapInfo>true</OlapInfo><Collations><Collation>Latin1_General</Collation></Collations><Languages><Language>1033</Language></Languages><FileGroups>",
    );
    for (_, class, name, object_id, paths) in groups {
        let persist = "Model.1.db";
        let version = if name == "SCRIPT" { 0 } else { 1 };
        let persist_location = 1;
        let object_name = match name {
            "DB" => "Database",
            "DS" => "Source",
            "DSV" => "View",
            "CUBE" => "Cube",
            "DIM" => "T1",
            "SCRIPT" => "MdxScript",
            "MG" | "PART" => "T1",
            _ => name,
        };
        result.push_str(&format!(
                "<FileGroup><Class>{class}</Class><ID>{object_id}</ID><Name>{object_name}</Name><ObjectVersion>{version}</ObjectVersion><PersistLocation>{persist_location}</PersistLocation><PersistLocationPath>{persist}</PersistLocationPath><StorageLocationPath></StorageLocationPath><ObjectID>{object_id}</ObjectID><FileList>"
            ));
        for path in paths {
            let size = entries
                .iter()
                .find(|(entry_path, _)| *entry_path == *path)
                .map_or(0, |(_, payload)| payload.len());
            result.push_str(&format!(
                    "<BackupFile><Path>C:\\inert\\{path}</Path><StoragePath>{path}</StoragePath><LastWriteTime>0</LastWriteTime><Size>{size}</Size></BackupFile>"
                ));
        }
        result.push_str("</FileList></FileGroup>");
    }
    result.push_str("</FileGroups></BackupLog>");
    result
}

fn partitions_fixture() -> String {
    "<Partitions><Partition><ObjectPath></ObjectPath><Name></Name><DataSize>0</DataSize><Location></Location><DataSourceID></DataSourceID><ConnectionString></ConnectionString></Partition></Partitions>".into()
}

fn olap_definition_fixture(
    kind: &str,
    path_id: &str,
    id: &str,
    parent: &str,
    cube: Option<&str>,
    prefix: &str,
    data_files: &str,
    extras: &str,
) -> String {
    let object_id = match id {
        "11111111-2222-3333-4444-555555555551"
        | "11111111-2222-3333-4444-555555555552"
        | "11111111-2222-3333-4444-555555555553"
        | "11111111-2222-3333-4444-555555555554"
        | "11111111-2222-3333-4444-555555555555"
        | "11111111-2222-3333-4444-555555555556"
        | "11111111-2222-3333-4444-555555555557"
        | "11111111-2222-3333-4444-555555555558" => id,
        _ => id,
    };
    let parent = match cube {
        Some(cube) => format!(
            "<ParentObject><DatabaseID>{parent}</DatabaseID><CubeID>{cube}</CubeID></ParentObject>"
        ),
        None if parent.is_empty() => "<ParentObject></ParentObject>".into(),
        None => format!("<ParentObject><DatabaseID>{parent}</DatabaseID></ParentObject>"),
    };
    format!(
        "<Load>{parent}<ObjectDefinition><{kind}><ID>{id}</ID><ObjectID>{object_id}</ObjectID><Name>{path_id}</Name>{prefix}<Ordinal>0</Ordinal><ObjectVersion>1</ObjectVersion><PersistLocation>0</PersistLocation><System>false</System><DataFileList>{data_files}</DataFileList>{extras}</{kind}></ObjectDefinition></Load>"
    )
}

fn database_definition_fixture() -> String {
    olap_definition_fixture(
        "Database",
        "Database",
        "11111111-2222-3333-4444-555555555551",
        "",
        None,
        "<DbStorageLocation>Model.1.db</DbStorageLocation>",
        "",
        "",
    )
    .replace(
        "<PersistLocation>0</PersistLocation>",
        "<PersistLocation>1</PersistLocation>",
    )
}

fn datasource_definition_fixture() -> String {
    olap_definition_fixture(
        "DataSource",
        "Source",
        "11111111-2222-3333-4444-555555555552",
        "11111111-2222-3333-4444-555555555551",
        None,
        "",
        "",
        "<PermissionFileList></PermissionFileList>",
    )
}

fn datasource_view_definition_fixture() -> String {
    olap_definition_fixture(
        "DataSourceView",
        "View",
        "11111111-2222-3333-4444-555555555553",
        "11111111-2222-3333-4444-555555555551",
        None,
        "",
        "",
        "",
    )
}

fn cube_definition_fixture() -> String {
    olap_definition_fixture(
        "Cube",
        "Cube",
        "11111111-2222-3333-4444-555555555554",
        "11111111-2222-3333-4444-555555555551",
        None,
        "<Dimensions><Dimension xsi:type=\"CubeDimension\"><DimensionID>11111111-2222-3333-4444-555555555555</DimensionID><Attributes><Attribute xsi:type=\"CubeAttribute\"><AttributeID>Key</AttributeID></Attribute><Attribute xsi:type=\"CubeAttribute\"><AttributeID>Year</AttributeID></Attribute></Attributes></Dimension></Dimensions>",
        "",
        "<PermissionFileList></PermissionFileList><MeasureGroupFileList>T1.1.det.xml</MeasureGroupFileList><PerspectiveFileList></PerspectiveFileList><AssemblyFileList></AssemblyFileList>",
    )
}

fn dimension_definition_fixture(with_relationship: bool) -> String {
    let relationships = if with_relationship {
        "<Relationships><Relationship><ID>RelA</ID></Relationship></Relationships>"
    } else {
        ""
    };
    let data_files = if with_relationship {
        "1.T1.Key.0.idf;1.T1.Year.0.idf;1.H$T1$Key.POS_TO_ID.0.idf;1.H$T1$Year.POS_TO_ID.0.idf;1.R$T1$RelA.INDEX.0.idf"
    } else {
        "1.T1.Key.0.idf;1.T1.Year.0.idf;1.H$T1$Key.POS_TO_ID.0.idf;1.H$T1$Year.POS_TO_ID.0.idf"
    };
    olap_definition_fixture(
        "Dimension",
        "T1",
        "11111111-2222-3333-4444-555555555555",
        "11111111-2222-3333-4444-555555555551",
        None,
        &format!(
            "<Attributes><Attribute xsi:type=\"DimensionAttribute\"><ID>Key</ID></Attribute><Attribute xsi:type=\"DimensionAttribute\"><ID>Year</ID></Attribute></Attributes>{relationships}"
        ),
        data_files,
        "<PermissionFileList></PermissionFileList>",
    )
}

fn relationship_definition_fixture() -> String {
    let metadata = column_metadata_fixture();
    let column_start = metadata.find("<XMObject class=\"XMRawColumn\"").unwrap();
    let column_end = metadata
        .find("</XMObject><XMObject class=\"XMRawColumn\"")
        .unwrap();
    let mut column = metadata[column_start..column_end + "</XMObject>".len()].to_owned();
    column = column.replacen("name=\"Key\"", "name=\"RelA\"", 1);
    column = column.replacen("<Settings>0</Settings>", "<Settings>5</Settings>", 1);
    format!(
        "<XMObject class=\"XMSimpleTable\" name=\"OldName\"><Properties><Version>1</Version><Settings>0</Settings><RIViolationCount>0</RIViolationCount></Properties><Members><Member><Name>SegmentMap</Name><XMObject class=\"XMSegment1Map\"><Properties><Records>1</Records></Properties></XMObject></Member><Member><Name>TableStats</Name><XMObject class=\"XMTableStats\"><Properties><SegmentSize>1</SegmentSize><Usage>0</Usage></Properties></XMObject></Member></Members><Collections><Collection><Name>Partitions</Name></Collection><Collection><Name>Columns</Name>{column}</Collection><Collection><Name>Relationships</Name><XMObject class=\"XMRelationship\" name=\"RelA\"><Properties><PrimaryTable>OldName</PrimaryTable><PrimaryColumn>Key</PrimaryColumn><ForeignColumn>Key</ForeignColumn></Properties><DataObjects><DataObject><XMObject class=\"XMRelationshipIndex123DIDs\"/></DataObject></DataObjects></XMObject></Collection><Collection><Name>UserHierarchies</Name></Collection></Collections></XMObject>"
    )
}

fn mdx_script_definition_fixture() -> String {
    olap_definition_fixture(
        "MdxScript",
        "MdxScript",
        "11111111-2222-3333-4444-555555555556",
        "11111111-2222-3333-4444-555555555551",
        Some("11111111-2222-3333-4444-555555555554"),
        "",
        "",
        "",
    )
    .replace(
        "<ObjectVersion>1</ObjectVersion>",
        "<ObjectVersion>0</ObjectVersion>",
    )
}

fn measure_group_definition_fixture(with_relationship: bool) -> String {
    let reference_dimension = if with_relationship {
        "<Dimension xsi:type=\"ReferenceMeasureGroupDimension\"><CubeDimensionID>11111111-2222-3333-4444-555555555555</CubeDimensionID><Attributes><Attribute xsi:type=\"MeasureGroupDimensionAttribute\"><AttributeID>Key</AttributeID></Attribute><Attribute xsi:type=\"MeasureGroupDimensionAttribute\"><AttributeID>Year</AttributeID></Attribute></Attributes></Dimension>"
    } else {
        ""
    };
    olap_definition_fixture(
        "MeasureGroup",
        "T1",
        "11111111-2222-3333-4444-555555555557",
        "11111111-2222-3333-4444-555555555551",
        Some("11111111-2222-3333-4444-555555555554"),
        &format!(
            "<Dimensions><Dimension xsi:type=\"DegenerateMeasureGroupDimension\"><CubeDimensionID>11111111-2222-3333-4444-555555555555</CubeDimensionID><Attributes><Attribute xsi:type=\"MeasureGroupDimensionAttribute\"><AttributeID>Key</AttributeID></Attribute><Attribute xsi:type=\"MeasureGroupDimensionAttribute\"><AttributeID>Year</AttributeID></Attribute></Attributes></Dimension>{reference_dimension}</Dimensions>"
        ),
        "",
        "<AggregationDesignFileList></AggregationDesignFileList><PartitionFileList>T1.1.prt.xml</PartitionFileList>",
    )
}

fn partition_definition_fixture() -> String {
    olap_definition_fixture(
        "Partition",
        "T1",
        "11111111-2222-3333-4444-555555555558",
        "11111111-2222-3333-4444-555555555551",
        Some("11111111-2222-3333-4444-555555555554"),
        "",
        "",
        "",
    )
}

#[test]
fn complete_xldm140_rename_updates_outer_identity_relationships_and_inner_closure() {
    let source = complete_package(true);
    let source_bytes = source.to_bytes().expect("source bytes");
    let preserved = preserved_members(&source);
    let inner_before = inner_member_bytes(&source);
    let before = source.data_model().expect("source snapshot");
    assert_eq!(before.definition().unwrap().tables[0].name, OLD_TABLE_NAME);
    assert_eq!(
        before.definition().unwrap().relationships[0].from_table,
        OLD_TABLE_NAME
    );

    let mut transaction = before.edit();
    assert!(
        transaction
            .rename_table(TABLE_ID, NEW_TABLE_NAME)
            .expect("stage table rename")
    );
    let commit = transaction.commit().expect("commit table rename");
    assert!(commit.changed());

    let changed = source.apply_data_model(&commit).expect("apply rename");
    let after = changed.data_model().expect("changed snapshot");
    let definition = after.definition().expect("changed definition");
    assert_eq!(definition.tables[0].id, TABLE_ID);
    assert_eq!(definition.tables[0].name, NEW_TABLE_NAME);
    assert_eq!(definition.relationships[0].from_table, NEW_TABLE_NAME);
    assert_eq!(definition.relationships[0].to_table, NEW_TABLE_NAME);
    assert_inner_table_name(&changed, NEW_TABLE_NAME);
    let inner_after = inner_member_bytes(&changed);
    for (path, bytes) in &inner_before {
        if path != "Model.1.db/T1.0.dim/T1.1.tbl.xml" && path != RELATION_METADATA_PATH {
            assert_eq!(inner_after.get(path), Some(bytes), "member {path}");
        }
    }
    let relation_before = inner_before
        .get(RELATION_METADATA_PATH)
        .expect("relationship metadata before rename");
    let relation_after = inner_after
        .get(RELATION_METADATA_PATH)
        .expect("relationship metadata after rename");
    let relation_expected = replace_bytes_once(
        &replace_bytes_once(relation_before, b"name=\"OldName\"", b"name=\"Renamed\""),
        b"<PrimaryTable>OldName</PrimaryTable>",
        b"<PrimaryTable>Renamed</PrimaryTable>",
    );
    assert_eq!(relation_after, &relation_expected);
    assert_eq!(
        inner_after.get("Model.1.db/T1.0.dim/1.R$T1$RelA.INDEX.0.idf"),
        inner_before.get("Model.1.db/T1.0.dim/1.R$T1$RelA.INDEX.0.idf")
    );
    assert_eq!(preserved_members(&changed), preserved);
    assert_ne!(changed.to_bytes().unwrap(), source_bytes);

    let reopened = Package::from_bytes(changed.to_bytes().unwrap()).expect("reopen changed");
    assert_eq!(
        reopened.data_model().unwrap().definition().unwrap().tables[0].name,
        NEW_TABLE_NAME
    );
    assert_inner_table_name(&reopened, NEW_TABLE_NAME);
}

#[test]
fn exact_same_name_is_a_noop_and_preserves_package_bytes() {
    let source = complete_package(false);
    let source_bytes = source.to_bytes().expect("source bytes");
    let before = source.data_model().expect("snapshot");
    let mut transaction = before.edit();
    assert!(
        !transaction
            .rename_table(TABLE_ID, OLD_TABLE_NAME)
            .expect("exact name no-op")
    );
    let commit = transaction.commit().expect("no-op commit");
    assert!(!commit.changed());
    let applied = source.apply_data_model(&commit).expect("no-op apply");
    assert_eq!(applied.to_bytes().unwrap(), source_bytes);
    assert!(commit.patch().is_empty());
}

#[test]
fn rename_patch_inverse_restores_exact_source_and_rejects_stale_source() {
    let source = complete_package(false);
    let source_bytes = source.to_bytes().expect("source bytes");
    let before = source.data_model().expect("snapshot");
    let mut transaction = before.edit();
    transaction
        .rename_table(TABLE_ID, NEW_TABLE_NAME)
        .expect("stage rename");
    let commit = transaction.commit().expect("commit");
    let changed = source.apply_data_model(&commit).expect("apply");
    let restored = changed
        .apply_data_model_patch(&commit.patch().inverse())
        .expect("inverse");
    assert_eq!(restored.to_bytes().unwrap(), source_bytes);

    let stale = package_with_corrupted_model_source(source.clone());
    assert!(
        stale.apply_data_model(&commit).is_err(),
        "stale source must be refused"
    );
    assert_eq!(
        stale.to_bytes().unwrap(),
        package_with_corrupted_model_source(source)
            .to_bytes()
            .unwrap()
    );
}

#[test]
fn signed_source_allows_exact_noop_but_refuses_identity_change() {
    let signed = signed_complete_package();
    assert!(signed.opc_package().is_signed());
    let source_bytes = signed.to_bytes().expect("signed source bytes");

    let snapshot = signed.data_model().expect("signed snapshot");
    let noop = snapshot.edit().commit().expect("signed no-op commit");
    signed
        .apply_data_model(&noop)
        .expect("signed no-op application");
    assert!(signed.opc_package().is_signed());
    assert_eq!(signed.to_bytes().unwrap(), source_bytes);

    let mut transaction = signed.data_model().expect("signed snapshot").edit();
    transaction
        .rename_table(TABLE_ID, NEW_TABLE_NAME)
        .expect("stage signed rename");
    assert!(transaction.commit().is_err());
    assert!(signed.opc_package().is_signed());
    assert_eq!(signed.to_bytes().unwrap(), source_bytes);
}

#[test]
fn unknown_table_id_and_invalid_names_refuse_without_draft_mutation() {
    let source = complete_package(false);
    let before = source.data_model().expect("snapshot");
    let mut transaction = before.edit();
    assert!(transaction.rename_table("missing", NEW_TABLE_NAME).is_err());
    assert!(transaction.rename_table(TABLE_ID, "").is_err());
    assert!(transaction.rename_table(TABLE_ID, "bad\0name").is_err());
    assert_eq!(
        transaction.definition().expect("draft definition").tables[0].name,
        OLD_TABLE_NAME
    );
}

#[test]
fn duplicate_outer_name_refuses_atomically() {
    let source = duplicate_name_package();
    let before = source.data_model().expect("snapshot");
    let mut transaction = before.edit();
    assert!(transaction.rename_table(TABLE_ID, "Other").is_err());
    assert_eq!(
        transaction.definition().unwrap().tables[0].name,
        OLD_TABLE_NAME
    );
}

#[test]
fn case_variant_outer_relationship_endpoints_are_rewritten() {
    let source = case_variant_endpoint_package();
    let before = source.data_model().expect("case-variant source snapshot");
    let mut transaction = before.edit();
    transaction
        .rename_table(TABLE_ID, NEW_TABLE_NAME)
        .expect("case-variant endpoint rename");
    let commit = transaction.commit().expect("case-variant commit");
    let changed = source
        .apply_data_model(&commit)
        .expect("case-variant apply");
    let changed_snapshot = changed.data_model().unwrap();
    let definition = changed_snapshot.definition().unwrap();
    assert_eq!(definition.tables[0].id, TABLE_ID);
    assert_eq!(definition.tables[0].name, NEW_TABLE_NAME);
    assert_eq!(definition.relationships[0].from_table, NEW_TABLE_NAME);
    assert_eq!(definition.relationships[0].to_table, NEW_TABLE_NAME);
    assert_inner_table_name(&changed, NEW_TABLE_NAME);
}

#[test]
fn outer_inner_name_mismatch_refuses_without_repairing_source() {
    let source = package_with_model(
        model_definition("Sales", false),
        build_identity_storage(&complete_distinct_id_storage_entries()),
    );
    let source_bytes = source.to_bytes().expect("source bytes");
    let before = source.data_model().expect("mismatched source snapshot");
    let mut transaction = before.edit();
    assert!(transaction.rename_table(TABLE_ID, NEW_TABLE_NAME).is_err());
    assert_eq!(transaction.definition().unwrap().tables[0].name, "Sales");
    assert_eq!(
        transaction.payload().unwrap(),
        package_part_bytes(&source, MODEL_URI)
    );
    assert_eq!(source.to_bytes().unwrap(), source_bytes);
}

#[test]
fn escaped_table_name_updates_outer_and_inner_xml() {
    let source = complete_package(false);
    let before = source.data_model().expect("source snapshot");
    let escaped = "A&B<escaped-name>";
    let mut transaction = before.edit();
    transaction
        .rename_table(TABLE_ID, escaped)
        .expect("escaped table name");
    let commit = transaction.commit().expect("escaped commit");
    let changed = source.apply_data_model(&commit).expect("escaped apply");
    let changed_snapshot = changed.data_model().unwrap();
    assert_eq!(
        changed_snapshot.definition().unwrap().tables[0].name,
        escaped
    );
    assert_inner_table_name(&changed, escaped);
    let model_bytes = package_part_bytes(&changed, MODEL_URI);
    let escaped_xml = b"A&amp;B&lt;escaped-name>";
    assert!(
        model_bytes
            .windows(escaped_xml.len())
            .any(|window| window == escaped_xml),
        "escaped name must remain XML-escaped in the opaque model source"
    );
}

#[test]
fn oversized_table_name_refuses_without_draft_mutation() {
    let source = complete_package(false);
    let before = source.data_model().expect("source snapshot");
    let too_long = "x".repeat(32 * 1024 + 1);
    let mut transaction = before.edit();
    assert!(transaction.rename_table(TABLE_ID, &too_long).is_err());
    assert_eq!(
        transaction.definition().unwrap().tables[0].name,
        OLD_TABLE_NAME
    );
}

#[test]
fn rename_then_time_grouping_updates_both_outer_and_inner_identity() {
    let source = complete_package(false);
    let before = source.data_model().expect("source snapshot");
    let mut transaction = before.edit();
    transaction
        .rename_table(TABLE_ID, NEW_TABLE_NAME)
        .expect("rename first");
    transaction
        .add_time_grouping(time_grouping(NEW_TABLE_NAME))
        .expect("time grouping after rename");
    let commit = transaction.commit().expect("combined commit");
    let changed = source.apply_data_model(&commit).expect("combined apply");
    let changed_snapshot = changed.data_model().unwrap();
    let definition = changed_snapshot.definition().unwrap();
    assert_eq!(definition.time_groupings[0].table_name, NEW_TABLE_NAME);
    assert_inner_table_name(&changed, NEW_TABLE_NAME);
}

#[test]
fn time_grouping_then_rename_updates_both_outer_and_inner_identity() {
    let source = complete_package(false);
    let before = source.data_model().expect("source snapshot");
    let mut transaction = before.edit();
    transaction
        .add_time_grouping(time_grouping(OLD_TABLE_NAME))
        .expect("time grouping first");
    transaction
        .rename_table(TABLE_ID, NEW_TABLE_NAME)
        .expect("rename after time grouping");
    let commit = transaction.commit().expect("combined commit");
    let changed = source.apply_data_model(&commit).expect("combined apply");
    let changed_snapshot = changed.data_model().unwrap();
    let definition = changed_snapshot.definition().unwrap();
    assert_eq!(definition.time_groupings[0].table_name, NEW_TABLE_NAME);
    assert_inner_table_name(&changed, NEW_TABLE_NAME);
}

#[test]
fn opaque_payload_replacement_then_rename_refuses_without_source_repair() {
    let source = complete_package(false);
    let replacement = complete_payload_named("Other");
    let before = source.data_model().expect("source snapshot");
    let mut transaction = before.edit();
    transaction
        .replace_payload(replacement.clone())
        .expect("stage opaque replacement");
    assert!(transaction.rename_table(TABLE_ID, NEW_TABLE_NAME).is_err());
    assert_eq!(
        transaction.definition().unwrap().tables[0].name,
        OLD_TABLE_NAME
    );
    assert_eq!(transaction.payload().unwrap(), replacement.as_slice());
}

#[test]
fn incomplete_opaque_model_refuses_table_identity_edits_and_preserves_source() {
    let source = opaque_package();
    let source_bytes = source.to_bytes().expect("source bytes");
    let before = source.data_model().expect("snapshot");
    let mut transaction = before.edit();
    assert!(transaction.rename_table(TABLE_ID, NEW_TABLE_NAME).is_err());
    assert_eq!(
        transaction.definition().unwrap().tables[0].name,
        OLD_TABLE_NAME
    );
    assert_eq!(source.to_bytes().unwrap(), source_bytes);
}

#[test]
fn tabular150_model_profile_refuses_xldm140_identity_edit() {
    let source = complete_package(false);
    let mut opc = source.into_opc();
    let model = PackURI::new(MODEL_URI).expect("model URI");
    let part = opc.get_part_mut(&model).expect("model");
    let mut bytes = part.blob().to_vec();
    replace_utf16_once(&mut bytes, ">140<", ">150<");
    part.set_blob(bytes);
    let source = Package::from_opc(opc).expect("tabular profile package");
    let before = source.data_model().expect("tabular source snapshot");
    let mut transaction = before.edit();
    assert!(transaction.rename_table(TABLE_ID, NEW_TABLE_NAME).is_err());
}

#[test]
fn caller_rewrite_limit_refuses_before_publication() {
    let source = complete_package(false);
    let source_bytes = source.to_bytes().expect("source bytes");
    let mut limits = ReadLimits::DEFAULT;
    limits.max_rewrite_bytes = package_part_bytes(&source, WORKBOOK_URI).len();
    let before = source
        .data_model_with_limits(limits)
        .expect("source under caller rewrite limit");
    let mut transaction = before.edit();
    transaction
        .rename_table(TABLE_ID, "A considerably longer table name")
        .expect("stage rename");
    assert!(transaction.commit().is_err());
    assert_eq!(source.to_bytes().unwrap(), source_bytes);
}

#[test]
fn caller_model_part_limit_refuses_before_draft_change() {
    let source = complete_package(false);
    let source_bytes = source.to_bytes().expect("source bytes");
    let payload_len = package_part_bytes(&source, MODEL_URI).len();
    let long_name = "x".repeat(32 * 1024 - 1);
    let mut limits = ReadLimits::DEFAULT;
    limits.max_part_bytes = payload_len;
    let before = source
        .data_model_with_limits(limits)
        .expect("source under exact model-part limit");
    let mut transaction = before.edit();
    assert!(transaction.rename_table(TABLE_ID, &long_name).is_err());
    assert_eq!(
        transaction.definition().unwrap().tables[0].name,
        OLD_TABLE_NAME
    );
    assert_eq!(source.to_bytes().unwrap(), source_bytes);
}

#[test]
fn native_date_fixture_has_no_model_and_is_not_positive_identity_evidence() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ooxml/xlsb/date.xlsb");
    let bytes = std::fs::read(path).expect("native date fixture");
    let package = Package::from_bytes(bytes).expect("native package");
    let snapshot = package.data_model().expect("native snapshot");
    assert!(!snapshot.is_present());
}
