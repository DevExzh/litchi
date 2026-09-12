// Deterministic complete-XLDM-140 fixture generation for matrix correctness.
//
// The source test fixture included immediately before this file supplies the
// validated one-table XML templates and workbook envelope.  This file only
// expands those templates into a larger, fully linked closure; it does not
// manufacture an opaque payload or bypass the neutral XLDM proof.

use litchi_xlsb::data_model::{TimeGrouping, TimeGroupingColumn, TimeGroupingContentType};

const DATABASE_ID: &str = "11111111-2222-3333-4444-555555555551";
const DATASOURCE_ID: &str = "11111111-2222-3333-4444-555555555552";
const DATASOURCE_VIEW_ID: &str = "11111111-2222-3333-4444-555555555553";
const CUBE_ID: &str = "11111111-2222-3333-4444-555555555554";
const SCRIPT_ID: &str = "11111111-2222-3333-4444-555555555556";

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum Layout {
    SelectedTable,
    Distributed,
}

impl Layout {
    fn parse(value: &str) -> Result<Self, String> {
        match value {
            "selected_table" => Ok(Self::SelectedTable),
            "distributed" => Ok(Self::Distributed),
            _ => Err(format!("unknown endpoint layout: {value}")),
        }
    }
}

#[derive(Clone, Debug)]
struct TableSpec {
    id: String,
    name: String,
    dimension_id: String,
    measure_group_id: String,
    partition_id: String,
    has_relationship: bool,
    relationship_ids: Vec<String>,
}

#[derive(Clone, Debug)]
struct RelationSpec {
    id: String,
    containing_table: usize,
    primary_table: usize,
}

/// Build a complete, internally linked synthetic XLDM 140 package.
pub fn complete_scaled(
    table_count: usize,
    relationship_count: usize,
    endpoint_layout: &str,
) -> Result<Package, String> {
    if table_count == 0 || table_count > 64 {
        return Err(format!("table count outside matrix: {table_count}"));
    }
    if relationship_count > 256 {
        return Err(format!("relationship count outside matrix: {relationship_count}"));
    }
    let layout = Layout::parse(endpoint_layout)?;
    let mut tables = (1..=table_count)
        .map(|index| TableSpec {
            id: format!("T{index}"),
            name: format!("Table{index}"),
            dimension_id: format!("11111111-2222-3333-4444-{:012X}", 0x5555_5555_5500 + index),
            measure_group_id: format!("11111111-2222-3333-4444-{:012X}", 0x6666_6666_6600 + index),
            partition_id: format!("11111111-2222-3333-4444-{:012X}", 0x7777_7777_7700 + index),
            has_relationship: false,
            relationship_ids: Vec::new(),
        })
        .collect::<Vec<_>>();
    let mut pairs = (0..table_count)
        .flat_map(|containing_table| {
            (0..table_count)
                .filter(move |primary_table| *primary_table != containing_table)
                .map(move |primary_table| (containing_table, primary_table))
        })
        .collect::<Vec<_>>();
    if matches!(layout, Layout::SelectedTable) {
        pairs.sort_by_key(|(containing_table, primary_table)| {
            (usize::from(*containing_table != 0), *containing_table, *primary_table)
        });
    } else {
        // The distributed layout starts with the directed ring edges and
        // then widens the ring distance.  This spreads containing/primary
        // incidence across the table set while retaining one deterministic,
        // unique pair for every relationship.
        pairs.sort_by_key(|(containing_table, primary_table)| {
            (
                (primary_table + table_count - containing_table) % table_count,
                *containing_table,
                *primary_table,
            )
        });
    }
    if relationship_count > pairs.len() {
        return Err(format!(
            "relationship count {} exceeds unique table-pair capacity {}",
            relationship_count,
            pairs.len()
        ));
    }
    let relationships = pairs
        .into_iter()
        .take(relationship_count)
        .enumerate()
        .map(|(index, (containing_table, primary_table))| {
            let id = format!("Rel{}", index + 1);
            tables[containing_table].has_relationship = true;
            tables[containing_table].relationship_ids.push(id.clone());
            RelationSpec {
                id,
                containing_table,
                primary_table,
            }
        })
        .collect::<Vec<_>>();

    // Keep the optional Year time-grouping closure on the small/medium
    // points.  Large points deliberately exercise the table/relationship
    // graph at the host's bounded proof ceiling without adding a second
    // column family unrelated to the identity rename.
    let include_time_grouping = table_count <= 16;
    let payload = build_storage(&tables, &relationships, include_time_grouping)?;
    let definition = Definition {
        min_version_load: 5,
        tables: tables
            .iter()
            .map(|table| Table {
                id: table.id.clone(),
                name: table.name.clone(),
                connection: CONNECTION_NAME.to_owned(),
            })
            .collect(),
        relationships: relationships
            .iter()
            .map(|relationship| Relationship {
                from_table: tables[relationship.containing_table].name.clone(),
                from_column: "Key".to_owned(),
                to_table: tables[relationship.primary_table].name.clone(),
                to_column: "Key".to_owned(),
            })
            .collect(),
        time_groupings: if include_time_grouping {
                tables
                    .iter()
                    .map(|table| TimeGrouping {
                        table_name: table.name.clone(),
                        column_name: "Key".to_owned(),
                        column_id: "Key".to_owned(),
                        columns: vec![TimeGroupingColumn {
                            is_selected: true,
                            content_type: TimeGroupingContentType::Years,
                            column_name: "Year".to_owned(),
                            column_id: "Year".to_owned(),
                        }],
                    })
                    .collect()
            } else {
                Vec::new()
            },
    };
    scaled_package_with_model(definition, payload)
}

fn scaled_package_with_model(definition: Definition, payload: Vec<u8>) -> Result<Package, String> {
    let mut writer = WorkbookWriter::new();
    writer.add_worksheet(MutableWorksheet::new("Sheet1"));
    writer
        .set_connections(connections())
        .map_err(|error| error.to_string())?;
    let model = Model::from_bytes(definition, payload).map_err(|error| error.to_string())?;
    writer
        .set_data_model(model)
        .map_err(|error| format!("set_data_model: {error}"))?;
    let mut output = Cursor::new(Vec::new());
    writer.save(&mut output).map_err(|error| error.to_string())?;
    let bytes = output.into_inner();
    Package::from_bytes(bytes.clone()).map_err(|error| error.to_string())?;
    Package::from_bytes(bytes).map_err(|error| error.to_string())
}

fn build_storage(
    tables: &[TableSpec],
    relationships: &[RelationSpec],
    include_time_grouping: bool,
) -> Result<Vec<u8>, String> {
    let mut entries = vec![
        (
        "Partitions".to_owned(),
        partitions_fixture().into_bytes(),
        ),
        (
        "Model.1.db.xml".to_owned(),
        database_definition_fixture().into_bytes(),
        ),
        (
        "Model.1.db/Source.1.ds.xml".to_owned(),
        datasource_definition_fixture().into_bytes(),
        ),
        (
        "Model.1.db/View.1.dsv.xml".to_owned(),
        datasource_view_definition_fixture().into_bytes(),
        ),
        (
        "Model.1.db/C.1.cub.xml".to_owned(),
        cube_definition(tables, include_time_grouping).into_bytes(),
        ),
        (
        "Model.1.db/C.0.cub/MdxScript.0.scr.xml".to_owned(),
        mdx_script_definition_fixture().into_bytes(),
        ),
    ];

    for (index, table) in tables.iter().enumerate() {
        entries.push((
            format!("Model.1.db/{}.1.dim.xml", table.id),
            dimension_definition(table, index, relationships, include_time_grouping).into_bytes(),
        ));
        entries.push((
            format!("Model.1.db/C.0.cub/{}.1.det.xml", table.id),
            measure_group_definition(
                table,
                index,
                relationships,
                tables,
                include_time_grouping,
            )
            .into_bytes(),
        ));
        entries.push((
            format!("Model.1.db/C.0.cub/{}.0.det/{}.1.prt.xml", table.id, table.id),
            partition_definition(table, index).into_bytes(),
        ));
        entries.push((
            format!("Model.1.db/{}.0.dim/{}.1.tbl.xml", table.id, table.id),
            table_metadata(table, include_time_grouping).into_bytes(),
        ));
        entries.push((
            format!("Model.1.db/{}.0.dim/1.{}.Key.0.idf", table.id, table.id),
            column_data_fixture(),
        ));
        entries.push((
            format!("Model.1.db/{}.0.dim/1.H${}$Key.POS_TO_ID.0.idf", table.id, table.id),
            generated_mapping_fixture(),
        ));
        if include_time_grouping {
            entries.push((
                format!("Model.1.db/{}.0.dim/1.{}.Year.0.idf", table.id, table.id),
                column_data_fixture(),
            ));
            entries.push((
                format!(
                    "Model.1.db/{}.0.dim/1.H${}$Year.POS_TO_ID.0.idf",
                    table.id, table.id
                ),
                generated_mapping_fixture(),
            ));
        }
    }
    for relationship in relationships {
        let containing = &tables[relationship.containing_table];
        entries.push((
            format!(
                "Model.1.db/{}.0.dim/R${}${}.1.tbl.xml",
                containing.id, containing.id, relationship.id
            ),
            relationship_definition(relationship, tables).into_bytes(),
        ));
        entries.push((
            format!(
                "Model.1.db/{}.0.dim/1.R${}${}.INDEX.0.idf",
                containing.id, containing.id, relationship.id
            ),
            column_data_fixture(),
        ));
    }
    let backup_log = scaled_backup_log_fixture(&entries, tables);
    entries.push(("BackupLog".to_owned(), backup_log.into_bytes()));
    scaled_build_identity_storage(&entries)
}

fn table_metadata(table: &TableSpec, include_time_grouping: bool) -> String {
    let mut result = column_metadata_fixture()
        .replace("OldName", &table.name)
        .replace("T1", &table.id)
        .replace(
            "<TableStore>Key</TableStore>",
            &format!("<TableStore>{}$Key</TableStore>", table.id),
        )
        .replace(
            "<TableStore>Year</TableStore>",
            &format!("<TableStore>{}$Year</TableStore>", table.id),
        );
    if !include_time_grouping {
        let year_start = result
            .find("<XMObject class=\"XMRawColumn\" name=\"Year\">")
            .expect("Year column");
        let year_end = result[year_start..]
            .find("</XMObject></Collection><Collection><Name>Relationships")
            .expect("Year column end")
            + year_start
            + "</XMObject>".len();
        result.replace_range(year_start..year_end, "");
    }
    result
}

fn dimension_definition(
    table: &TableSpec,
    index: usize,
    relationships: &[RelationSpec],
    include_time_grouping: bool,
) -> String {
    let mut result = dimension_definition_fixture(false)
        .replace("11111111-2222-3333-4444-555555555555", &table.dimension_id)
        .replace("T1", &table.id)
        .replace("<Ordinal>0</Ordinal>", &format!("<Ordinal>{index}</Ordinal>"));
    if !include_time_grouping {
        result = result.replace(
            "<Attribute xsi:type=\"DimensionAttribute\"><ID>Year</ID></Attribute>",
            "",
        );
    }
    let relation_ids = relationships
        .iter()
        .filter(|relationship| relationship.containing_table == index)
        .map(|relationship| relationship.id.as_str())
        .collect::<Vec<_>>();
    let dimension_relationship_ids = relationships
        .iter()
        .filter(|relationship| relationship.primary_table == index)
        .map(|relationship| relationship.id.as_str())
        .collect::<Vec<_>>();
    let base_data_files = if include_time_grouping {
        format!(
            "1.{}.Key.0.idf;1.{}.Year.0.idf;1.H${}$Key.POS_TO_ID.0.idf;1.H${}$Year.POS_TO_ID.0.idf",
            table.id, table.id, table.id, table.id
        )
    } else {
        format!(
            "1.{}.Key.0.idf;1.H${}$Key.POS_TO_ID.0.idf",
            table.id, table.id
        )
    };
    let data_files = relation_ids.iter().fold(base_data_files, |mut value, relation_id| {
            value.push_str(&format!(";1.R${}${}.INDEX.0.idf", table.id, relation_id));
            value
        });
    // The template always contains the complete Key+Year list.  When the
    // large relationship points intentionally omit the optional Year
    // grouping, replace that complete template list with the Key-only list;
    // using the already-shortened list as the search pattern would leave the
    // stale Year member in the serialized dimension definition.
    let old_data_files = format!(
        "1.{}.Key.0.idf;1.{}.Year.0.idf;1.H${}$Key.POS_TO_ID.0.idf;1.H${}$Year.POS_TO_ID.0.idf",
        table.id, table.id, table.id, table.id
    );
    result = result.replace(&old_data_files, &data_files);
    let refs = dimension_relationship_ids
        .iter()
        .map(|id| format!("<Relationship><ID>{id}</ID></Relationship>"))
        .collect::<String>();
    result = result.replace(
        "<Ordinal>",
        &format!("<Relationships>{refs}</Relationships><Ordinal>"),
    );
    result
}

fn relationship_definition(relationship: &RelationSpec, tables: &[TableSpec]) -> String {
    let containing = &tables[relationship.containing_table];
    let primary = &tables[relationship.primary_table];
    let column = compact_relationship_column(&containing.id, &relationship.id);
    format!(
        "<XMObject class=\"XMSimpleTable\" name=\"{}\"><Properties><Version>1</Version><Settings>0</Settings><RIViolationCount>0</RIViolationCount></Properties><Members><Member><Name>SegmentMap</Name><XMObject class=\"XMSegment1Map\"><Properties><Records>1</Records></Properties></XMObject></Member><Member><Name>TableStats</Name><XMObject class=\"XMTableStats\"><Properties><SegmentSize>1</SegmentSize><Usage>0</Usage></Properties></XMObject></Member></Members><Collections><Collection><Name>Partitions</Name></Collection><Collection><Name>Columns</Name>{column}</Collection><Collection><Name>Relationships</Name><XMObject class=\"XMRelationship\" name=\"{}\"><Properties><PrimaryTable>{}</PrimaryTable><PrimaryColumn>Key</PrimaryColumn><ForeignColumn>Key</ForeignColumn></Properties><DataObjects><DataObject><XMObject class=\"XMRelationshipIndex123DIDs\"/></DataObject></DataObjects></XMObject></Collection><Collection><Name>UserHierarchies</Name></Collection></Collections></XMObject>",
        containing.name, relationship.id, primary.name
    )
}

fn compact_relationship_column(table_id: &str, relationship_id: &str) -> String {
    let stats = "<Properties><DistinctStates>1</DistinctStates><MinDataID>0</MinDataID><MaxDataID>0</MaxDataID><OriginalMinSegmentDataID>0</OriginalMinSegmentDataID><RLESortOrder>-1</RLESortOrder><RowCount>1</RowCount><HasNulls>false</HasNulls><RLERuns>0</RLERuns><OthersRLERuns>0</OthersRLERuns><Usage>0</Usage><DBType>7</DBType><XMType>0</XMType><CompressionType>0</CompressionType><CompressionParam>0</CompressionParam><EncodingHint>0</EncodingHint><AggCounter>0</AggCounter><WhereCounter>0</WhereCounter><OrderByCounter>0</OrderByCounter></Properties>";
    let segment_stats = "<XMObject class=\"XMColumnSegmentStats\"><Properties><DistinctStates>1</DistinctStates><MinDataID>0</MinDataID><MaxDataID>0</MaxDataID><OriginalMinSegmentDataID>0</OriginalMinSegmentDataID><RLESortOrder>-1</RLESortOrder><RowCount>1</RowCount><HasNulls>false</HasNulls><RLERuns>0</RLERuns><OthersRLERuns>0</OthersRLERuns></Properties></XMObject>".to_owned();
    let compression =
        "<XMObject class=\"XM123CompressionInfo\"><Properties><Min>0</Min></Properties></XMObject>";
    let subsegment = format!(
        "<XMObject class=\"XMColumnSegment\"><Properties><Records>1</Records><Mask>0</Mask></Properties><Members><Member><Name>CompressionInfo</Name>{compression}</Member><Member><Name>ColumnSegmentStats</Name>{segment_stats}</Member></Members></XMObject>"
    );
    format!(
        "<XMObject class=\"XMRawColumn\" name=\"Rel{relationship_id}\"><Properties><Settings>5</Settings><ColumnFlags>8</ColumnFlags><Collation></Collation><OrderByColumn></OrderByColumn><Locale>0</Locale><BinaryCharacters>0</BinaryCharacters></Properties><Members><Member><Name>IntrinsicHierarchy</Name><XMObject class=\"XMHierarchy\"><Properties><SortOrder>0</SortOrder><IsProcessed>false</IsProcessed><TypeMaterialization>0</TypeMaterialization><ColumnPosition2DataID>-1</ColumnPosition2DataID><ColumnDataID2Position>-1</ColumnDataID2Position><DistinctDataIDs>1</DistinctDataIDs><TableStore>R${table_id}${relationship_id}$Key</TableStore></Properties></XMObject></Member><Member><Name>ColumnStats</Name><XMObject class=\"XMColumnStats\">{stats}</XMObject></Member></Members><Collections><Collection><Name>Segments</Name><XMObject class=\"XMColumnSegment\"><Properties><Records>1</Records><Mask>0</Mask></Properties><Members><Member><Name>SubSegment</Name>{subsegment}</Member><Member><Name>CompressionInfo</Name>{compression}</Member><Member><Name>ColumnSegmentStats</Name>{segment_stats}</Member></Members></XMObject></Collection></Collections><DataObjects><DataObject><XMObject class=\"XMRawColumnPartitionDataObject\" name=\"1.{table_id}.Key.0.idf\"><Properties><DataVersion>0</DataVersion><Partition>0</Partition><SegmentCount>1</SegmentCount></Properties></XMObject></DataObject><DataObject><XMObject class=\"XMValueDataDictionary&lt;XM_Long&gt;\" name=\"1.{table_id}.Key.dictionary\"><Properties><DataVersion>0</DataVersion><BaseId>0</BaseId><Magnitude>0</Magnitude></Properties></XMObject></DataObject></DataObjects></XMObject>"
    )
}

fn measure_group_definition(
    table: &TableSpec,
    index: usize,
    relationships: &[RelationSpec],
    tables: &[TableSpec],
    include_time_grouping: bool,
) -> String {
    let mut result = measure_group_definition_fixture(false)
        .replace("11111111-2222-3333-4444-555555555555", &table.dimension_id)
        .replace("11111111-2222-3333-4444-555555555557", &table.measure_group_id)
        .replace("T1", &table.id)
        .replace("<Ordinal>0</Ordinal>", &format!("<Ordinal>{index}</Ordinal>"));
    result = result.replace(
        &format!("<Name>T{index}</Name>"),
        &format!("<Name>{}</Name>", table.id),
    );
    let attributes = if include_time_grouping {
        "<Attribute xsi:type=\"MeasureGroupDimensionAttribute\"><AttributeID>Key</AttributeID></Attribute><Attribute xsi:type=\"MeasureGroupDimensionAttribute\"><AttributeID>Year</AttributeID></Attribute>"
    } else {
        "<Attribute xsi:type=\"MeasureGroupDimensionAttribute\"><AttributeID>Key</AttributeID></Attribute>"
    };
    let references = relationships
        .iter()
        .filter(|relationship| relationship.primary_table == index)
        .map(|relationship| {
            let related = &tables[relationship.containing_table];
            format!(
                "<Dimension xsi:type=\"ReferenceMeasureGroupDimension\"><CubeDimensionID>{}</CubeDimensionID><Attributes>{attributes}</Attributes></Dimension>",
                related.dimension_id,
            )
        })
        .collect::<String>();
    let own_attributes = if include_time_grouping {
        "<Attribute xsi:type=\"MeasureGroupDimensionAttribute\"><AttributeID>Key</AttributeID></Attribute><Attribute xsi:type=\"MeasureGroupDimensionAttribute\"><AttributeID>Year</AttributeID></Attribute>"
    } else {
        "<Attribute xsi:type=\"MeasureGroupDimensionAttribute\"><AttributeID>Key</AttributeID></Attribute>"
    };
    result = result.replace(
        "<Attributes><Attribute xsi:type=\"MeasureGroupDimensionAttribute\"><AttributeID>Key</AttributeID></Attribute><Attribute xsi:type=\"MeasureGroupDimensionAttribute\"><AttributeID>Year</AttributeID></Attribute></Attributes>",
        &format!("<Attributes>{own_attributes}</Attributes>"),
    );
    let dimensions_end = result
        .find("</Dimensions>")
        .expect("measure group dimensions");
    result.insert_str(dimensions_end, &references);
    result
}

fn partition_definition(table: &TableSpec, index: usize) -> String {
    partition_definition_fixture()
        .replace("11111111-2222-3333-4444-555555555558", &table.partition_id)
        .replace("11111111-2222-3333-4444-555555555555", &table.dimension_id)
        .replace("T1", &table.id)
        .replace("<Ordinal>0</Ordinal>", &format!("<Ordinal>{index}</Ordinal>"))
}

fn cube_definition(tables: &[TableSpec], include_time_grouping: bool) -> String {
    let mut result = cube_definition_fixture();
    let dimensions = tables
        .iter()
        .map(|table| {
            if include_time_grouping {
                format!(
                    "<Dimension xsi:type=\"CubeDimension\"><DimensionID>{}</DimensionID><Attributes><Attribute xsi:type=\"CubeAttribute\"><AttributeID>Key</AttributeID></Attribute><Attribute xsi:type=\"CubeAttribute\"><AttributeID>Year</AttributeID></Attribute></Attributes></Dimension>",
                    table.dimension_id
                )
            } else {
                format!(
                    "<Dimension xsi:type=\"CubeDimension\"><DimensionID>{}</DimensionID><Attributes><Attribute xsi:type=\"CubeAttribute\"><AttributeID>Key</AttributeID></Attribute></Attributes></Dimension>",
                    table.dimension_id
                )
            }
        })
        .collect::<String>();
    let start = result.find("<Dimensions>").expect("cube dimensions start");
    let end = result
        .find("</Dimensions>")
        .expect("cube dimensions end")
        + "</Dimensions>".len();
    result.replace_range(start..end, &format!("<Dimensions>{dimensions}</Dimensions>"));
    let measure_groups = tables
        .iter()
        .map(|table| format!("{}.1.det.xml", table.id))
        .collect::<Vec<_>>()
        .join(";");
    result.replace(
        "<MeasureGroupFileList>T1.1.det.xml</MeasureGroupFileList>",
        &format!("<MeasureGroupFileList>{measure_groups}</MeasureGroupFileList>"),
    )
}

fn scaled_backup_log_fixture(entries: &[(String, Vec<u8>)], tables: &[TableSpec]) -> String {
    let mut result = String::from(
        "<BackupLog><BackupRestoreSyncVersion>1153</BackupRestoreSyncVersion><ServerRoot>C:\\inert</ServerRoot><SvrEncryptPwdFlag>true</SvrEncryptPwdFlag><ServerEnableBinaryXML>false</ServerEnableBinaryXML><ServerEnableCompression>false</ServerEnableCompression><CompressionFlag>false</CompressionFlag><EncryptionFlag>false</EncryptionFlag><ObjectName>Model</ObjectName><ObjectId>11111111-2222-3333-4444-555555555500</ObjectId><Write>ReadWrite</Write><OlapInfo>true</OlapInfo><Collations><Collation>Latin1_General</Collation></Collations><Languages><Language>1033</Language></Languages><FileGroups>",
    );
    append_group(
        &mut result,
        "100002",
        100_002,
        "Database",
        DATABASE_ID,
        1,
        &["Model.1.db.xml"],
        entries,
    );
    append_group(
        &mut result,
        "100003",
        100_003,
        "Source",
        DATASOURCE_ID,
        1,
        &["Model.1.db/Source.1.ds.xml"],
        entries,
    );
    append_group(
        &mut result,
        "100053",
        100_053,
        "View",
        DATASOURCE_VIEW_ID,
        1,
        &["Model.1.db/View.1.dsv.xml"],
        entries,
    );
    append_group(
        &mut result,
        "100010",
        100_010,
        "Cube",
        CUBE_ID,
        1,
        &["Model.1.db/C.1.cub.xml"],
        entries,
    );
    for (index, table) in tables.iter().enumerate() {
        let dimension_path = format!("Model.1.db/{}.1.dim.xml", table.id);
        let mut dimension_members = vec![dimension_path.as_str()];
        dimension_members.extend(
            entries
            .iter()
            .filter(|(path, _)| {
                path.starts_with(&format!("Model.1.db/{}.0.dim/", table.id))
            })
            .map(|(path, _)| path.as_str())
        );
        append_group(
            &mut result,
            "100006",
            100_006 + i32::try_from(index).unwrap_or(i32::MAX),
            &table.id,
            &table.dimension_id,
            1,
            std::slice::from_ref(&dimension_path.as_str()),
            entries,
        );
        // The Dimension group's FileList owns all table-local native,
        // generated, and relationship-index members.
        replace_group_members(
            &mut result,
            "100006",
            &dimension_members,
            entries,
        );
        let measure_path = format!("Model.1.db/C.0.cub/{}.1.det.xml", table.id);
        append_group(
            &mut result,
            "100016",
            100_016 + i32::try_from(index).unwrap_or(i32::MAX),
            &table.id,
            &table.measure_group_id,
            1,
            std::slice::from_ref(&measure_path.as_str()),
            entries,
        );
        let partition_path = format!(
            "Model.1.db/C.0.cub/{}.0.det/{}.1.prt.xml",
            table.id, table.id
        );
        append_group(
            &mut result,
            "100021",
            100_021 + i32::try_from(index).unwrap_or(i32::MAX),
            &table.id,
            &table.partition_id,
            1,
            std::slice::from_ref(&partition_path.as_str()),
            entries,
        );
    }
    let script_path = "Model.1.db/C.0.cub/MdxScript.0.scr.xml";
    append_group(
        &mut result,
        "100060",
        100_060,
        "MdxScript",
        SCRIPT_ID,
        0,
        &[script_path],
        entries,
    );
    result.push_str("</FileGroups></BackupLog>");
    result
}

fn replace_group_members(
    result: &mut String,
    class: &str,
    members: &[&str],
    entries: &[(String, Vec<u8>)],
) {
    // `append_group` writes each FileList in one contiguous span.  The
    // generated group has already been appended, so replace the last DIM
    // group's empty/one-member list using its definition path marker.
    let marker = format!("<FileGroup><Class>{class}</Class>");
    let Some(group_start) = result.rfind(&marker) else {
        return;
    };
    let Some(file_list_start_rel) = result[group_start..].find("<FileList>") else {
        return;
    };
    let file_list_start = group_start + file_list_start_rel;
    let Some(file_list_end_rel) = result[file_list_start..].find("</FileList>") else {
        return;
    };
    let file_list_end = file_list_start + file_list_end_rel + "</FileList>".len();
    let list = members
        .iter()
        .map(|path| {
            let size = entries
                .iter()
                .find(|(entry_path, _)| entry_path == path)
                .map_or(0, |(_, payload)| payload.len());
            format!("<BackupFile><Path>C:\\inert\\{path}</Path><StoragePath>{path}</StoragePath><LastWriteTime>0</LastWriteTime><Size>{size}</Size></BackupFile>")
        })
        .collect::<String>();
    result.replace_range(file_list_start..file_list_end, &format!("<FileList>{list}</FileList>"));
}

fn append_group(
    result: &mut String,
    class: &str,
    class_id: i32,
    name: &str,
    object_id: &str,
    object_version: i32,
    paths: &[&str],
    entries: &[(String, Vec<u8>)],
) {
    result.push_str(&format!(
        "<FileGroup><Class>{class}</Class><ID>{object_id}</ID><Name>{name}</Name><ObjectVersion>{object_version}</ObjectVersion><PersistLocation>1</PersistLocation><PersistLocationPath>Model.1.db</PersistLocationPath><StorageLocationPath></StorageLocationPath><ObjectID>{object_id}</ObjectID><FileList>"
    ));
    for path in paths {
        let size = entries
            .iter()
            .find(|(entry_path, _)| entry_path == path)
            .map_or(0, |(_, payload)| payload.len());
        result.push_str(&format!(
            "<BackupFile><Path>C:\\inert\\{path}</Path><StoragePath>{path}</StoragePath><LastWriteTime>0</LastWriteTime><Size>{size}</Size></BackupFile>"
        ));
    }
    let _ = class_id;
    result.push_str("</FileList></FileGroup>");
}

fn scaled_build_identity_storage(entries: &[(String, Vec<u8>)]) -> Result<Vec<u8>, String> {
    let mut bytes = vec![0; XLDM_PAGE_SIZE];
    bytes.extend_from_slice(&BOM);
    let mut allocations = Vec::with_capacity(entries.len());
    for (index, (_, payload)) in entries.iter().enumerate() {
        if index + 1 == entries.len() {
            bytes.extend_from_slice(&BOM);
        }
        let offset = bytes.len();
        bytes.extend_from_slice(payload);
        bytes.extend_from_slice(&scaled_crc32(payload).to_le_bytes());
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
    let directory_bytes = scaled_utf16le(&directory);
    bytes.extend_from_slice(&directory_bytes);
    bytes.resize(bytes.len().div_ceil(XLDM_PAGE_SIZE) * XLDM_PAGE_SIZE, 0);
    let header = format!(
        "<BackupLog><BackupRestoreSyncVersion>140</BackupRestoreSyncVersion><Fault>false</Fault><faultcode>0</faultcode><ErrorCode>true</ErrorCode><EncryptionFlag>false</EncryptionFlag><EncryptionKey>0</EncryptionKey><ApplyCompression>true</ApplyCompression><m_cbOffsetHeader>{directory_offset}</m_cbOffsetHeader><DataSize>{}</DataSize><Files>{}</Files><ObjectID>11111111-2222-3333-4444-555555555500</ObjectID><m_cbOffsetData>4096</m_cbOffsetData></BackupLog>",
        directory_bytes.len(),
        entries.len()
    );
    let mut page = Vec::new();
    page.extend_from_slice(&BOM);
    page.extend_from_slice(&scaled_utf16le(XLDM_STREAM_SIGNATURE));
    page.extend_from_slice(&scaled_utf16le(&header));
    if page.len() > XLDM_PAGE_SIZE {
        return Err(String::from("XLDM header exceeds one page"));
    }
    page.resize(XLDM_PAGE_SIZE, 0);
    bytes[..XLDM_PAGE_SIZE].copy_from_slice(&page);
    Ok(bytes)
}

fn scaled_utf16le(value: &str) -> Vec<u8> {
    value.encode_utf16().flat_map(u16::to_le_bytes).collect()
}

fn scaled_crc32(bytes: &[u8]) -> u32 {
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
