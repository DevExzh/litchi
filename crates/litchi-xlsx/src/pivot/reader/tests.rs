//! Focused pivot reader regression and package-graph tests.

use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;
use litchi_opc::{OpcPackage, PackURI, Part};

use super::super::PivotValueFunction;
use super::super::cache::Item;
use super::codec::{
    read_pivot_cache_definition, read_pivot_cache_records, read_pivot_table_definition,
};
use super::package::read_pivot_tables;

fn package_with_pivot_table() -> (OpcPackage, PackURI) {
    let mut package = OpcPackage::new();
    let workbook_uri = PackURI::new("/custom/book.xml").unwrap();
    let worksheet_uri = PackURI::new("/custom/sheets/data.xml").unwrap();
    let source_uri = PackURI::new("/custom/sheets/source.xml").unwrap();
    let table_uri = PackURI::new("/custom/pivots/table.xml").unwrap();
    let cache_uri = PackURI::new("/custom/cache/cache.xml").unwrap();
    let records_uri = PackURI::new("/custom/cache/records.xml").unwrap();
    let mut workbook_part = BlobPart::new(
        workbook_uri,
        ct::SML_SHEET_MAIN.to_string(),
        br#"<workbook xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"
                xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships">
                <sheets><sheet name="Pivot" sheetId="1" r:id="rId1"/>
                    <sheet name="Source" sheetId="2" r:id="rId2"/></sheets>
                <pivotCaches><pivotCache cacheId="7" r:id="rId3"/></pivotCaches>
            </workbook>"#
            .to_vec(),
    );
    workbook_part.relate_to("sheets/data.xml", rt::STRICT_WORKSHEET);
    workbook_part.relate_to("sheets/source.xml", rt::STRICT_WORKSHEET);
    workbook_part.relate_to("cache/cache.xml", rt::STRICT_PIVOT_CACHE_DEFINITION);
    let mut worksheet_part = BlobPart::new(
        worksheet_uri.clone(),
        ct::SML_WORKSHEET.to_string(),
        Vec::new(),
    );
    worksheet_part.relate_to("../pivots/table.xml", rt::STRICT_PIVOT_TABLE);
    let mut cache_part = BlobPart::new(
        cache_uri,
        ct::SML_PIVOT_CACHE_DEFINITION.to_string(),
        br#"<pivotCacheDefinition xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"
                xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
                r:id="rId1" recordCount="2">
                <cacheSource type="worksheet"><worksheetSource ref="$A$1:$B$3" r:id="rId2"/></cacheSource>
                <cacheFields count="1"><cacheField name="Cache Region"/></cacheFields>
            </pivotCacheDefinition>"#
            .to_vec(),
    );
    cache_part.relate_to("records.xml", rt::STRICT_PIVOT_CACHE_RECORDS);
    cache_part.relate_to("../sheets/source.xml", rt::STRICT_WORKSHEET);
    let mut table_part = BlobPart::new(
        table_uri,
        ct::SML_PIVOT_TABLE.to_string(),
        br#"<pivotTableDefinition xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"
                name="PivotOne" cacheId="7" dataCaption="Values">
                <location ref="A1:C5" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
                <pivotFields count="1"><pivotField/></pivotFields>
                <rowFields count="1"><field x="0"/></rowFields>
            </pivotTableDefinition>"#
            .to_vec(),
    );
    table_part.relate_to("../cache/cache.xml", rt::STRICT_PIVOT_CACHE_DEFINITION);
    package.relate_to("custom/book.xml", rt::STRICT_OFFICE_DOCUMENT);
    package.add_part(Box::new(workbook_part));
    package.add_part(Box::new(worksheet_part));
    package.add_part(Box::new(BlobPart::new(
        source_uri,
        ct::SML_WORKSHEET.to_string(),
        Vec::new(),
    )));
    package.add_part(Box::new(cache_part));
    package.add_part(Box::new(BlobPart::new(
        records_uri,
        ct::SML_PIVOT_CACHE_RECORDS.to_string(),
        br#"<pivotCacheRecords xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" count="2">
                <r><s v="North"/></r><r><s v="South"/></r>
            </pivotCacheRecords>"#
            .to_vec(),
    )));
    package.add_part(Box::new(table_part));
    (package, worksheet_uri)
}

#[test]
fn resolves_strict_custom_pivot_table_parts() {
    let (package, _) = package_with_pivot_table();
    let tables = read_pivot_tables(&package).unwrap();

    assert_eq!(tables.len(), 1);
    assert_eq!(tables[0].name, "PivotOne");
    assert_eq!(tables[0].sheet_name, "Pivot");
    assert_eq!(tables[0].source_sheet.as_deref(), Some("Source"));
    assert_eq!(tables[0].source_ref.as_deref(), Some("$A$1:$B$3"));
    assert_eq!(tables[0].field_names, ["Cache Region"]);
    assert_eq!(tables[0].row_fields[0].field_name, "Cache Region");
}

#[test]
fn tolerates_chartsheet_entries_in_sheet_walk() {
    let (mut package, _) = package_with_pivot_table();
    let workbook_uri = PackURI::new("/custom/book.xml").unwrap();
    let workbook_part = package.get_part_mut(&workbook_uri).unwrap();
    let updated = std::str::from_utf8(workbook_part.blob()).unwrap().replace(
        "</sheets>",
        r#"<sheet name="Chart1" sheetId="3" r:id="rId4"/></sheets>"#,
    );
    workbook_part.set_blob(updated.into_bytes());
    workbook_part.relate_to(
        "chartsheets/chart1.xml",
        "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet",
    );
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/custom/chartsheets/chart1.xml").unwrap(),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml".to_string(),
        Vec::new(),
    )));

    let tables = read_pivot_tables(&package).unwrap();
    assert_eq!(tables.len(), 1);
    assert_eq!(tables[0].name, "PivotOne");
}

#[test]
fn poi_fixture_with_chartsheet_reads_pivot_tables() {
    const POI_CHARTSHEET: &[u8] =
        include_bytes!("../../../../../test-data/poi/test-data/spreadsheet/WithChartSheet.xlsx");
    let package = OpcPackage::from_bytes(POI_CHARTSHEET).unwrap();
    let tables = read_pivot_tables(&package).unwrap();
    assert_eq!(tables.len(), 5);
    assert!(
        tables
            .iter()
            .any(|table| table.name == "PivotTable2" && table.sheet_name == "Sheet2")
    );
}

#[test]
fn rejects_external_and_wrong_content_type_pivot_parts() {
    let (mut package, worksheet_uri) = package_with_pivot_table();
    let relationships = package.get_part_mut(&worksheet_uri).unwrap().rels_mut();
    relationships.remove("rId1").unwrap();
    relationships.add_relationship(
        rt::STRICT_PIVOT_TABLE.to_string(),
        "https://example.com/pivot.xml".to_string(),
        "rId1".to_string(),
        true,
    );
    assert!(read_pivot_tables(&package).is_err());

    let (mut package, _) = package_with_pivot_table();
    package.add_part(Box::new(BlobPart::new(
        PackURI::new("/custom/pivots/table.xml").unwrap(),
        ct::SML_WORKSHEET.to_string(),
        Vec::new(),
    )));
    assert!(read_pivot_tables(&package).is_err());
}

#[test]
fn rejects_invalid_pivot_cache_relationship_graphs() {
    let workbook_uri = PackURI::new("/custom/book.xml").unwrap();
    let table_uri = PackURI::new("/custom/pivots/table.xml").unwrap();
    let cache_uri = PackURI::new("/custom/cache/cache.xml").unwrap();
    let records_uri = PackURI::new("/custom/cache/records.xml").unwrap();

    let (mut package, _) = package_with_pivot_table();
    let relationships = package.get_part_mut(&workbook_uri).unwrap().rels_mut();
    relationships.remove("rId3").unwrap();
    relationships.add_relationship(
        rt::STRICT_PIVOT_CACHE_DEFINITION.to_string(),
        "https://example.com/cache.xml".to_string(),
        "rId3".to_string(),
        true,
    );
    assert!(read_pivot_tables(&package).is_err());

    let (mut package, _) = package_with_pivot_table();
    package
        .get_part_mut(&table_uri)
        .unwrap()
        .rels_mut()
        .add_relationship(
            rt::STRICT_PIVOT_CACHE_DEFINITION.to_string(),
            "../cache/duplicate.xml".to_string(),
            "duplicate".to_string(),
            false,
        );
    assert!(read_pivot_tables(&package).is_err());

    let (mut package, _) = package_with_pivot_table();
    let table_part = package.get_part_mut(&table_uri).unwrap();
    let changed = std::str::from_utf8(table_part.blob())
        .unwrap()
        .replace("cacheId=\"7\"", "cacheId=\"8\"");
    table_part.set_blob(changed.into_bytes());
    assert!(read_pivot_tables(&package).is_err());

    let (mut package, _) = package_with_pivot_table();
    package.add_part(Box::new(BlobPart::new(
        records_uri.clone(),
        ct::SML_WORKSHEET.to_string(),
        Vec::new(),
    )));
    assert!(read_pivot_tables(&package).is_err());

    let (mut package, _) = package_with_pivot_table();
    package.add_part(Box::new(BlobPart::new(
        records_uri,
        ct::SML_PIVOT_CACHE_RECORDS.to_string(),
        br#"<pivotCacheRecords xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" count="2">
                <r><x v="0"/></r><r><s v="South"/></r>
            </pivotCacheRecords>"#
            .to_vec(),
    )));
    assert!(read_pivot_tables(&package).is_err());

    let (mut package, _) = package_with_pivot_table();
    let relationships = package.get_part_mut(&cache_uri).unwrap().rels_mut();
    relationships.remove("rId2").unwrap();
    relationships.add_relationship(
        rt::STRICT_WORKSHEET.to_string(),
        "https://example.com/source.xml".to_string(),
        "rId2".to_string(),
        true,
    );
    assert!(read_pivot_tables(&package).is_err());

    let (mut package, _) = package_with_pivot_table();
    let workbook_part = package.get_part_mut(&workbook_uri).unwrap();
    let changed = std::str::from_utf8(workbook_part.blob()).unwrap().replace(
        r#"<pivotCaches><pivotCache cacheId="7" r:id="rId3"/></pivotCaches>"#,
        "",
    );
    workbook_part.set_blob(changed.into_bytes());
    assert!(read_pivot_tables(&package).is_err());

    let (mut package, _) = package_with_pivot_table();
    package.add_part(Box::new(BlobPart::new(
        cache_uri,
        ct::SML_WORKSHEET.to_string(),
        Vec::new(),
    )));
    assert!(read_pivot_tables(&package).is_err());
}

#[test]
fn parses_prefixed_pivot_table_definition() {
    let xml = r#"<p:pivotTableDefinition
            xmlns:p="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            xmlns:f="urn:foreign" name="Sales &amp; Margin" cacheId="4" dataCaption="Values">
            <f:location ref="XFE1"/>
            <p:location ref="$A$1:$C$5" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
            <p:pivotFields count="2"><p:pivotField name="Region"/><p:pivotField/></p:pivotFields>
            <p:rowFields count="2"><p:field x="-2"/><p:field x="0"/></p:rowFields>
            <p:colFields count="1"><p:field x="1"/></p:colFields>
            <p:pageFields count="1"><p:pageField fld="0"/></p:pageFields>
            <p:dataFields count="1"><p:dataField fld="1" subtotal="average" name="Average Margin"/></p:dataFields>
        </p:pivotTableDefinition>"#;
    let table = read_pivot_table_definition(xml).unwrap().unwrap();

    assert_eq!(table.name, "Sales & Margin");
    assert_eq!(table.location_ref, "$A$1:$C$5");
    assert_eq!(table.field_names, ["Region", "Field1"]);
    assert_eq!(table.row_fields.len(), 1);
    assert_eq!(table.column_fields[0].field_name, "Field1");
    assert_eq!(table.filter_fields[0].field_name, "Region");
    assert_eq!(table.data_fields[0].function, PivotValueFunction::Average);
}

#[test]
fn rejects_malformed_pivot_table_definitions() {
    const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    let root = |body: &str, attributes: &str| {
        format!(
            r#"<pivotTableDefinition xmlns="{S}" name="P" cacheId="1" dataCaption="V" {attributes}>{body}</pivotTableDefinition>"#
        )
    };
    let location =
        r#"<location ref="A1:B2" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>"#;
    let invalid = [
        format!(
            r#"<pivotTableDefinition xmlns="{S}" name="P" cacheId="1">{location}</pivotTableDefinition>"#
        ),
        root("", ""),
        root(
            r#"<location ref="XFE1" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>"#,
            "",
        ),
        root(
            &format!(r#"{location}<pivotFields count="2"><pivotField/></pivotFields>"#),
            "",
        ),
        root(
            &format!(
                r#"{location}<location ref="A1" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>"#
            ),
            "",
        ),
        root(
            &format!(
                r#"{location}<dataFields><dataField fld="0" subtotal="median"/></dataFields>"#
            ),
            "",
        ),
        root(
            &format!(
                r#"{location}<pivotFields><pivotField/></pivotFields><rowFields><field x="1"/></rowFields>"#
            ),
            "",
        ),
    ];
    for xml in invalid {
        assert!(read_pivot_table_definition(&xml).is_err(), "accepted {xml}");
    }
    assert!(
        read_pivot_table_definition(r#"<pivotTableDefinition xmlns="urn:foreign"/>"#)
            .unwrap()
            .is_none()
    );
}

#[test]
fn parses_prefixed_pivot_cache_definition_and_shared_items() {
    let xml = r##"<p:pivotCacheDefinition
            xmlns:p="http://purl.oclc.org/ooxml/spreadsheetml/main"
            xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"
            xmlns:f="urn:foreign" r:id="records" invalid="true" saveData="0"
            refreshOnLoad="1" optimizeMemory="false" enableRefresh="0"
            refreshedBy="Alice &amp; Bob" refreshedDate="42.5" backgroundQuery="true"
            missingItemsLimit="10" createdVersion="7" recordCount="6"
            upgradeOnRefresh="1" tupleCache="0" supportSubquery="true">
            <f:cacheSource type="worksheet"><p:worksheetSource ref="XFE1"/></f:cacheSource>
            <p:cacheSource type="worksheet" connectionId="8"><p:worksheetSource
                sheet="Data &amp; More" ref="$A$1:$B$4" r:id="source-sheet"/></p:cacheSource>
            <p:cacheFields count="2">
                <p:cacheField name="Region" caption="Area" databaseField="false"
                        uniqueList="0" numFmtId="4" formula="x" sqlType="-1"
                        hierarchy="2" level="3" mappingCount="4" memberPropertyField="true">
                    <p:sharedItems count="6"><p:m/><p:n v="2.5"/><p:b v="true"/>
                        <p:e v="#N/A"/><p:s v="North &amp; West"/><p:d v="2026-07-14T00:00:00Z"/>
                    </p:sharedItems>
                </p:cacheField>
                <p:cacheField name="Sales"/>
            </p:cacheFields>
        </p:pivotCacheDefinition>"##;
    let cache = read_pivot_cache_definition(xml).unwrap().unwrap();

    assert_eq!(cache.id.as_deref(), Some("records"));
    assert!(cache.invalid);
    assert!(!cache.save_data);
    assert_eq!(cache.refreshed_by.as_deref(), Some("Alice & Bob"));
    assert_eq!(cache.source_worksheet.as_deref(), Some("Data & More"));
    assert_eq!(cache.source_ref.as_deref(), Some("$A$1:$B$4"));
    assert_eq!(cache.source_connection_id, Some(8));
    assert_eq!(
        cache.source_relationship_id.as_deref(),
        Some("source-sheet")
    );
    assert_eq!(cache.cache_fields.len(), 2);
    let field = &cache.cache_fields[0];
    assert!(!field.database_field);
    assert_eq!(field.sql_type, Some(-1));
    assert_eq!(field.mapping_count, Some(4));
    assert_eq!(field.member_property_field, Some(true));
    assert_eq!(field.shared_items.len(), 6);
    assert!(matches!(field.shared_items[0], Item::Missing));
    assert!(matches!(field.shared_items[1], Item::Number(2.5)));
    assert!(matches!(field.shared_items[2], Item::Boolean(true)));
    assert!(matches!(&field.shared_items[4], Item::String(value) if value == "North & West"));
}

#[test]
fn rejects_malformed_pivot_cache_definitions() {
    const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let valid_source = r#"<cacheSource type="worksheet"><worksheetSource sheet="Data" ref="A1:B2"/></cacheSource>"#;
    let invalid = [
        format!(r#"<pivotCacheDefinition xmlns="{S}"><cacheFields/></pivotCacheDefinition>"#),
        format!(
            r#"<pivotCacheDefinition xmlns="{S}" invalid="yes">{valid_source}<cacheFields/></pivotCacheDefinition>"#
        ),
        format!(
            r#"<pivotCacheDefinition xmlns="{S}"><cacheSource type="bad"/><cacheFields/></pivotCacheDefinition>"#
        ),
        format!(
            r#"<pivotCacheDefinition xmlns="{S}">{valid_source}<cacheFields count="2"><cacheField name="One"/></cacheFields></pivotCacheDefinition>"#
        ),
        format!(
            r#"<pivotCacheDefinition xmlns="{S}">{valid_source}<cacheSource type="worksheet"/><cacheFields/></pivotCacheDefinition>"#
        ),
        format!(
            r#"<pivotCacheDefinition xmlns="{S}"><cacheFields/>{valid_source}</pivotCacheDefinition>"#
        ),
        format!(
            r#"<pivotCacheDefinition xmlns="{S}">{valid_source}<cacheFields><cacheField/></cacheFields></pivotCacheDefinition>"#
        ),
        format!(
            r#"<pivotCacheDefinition xmlns="{S}">{valid_source}<cacheFields><cacheField name="One"><sharedItems><n v="NaN"/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>"#
        ),
        format!(
            r#"<pivotCacheDefinition xmlns="{S}">{valid_source}<cacheFields><cacheField name="One"><sharedItems count="1"/></cacheField></cacheFields></pivotCacheDefinition>"#
        ),
        format!(
            r#"<pivotCacheDefinition xmlns="{S}" xmlns:r="{R}" r:id="">{valid_source}<cacheFields/></pivotCacheDefinition>"#
        ),
        format!(
            r#"<pivotCacheDefinition xmlns="{S}" xmlns:r="{R}"><cacheSource type="worksheet"><worksheetSource r:id=""/></cacheSource><cacheFields/></pivotCacheDefinition>"#
        ),
    ];
    for xml in invalid {
        assert!(read_pivot_cache_definition(&xml).is_err(), "accepted {xml}");
    }
    assert!(
        read_pivot_cache_definition(r#"<pivotCacheDefinition xmlns="urn:foreign"/>"#)
            .unwrap()
            .is_none()
    );
}

#[test]
fn parses_prefixed_pivot_cache_records() {
    let xml = r##"<p:pivotCacheRecords
            xmlns:p="http://purl.oclc.org/ooxml/spreadsheetml/main"
            xmlns:f="urn:foreign" count="2">
            <f:r><p:n v="99"/></f:r>
            <p:r><p:x v="3"/><p:m/><p:n v="2.5"/><p:b v="false"/>
                <p:e v="#N/A"/><p:s v="North &amp; West"/><p:d v="2026-07-14T00:00:00Z"/></p:r>
            <p:r/>
        </p:pivotCacheRecords>"##;
    let records = read_pivot_cache_records(xml).unwrap().unwrap();

    assert_eq!(records.records.len(), 2);
    let values = &records.records[0].values;
    assert_eq!(values.len(), 7);
    assert!(matches!(values[0], Item::Index(3)));
    assert!(matches!(values[1], Item::Missing));
    assert!(matches!(values[2], Item::Number(2.5)));
    assert!(matches!(values[3], Item::Boolean(false)));
    assert!(matches!(&values[5], Item::String(value) if value == "North & West"));
    assert!(records.records[1].values.is_empty());
}

#[test]
fn rejects_malformed_pivot_cache_records() {
    const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    for xml in [
        format!(r#"<pivotCacheRecords xmlns="{S}"/>"#),
        format!(r#"<pivotCacheRecords xmlns="{S}" count="2"><r/></pivotCacheRecords>"#),
        format!(r#"<pivotCacheRecords xmlns="{S}" count="1"><r><x/></r></pivotCacheRecords>"#),
        format!(
            r#"<pivotCacheRecords xmlns="{S}" count="1"><r><b v="yes"/></r></pivotCacheRecords>"#
        ),
        format!(
            r#"<pivotCacheRecords xmlns="{S}" count="1"><r><n v="NaN"/></r></pivotCacheRecords>"#
        ),
    ] {
        assert!(read_pivot_cache_records(&xml).is_err(), "accepted {xml}");
    }
    assert!(
        read_pivot_cache_records(r#"<pivotCacheRecords xmlns="urn:foreign" count="0"/>"#)
            .unwrap()
            .is_none()
    );
}
