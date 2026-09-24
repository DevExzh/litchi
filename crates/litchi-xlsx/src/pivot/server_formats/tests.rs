use super::*;
use litchi_opc::TargetMode;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::BlobPart;

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const X15: &str = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
const X14: &str = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

fn package() -> OpcPackage {
    let mut package = OpcPackage::new();
    let workbook_uri = PackURI::new("/xl/workbook.xml").unwrap();
    let worksheet_uri = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
    let table_uri = PackURI::new("/xl/pivotTables/pivotTable1.xml").unwrap();
    let cache_uri = PackURI::new("/xl/pivotCache/pivotCacheDefinition1.xml").unwrap();
    let connections_uri = PackURI::new("/xl/connections.xml").unwrap();

    let workbook_xml = format!(
        r#"<workbook xmlns="{SML}" xmlns:r="{REL}" xmlns:x15="{X15}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets><pivotCaches><pivotCache cacheId="7" r:id="rId2"/></pivotCaches><extLst><ext uri="{PIVOT_TABLE_REFERENCES_URI}"><x15:pivotTableReferences><x15:pivotTableReference r:id="rId3"/></x15:pivotTableReferences></ext></extLst></workbook>"#
    );
    let mut workbook = BlobPart::new(
        workbook_uri,
        ct::SML_SHEET_MAIN.into(),
        workbook_xml.into_bytes(),
    );
    workbook.relate_to("worksheets/sheet1.xml", rt::WORKSHEET);
    workbook.relate_to(
        "pivotCache/pivotCacheDefinition1.xml",
        rt::PIVOT_CACHE_DEFINITION,
    );
    workbook.relate_to("pivotTables/pivotTable1.xml", rt::PIVOT_TABLE);
    workbook.relate_to("connections.xml", CONNECTIONS_RELATIONSHIP);

    let worksheet = BlobPart::new(
        worksheet_uri,
        ct::SML_WORKSHEET.into(),
        format!(r#"<worksheet xmlns="{SML}"><sheetData/></worksheet>"#).into_bytes(),
    );
    let cache_xml = format!(
        r#"<pivotCacheDefinition xmlns="{SML}" xmlns:x15="{X15}" xmlns:x14="{X14}"><cacheSource type="external"><extLst><ext uri="{CACHE_SOURCE_URI}"><x14:sourceConnection name="External"/></ext></extLst></cacheSource><extLst><ext uri="{PIVOT_CACHE_ID_VERSION_URI}"><x15:pivotCacheIdVersion cacheIdSupportedVersion="15" cacheIdCreatedVersion="15"/></ext></extLst></pivotCacheDefinition>"#
    );
    let cache = BlobPart::new(
        cache_uri,
        ct::SML_PIVOT_CACHE_DEFINITION.into(),
        cache_xml.into_bytes(),
    );
    let table_xml = format!(
        r#"<pivotTableDefinition xmlns="{SML}" xmlns:x15="{X15}" name="Pivot" cacheId="7"><location ref="A1:C5"/><extLst><ext uri="{PIVOT_TABLE_SERVER_FORMATS_URI}"><x15:pivotTableServerFormats count="2"><x15:serverFormat culture="en-US" format="0.00"/><x15:serverFormat/></x15:pivotTableServerFormats></ext></extLst></pivotTableDefinition>"#
    );
    let mut table = BlobPart::new(
        table_uri,
        ct::SML_PIVOT_TABLE.into(),
        table_xml.into_bytes(),
    );
    table.relate_to(
        "../pivotCache/pivotCacheDefinition1.xml",
        rt::PIVOT_CACHE_DEFINITION,
    );
    let connections = BlobPart::new(
        connections_uri,
        CONNECTIONS_CONTENT_TYPE.into(),
        format!(r#"<connections xmlns="{SML}"><connection id="8" name="External"/></connections>"#)
            .into_bytes(),
    );

    package.relate_to("xl/workbook.xml", rt::OFFICE_DOCUMENT);
    package.add_part(Box::new(workbook));
    package.add_part(Box::new(worksheet));
    package.add_part(Box::new(cache));
    package.add_part(Box::new(table));
    package.add_part(Box::new(connections));
    package
}

#[test]
fn reads_and_edits_server_formats() {
    let mut package = package();
    let snapshot = Snapshot::load(&package, PivotTableSelector::Name("Pivot")).unwrap();
    assert_eq!(snapshot.formats()[0].culture.as_deref(), Some("en-US"));
    assert_eq!(snapshot.formats()[1], ServerFormat::new(None, None));

    let before = package
        .get_part(snapshot.table_part())
        .unwrap()
        .blob()
        .to_vec();
    let mut transaction = Transaction::new(&mut package, "Pivot").unwrap();
    transaction
        .update_server_format(
            0,
            ServerFormatEdit {
                culture: AttributeEdit::Clear,
                format: AttributeEdit::Set("#,##0".to_owned()),
            },
        )
        .unwrap();
    transaction.set_format(1, Some("0".to_owned())).unwrap();
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    let changed = package.get_part(snapshot.table_part()).unwrap().blob();
    assert!(std::str::from_utf8(changed).unwrap().contains(r#"format="#));
    assert!(
        !std::str::from_utf8(changed)
            .unwrap()
            .contains(r#"culture="en-US"#)
    );
    assert_ne!(changed, before.as_slice());
    assert_eq!(
        commit.snapshot().formats()[0].format.as_deref(),
        Some("#,##0")
    );

    let patch = commit.patch().clone();
    patch.inverse().apply(&mut package).unwrap();
    assert_eq!(
        package.get_part(snapshot.table_part()).unwrap().blob(),
        before
    );
    patch.apply(&mut package).unwrap();
    assert_eq!(
        Snapshot::load(&package, "Pivot").unwrap().formats()[1]
            .format
            .as_deref(),
        Some("0")
    );
}

#[test]
fn source_noop_is_empty() {
    let mut package = package();
    let original = package
        .get_part(&PackURI::new("/xl/pivotTables/pivotTable1.xml").unwrap())
        .unwrap()
        .blob()
        .to_vec();
    let transaction = Transaction::new(&mut package, "Pivot").unwrap();
    let commit = transaction.commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    assert_eq!(
        package
            .get_part(&PackURI::new("/xl/pivotTables/pivotTable1.xml").unwrap())
            .unwrap()
            .blob(),
        original
    );
}

#[test]
fn ordered_list_edits_update_count_and_round_trip_exactly() {
    let mut package = package();
    let table_uri = PackURI::new("/xl/pivotTables/pivotTable1.xml").unwrap();
    let original = package.get_part(&table_uri).unwrap().blob().to_vec();

    let mut transaction = Transaction::new(&mut package, "Pivot").unwrap();
    transaction
        .insert_server_format(
            1,
            ServerFormat::new(Some("fr-FR".to_owned()), Some("& <".to_owned())),
        )
        .unwrap();
    transaction.move_server_format(2, 0).unwrap();
    let commit = transaction.commit().unwrap();
    let changed = package.get_part(&table_uri).unwrap().blob();
    let changed_text = std::str::from_utf8(changed).unwrap();
    assert!(changed_text.contains("count=\"3\""));
    assert!(changed_text.contains("culture=\"fr-FR\""));
    assert!(changed_text.contains("format=\"&amp; &lt;\""));
    assert_eq!(commit.snapshot().formats().len(), 3);
    assert_eq!(
        commit.snapshot().formats()[0],
        ServerFormat::new(None, None)
    );

    commit.patch().inverse().apply(&mut package).unwrap();
    assert_eq!(package.get_part(&table_uri).unwrap().blob(), original);
}

#[test]
fn ordered_list_reorder_preserves_leaf_source_and_count_lexical_bytes() {
    let mut package = package();
    let table_uri = PackURI::new("/xl/pivotTables/pivotTable1.xml").unwrap();
    let original = package.get_part(&table_uri).unwrap().blob().to_vec();
    let mut transaction = Transaction::new(&mut package, "Pivot").unwrap();
    transaction.reorder_server_formats(&[1, 0]).unwrap();
    let commit = transaction.commit().unwrap();
    let changed = package.get_part(&table_uri).unwrap().blob();
    let original_text = std::str::from_utf8(&original).unwrap();
    let changed_text = std::str::from_utf8(changed).unwrap();
    assert_eq!(changed_text.matches("count=\"2\"").count(), 1);
    assert_eq!(changed_text.matches("culture=\"en-US\"").count(), 1);
    assert!(
        changed_text.find("<x15:serverFormat/>").unwrap()
            < changed_text.find("culture=\"en-US\"").unwrap()
    );
    assert!(original_text.contains("<x15:serverFormat culture=\"en-US\" format=\"0.00\"/>"));
    assert_eq!(
        commit.snapshot().formats()[0],
        ServerFormat::new(None, None)
    );
}

#[test]
fn ordered_list_refuses_removing_required_final_child() {
    let mut package = package();
    let mut transaction = Transaction::new(&mut package, "Pivot").unwrap();
    transaction.remove_server_format(0).unwrap();
    let error = transaction.remove_server_format(0).unwrap_err();
    assert!(error.to_string().contains("retain one server-format child"));
}

#[test]
fn ordinary_workbook_structural_edit_uses_same_public_path() {
    let workbook = Workbook::from_package(package()).unwrap();
    let mut edit = workbook.edit_pivot_table("Pivot").unwrap();
    edit.push_server_format(ServerFormat::new(None, Some("0".to_owned())))
        .unwrap();
    edit.reorder_server_formats(&[2, 0, 1]).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.snapshot().formats().len(), 3);
    assert_eq!(commit.snapshot().formats()[0].format.as_deref(), Some("0"));
    let round_trip = commit.patch().inverse().apply(commit.workbook()).unwrap();
    assert_eq!(round_trip.snapshot().formats().len(), 2);
}

#[test]
fn worksheet_metadata_deduplicates_canonical_target_edges() {
    let mut package = package();
    let worksheet_uri = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
    package
        .get_part_mut(&worksheet_uri)
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::PIVOT_TABLE.to_owned(),
            "../pivotTables/pivotTable1.xml".to_owned(),
            "rIdWorksheetPivotA".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();
    package
        .get_part_mut(&worksheet_uri)
        .unwrap()
        .rels_mut()
        .try_add_relationship(
            rt::PIVOT_TABLE.to_owned(),
            "../pivotTables/pivotTable1.xml".to_owned(),
            "rIdWorksheetPivotB".to_owned(),
            TargetMode::Internal,
        )
        .unwrap();

    let workbook = package.main_document_part().unwrap();
    let catalog = raw::parse_catalog(workbook.blob()).unwrap();
    let relationship_index = RelationshipIndex::build(&package).unwrap();
    let names = worksheet_pivot_names(&package, &catalog, &relationship_index).unwrap();
    assert_eq!(names, vec!["Pivot"]);
}

#[test]
fn workbook_contextual_edit_returns_immutable_commit_and_patch() {
    let workbook = Workbook::from_package(package()).unwrap();
    let view = workbook.pivot_table("Pivot").unwrap();
    assert_eq!(view.table_name(), "Pivot");
    assert_eq!(
        view.server_formats().formats()[0].culture.as_deref(),
        Some("en-US")
    );

    let mut edit = workbook.edit_pivot_table("Pivot").unwrap();
    edit.update_server_format(
        0,
        ServerFormatEdit {
            culture: AttributeEdit::Clear,
            format: AttributeEdit::Set("#,##0".to_owned()),
        },
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(
        commit.snapshot().formats()[0].format.as_deref(),
        Some("#,##0")
    );

    let next = commit.workbook();
    assert_eq!(
        next.pivot_table("Pivot").unwrap().formats()[0]
            .format
            .as_deref(),
        Some("#,##0")
    );
    let inverse = commit.patch().inverse();
    let restored = inverse.apply(next).unwrap();
    assert_eq!(
        restored.snapshot().formats()[0].culture.as_deref(),
        Some("en-US")
    );
}

#[test]
fn scanner_admits_tree_before_resolving_a_second_root() {
    let source = br#"<pivotTableDefinition/><bad:second xmlns:xml="urn:not-xml"/>"#;
    let error = scan_xml(source, "pivotTableDefinition", ReadLimits::default()).unwrap_err();
    assert!(
        error.to_string().contains("multiple roots"),
        "second-root admission must precede namespace validation: {error}"
    );
}
