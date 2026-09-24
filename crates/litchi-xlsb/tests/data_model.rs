use std::io::Cursor;

use litchi_opc::{BlobPart, PackURI, Part, TargetMode};
use litchi_xlsb::Package;
use litchi_xlsb::data_model::{
    ContentType, DATA_MODEL_CONTENT_TYPE, DATA_MODEL_PART_NAME, Definition, Model, Table,
    TimeGrouping, TimeGroupingColumn,
};
use litchi_xlsb::package::connections::{Connection, Connections, SourceType};
use litchi_xlsb::writer::{MutableWorksheet, WorkbookWriter};

const EMPTY_RELATIONSHIPS: &[u8] =
    br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>"#;

fn definition() -> Definition {
    Definition {
        min_version_load: 5,
        tables: vec![Table {
            id: "table-1".to_string(),
            name: "Sales".to_string(),
            connection: "Connection".to_string(),
        }],
        relationships: Vec::new(),
        time_groupings: Vec::new(),
    }
}

fn time_grouping() -> TimeGrouping {
    TimeGrouping {
        table_name: "Sales".into(),
        column_name: "Date".into(),
        column_id: "Date".into(),
        columns: vec![TimeGroupingColumn {
            is_selected: true,
            content_type: ContentType::Years,
            column_name: "Date.Year".into(),
            column_id: "Date.Year".into(),
        }],
    }
}

fn connections() -> Connections {
    Connections {
        connections: vec![Connection {
            connection_id: 1,
            source_type: SourceType::Odbc,
            name: "Connection".to_string(),
            ..Connection::default()
        }],
    }
}

fn package() -> Package {
    let mut writer = WorkbookWriter::new();
    writer.add_worksheet(MutableWorksheet::new("Sheet1"));
    writer.set_connections(connections()).expect("connections");
    writer
        .set_data_model(Model::from_bytes(definition(), vec![1, 2, 3, 4]).expect("model"))
        .expect("set model");
    let mut output = Cursor::new(Vec::new());
    writer.save(&mut output).expect("save");
    Package::from_bytes(output.into_inner()).expect("package")
}

fn empty_package() -> Package {
    let mut writer = WorkbookWriter::new();
    writer.add_worksheet(MutableWorksheet::new("Sheet1"));
    writer.set_connections(connections()).expect("connections");
    let mut output = Cursor::new(Vec::new());
    writer.save(&mut output).expect("save");
    Package::from_bytes(output.into_inner()).expect("package")
}

fn package_with_case_variant_model_part() -> Package {
    let mut opc = package().into_opc();
    let canonical = PackURI::new(DATA_MODEL_PART_NAME).expect("canonical model URI");
    let variant = PackURI::new("/xl/MODEL/ITEM.DATA").expect("variant model URI");
    let (content_type, bytes) = {
        let part = opc.get_part(&canonical).expect("model part");
        (part.content_type().to_owned(), part.blob().to_vec())
    };
    assert!(opc.remove_part(&canonical));
    opc.add_part(Box::new(BlobPart::new(variant, content_type, bytes)));
    Package::from_opc(opc).expect("case-equivalent model package")
}

fn package_with_case_variant_connection_part() -> Package {
    let mut opc = package().into_opc();
    let canonical = PackURI::new("/xl/connections.bin").expect("canonical connections URI");
    let variant = PackURI::new("/xl/CONNECTIONS.BIN").expect("variant connections URI");
    let (content_type, bytes) = {
        let part = opc.get_part(&canonical).expect("connections part");
        (part.content_type().to_owned(), part.blob().to_vec())
    };
    assert!(opc.remove_part(&canonical));
    opc.add_part(Box::new(BlobPart::new(variant, content_type, bytes)));
    Package::from_opc(opc).expect("case-equivalent connections package")
}

fn package_with_explicit_empty_model_relationships() -> Package {
    package_with_explicit_empty_model_relationships_with_padding(0)
}

fn package_with_explicit_empty_model_relationships_with_padding(padding: usize) -> Package {
    let package = package().into_opc();
    let content_types = noncanonical_content_types_with_padding(
        package
            .source_content_types()
            .expect("content types")
            .bytes(),
        padding,
    );
    let root = PackURI::new("/").expect("root URI");
    let root_relationships = package.source_relationships(&root).expect("root rels");
    let model = PackURI::new(DATA_MODEL_PART_NAME).expect("model URI");
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", &content_types)
        .expect("content types member");
    writer
        .write_stored("_rels/.rels", root_relationships.bytes())
        .expect("root relationships member");
    let mut parts = package
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
        .collect::<Vec<_>>();
    parts.sort_unstable_by_key(|part| part.partname().as_str());
    for part in parts {
        writer
            .write_stored(part.partname().membername(), part.blob())
            .expect("part member");
        let relationships = package
            .source_relationships(part.partname())
            .expect("part relationships");
        if part.partname().is_equivalent_to(&model) {
            writer
                .write_stored(
                    part.partname()
                        .rels_uri()
                        .expect("model rels URI")
                        .membername(),
                    EMPTY_RELATIONSHIPS,
                )
                .expect("explicit empty model relationships member");
        } else if relationships.member_present() || !part.rels().is_empty() {
            writer
                .write_stored(
                    part.partname()
                        .rels_uri()
                        .expect("relationships URI")
                        .membername(),
                    relationships.bytes(),
                )
                .expect("relationships member");
        }
    }
    let zip = writer.finish_to_bytes().expect("opened ZIP fixture bytes");
    let opc = if content_types.len() > litchi_opc::ReadLimits::default().max_content_types_bytes() {
        litchi_opc::OpcPackage::from_bytes_with_limits(
            &zip,
            elevated_opc_limits(content_types.len()),
        )
        .expect("open elevated ZIP fixture")
    } else {
        litchi_opc::OpcPackage::from_bytes(&zip).expect("open ZIP fixture")
    };
    Package::from_opc(opc).expect("package from ZIP fixture")
}

fn noncanonical_content_types_with_padding(source: &[u8], padding: usize) -> Vec<u8> {
    let source = std::str::from_utf8(source).expect("content types UTF-8");
    let root = source.find("<Types").expect("content types root");
    let mut xml = String::with_capacity(source.len() + padding + 64);
    xml.push_str(&source[..root]);
    xml.push_str("<?litchi-preserve?><!---->");
    xml.extend(std::iter::repeat_n(' ', padding));
    xml.push_str(&source[root..]);
    let xml = swap_content_type_attributes(xml, "Default", "Extension", "ContentType");
    swap_content_type_attributes(xml, "Override", "PartName", "ContentType").into_bytes()
}

fn elevated_opc_limits(content_types_bytes: usize) -> litchi_opc::ReadLimits {
    litchi_opc::ReadLimits::builder()
        .max_content_types_bytes(content_types_bytes)
        .expect("content types limit")
        .build()
        .expect("elevated OPC limits")
}

fn swap_content_type_attributes(
    mut xml: String,
    element: &str,
    first_attribute: &str,
    second_attribute: &str,
) -> String {
    let prefix = format!("<{element} {first_attribute}=\"");
    let start = xml.find(&prefix).expect("content type declaration");
    let first_start = start + prefix.len();
    let first_end = first_start + xml[first_start..].find('"').expect("first attribute");
    let second_marker = format!(" {second_attribute}=\"");
    let second_start = first_end
        + xml[first_end..]
            .find(&second_marker)
            .expect("second attribute")
        + second_marker.len();
    let second_end = second_start + xml[second_start..].find('"').expect("second value");
    let element_end = second_end + xml[second_end..].find("/>").expect("element end") + 2;
    let replacement = format!(
        "<{element} {second_attribute}=\"{}\" {first_attribute}=\"{}\"/>",
        &xml[second_start..second_end],
        &xml[first_start..first_end],
    );
    xml.replace_range(start..element_end, &replacement);
    xml
}

fn authored_opc(source: litchi_opc::OpcPackage) -> litchi_opc::OpcPackage {
    let mut authored = litchi_opc::OpcPackage::new();
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

fn signed_package() -> Package {
    let mut package = authored_opc(package().into_opc());
    let origin = PackURI::new("/_xmlsignatures/origin.sigs").expect("signature origin URI");
    let signature = PackURI::new("/_xmlsignatures/sig1.xml").expect("signature URI");
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
    Package::from_opc(package).expect("signed package")
}

#[test]
fn writer_emits_relationship_free_model_part_and_typed_workbook_records() {
    let package = package();
    let snapshot = package.data_model().expect("data model");
    let model = snapshot.model().expect("model present");
    assert_eq!(model.definition, definition());
    assert_eq!(model.part.bytes(), &[1, 2, 3, 4]);
    assert_eq!(
        snapshot.connection_names(),
        Some(&["Connection".to_string()][..])
    );
    assert!(
        package
            .opc_package()
            .get_part(&PackURI::new("/xl/model/item.data").expect("URI"))
            .expect("part")
            .rels()
            .is_empty()
    );
}

#[test]
fn writer_requires_a_connection_for_each_model_table() {
    let mut writer = WorkbookWriter::new();
    writer.add_worksheet(MutableWorksheet::new("Sheet1"));
    writer
        .set_data_model(Model::from_bytes(definition(), vec![1, 2, 3]).expect("model"))
        .expect("set model before its owner");
    let mut output = Cursor::new(Vec::new());
    let error = writer.save(&mut output).expect_err("missing owner");
    assert!(error.to_string().contains("External Data Connections part"));
}

#[test]
fn table_connection_binding_is_case_insensitive() {
    let mut value = definition();
    value.tables[0].connection = "connection".to_string();
    let mut writer = WorkbookWriter::new();
    writer.add_worksheet(MutableWorksheet::new("Sheet1"));
    writer.set_connections(connections()).expect("connections");
    writer
        .set_data_model(Model::from_bytes(value.clone(), vec![1, 2, 3]).expect("model"))
        .expect("case-insensitive binding");
    let mut output = Cursor::new(Vec::new());
    writer.save(&mut output).expect("save");
    let reopened = Package::from_bytes(output.into_inner()).expect("package");
    assert_eq!(
        reopened
            .data_model()
            .expect("snapshot")
            .definition()
            .expect("definition"),
        &value
    );
}

#[test]
fn load_version_edit_is_source_checked_and_reversible() {
    let package = package();
    let before = package.data_model().expect("before");
    let mut transaction = before.edit();
    transaction
        .edit_definition(|definition| {
            definition.min_version_load = 6;
            Ok(())
        })
        .expect("version edit");
    let commit = transaction.commit().expect("commit");
    let changed = package.apply_data_model(&commit).expect("apply");
    assert_eq!(
        changed
            .data_model()
            .expect("readback")
            .definition()
            .expect("definition")
            .min_version_load,
        6
    );

    let restored = changed
        .apply_data_model_patch(&commit.patch().inverse())
        .expect("inverse");
    assert_eq!(restored.data_model().expect("restored"), before);
}

#[test]
fn empty_transaction_is_an_exact_noop() {
    let package = package();
    let before = package.data_model().expect("before");
    let commit = before.edit().commit().expect("no-op commit");
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    let after = package.apply_data_model(&commit).expect("no-op apply");
    assert_eq!(after.data_model().expect("after"), before);
}

#[test]
fn identity_edits_are_refused_while_payload_is_opaque() {
    let package = package();
    let before = package.data_model().expect("before");
    let mut transaction = before.edit();
    let error = transaction
        .edit_definition(|definition| {
            definition.tables[0].name = "Renamed".to_string();
            Ok(())
        })
        .expect_err("opaque payload dependency refusal");
    assert!(error.to_string().contains("opaque"));
}

#[test]
fn connection_source_change_is_stale_even_when_model_and_workbook_bytes_are_unchanged() {
    let package = package();
    let before = package.data_model().expect("before");
    let mut transaction = before.edit();
    transaction
        .edit_definition(|definition| {
            definition.min_version_load = 6;
            Ok(())
        })
        .expect("version edit");
    let commit = transaction.commit().expect("commit");

    let workbook_uri = PackURI::new("/xl/workbook.bin").expect("workbook URI");
    let model_uri = PackURI::new(DATA_MODEL_PART_NAME).expect("model URI");
    let workbook_bytes = package
        .opc_package()
        .get_part(&workbook_uri)
        .expect("workbook")
        .blob()
        .to_vec();
    let model_bytes = package
        .opc_package()
        .get_part(&model_uri)
        .expect("model")
        .blob()
        .to_vec();

    let mut workbook = package.workbook().expect("workbook");
    let mut changed = workbook.connections().expect("connections").clone();
    changed.connections[0].description = Some("changed source metadata".to_string());
    workbook.set_connections(changed).expect("change source");
    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).expect("save changed source");
    let changed_package = Package::from_bytes(output.into_inner()).expect("reopen");

    assert_eq!(
        changed_package
            .opc_package()
            .get_part(&workbook_uri)
            .expect("workbook")
            .blob(),
        workbook_bytes.as_slice()
    );
    assert_eq!(
        changed_package
            .opc_package()
            .get_part(&model_uri)
            .expect("model")
            .blob(),
        model_bytes.as_slice()
    );
    assert!(changed_package.apply_data_model(&commit).is_err());
}

#[test]
fn creation_removal_and_stale_source_checks_are_atomic() {
    let empty = empty_package();
    assert!(!empty.data_model().expect("empty model").is_present());

    let mut create = empty.data_model().expect("empty model").edit();
    create
        .replace_model(Some(
            Model::from_bytes(definition(), vec![9, 8, 7]).expect("model"),
        ))
        .expect("create");
    let create_commit = create.commit().expect("create commit");
    let created = empty
        .apply_data_model(&create_commit)
        .expect("create apply");
    assert!(created.data_model().expect("created model").is_present());
    assert!(created.apply_data_model(&create_commit).is_err());

    let mut remove = created.data_model().expect("created model").edit();
    remove.remove_model().expect("remove");
    let remove_commit = remove.commit().expect("remove commit");
    let removed = created
        .apply_data_model(&remove_commit)
        .expect("remove apply");
    assert!(!removed.data_model().expect("removed model").is_present());
    let restored = removed
        .apply_data_model_patch(&remove_commit.patch().inverse())
        .expect("inverse");
    assert_eq!(
        restored.data_model().expect("restored"),
        created.data_model().expect("created")
    );
}

#[test]
fn workbook_records_and_model_part_are_a_strict_singleton_pair() {
    let with_model = package();
    let mut missing_part = with_model.clone().into_opc();
    missing_part.remove_part(&PackURI::new(DATA_MODEL_PART_NAME).expect("URI"));
    let missing_part = Package::from_opc(missing_part).expect("generic package");
    assert!(missing_part.data_model().is_err());

    let mut orphan_part = empty_package().into_opc();
    orphan_part.add_part(Box::new(BlobPart::new(
        PackURI::new(DATA_MODEL_PART_NAME).expect("URI"),
        DATA_MODEL_CONTENT_TYPE.to_string(),
        vec![1, 2, 3],
    )));
    let orphan_part = Package::from_opc(orphan_part).expect("generic package");
    assert!(orphan_part.data_model().is_err());

    let mut inbound = package().into_opc();
    let worksheet = PackURI::new("/xl/worksheets/sheet1.bin").expect("worksheet URI");
    inbound
        .get_part_mut(&worksheet)
        .expect("worksheet")
        .relate_to(DATA_MODEL_PART_NAME, "http://example.invalid/model");
    let inbound = Package::from_opc(inbound).expect("generic package");
    assert!(inbound.data_model().is_err());
}

#[test]
fn case_equivalent_model_and_connection_owners_are_resolved_by_physical_name() {
    let variant_model = package_with_case_variant_model_part();
    assert_eq!(
        variant_model
            .data_model()
            .expect("variant model snapshot")
            .part()
            .expect("variant model")
            .part_name(),
        "/xl/MODEL/ITEM.DATA"
    );
    let before = variant_model.data_model().expect("before");
    let mut transaction = before.edit();
    transaction.remove_model().expect("remove variant model");
    let commit = transaction.commit().expect("variant remove commit");
    let removed = variant_model
        .apply_data_model(&commit)
        .expect("remove variant model");
    let restored = removed
        .apply_data_model_patch(&commit.patch().inverse())
        .expect("restore variant model");
    assert_eq!(
        restored
            .data_model()
            .expect("restored variant model")
            .part()
            .expect("restored model")
            .part_name(),
        "/xl/MODEL/ITEM.DATA"
    );
    let variant_connection = package_with_case_variant_connection_part();
    assert_eq!(
        variant_connection
            .data_model()
            .expect("variant connection snapshot")
            .connection_names(),
        Some(&["Connection".to_owned()][..])
    );
    let mut workbook = variant_connection.workbook().expect("variant workbook");
    workbook
        .set_connections(connections())
        .expect("replace physical case-variant connections owner");
    assert_eq!(
        workbook
            .connections()
            .map(|value| value.connections[0].name.as_str()),
        Some("Connection")
    );
    assert!(
        workbook
            .remove_connections()
            .expect("remove connections owner")
    );
    assert!(workbook.connections().is_none());
}

#[test]
fn source_tokens_restore_explicit_empty_model_relationships_after_reopened_zip_inverse() {
    let package = package_with_explicit_empty_model_relationships();
    let model_uri = PackURI::new(DATA_MODEL_PART_NAME).expect("model URI");
    let model_rels_member = model_uri
        .rels_uri()
        .expect("model relationships URI")
        .membername()
        .to_owned();
    let before_zip = package.to_bytes().expect("opened ZIP bytes");
    let before_archive =
        soapberry_zip::office::ArchiveReader::new(&before_zip).expect("opened ZIP archive");
    let before_content_types = before_archive
        .read("[Content_Types].xml")
        .expect("content types member");
    let before_relationships = before_archive
        .read(&model_rels_member)
        .expect("model relationships member");
    assert!(
        before_content_types
            .windows(b"<?litchi-preserve?>".len())
            .any(|window| { window == b"<?litchi-preserve?>" })
    );
    assert!(
        before_content_types
            .windows(b"<Default ContentType=\"".len())
            .any(|window| window == b"<Default ContentType=\"")
    );
    assert!(
        before_content_types
            .windows(b"<Override ContentType=\"".len())
            .any(|window| window == b"<Override ContentType=\"")
    );
    assert_eq!(before_relationships, EMPTY_RELATIONSHIPS);
    let before = package.data_model().expect("before snapshot");
    let mut transaction = before.edit();
    transaction.remove_model().expect("remove model");
    let commit = transaction.commit().expect("remove commit");
    let removed = package.apply_data_model(&commit).expect("remove model");
    let removed_zip = removed.to_bytes().expect("removed ZIP bytes");
    let reopened_removed = Package::from_bytes(removed_zip).expect("reopen removed ZIP");
    let restored = reopened_removed
        .apply_data_model_patch(&commit.patch().inverse())
        .expect("inverse model");
    let restored_zip = restored.to_bytes().expect("restored ZIP bytes");
    let restored_archive =
        soapberry_zip::office::ArchiveReader::new(&restored_zip).expect("restored ZIP archive");
    assert_eq!(
        restored_archive
            .read("[Content_Types].xml")
            .expect("restored content types member"),
        before_content_types
    );
    assert_eq!(
        restored_archive
            .read(&model_rels_member)
            .expect("restored model relationships member"),
        EMPTY_RELATIONSHIPS
    );
    let reopened = Package::from_bytes(restored_zip).expect("reopen restored ZIP");
    assert_eq!(reopened.data_model().expect("reopened snapshot"), before);
}

#[test]
fn elevated_source_tokens_survive_reopened_inverse_with_large_content_types() {
    const PADDING: usize = 8 * 1024 * 1024;
    let package = package_with_explicit_empty_model_relationships_with_padding(PADDING);
    let model_uri = PackURI::new(DATA_MODEL_PART_NAME).expect("model URI");
    let model_rels_member = model_uri
        .rels_uri()
        .expect("model relationships URI")
        .membername()
        .to_owned();
    let before_zip = package.to_bytes().expect("large source ZIP bytes");
    let before_archive =
        soapberry_zip::office::ArchiveReader::new(&before_zip).expect("large source archive");
    let before_content_types = before_archive
        .read("[Content_Types].xml")
        .expect("large content types member");
    assert!(
        before_content_types.len() > litchi_opc::ReadLimits::default().max_content_types_bytes()
    );
    assert_eq!(
        before_archive
            .read(&model_rels_member)
            .expect("model relationships member"),
        EMPTY_RELATIONSHIPS
    );

    let mut limits = litchi_xlsb::data_model::ReadLimits::DEFAULT;
    limits.max_part_bytes = 16 * 1024 * 1024;
    limits.max_metadata_bytes = 16 * 1024 * 1024;
    let before = package
        .data_model_with_limits(limits)
        .expect("elevated Data Model snapshot");
    let mut transaction = before.edit();
    transaction.remove_model().expect("remove model");
    let commit = transaction.commit().expect("remove commit");
    let removed = package.apply_data_model(&commit).expect("remove model");
    let removed_zip = removed.to_bytes().expect("removed ZIP bytes");
    let reopened_removed_opc = litchi_opc::OpcPackage::from_bytes_with_limits(
        &removed_zip,
        elevated_opc_limits(before_content_types.len()),
    )
    .expect("reopen removed ZIP with elevated limits");
    let reopened_removed = Package::from_opc(reopened_removed_opc).expect("removed package");
    let restored = reopened_removed
        .apply_data_model_patch(&commit.patch().inverse())
        .expect("inverse model");
    let restored_zip = restored.to_bytes().expect("restored ZIP bytes");
    let restored_archive =
        soapberry_zip::office::ArchiveReader::new(&restored_zip).expect("restored ZIP archive");
    assert_eq!(
        restored_archive
            .read("[Content_Types].xml")
            .expect("restored content types member"),
        before_content_types
    );
    assert_eq!(
        restored_archive
            .read(&model_rels_member)
            .expect("restored model relationships member"),
        EMPTY_RELATIONSHIPS
    );
    let reopened_opc = litchi_opc::OpcPackage::from_bytes_with_limits(
        &restored_zip,
        elevated_opc_limits(before_content_types.len()),
    )
    .expect("reopen restored ZIP with elevated limits");
    let reopened = Package::from_opc(reopened_opc).expect("restored package");
    assert_eq!(
        reopened
            .data_model_with_limits(limits)
            .expect("reopened elevated snapshot"),
        before
    );
}

#[test]
fn data_model_limits_bound_borrowed_graph_metadata_before_snapshot_ownership() {
    let package = package();
    let mut limits = litchi_xlsb::data_model::ReadLimits::DEFAULT;
    limits.max_metadata_bytes = 1;
    let error = package
        .data_model_with_limits(limits)
        .expect_err("metadata admission must be bounded");
    assert!(error.to_string().contains("metadata"));
}

#[test]
fn data_model_limits_bound_connection_count_before_typed_load() {
    let package = package();
    let mut limits = litchi_xlsb::data_model::ReadLimits::DEFAULT;
    limits.max_connections = 0;
    let error = package
        .data_model_with_limits(limits)
        .expect_err("connection admission must be bounded");
    assert!(error.to_string().contains("External Data Connections"));
}

#[test]
fn signed_noop_preserves_and_signed_change_requires_explicit_disposition() {
    let signed = signed_package();
    assert!(signed.opc_package().is_signed());
    let before_bytes = signed.to_bytes().expect("signed bytes");
    let noop = signed
        .data_model()
        .expect("signed snapshot")
        .edit()
        .commit()
        .expect("noop");
    let unchanged = signed.apply_data_model(&noop).expect("signed noop");
    assert!(unchanged.opc_package().is_signed());
    assert_eq!(unchanged.to_bytes().expect("noop bytes"), before_bytes);

    let mut edit = signed.data_model().expect("signed snapshot").edit();
    edit.edit_definition(|definition| {
        definition.min_version_load = 6;
        Ok(())
    })
    .expect("stage signed edit");
    let error = edit.commit().expect_err("signed change must refuse");
    assert!(matches!(
        error,
        litchi_xlsb::data_model::Error::Opc(
            litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy
        )
    ));

    let mut workbook = signed.workbook().expect("signed workbook facade");
    let workbook_before = {
        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).expect("signed workbook bytes");
        output.into_inner()
    };
    let workbook_noop = workbook
        .data_model()
        .expect("signed workbook snapshot")
        .edit()
        .commit()
        .expect("signed workbook noop");
    workbook
        .apply_data_model(&workbook_noop)
        .expect("signed workbook noop apply");
    assert!(workbook.opc_package().is_signed());
    let workbook_after = {
        let mut output = Cursor::new(Vec::new());
        workbook.save(&mut output).expect("signed workbook bytes");
        output.into_inner()
    };
    assert_eq!(workbook_after, workbook_before);

    let mut workbook_edit = workbook
        .data_model()
        .expect("signed workbook snapshot")
        .edit();
    workbook_edit
        .edit_definition(|definition| {
            definition.min_version_load = 6;
            Ok(())
        })
        .expect("stage signed workbook edit");
    assert!(matches!(
        workbook_edit.commit(),
        Err(litchi_xlsb::data_model::Error::Opc(
            litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy
        ))
    ));
}

#[test]
fn time_grouping_edit_requires_inner_xldm_identity_proof_before_draft_change() {
    let package = package();
    let mut edit = package.data_model().expect("Data Model snapshot").edit();
    let error = edit
        .add_time_grouping(time_grouping())
        .expect_err("opaque bytes without XLDM metadata cannot admit a grouping");
    assert!(error.to_string().contains("XLDM outer proof failed"));
    assert!(
        edit.definition()
            .expect("staged Data Model definition")
            .time_groupings
            .is_empty()
    );
}

#[test]
fn writer_rejects_time_grouping_before_retaining_unproven_opaque_payload() {
    let mut writer = WorkbookWriter::new();
    writer.add_worksheet(MutableWorksheet::new("Sheet1"));
    writer.set_connections(connections()).expect("connections");
    let mut definition = definition();
    definition.time_groupings.push(time_grouping());
    let error = match writer
        .set_data_model(Model::from_bytes(definition, vec![1, 2, 3, 4]).expect("model"))
    {
        Ok(_) => panic!("writer must prove retained groupings before staging"),
        Err(error) => error,
    };
    assert!(error.to_string().contains("XLDM outer proof failed"));
    assert!(writer.data_model().is_none());
}
