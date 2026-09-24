//! Native version-150 read, descriptor publication, and compatibility boundaries.
#![allow(clippy::unwrap_used, reason = "test assertions panic on failure")]

use litchi_opc::{OpcPackage, PackURI};
use litchi_xlsx::{Error, Package};

const SOURCE: &[u8] = include_bytes!("data/data_model/tdf167689_x15_namespace.xlsx");

#[test]
fn unknown_model_profile_is_retained_on_package_write() {
    let mut raw = OpcPackage::from_bytes(SOURCE).unwrap();
    let name = PackURI::new("/xl/model/item.data").unwrap();
    let mut unknown = raw.get_part(&name).unwrap().blob().to_vec();
    replace_utf16(&mut unknown[..4096], ">150<", ">151<");
    raw.get_part_mut(&name).unwrap().set_blob(unknown);
    let source = raw.get_part(&name).unwrap();
    assert!(matches!(
        litchi_xlsx::package::xldm::inspect(source.blob()),
        Err(Error::Unsupported {
            feature: "XLDM header version"
        })
    ));
    let mut package = Package::from_opc(raw.clone()).unwrap();
    assert!(matches!(
        package.data_model(),
        Err(Error::Unsupported {
            feature: "XLDM header version"
        })
    ));
    assert!(matches!(
        package.edit_data_model(),
        Err(Error::Unsupported {
            feature: "XLDM header version"
        })
    ));
    let mut output = Vec::new();
    package.write_to(&mut output).unwrap();
    let reopened = OpcPackage::from_bytes(&output).unwrap();
    assert_eq!(reopened.get_part(&name).unwrap().blob(), source.blob());
    // Failed feature access must also retain the source descriptor and its
    // linked-table connection, including unsupported extension declarations.
    for path in ["/xl/workbook.xml", "/xl/connections.xml"] {
        let name = PackURI::new(path).unwrap();
        assert_eq!(
            reopened.get_part(&name).unwrap().blob(),
            raw.get_part(&name).unwrap().blob()
        );
    }
}

fn replace_utf16(bytes: &mut [u8], before: &str, after: &str) {
    let encode = |value: &str| {
        value
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect::<Vec<_>>()
    };
    let before = encode(before);
    let after = encode(after);
    assert_eq!(before.len(), after.len());
    let position = bytes
        .windows(before.len())
        .position(|value| value == before)
        .unwrap();
    bytes[position..position + before.len()].copy_from_slice(&after);
}

#[test]
fn native_model_read_descriptor_edit_reopen_and_exact_inverse() {
    check_native_model_descriptor_round_trip(OpcPackage::from_vec(SOURCE.to_vec()).unwrap());
}

#[test]
fn borrowed_native_model_preserves_relationship_xml_through_edit_and_inverse() {
    check_native_model_descriptor_round_trip(OpcPackage::from_bytes(SOURCE).unwrap());
}

#[test]
fn native_relationship_token_removal_and_inverse_preserve_exact_xml() {
    let owner = PackURI::new("/xl/workbook.xml").unwrap();
    let edge = r#"<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/powerPivotData" Target="model/item.data"/>"#;
    for owned in [false, true] {
        let mut package = if owned {
            OpcPackage::from_vec(SOURCE.to_vec()).unwrap()
        } else {
            OpcPackage::from_bytes(SOURCE).unwrap()
        };
        let before = package.source_relationships(&owner).unwrap();
        let after = before
            .without_relationship("rId6", before.bytes().len())
            .unwrap();
        assert_eq!(
            after.bytes(),
            String::from_utf8(before.bytes().to_vec())
                .unwrap()
                .replace(edge, "")
                .as_bytes()
        );
        assert!(package.try_replace_relationships(&before, &after).unwrap());
        // This exercises the OPC primitive; complete model removal also owns
        // workbook definitions, connections, names, and vendor metadata.
        let output = litchi_opc::PackageWriter::to_bytes(&package).unwrap();
        let mut reopened = OpcPackage::from_vec(output).unwrap();
        assert_eq!(reopened.source_relationships(&owner).unwrap(), after);
        assert!(reopened.try_replace_relationships(&after, &before).unwrap());
        let restored = litchi_opc::PackageWriter::to_bytes(&reopened).unwrap();
        let archive = soapberry_zip::office::ArchiveReader::new(&restored).unwrap();
        assert_eq!(
            archive.read("xl/_rels/workbook.xml.rels").unwrap(),
            before.bytes()
        );
        let original = OpcPackage::from_bytes(SOURCE).unwrap();
        for part in original
            .try_iter_parts()
            .map(|part| part.expect("part payload decodes"))
        {
            assert_eq!(
                archive.read(part.partname().membername()).unwrap(),
                part.blob()
            );
        }
    }
}

#[test]
fn native_descriptor_patch_rejects_relationship_lexical_source_changes() {
    let mut package = Package::from_bytes(SOURCE.to_vec()).unwrap();
    let mut edit = package.edit_data_model().unwrap();
    edit.edit_definition(|definition| {
        definition.min_version_load += 1;
        Ok(())
    })
    .unwrap();
    let patch = edit.commit().unwrap().patch().clone();
    let original = soapberry_zip::office::ArchiveReader::new(SOURCE).unwrap();
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    for member in original.file_names() {
        let mut bytes = original.read(member).unwrap();
        if member == "xl/_rels/workbook.xml.rels" {
            bytes = String::from_utf8(bytes)
                .unwrap()
                .replace(
                    "</Relationships>",
                    "<!-- concurrent context --></Relationships>",
                )
                .into_bytes();
        }
        writer.write_deflated(member, &bytes).unwrap();
    }
    let changed = writer.finish_to_bytes().unwrap();
    let mut target = OpcPackage::from_vec(changed.clone()).unwrap();
    assert!(matches!(
        patch.apply(&mut target),
        Err(Error::PatchConflict { .. })
    ));
    assert_eq!(
        litchi_opc::PackageWriter::to_bytes(&target).unwrap(),
        changed
    );
}

#[test]
fn public_structural_descriptor_and_payload_replacement_is_refused_without_inner_identity_proof() {
    let mut package = Package::from_bytes(SOURCE.to_vec()).unwrap();
    let mut edit = package.edit_data_model().unwrap();
    let mut replacement = edit.model().unwrap().to_owned();
    replacement.definition.tables[0].id.push_str("-replacement");
    replacement.definition.extension_list = Some(litchi_xlsx::DataModelOpaqueXml {
        xml: br#"<x15:extLst xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:ext uri="urn:replacement"/></x15:extLst>"#.to_vec(),
    });
    replace_utf16(&mut replacement.payload.data[..4096], ">150<", ">151<");

    assert!(matches!(
        edit.set(replacement),
        Err(Error::Unsupported {
            feature: "structural Data Model edits require validated inner XLDM identity closure"
        })
    ));
    assert!(!edit.is_changed());
}

#[test]
fn public_data_model_writer_enforces_xml10_character_boundaries() {
    let definition = || litchi_xlsx::DataModelDefinition {
        min_version_load: 5,
        tables: vec![litchi_xlsx::DataModelTable {
            id: "id".into(),
            name: "table".into(),
            connection: "connection".into(),
        }],
        relationships: Vec::new(),
        extension_list: None,
    };

    for forbidden in ['\0', '\u{1}', '\u{B}', '\u{FFFE}'] {
        let mut value = definition();
        value.tables[0].name = format!("table{forbidden}");
        assert!(matches!(
            litchi_xlsx::write_data_model(&value),
            Err(Error::Invalid(message)) if message.contains("XML 1.0-forbidden character")
        ));
    }

    let valid = "\t\n\r \u{D7FF}\u{E000}\u{FFFD}\u{10000}\u{10FFFF}";
    let mut value = definition();
    value.tables[0].id = format!("id{valid}");
    value.tables[0].name = format!("table{valid}");
    value.tables[0].connection = format!("connection{valid}");
    let encoded = litchi_xlsx::write_data_model(&value).unwrap();
    let decoded = litchi_xlsx::parse_data_model(&encoded).unwrap();
    assert_eq!(decoded.tables[0], value.tables[0]);
}

#[test]
fn native_model_removal_retains_connections_and_restores_exact_source() {
    for override_model in [false, true] {
        let source = if override_model {
            native_source_with_xml_change("[Content_Types].xml", |xml| {
                xml.replace(
                    "<Default Extension=\"data\"",
                    "<Override PartName=\"/xl/model/item.data\"",
                )
            })
        } else {
            SOURCE.to_vec()
        };
        for owned in [false, true] {
            let raw = if owned {
                OpcPackage::from_vec(source.clone()).unwrap()
            } else {
                OpcPackage::from_bytes(&source).unwrap()
            };
            let mut package = Package::from_opc(raw.clone()).unwrap();
            let mut edit = package.edit_data_model().unwrap();
            assert!(edit.remove().unwrap());
            let commit = edit.commit().unwrap();
            assert!(commit.snapshot().model().is_none());
            let patch = commit.patch().clone();
            let mut output = Vec::new();
            package.write_to(&mut output).unwrap();
            let mut reopened = Package::from_bytes(output).unwrap().into_plain_opc();
            assert!(
                reopened
                    .get_part(&PackURI::new("/xl/model/item.data").unwrap())
                    .is_err()
            );
            assert!(
                !reopened
                    .get_part(&PackURI::new("/xl/workbook.xml").unwrap())
                    .unwrap()
                    .rels()
                    .iter()
                    .any(|rel| rel.r_id() == "rId6")
            );
            for part in raw
                .try_iter_parts()
                .map(|part| part.expect("part payload decodes"))
                .filter(|part| {
                    !matches!(
                        part.partname().as_str(),
                        "/xl/workbook.xml" | "/xl/model/item.data"
                    )
                })
            {
                assert_eq!(
                    reopened.get_part(part.partname()).unwrap().blob(),
                    part.blob()
                );
            }
            patch.inverse().apply(&mut reopened).unwrap();
            let restored = litchi_opc::PackageWriter::to_bytes(&reopened).unwrap();
            let archive = soapberry_zip::office::ArchiveReader::new(&restored).unwrap();
            for part in raw
                .try_iter_parts()
                .map(|part| part.expect("part payload decodes"))
            {
                assert_eq!(
                    archive.read(part.partname().membername()).unwrap(),
                    part.blob()
                );
            }
            let original = soapberry_zip::office::ArchiveReader::new(&source).unwrap();
            for member in ["xl/_rels/workbook.xml.rels", "[Content_Types].xml"] {
                assert_eq!(
                    archive.read(member).unwrap(),
                    original.read(member).unwrap(),
                    "owned={owned}, member={member}"
                );
            }
        }
    }
}

#[test]
fn native_model_removal_preserves_connections_processing_instructions_and_tokens() {
    let source = native_source_with_xml_change("xl/connections.xml", |xml| {
        xml.replace(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><?catalog keep-before?>",
        )
        .replace(
            "<connection id=\"3\"",
            "<?catalog keep-inside?>\n<connection id=\"3\"",
        )
        .replace(
            "</connection><connection id=\"4\"",
            "</connection><?catalog keep-between?>\n<connection id=\"4\"",
        )
        .replace(
            "id=\"\" model=\"1\"",
            "id=\"\" model=\"1\" autoDelete=\"true\"",
        )
    });
    let original = OpcPackage::from_vec(source.clone()).unwrap();
    let workbook = PackURI::new("/xl/workbook.xml").unwrap();
    let connections_name = PackURI::new("/xl/connections.xml").unwrap();
    let original_connections = String::from_utf8(
        original
            .get_part(&connections_name)
            .unwrap()
            .blob()
            .to_vec(),
    )
    .unwrap();
    let removed_start = original_connections.find("<connection id=\"4\"").unwrap();
    let removed_end = original_connections[removed_start..]
        .find("</connection>")
        .map(|offset| removed_start + offset + "</connection>".len())
        .unwrap();
    let expected_connections = format!(
        "{}{}",
        &original_connections[..removed_start],
        &original_connections[removed_end..]
    );
    let before_relationships = original.source_relationships(&workbook).unwrap();
    let expected_relationships = before_relationships
        .without_relationship("rId6", before_relationships.bytes().len())
        .unwrap();

    let mut package = Package::from_bytes(source.clone()).unwrap();
    let mut edit = package.edit_data_model().unwrap();
    assert!(edit.remove().unwrap());
    let patch = edit.commit().unwrap().patch().clone();

    let mut output = Vec::new();
    package.write_to(&mut output).unwrap();
    let mut reopened = Package::from_bytes(output).unwrap().into_plain_opc();
    let connections = reopened.get_part(&connections_name).unwrap().blob();
    assert_eq!(connections, expected_connections.as_bytes());
    assert!(
        connections
            .windows(b"<?catalog keep-before?>".len())
            .any(|window| window == b"<?catalog keep-before?>")
    );
    assert!(
        connections
            .windows(b"<?catalog keep-inside?>".len())
            .any(|window| window == b"<?catalog keep-inside?>")
    );
    assert!(
        connections
            .windows(b"<?catalog keep-between?>".len())
            .any(|window| window == b"<?catalog keep-between?>")
    );
    assert!(
        !connections
            .windows(b"<connection id=\"4\"".len())
            .any(|window| window == b"<connection id=\"4\"")
    );
    assert_eq!(
        reopened.source_relationships(&workbook).unwrap().bytes(),
        expected_relationships.bytes()
    );

    patch.inverse().apply(&mut reopened).unwrap();
    let restored = litchi_opc::PackageWriter::to_bytes(&reopened).unwrap();
    let actual = soapberry_zip::office::ArchiveReader::new(&restored).unwrap();
    let expected = soapberry_zip::office::ArchiveReader::new(&source).unwrap();
    for member in expected.file_names() {
        assert_eq!(
            actual.read(member).unwrap(),
            expected.read(member).unwrap(),
            "{member}"
        );
    }
}

fn native_source_with_xml_change(member: &str, change: impl FnOnce(String) -> String) -> Vec<u8> {
    source_with_xml_change(SOURCE, member, change)
}

fn source_with_xml_change(
    source: &[u8],
    member: &str,
    change: impl FnOnce(String) -> String,
) -> Vec<u8> {
    let archive = soapberry_zip::office::ArchiveReader::new(source).unwrap();
    let updated = change(String::from_utf8(archive.read(member).unwrap()).unwrap()).into_bytes();
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    for name in archive.file_names() {
        let bytes = if name == member {
            updated.clone()
        } else {
            archive.read(name).unwrap()
        };
        writer.write_deflated(name, &bytes).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

#[test]
fn native_removal_honors_auto_delete_and_addin_retention_with_exact_inverse() {
    for addin in [false, true] {
        let source = native_source_with_xml_change("xl/connections.xml", |xml| {
            xml.replace(
                "id=\"Tabelle1\"",
                &format!(
                    "id=\"Tabelle1\" autoDelete=\"1\" usedByAddin=\"{}\"",
                    u8::from(addin)
                ),
            )
            .replace(
                "id=\"\" model=\"1\"",
                "id=\"\" model=\"1\" autoDelete=\"true\"",
            )
        });
        let original = OpcPackage::from_vec(source.clone()).unwrap();
        let mut package = Package::from_bytes(source).unwrap();
        let mut edit = package.edit_data_model().unwrap();
        edit.remove().unwrap();
        let patch = edit.commit().unwrap().patch().clone();
        let mut output = Vec::new();
        package.write_to(&mut output).unwrap();
        let mut reopened = Package::from_bytes(output).unwrap().into_plain_opc();
        let connections = litchi_xlsx::connections::Connections::parse(
            reopened
                .get_part(&PackURI::new("/xl/connections.xml").unwrap())
                .unwrap()
                .blob(),
        )
        .unwrap();
        assert_eq!(
            connections
                .connections
                .iter()
                .map(|connection| connection.id)
                .collect::<Vec<_>>(),
            if addin { vec![1, 2, 3] } else { vec![1, 2] }
        );
        for part in original
            .try_iter_parts()
            .map(|part| part.expect("part payload decodes"))
            .filter(|part| {
                !matches!(
                    part.partname().as_str(),
                    "/xl/workbook.xml" | "/xl/model/item.data" | "/xl/connections.xml"
                )
            })
        {
            assert_eq!(
                reopened.get_part(part.partname()).unwrap().blob(),
                part.blob()
            );
        }
        patch.inverse().apply(&mut reopened).unwrap();
        for part in original
            .try_iter_parts()
            .map(|part| part.expect("part payload decodes"))
        {
            assert_eq!(
                reopened.get_part(part.partname()).unwrap().blob(),
                part.blob()
            );
        }
    }
}

#[test]
fn native_model_removal_refuses_query_pivot_and_cube_dependencies_atomically() {
    for consumer in 0..5 {
        let mut raw = OpcPackage::from_vec(SOURCE.to_vec()).unwrap();
        if consumer == 0 {
            let name = PackURI::new("/xl/queryTables/queryTable1.xml").unwrap();
            let source = String::from_utf8(raw.get_part(&name).unwrap().blob().to_vec()).unwrap();
            assert!(source.contains("connectionId=\"2\""));
            raw.get_part_mut(&name).unwrap().set_blob(
                source
                    .replace("connectionId=\"2\"", "connectionId=\"3\"")
                    .into_bytes(),
            );
        } else if consumer == 1 {
            raw.try_add_part(Box::new(litchi_opc::BlobPart::new(PackURI::new("/xl/pivotCache/pivotCacheDefinition1.xml").unwrap(), "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml".into(), br#"<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="external" connectionId="4"/></pivotCacheDefinition>"#.to_vec()))).unwrap();
        } else if consumer == 2 {
            let name = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
            raw.get_part_mut(&name).unwrap().set_blob(br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f>CUBEVALUE("ThisWorkbookDataModel","[Measures].[Count]")</f></c></row></sheetData></worksheet>"#.to_vec());
        } else if consumer == 3 {
            let name = PackURI::new("/xl/tables/table1.xml").unwrap();
            raw.get_part_mut(&name).unwrap().set_blob(br#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><tableColumns><tableColumn><calculatedColumnFormula>CUBEVALUE(A1)</calculatedColumnFormula></tableColumn></tableColumns></table>"#.to_vec());
        } else {
            let name = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
            raw.get_part_mut(&name).unwrap().set_blob(br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><extLst><xm:f>CUBEVALUE(A1)</xm:f></extLst></worksheet>"#.to_vec());
        }
        let expected = raw.clone();
        let mut package = Package::from_opc(raw).unwrap();
        let before = package.data_model().unwrap();
        let mut edit = package.edit_data_model().unwrap();
        edit.remove().unwrap();
        assert!(
            matches!(edit.commit(), Err(Error::Unsupported { .. })),
            "consumer {consumer}"
        );
        assert_eq!(package.data_model().unwrap().model(), before.model());
        let actual = package.into_plain_opc();
        assert_eq!(actual.part_count(), expected.part_count());
        for part in expected
            .try_iter_parts()
            .map(|part| part.expect("part payload decodes"))
        {
            let current = actual.get_part(part.partname()).unwrap();
            assert_eq!(current.blob(), part.blob());
            assert_eq!(current.rels().to_xml(), part.rels().to_xml());
        }
    }
}

#[test]
fn native_removal_patch_rechecks_changed_dependent_parts_before_publication() {
    let mut original = OpcPackage::from_vec(SOURCE.to_vec()).unwrap();
    let name = PackURI::new("/xl/pivotCache/pivotCacheDefinition1.xml").unwrap();
    original.try_add_part(Box::new(litchi_opc::BlobPart::new(name.clone(), "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml".into(), br#"<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="external" connectionId="1"/></pivotCacheDefinition>"#.to_vec()))).unwrap();
    let mut package = Package::from_opc(original.clone()).unwrap();
    let mut edit = package.edit_data_model().unwrap();
    edit.remove().unwrap();
    let patch = edit.commit().unwrap().patch().clone();
    let mut target = original;
    // Keep the part/type set unchanged so the source manifest guard passes;
    // the dependency check must notice the newly model-bound cache itself.
    target.get_part_mut(&name).unwrap().set_blob(br#"<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="external" connectionId="4"/></pivotCacheDefinition>"#.to_vec());
    let before = target.clone();
    assert!(matches!(
        patch.apply(&mut target),
        Err(Error::Unsupported { .. })
    ));
    assert_eq!(target.part_count(), before.part_count());
    for part in before
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
    {
        assert_eq!(
            target.get_part(part.partname()).unwrap().blob(),
            part.blob()
        );
    }
}

fn check_native_model_descriptor_round_trip(raw: OpcPackage) {
    use litchi_xlsx::package::xldm::{self, GeneratedNameKind, StorageProfile};
    let name = PackURI::new("/xl/model/item.data").unwrap();
    let source = raw.get_part(&name).unwrap();
    let storage = xldm::inspect(source.blob()).unwrap();
    assert_eq!(storage.profile(), StorageProfile::Tabular150);
    assert_eq!(storage.files.len(), 48);
    assert_eq!(storage.partition_marker.partition_count, 1);
    assert_eq!(storage.backup_log.file_groups.len(), 6);
    assert_eq!(storage.backup_log.backup_restore_sync_version, 1153);
    let files = storage
        .backup_log
        .file_groups
        .iter()
        .flat_map(|group| &group.files)
        .collect::<Vec<_>>();
    assert_eq!(files.len(), 46);
    assert_eq!(files[1].generated.kind, GeneratedNameKind::CryptographicKey);
    assert!(
        files
            .iter()
            .any(|file| file.generated.normalized_path.contains("TT_Prüfung"))
    );
    assert!(
        files
            .iter()
            .any(|file| file.generated.normalized_path.contains("Tage heute"))
    );
    assert_eq!(xldm::write(&storage).unwrap(), source.blob());
    // The borrowed inner codecs cannot decode native compressed members yet.
    // They must not silently classify the hashed storage names as an empty model.
    assert!(xldm::native::inspect(&storage, &Default::default()).is_err());
    assert!(xldm::generated::inspect_system_generated(&storage).is_err());
    assert!(xldm::metadata::inspect(&storage).is_err());
    let empty = xldm::metadata::MetadataModel {
        files: vec![],
        columns: vec![],
        relationships: vec![],
        hierarchies: vec![],
    };
    assert!(xldm::olap::inspect(&storage, &empty).is_err());
    let mut package = Package::from_opc(raw.clone()).unwrap();
    let before = package.data_model().unwrap();
    let model = before.model().unwrap();
    assert_eq!(model.payload().as_ptr(), source.blob().as_ptr());
    assert_eq!(
        model.definition().tables[0].connection,
        "LinkedTable_Tabelle1"
    );
    assert!(
        !package
            .edit_data_model()
            .unwrap()
            .commit()
            .unwrap()
            .changed()
    );
    let version = model.definition().min_version_load + 1;
    let mut transaction = package.edit_data_model().unwrap();
    assert!(
        transaction
            .edit_definition(|definition| {
                definition.tables[0].name.push_str("changed");
                Ok(())
            })
            .is_err()
    );
    assert!(!transaction.is_changed());
    assert!(
        transaction
            .edit_definition(|definition| {
                definition.min_version_load = version;
                Ok(())
            })
            .unwrap()
    );
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(
        commit.snapshot().model().unwrap().payload().as_ptr(),
        model.payload().as_ptr()
    );
    let patch = commit.patch().clone();
    let mut output = Vec::new();
    package.write_to(&mut output).unwrap();
    let relationship_xml = soapberry_zip::office::ArchiveReader::new(SOURCE)
        .unwrap()
        .read("xl/_rels/workbook.xml.rels")
        .unwrap();
    assert_eq!(
        soapberry_zip::office::ArchiveReader::new(&output)
            .unwrap()
            .read("xl/_rels/workbook.xml.rels")
            .unwrap(),
        relationship_xml
    );
    let reopened = Package::from_bytes(output).unwrap();
    assert_eq!(
        reopened
            .data_model()
            .unwrap()
            .model()
            .unwrap()
            .definition()
            .min_version_load,
        version
    );
    let mut after = reopened.into_plain_opc();
    let workbook_name = PackURI::new("/xl/workbook.xml").unwrap();
    assert_eq!(
        after.get_part(&workbook_name).unwrap().rels().to_xml(),
        raw.get_part(&workbook_name).unwrap().rels().to_xml()
    );
    for part in raw
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
    {
        if part.partname().as_str() != "/xl/workbook.xml" {
            assert_eq!(after.get_part(part.partname()).unwrap().blob(), part.blob());
        }
    }
    patch.inverse().apply(&mut after).unwrap();
    let restored = litchi_opc::PackageWriter::to_bytes(&after).unwrap();
    assert_eq!(
        soapberry_zip::office::ArchiveReader::new(&restored)
            .unwrap()
            .read("xl/_rels/workbook.xml.rels")
            .unwrap(),
        relationship_xml
    );
    for part in raw
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
    {
        assert_eq!(after.get_part(part.partname()).unwrap().blob(), part.blob());
    }
}

#[test]
fn native_workbook_relationship_compatibility_has_exact_owner_and_target() {
    const REL: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/powerPivotData";
    let source = OpcPackage::from_bytes(SOURCE).unwrap();
    let workbook = PackURI::new("/xl/workbook.xml").unwrap();
    let edge_id = source
        .get_part(&workbook)
        .unwrap()
        .rels()
        .iter()
        .find(|relationship| relationship.reltype() == REL)
        .unwrap()
        .r_id()
        .to_owned();
    for (reltype, target, external) in [
        (REL, "model/item.data", true),
        (REL, "model/item.data?query", false),
        (REL, "model/item.data#fragment", false),
        (REL, "tables/table1.xml", false),
        ("urn:unmodeled", "model/item.data", false),
    ] {
        let mut changed = source.clone();
        let rels = changed.get_part_mut(&workbook).unwrap().rels_mut();
        rels.remove(&edge_id);
        rels.add_relationship(reltype.into(), target.into(), edge_id.clone(), external);
        let mut package = Package::from_opc(changed).unwrap();
        assert!(package.data_model().is_err());
        assert!(package.edit_data_model().is_err());
    }
    let mut duplicate = source.clone();
    duplicate
        .get_part_mut(&workbook)
        .unwrap()
        .rels_mut()
        .add_relationship(
            REL.into(),
            "model/item.data".into(),
            "rIdDuplicate".into(),
            false,
        );
    assert!(Package::from_opc(duplicate).unwrap().data_model().is_err());
    let mut foreign = source.clone();
    foreign
        .get_part_mut(&workbook)
        .unwrap()
        .rels_mut()
        .remove(&edge_id);
    foreign.rels_mut().add_relationship(
        REL.into(),
        "xl/model/item.data".into(),
        "rIdForeign".into(),
        false,
    );
    assert!(Package::from_opc(foreign).unwrap().data_model().is_err());
}

fn repair_native_crc(bytes: &mut [u8], start: usize, end: usize) {
    let mut crc = u32::MAX;
    for byte in &bytes[start..end - 4] {
        crc ^= u32::from(*byte) << 24;
        for _ in 0..8 {
            crc = (crc << 1)
                ^ if crc & 0x8000_0000 != 0 {
                    0x04c1_1db7
                } else {
                    0
                };
        }
    }
    bytes[end - 4..end].copy_from_slice(&(!crc).to_le_bytes());
}

#[test]
fn native_profile_rejects_corrupt_boundaries_frames_and_log_bindings() {
    use litchi_xlsx::package::xldm;
    let raw = OpcPackage::from_bytes(SOURCE).unwrap();
    let bytes = raw
        .get_part(&PackURI::new("/xl/model/item.data").unwrap())
        .unwrap()
        .blob();
    for (start, end) in [(4096, 4584), (51593, 88227)] {
        let mut changed = bytes.to_vec();
        changed[start] = 0;
        repair_native_crc(&mut changed, start, end);
        assert!(xldm::inspect(&changed).is_err(), "missing owned marker BOM");
    }
    let mut changed = bytes.to_vec();
    replace_utf16(&mut changed[90112..116674], ">4584<", ">4585<");
    assert!(xldm::inspect(&changed).is_err(), "overlapping allocations");
    let mut changed = bytes.to_vec();
    changed[4584..4586].copy_from_slice(&3878u16.to_le_bytes());
    repair_native_crc(&mut changed, 4584, 5665);
    assert!(
        xldm::inspect(&changed).is_err(),
        "frame/log size disagreement"
    );
    for (before, after) in [
        ("<Size>3877</Size>", "<Size>3878</Size>"),
        (
            "<StoragePath>C7CCDAAB8C848F1B7A0",
            "<StoragePath>D7CCDAAB8C848F1B7A0",
        ),
        ("<Class>100002</Class>", "<Class>100003</Class>"),
        (
            "BCC901AC15414BF88640.0.db.xml",
            "../901AC15414BF88640.0.db.xml",
        ),
    ] {
        let mut changed = bytes.to_vec();
        replace_utf16(&mut changed[51593..88223], before, after);
        repair_native_crc(&mut changed, 51593, 88227);
        assert!(
            xldm::inspect(&changed).is_err(),
            "invalid native log: {after}"
        );
    }
    let mut changed = bytes.to_vec();
    replace_utf16(&mut changed[..4096], ">150<", ">140<");
    assert!(
        xldm::inspect(&changed).is_err(),
        "native layout cannot enter canonical profile"
    );
}

#[test]
fn native_removal_refuses_named_extension_and_attribute_formula_consumers() {
    let cases = [
        (
            "pivotCache/pivotCacheDefinition9.xml",
            "pivotCacheDefinition",
            r#"<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><cacheSource type="external"><extLst><ext uri="{F057638F-6D5F-4E77-A914-E7F072B9BCA8}"><x14:sourceConnection name="ThisWorkbookDataModel"/></ext></extLst></cacheSource></pivotCacheDefinition>"#,
        ),
        (
            "queryTables/queryTable9.xml",
            "queryTable",
            r#"<queryTable xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" connectionId="1"><extLst><ext uri="{883FBD77-0823-4A55-B5E3-86C4891E6966}"><x15:queryTable sourceDataName="ThisWorkbookDataModel"/></ext></extLst></queryTable>"#,
        ),
        (
            "pivotTables/pivotTable9.xml",
            "pivotTable",
            r#"<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><extLst><ext uri="{E67621CE-5B39-4880-91FE-76760E9C1902}"><x15:pivotTableUISettings sourceDataName="ThisWorkbookDataModel"/></ext></extLst></pivotTableDefinition>"#,
        ),
        (
            "worksheets/sheet9.xml",
            "worksheet",
            r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><conditionalFormatting><cfRule><colorScale><cfvo type="formula" val="C&#85;BEVALUE(A1)"/></colorScale></cfRule></conditionalFormatting></worksheet>"#,
        ),
        (
            "pivotCache/pivotCacheDefinition9.xml",
            "pivotCacheDefinition",
            r#"<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="worksheet"/><cacheFields><cacheField formula="CUBEVALUE(A1)"/></cacheFields></pivotCacheDefinition>"#,
        ),
        (
            "pivotCache/pivotCacheDefinition9.xml",
            "pivotCacheDefinition",
            r#"<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="worksheet"/><calculatedItems><calculatedItem formula="CUBEMEMBER(A1)"/></calculatedItems></pivotCacheDefinition>"#,
        ),
        (
            "revisions/revisionLog9.xml",
            "revisionLog",
            r#"<revisions xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><rdn><oldFormula>CUBEVALUE(A1)</oldFormula></rdn></revisions>"#,
        ),
    ];
    for (name, kind, xml) in cases {
        let mut raw = OpcPackage::from_vec(SOURCE.to_vec()).unwrap();
        raw.try_add_part(Box::new(litchi_opc::BlobPart::new(
            PackURI::new(format!("/xl/{name}")).unwrap(),
            format!("application/vnd.openxmlformats-officedocument.spreadsheetml.{kind}+xml"),
            xml.as_bytes().to_vec(),
        )))
        .unwrap();
        let before = raw.clone();
        let mut package = Package::from_opc(raw).unwrap();
        let mut edit = package.edit_data_model().unwrap();
        edit.remove().unwrap();
        assert!(
            matches!(edit.commit(), Err(Error::Unsupported { .. })),
            "{name}"
        );
        let after = package.into_plain_opc();
        assert_eq!(after.part_count(), before.part_count());
        for part in before
            .try_iter_parts()
            .map(|part| part.expect("part payload decodes"))
        {
            assert_eq!(after.get_part(part.partname()).unwrap().blob(), part.blob());
        }
    }
}

fn all_auto_delete_connections_source() -> Vec<u8> {
    let source = native_source_with_xml_change("xl/connections.xml", |mut xml| {
        for id in [1, 2] {
            let start = xml.find(&format!("<connection id=\"{id}\"")).unwrap();
            let end = start + xml[start..].find("</connection>").unwrap() + "</connection>".len();
            xml.replace_range(start..end, "");
        }
        xml.replace("id=\"Tabelle1\"", "id=\"Tabelle1\" autoDelete=\"true\"")
            .replace(
                "id=\"\" model=\"1\"",
                "id=\"\" model=\"1\" autoDelete=\"true\"",
            )
    });
    // This derived source retains the ordinary query outputs as static tables,
    // removing their query owners and declarations along with connections 1/2.
    let archive = soapberry_zip::office::ArchiveReader::new(&source).unwrap();
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    for name in archive.file_names() {
        if name.starts_with("xl/queryTables/")
            || matches!(
                name,
                "xl/tables/_rels/table2.xml.rels" | "xl/tables/_rels/table3.xml.rels"
            )
        {
            continue;
        }
        let bytes = archive.read(name).unwrap();
        let bytes = if matches!(name, "xl/tables/table2.xml" | "xl/tables/table3.xml") {
            String::from_utf8(bytes)
                .unwrap()
                .replace(" tableType=\"queryTable\"", "")
                .into_bytes()
        } else if name == "[Content_Types].xml" {
            let mut xml = String::from_utf8(bytes).unwrap();
            for id in [1, 2] {
                let start = xml
                    .find(&format!(
                        "<Override PartName=\"/xl/queryTables/queryTable{id}.xml\""
                    ))
                    .unwrap();
                let end = start + xml[start..].find("/>").unwrap() + 2;
                xml.replace_range(start..end, "");
            }
            xml.into_bytes()
        } else {
            bytes
        };
        writer.write_deflated(name, &bytes).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

#[test]
fn native_removal_deletes_final_connections_owner_and_restores_exact_inverse() {
    let source = all_auto_delete_connections_source();
    let name = PackURI::new("/xl/connections.xml").unwrap();
    for borrowed in [false, true] {
        let original = if borrowed {
            OpcPackage::from_bytes(&source).unwrap()
        } else {
            OpcPackage::from_vec(source.clone()).unwrap()
        };
        let mut package = Package::from_opc(original.clone()).unwrap();
        let mut edit = package.edit_data_model().unwrap();
        edit.remove().unwrap();
        let patch = edit.commit().unwrap().patch().clone();
        let mut output = Vec::new();
        package.write_to(&mut output).unwrap();
        let mut reopened = OpcPackage::from_vec(output).unwrap();
        assert!(reopened.get_part(&name).is_err());
        let workbook = PackURI::new("/xl/workbook.xml").unwrap();
        assert!(
            !reopened
                .get_part(&workbook)
                .unwrap()
                .rels()
                .iter()
                .any(|rel| rel.reltype().ends_with("/connections"))
        );
        let mut forward = original.clone();
        patch.apply(&mut forward).unwrap();
        assert!(forward.get_part(&name).is_err());
        for part in original
            .try_iter_parts()
            .map(|part| part.expect("part payload decodes"))
            .filter(|part| {
                !matches!(
                    part.partname().as_str(),
                    "/xl/connections.xml" | "/xl/model/item.data" | "/xl/workbook.xml"
                )
            })
        {
            assert_eq!(
                reopened.get_part(part.partname()).unwrap().blob(),
                part.blob()
            );
        }
        patch.inverse().apply(&mut reopened).unwrap();
        let restored = litchi_opc::PackageWriter::to_bytes(&reopened).unwrap();
        let actual = soapberry_zip::office::ArchiveReader::new(&restored).unwrap();
        let expected = soapberry_zip::office::ArchiveReader::new(&source).unwrap();
        for member in expected.file_names() {
            assert_eq!(
                actual.read(member).unwrap(),
                expected.read(member).unwrap(),
                "{member}"
            );
        }
    }
}

#[test]
fn final_connections_owner_removal_refuses_unowned_content_and_edges_atomically() {
    let source = all_auto_delete_connections_source();
    for case in 0..7 {
        let bytes = if case == 0 {
            source_with_xml_change(&source, "xl/connections.xml", |xml| {
                xml.replace(
                    "</connections>",
                    "<v:metadata xmlns:v=\"urn:vendor\"/></connections>",
                )
            })
        } else if case == 3 {
            source_with_xml_change(&source, "xl/connections.xml", |xml| {
                xml.replace(
                    "<connections ",
                    "<connections xmlns:v=\"urn:vendor\" v:catalog=\"retain\" ",
                )
            })
        } else if case == 4 {
            source_with_xml_change(&source, "xl/connections.xml", |xml| {
                xml.replace("<connections ", "<connections catalog=\"retain\" ")
            })
        } else {
            source.clone()
        };
        let mut raw = OpcPackage::from_vec(bytes).unwrap();
        let name = PackURI::new("/xl/connections.xml").unwrap();
        if case == 1 || case == 5 {
            raw.rels_mut().add_relationship(
                "urn:vendor".into(),
                if case == 5 {
                    "xl/CONNECTIONS.XML"
                } else {
                    "xl/connections.xml"
                }
                .into(),
                "rIdForeign".into(),
                false,
            );
        } else if case == 2 {
            raw.get_part_mut(&name)
                .unwrap()
                .rels_mut()
                .add_relationship(
                    "urn:vendor".into(),
                    "https://example.invalid/inert".into(),
                    "rIdVendor".into(),
                    true,
                );
        }
        if case == 6 {
            raw.get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
                .unwrap()
                .rels_mut()
                .add_relationship(
                    "urn:vendor".into(),
                    "../CONNECTIONS.XML".into(),
                    "rIdForeign".into(),
                    false,
                );
        }
        let before = raw.clone();
        let mut package = Package::from_opc(raw).unwrap();
        if case == 2 {
            assert!(package.edit_data_model().is_err());
        } else {
            let mut edit = package.edit_data_model().unwrap();
            edit.remove().unwrap();
            assert!(
                matches!(edit.commit(), Err(Error::Unsupported { .. })),
                "{case}"
            );
        }
        let after = package.into_plain_opc();
        assert_eq!(after.part_count(), before.part_count());
        for part in before
            .try_iter_parts()
            .map(|part| part.expect("part payload decodes"))
        {
            assert_eq!(after.get_part(part.partname()).unwrap().blob(), part.blob());
            assert_eq!(
                after.get_part(part.partname()).unwrap().rels().to_xml(),
                part.rels().to_xml()
            );
        }
    }
}

#[test]
fn connections_owner_removal_patch_rechecks_new_foreign_references() {
    let source = all_auto_delete_connections_source();
    let original = OpcPackage::from_vec(source).unwrap();
    let mut package = Package::from_opc(original.clone()).unwrap();
    let mut edit = package.edit_data_model().unwrap();
    edit.remove().unwrap();
    let patch = edit.commit().unwrap().patch().clone();
    let mut target = original;
    target.rels_mut().add_relationship(
        "urn:vendor".into(),
        "xl/connections.xml".into(),
        "rIdForeign".into(),
        false,
    );
    let before = target.clone();
    assert!(matches!(
        patch.apply(&mut target),
        Err(Error::Unsupported {
            feature: "Data Model removal with foreign Connections references"
        })
    ));
    assert_eq!(target.rels().to_xml(), before.rels().to_xml());
    for part in before
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
    {
        assert_eq!(
            target.get_part(part.partname()).unwrap().blob(),
            part.blob()
        );
    }
}

#[test]
fn connections_owner_inverse_rejects_uri_reuse_even_with_unchanged_manifest() {
    let source = all_auto_delete_connections_source();
    // An unused Default can remain after deletion and cover a later orphan at
    // the same URI, so manifest equality alone cannot authorize restoration.
    let source = source_with_xml_change(&source, "[Content_Types].xml", |xml| {
        xml.replace(
            "<Override PartName=\"/xl/connections.xml\"",
            "<Default Extension=\"cnx\"",
        )
    });
    let source = source_with_xml_change(&source, "xl/_rels/workbook.xml.rels", |xml| {
        xml.replace("Target=\"connections.xml\"", "Target=\"connections.cnx\"")
    });
    let archive = soapberry_zip::office::ArchiveReader::new(&source).unwrap();
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    for member in archive.file_names() {
        writer
            .write_deflated(
                if member == "xl/connections.xml" {
                    "xl/connections.cnx"
                } else {
                    member
                },
                &archive.read(member).unwrap(),
            )
            .unwrap();
    }
    let source = writer.finish_to_bytes().unwrap();
    let mut package = Package::from_bytes(source).unwrap();
    let mut edit = package.edit_data_model().unwrap();
    edit.remove().unwrap();
    let inverse = edit.commit().unwrap().patch().inverse();
    let mut target = package.into_plain_opc();
    let manifest = target.source_content_types().unwrap();
    let name = PackURI::new("/xl/connections.cnx").unwrap();
    target
        .try_add_part(Box::new(litchi_opc::BlobPart::new(
            name.clone(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml".into(),
            b"new unrelated owner".to_vec(),
        )))
        .unwrap();
    assert_eq!(target.source_content_types().unwrap(), manifest);
    let before = target.clone();
    assert!(
        matches!(inverse.apply(&mut target), Err(Error::PatchConflict { part }) if part == "Data Model restored Connections owner")
    );
    assert_eq!(target.part_count(), before.part_count());
    for part in before
        .try_iter_parts()
        .map(|part| part.expect("part payload decodes"))
    {
        assert_eq!(
            target.get_part(part.partname()).unwrap().blob(),
            part.blob()
        );
    }
}

#[test]
fn connections_owner_inverse_rejects_new_incoming_edges_to_absent_uri() {
    let source = all_auto_delete_connections_source();
    let mut package = Package::from_bytes(source).unwrap();
    let mut edit = package.edit_data_model().unwrap();
    edit.remove().unwrap();
    let inverse = edit.commit().unwrap().patch().inverse();
    let removed = package.into_plain_opc();
    let name = PackURI::new("/xl/connections.xml").unwrap();
    for (from_root, target_name) in [
        (false, "connections.xml"),
        (true, "connections.xml"),
        (false, "CONNECTIONS.XML"),
        (true, "CONNECTIONS.XML"),
    ] {
        let mut target = removed.clone();
        if from_root {
            target.rels_mut().add_relationship(
                "urn:vendor".into(),
                format!("xl/{target_name}"),
                "rIdForeign".into(),
                false,
            );
        } else {
            target
                .get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
                .unwrap()
                .rels_mut()
                .add_relationship(
                    "urn:vendor".into(),
                    format!("../{target_name}"),
                    "rIdForeign".into(),
                    false,
                );
        }
        assert!(target.get_part(&name).is_err());
        assert_eq!(
            target.source_content_types().unwrap(),
            removed.source_content_types().unwrap()
        );
        let before = target.clone();
        assert!(
            matches!(inverse.apply(&mut target), Err(Error::PatchConflict { part }) if part == "Data Model restored Connections references")
        );
        assert_eq!(target.rels().to_xml(), before.rels().to_xml());
        assert_eq!(target.part_count(), before.part_count());
        for part in before
            .try_iter_parts()
            .map(|part| part.expect("part payload decodes"))
        {
            let after = target.get_part(part.partname()).unwrap();
            assert_eq!(after.blob(), part.blob());
            assert_eq!(after.rels().to_xml(), part.rels().to_xml());
        }
    }
}

#[test]
fn final_connections_owner_resolves_case_variant_owning_target() {
    let source = all_auto_delete_connections_source();
    let source = source_with_xml_change(&source, "xl/_rels/workbook.xml.rels", |xml| {
        xml.replace("Target=\"connections.xml\"", "Target=\"CONNECTIONS.XML\"")
    });
    let mut package = Package::from_bytes(source.clone()).unwrap();
    let mut edit = package.edit_data_model().unwrap();
    edit.remove().unwrap();
    let patch = edit.commit().unwrap().patch().clone();
    let mut output = Vec::new();
    package.write_to(&mut output).unwrap();
    let mut reopened = OpcPackage::from_vec(output).unwrap();
    assert!(
        reopened
            .get_part(&PackURI::new("/xl/connections.xml").unwrap())
            .is_err()
    );
    patch.inverse().apply(&mut reopened).unwrap();
    let restored = litchi_opc::PackageWriter::to_bytes(&reopened).unwrap();
    let actual = soapberry_zip::office::ArchiveReader::new(&restored).unwrap();
    let original = soapberry_zip::office::ArchiveReader::new(&source).unwrap();
    for member in original.file_names() {
        assert_eq!(
            actual.read(member).unwrap(),
            original.read(member).unwrap(),
            "{member}"
        );
    }
}

#[test]
fn native_model_rejects_case_equivalent_foreign_targets() {
    for from_root in [false, true] {
        let mut raw = OpcPackage::from_vec(SOURCE.to_vec()).unwrap();
        if from_root {
            raw.rels_mut().add_relationship(
                "urn:vendor".into(),
                "xl/model/ITEM.DATA".into(),
                "rIdForeign".into(),
                false,
            );
        } else {
            raw.get_part_mut(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
                .unwrap()
                .rels_mut()
                .add_relationship(
                    "urn:vendor".into(),
                    "../model/ITEM.DATA".into(),
                    "rIdForeign".into(),
                    false,
                );
        }
        assert!(Package::from_opc(raw).unwrap().data_model().is_err());
    }
}

#[test]
fn native_model_workbook_edge_resolves_case_equivalent_target() {
    let source = native_source_with_xml_change("xl/_rels/workbook.xml.rels", |xml| {
        xml.replace("Target=\"model/item.data\"", "Target=\"model/ITEM.DATA\"")
    });
    let mut package = Package::from_bytes(source.clone()).unwrap();
    assert!(package.data_model().unwrap().model().is_some());
    let mut edit = package.edit_data_model().unwrap();
    edit.remove().unwrap();
    let inverse = edit.commit().unwrap().patch().inverse();
    let mut output = Vec::new();
    package.write_to(&mut output).unwrap();
    let mut reopened = OpcPackage::from_vec(output).unwrap();
    inverse.apply(&mut reopened).unwrap();
    let restored = litchi_opc::PackageWriter::to_bytes(&reopened).unwrap();
    let actual = soapberry_zip::office::ArchiveReader::new(&restored).unwrap();
    let original = soapberry_zip::office::ArchiveReader::new(&source).unwrap();
    for member in original.file_names() {
        assert_eq!(
            actual.read(member).unwrap(),
            original.read(member).unwrap(),
            "{member}"
        );
    }
}
