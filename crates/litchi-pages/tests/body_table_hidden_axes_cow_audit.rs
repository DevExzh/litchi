//! Identity and copy-on-write audit for the Pages hidden-state graph.
//!
//! The main hidden-axis fixture covers semantic set/clear behavior.  This
//! companion binary deliberately checks the native objects that are private
//! to the package owner: existing-helper identity, metadata/member locality,
//! copy-on-write behavior, and refusal of an inbound/shared formula-owner
//! edge.  Pages currently edits existing hidden-state owners only; an absent
//! owner is read as empty and cannot be synthesized by this API.

mod fixture {
    include!("body_table_hidden_axes.rs");

    use std::collections::{HashMap, HashSet};

    const AUDIT_METADATA_MEMBER: &str = "Index/Metadata.iwa";
    const AUDIT_METADATA_OBJECT_ID: u64 = 50_000;
    const AUDIT_METADATA_ORPHAN_ID: u64 = 90_000;
    const AUDIT_EXTRA_MEMBER: &str = "Index/CalculationEngine.iwa";
    const AUDIT_EXTRA_OBJECT_ID: u64 = 80_000;
    const AUDIT_EXTRA_MESSAGE_TYPE: u32 = 9_999;

    #[derive(Debug, Clone, PartialEq, Eq)]
    struct EntryPhysicalSnapshot {
        name: String,
        raw_name: Vec<u8>,
        data: Vec<u8>,
        metadata: litchi_iwa_archive::package::EntryMetadata,
        local_record: Vec<u8>,
        central_record: Vec<u8>,
    }

    fn unchanged_entry_snapshots(package: &[u8]) -> TestResult<Vec<EntryPhysicalSnapshot>> {
        Ok(Catalog::from_bytes(package)?
            .iter()
            .filter(|entry| entry.name() != DOCUMENT_MEMBER && !PREVIEWS.contains(&entry.name()))
            .map(|entry| EntryPhysicalSnapshot {
                name: entry.name().to_owned(),
                raw_name: entry.raw_name().to_vec(),
                data: entry.data().to_vec(),
                metadata: entry.metadata().clone(),
                local_record: entry.raw_record().local_record().to_vec(),
                central_record: {
                    let mut record = entry.raw_record().central_directory_record().to_vec();
                    // A resized earlier member or removed preview relocates
                    // this local record. Only its required central-directory
                    // offset fixup is outside the byte-preservation contract.
                    assert!(record.len() >= 46);
                    assert_ne!(
                        &record[42..46],
                        &[0xff; 4],
                        "fixture must not use ZIP64 offsets"
                    );
                    record[42..46].fill(0);
                    record
                },
            })
            .collect())
    }

    fn archive_object_ids(archive: &Archive) -> HashSet<u64> {
        archive
            .objects
            .iter()
            .filter_map(|object| object.archive_info.identifier)
            .collect()
    }

    fn metadata_payload(package: &[u8]) -> TestResult<Vec<u8>> {
        let catalog = Catalog::from_bytes(package)?;
        let entry = catalog
            .iter()
            .find(|entry| entry.name() == AUDIT_METADATA_MEMBER)
            .ok_or("missing metadata member")?;
        let archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
        Ok(archive
            .object(AUDIT_METADATA_OBJECT_ID)
            .ok_or("missing metadata object")?
            .messages
            .iter()
            .find(|message| message.type_ == 11_006)
            .ok_or("missing metadata message")?
            .data
            .clone())
    }

    fn metadata_package(source: &[u8]) -> TestResult<Vec<u8>> {
        let catalog = Catalog::from_bytes(source)?;
        let document = catalog
            .iter()
            .find(|entry| entry.name() == DOCUMENT_MEMBER)
            .ok_or("missing document member")?;
        let document_archive =
            Archive::parse(SnappyStream::decompress(document.data())?.as_bytes())?;
        let mut identifiers = document_archive
            .objects
            .iter()
            .map(|object| {
                object
                    .archive_info
                    .identifier
                    .ok_or_else(|| {
                        std::io::Error::new(
                            std::io::ErrorKind::InvalidData,
                            "missing document object identifier",
                        )
                    })
                    .map_err(Into::into)
            })
            .collect::<TestResult<Vec<_>>>()?;
        identifiers.push(AUDIT_METADATA_OBJECT_ID);
        // Metadata UUID maps can retain an object identity after its native
        // payload is gone.  The allocator still has to treat that identity
        // as occupied, even when PackageMetadata.last_object_identifier is
        // stale relative to the map.
        identifiers.push(AUDIT_METADATA_ORPHAN_ID);
        identifiers.sort_unstable();
        let metadata = tsp::PackageMetadata {
            last_object_identifier: AUDIT_METADATA_OBJECT_ID,
            save_token: Some(1),
            components: vec![tsp::ComponentInfo {
                identifier: 1,
                preferred_locator: "Document".to_owned(),
                locator: Some("Document".to_owned()),
                save_token: Some(1),
                object_uuid_map_entries: identifiers
                    .iter()
                    .copied()
                    .map(|identifier| tsp::ObjectUuidMapEntry {
                        identifier,
                        uuid: tsp::Uuid {
                            lower: identifier.saturating_add(10_000),
                            upper: identifier.saturating_add(20_000),
                        },
                    })
                    .collect(),
                ..tsp::ComponentInfo::default()
            }],
            ..tsp::PackageMetadata::default()
        }
        .encode_to_vec();
        let metadata_archive = SnappyStream::compress(
            &Archive {
                objects: vec![object(AUDIT_METADATA_OBJECT_ID, 11_006, metadata, &[])?],
            }
            .to_bytes()?,
        )?;
        let mut members = catalog
            .iter()
            .filter(|entry| entry.name() != AUDIT_METADATA_MEMBER)
            .map(|entry| (entry.name().to_owned(), entry.data().to_vec()))
            .collect::<Vec<_>>();
        members.push((AUDIT_METADATA_MEMBER.to_owned(), metadata_archive));
        let references = members
            .iter()
            .map(|(name, data)| (name.as_str(), data.as_slice()))
            .collect::<Vec<_>>();
        Ok(litchi_iwa_archive::package::to_bytes(
            references,
            Limits::default(),
        )?)
    }

    fn add_extra_component(source: &[u8]) -> TestResult<Vec<u8>> {
        let extra = Archive {
            objects: vec![ArchiveObject::new(
                AUDIT_EXTRA_OBJECT_ID,
                vec![RawMessage {
                    type_: AUDIT_EXTRA_MESSAGE_TYPE,
                    data: vec![0xde, 0xad, 0xbe, 0xef],
                }],
            )?],
        };
        let compressed = SnappyStream::compress(&extra.to_bytes()?)?;
        Ok(
            Catalog::from_bytes(source)?.reassemble_with_insertions_to_bytes(
                &[litchi_iwa_archive::package::EntryInsertion::new(
                    AUDIT_EXTRA_MEMBER,
                    compressed.as_slice(),
                )],
                Limits::default(),
            )?,
        )
    }

    fn add_identifier(source: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
        rewrite_document_archive(source, |archive| {
            archive.insert_object(ArchiveObject::new(
                identifier,
                vec![RawMessage {
                    type_: AUDIT_EXTRA_MESSAGE_TYPE,
                    data: Vec::new(),
                }],
            )?)?;
            Ok(())
        })
    }

    fn alias_second_table_formula_owner(source: &[u8]) -> TestResult<Vec<u8>> {
        rewrite_document_archive(source, |archive| {
            let model = archive
                .object_mut(table_model(1))
                .ok_or("missing second model")?;
            let message_index = model
                .messages
                .iter()
                .position(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
                .ok_or("missing second model message")?;
            let mut decoded =
                tst::TableModelArchive::decode(model.messages[message_index].data.as_slice())?;
            decoded.hidden_state_formula_owner_for_columns =
                Some(reference(table_formula_object(0, true)));
            let mut payload = decoded.encode_to_vec();
            litchi_iwa_common::wire::append_varint_field(
                &mut payload,
                UNKNOWN_MODEL_FIELD,
                UNKNOWN_MODEL_VALUE,
            )?;
            model.replace_message(
                message_index,
                RawMessage {
                    type_: TABLE_MODEL_MESSAGE_TYPE,
                    data: payload,
                },
            )?;
            let info = &mut model.archive_info.message_infos[message_index];
            for identifier in &mut info.object_references {
                if *identifier == table_formula_object(1, true) {
                    *identifier = table_formula_object(0, true);
                }
            }
            for field in &mut info.field_infos {
                if field.path.as_slice() == [34]
                    && field.object_references.as_slice() == [table_formula_object(1, true)]
                {
                    field.object_references[0] = table_formula_object(0, true);
                }
            }
            Ok(())
        })
    }

    #[test]
    fn absent_owner_refuses_creation_and_preserves_cow() -> TestResult {
        let source = normal_package()?;
        let package = Package::from_bytes(&source)?;
        let source_ids = archive_object_ids(&document_archive(&source)?);
        let before = package.exact_bytes();
        let requested = HiddenAxes::new([AxisIndex::row(0), AxisIndex::column(3)])?;
        let result = package
            .edit_body_table_hidden_axes(1usize)?
            .set(requested)
            .commit();
        assert!(matches!(
            result,
            Err(Error::UnsupportedDependency | Error::UnsupportedSource)
        ));
        assert_eq!(package.exact_bytes(), before);
        assert_eq!(
            archive_object_ids(&document_archive(&package.exact_bytes())?),
            source_ids,
            "refusing an absent owner must not allocate helper objects"
        );

        let noop = package
            .edit_body_table_hidden_axes(1usize)?
            .clear()
            .commit()?;
        assert!(noop.patch().is_noop());
        assert_eq!(noop.package().exact_bytes(), source);
        Ok(())
    }

    #[test]
    fn clearing_user_hidden_axes_keeps_helper_identity_and_exact_inverse() -> TestResult {
        let source = normal_package()?;
        let package = Package::from_bytes(&source)?;
        let source_archive = document_archive(&source)?;
        let helper_ids = [
            table_formula_owner(0),
            table_formula_object(0, true),
            table_formula_object(0, false),
            table_filter_set(0, true),
            table_filter_set(0, false),
        ];
        let helper_objects = helper_ids
            .into_iter()
            .map(|identifier| {
                Ok::<_, Box<dyn StdError>>((
                    identifier,
                    source_archive
                        .object(identifier)
                        .ok_or_else(|| {
                            std::io::Error::new(std::io::ErrorKind::InvalidData, "missing helper")
                        })?
                        .clone(),
                ))
            })
            .collect::<TestResult<HashMap<_, _>>>()?;
        let commit = package
            .edit_body_table_hidden_axes(0usize)?
            .clear()
            .commit()?;
        let target = commit.package().exact_bytes();
        let target_archive = document_archive(&target)?;
        for (identifier, object) in helper_objects {
            let after = target_archive
                .object(identifier)
                .ok_or("missing preserved helper")?;
            assert!(object.same_content_ignoring_offsets(after));
            assert_eq!(
                object_header_bytes(&source, identifier)?,
                object_header_bytes(&target, identifier)?
            );
        }
        assert_eq!(
            archive_object_ids(&target_archive),
            archive_object_ids(&source_archive),
            "clearing axes must not cull native helper objects"
        );
        let restored = commit
            .package()
            .apply_body_table_hidden_axes(&commit.patch().inverse())?;
        assert_eq!(restored.package().exact_bytes(), source);
        Ok(())
    }

    #[test]
    fn shared_formula_owner_inbound_edge_is_rejected_without_mutation() -> TestResult {
        let source = synthetic_package(
            [
                TableOptions {
                    user_hidden: true,
                    ..TableOptions::default()
                },
                TableOptions {
                    user_hidden: true,
                    ..TableOptions::default()
                },
            ],
            ["Revenue", "Costs"],
        )?;
        let aliased = alias_second_table_formula_owner(&source)?;
        let package = Package::from_bytes(&aliased)?;
        let before = package.exact_bytes();
        let result = package
            .edit_body_table_hidden_axes(1usize)
            .and_then(|edit| edit.clear().commit());
        assert!(result.is_err(), "shared formula owner was mutated");
        assert_eq!(package.exact_bytes(), before);
        Ok(())
    }

    #[test]
    fn existing_owner_rewrite_stays_local_with_an_unrelated_component() -> TestResult {
        let source = add_extra_component(&normal_package()?)?;
        let extra_before = member_bytes(&source, AUDIT_EXTRA_MEMBER)?;
        let entry_snapshots_before = unchanged_entry_snapshots(&source)?;
        let package = Package::from_bytes(&source)?;
        let source_document_ids = archive_object_ids(&document_archive(&source)?);
        let commit = package
            .edit_body_table_hidden_axes(0usize)?
            .set(HiddenAxes::new([AxisIndex::row(0)])?)
            .commit()?;
        let target = commit.package().exact_bytes();
        assert_eq!(member_bytes(&target, AUDIT_EXTRA_MEMBER)?, extra_before);
        assert_eq!(
            unchanged_entry_snapshots(&target)?,
            entry_snapshots_before,
            "unselected ZIP names, metadata, records, and payloads changed"
        );
        let target_catalog = Catalog::from_bytes(&target)?;
        let extra = target_catalog
            .iter()
            .find(|entry| entry.name() == AUDIT_EXTRA_MEMBER)
            .ok_or("missing extra component")?;
        let decompressed = SnappyStream::decompress(extra.data())?;
        let extra_archive = Archive::parse(decompressed.as_bytes())?;
        assert!(extra_archive.object(AUDIT_EXTRA_OBJECT_ID).is_some());
        let target_ids = archive_object_ids(&document_archive(&target)?);
        assert_eq!(target_ids, source_document_ids);
        Ok(())
    }

    #[test]
    fn existing_owner_rewrite_preserves_unrelated_metadata_exactly() -> TestResult {
        let source = metadata_package(&normal_package()?)?;
        let entry_snapshots_before = unchanged_entry_snapshots(&source)?;
        let package = Package::from_bytes(&source)?;
        let metadata_before = member_bytes(&source, AUDIT_METADATA_MEMBER)?;
        let metadata_payload_before = metadata_payload(&source)?;
        let commit = package
            .edit_body_table_hidden_axes(0usize)?
            .set(HiddenAxes::new([AxisIndex::row(0), AxisIndex::column(3)])?)
            .commit()?;
        let target = commit.package().exact_bytes();
        assert_eq!(
            member_bytes(&target, AUDIT_METADATA_MEMBER)?,
            metadata_before
        );
        assert_eq!(metadata_payload(&target)?, metadata_payload_before);
        assert_eq!(
            unchanged_entry_snapshots(&target)?,
            entry_snapshots_before,
            "metadata or another unselected ZIP record changed"
        );
        let restored = commit
            .package()
            .apply_body_table_hidden_axes(&commit.patch().inverse())?;
        assert_eq!(restored.package().exact_bytes(), source);
        Ok(())
    }

    #[test]
    fn absent_owner_refusal_does_not_scan_or_allocate_near_max_ids() -> TestResult {
        let source = add_identifier(&normal_package()?, u64::MAX)?;
        let package = Package::from_bytes(&source)?;
        let before = package.exact_bytes();
        let result = package
            .edit_body_table_hidden_axes(1usize)?
            .set(HiddenAxes::new([AxisIndex::row(0)])?)
            .commit();
        assert!(matches!(
            result,
            Err(Error::UnsupportedDependency | Error::UnsupportedSource)
        ));
        assert_eq!(package.exact_bytes(), before);
        Ok(())
    }
}
