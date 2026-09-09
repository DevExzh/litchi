//! Exact-source integration coverage for direct Numbers cell-comment replies.
//!
//! The fixture is intentionally small, but it keeps the native ownership
//! graph intact: a tile cell points at a comment-list key, the list entry
//! points at a `TSD.CommentStorageArchive`, and that root archive points at
//! source-ordered reply archives.  The public tests only use selectors,
//! checked cell positions, and reply ordinals; native identifiers are used
//! only by the test oracle when checking COW and culling.

#![allow(deprecated)]

use std::io;

#[path = "support/numbers_comment_fixture.rs"]
mod fixture;
use fixture::*;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::{tsp, tst};
use litchi_numbers::cell::comment::CommentReplyIndex;
use litchi_numbers::{
    CellPosition, Package, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector,
    TableSelector,
};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn load_package(source: &[u8]) -> TestResult<Package> {
    Ok(Package::from_bytes(source)?)
}

fn read_all(package: &Package, row: usize) -> TestResult<Vec<String>> {
    let replies = package.table_cell_comment_replies(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(u32::try_from(row)?, 0),
    )?;
    Ok(replies
        .iter()
        .map(|reply| reply.text().to_owned())
        .collect())
}

fn read_one(package: &Package, row: usize, index: u32) -> TestResult<String> {
    Ok(package
        .table_cell_comment_reply(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(u32::try_from(row)?, 0),
            CommentReplyIndex::new(index),
        )?
        .text()
        .to_owned())
}

#[test]
fn root_comment_creation_reuses_strict_graph_and_is_exactly_reversible() -> TestResult {
    let source = fixture(FixtureMode::SingleRoot, None)?;
    let package = load_package(&source)?;
    assert_eq!(
        package.table_cell_comment(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(1, 0),
        )?,
        None,
    );

    let commit = package.set_table_cell_comment(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(1, 0),
        "new strict root",
    )?;
    assert_eq!(
        commit
            .package()
            .table_cell_comment(
                SheetSelector::index(0),
                TableSelector::index(0),
                CellPosition::new(1, 0),
            )?
            .as_ref()
            .map(|comment| comment.text()),
        Some("new strict root"),
    );
    assert_eq!(read_all(commit.package(), 0)?, ["first reply"]);
    assert_eq!(commit.diagnostics().deleted_previews(), 3);

    let replay = package.apply_table_cell_comment(commit.patch())?;
    assert_eq!(
        replay
            .package()
            .table_cell_comment(
                SheetSelector::index(0),
                TableSelector::index(0),
                CellPosition::new(1, 0),
            )?
            .as_ref()
            .map(|comment| comment.text()),
        Some("new strict root"),
    );
    assert!(
        commit
            .package()
            .apply_table_cell_comment(commit.patch())
            .is_err()
    );

    let inverse = commit.patch().inverse();
    assert_eq!(inverse.after(), None);
    assert_eq!(
        commit
            .patch()
            .inverse()
            .inverse()
            .after()
            .map(|comment| comment.text()),
        Some("new strict root")
    );
    let restored = commit.package().apply_table_cell_comment(&inverse)?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn root_comment_creation_populates_an_empty_author_storage_atomically() -> TestResult {
    let source = fixture(FixtureMode::Rootless, None)?;
    let package = load_package(&source)?;
    let commit = package.set_table_cell_comment(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(1, 0),
        "new root with generated author",
    )?;
    let created = commit
        .package()
        .table_cell_comment(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(1, 0),
        )?
        .ok_or_else(|| io::Error::other("generated root comment is missing"))?;
    assert_eq!(created.text(), "new root with generated author");
    assert_eq!(
        created.timestamp().map(|timestamp| timestamp.as_f64()),
        Some(0.0),
        "new root comments must retain the writer's zero-epoch timestamp"
    );
    let author = created
        .author()
        .ok_or_else(|| io::Error::other("generated root author is missing"))?;
    assert_eq!(author.display_name(), Some("litchi-iwa"));
    assert_eq!(author.public_id(), None);
    let candidate = exact_bytes(commit.package())?;
    let archive = member_archive(&candidate, DOCUMENT_MEMBER)?;
    let authors = archive
        .objects
        .iter()
        .filter(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == ANNOTATION_AUTHOR_TYPE)
        })
        .count();
    assert_eq!(authors, 1);
    let fresh_ids = archive
        .objects
        .iter()
        .filter_map(|object| object.archive_info.identifier)
        .filter(|identifier| *identifier > WATERMARK)
        .collect::<Vec<_>>();
    assert_eq!(fresh_ids.len(), 2);
    let storage = archive
        .object(AUTHOR_STORAGE_ID)
        .ok_or_else(|| io::Error::other("author storage is missing"))?;
    assert_eq!(
        storage
            .archive_info
            .message_infos
            .first()
            .map(|info| info.object_references.len()),
        Some(1),
    );
    let metadata_archive = member_archive(&candidate, METADATA_MEMBER)?;
    let metadata_payload = metadata_archive
        .objects
        .iter()
        .flat_map(|object| object.messages.iter())
        .find(|message| message.type_ == METADATA_TYPE)
        .ok_or_else(|| io::Error::other("package metadata is missing"))?;
    let metadata = tsp::PackageMetadata::decode(metadata_payload.data.as_slice())?;
    assert_eq!(metadata.last_object_identifier, WATERMARK + 2);
    let document = metadata
        .components
        .iter()
        .find(|component| component.identifier == 100)
        .ok_or_else(|| io::Error::other("Document metadata component is missing"))?;
    for identifier in &fresh_ids {
        assert!(
            document
                .object_uuid_map_entries
                .iter()
                .any(|entry| entry.identifier == *identifier),
            "fresh object {identifier} is missing current UUID ownership"
        );
    }
    let restored = commit
        .package()
        .apply_table_cell_comment(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn root_comment_creation_builds_a_missing_comment_list_atomically() -> TestResult {
    let source = fixture(FixtureMode::Rootless, Some(Corruption::MissingCommentList))?;
    let package = load_package(&source)?;
    let commit = package.set_table_cell_comment(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(1, 0),
        "new root with generated list",
    )?;
    assert_eq!(
        commit
            .package()
            .table_cell_comment(
                SheetSelector::index(0),
                TableSelector::index(0),
                CellPosition::new(1, 0),
            )?
            .as_ref()
            .map(|comment| comment.text()),
        Some("new root with generated list"),
    );
    let candidate = exact_bytes(commit.package())?;
    let archive = member_archive(&candidate, DOCUMENT_MEMBER)?;
    let comment_lists = archive
        .objects
        .iter()
        .filter(|object| {
            object.messages.iter().any(|message| {
                message.type_ == TABLE_DATA_LIST_TYPE
                    && tst::TableDataList::decode(message.data.as_slice())
                        .map(|list| {
                            list.list_type == tst::table_data_list::ListType::CommentStorage as i32
                        })
                        .unwrap_or(false)
            })
        })
        .count();
    assert_eq!(comment_lists, 1);
    let model = archive
        .object(TABLE_MODEL_ID)
        .ok_or_else(|| io::Error::other("table model is missing"))?;
    let decoded = tst::TableModelArchive::decode(model.messages[0].data.as_slice())?;
    let list_identifier = decoded
        .base_data_store
        .comment_storage_table
        .ok_or_else(|| io::Error::other("comment-list reference is missing"))?
        .identifier;
    assert!(list_identifier > WATERMARK);
    let model_info = &model.archive_info.message_infos[0];
    assert!(model_info.object_references.contains(&list_identifier));
    assert!(model_info.field_infos.iter().any(|field| {
        field.path.as_slice() == [4, 19] && field.object_references == [list_identifier]
    }));
    let fresh_ids = archive
        .objects
        .iter()
        .filter_map(|object| object.archive_info.identifier)
        .filter(|identifier| *identifier > WATERMARK)
        .collect::<Vec<_>>();
    assert_eq!(fresh_ids.len(), 3);
    let metadata_archive = member_archive(&candidate, METADATA_MEMBER)?;
    let metadata_payload = metadata_archive
        .objects
        .iter()
        .flat_map(|object| object.messages.iter())
        .find(|message| message.type_ == METADATA_TYPE)
        .ok_or_else(|| io::Error::other("package metadata is missing"))?;
    let metadata = tsp::PackageMetadata::decode(metadata_payload.data.as_slice())?;
    assert_eq!(metadata.last_object_identifier, WATERMARK + 3);
    let document = metadata
        .components
        .iter()
        .find(|component| component.identifier == 100)
        .ok_or_else(|| io::Error::other("Document metadata component is missing"))?;
    for identifier in &fresh_ids {
        assert!(
            document
                .object_uuid_map_entries
                .iter()
                .any(|entry| entry.identifier == *identifier),
            "fresh object {identifier} is missing current UUID ownership"
        );
    }
    let restored = commit
        .package()
        .apply_table_cell_comment(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn root_comment_creation_builds_a_missing_author_storage_atomically() -> TestResult {
    let source = fixture(
        FixtureMode::Rootless,
        Some(Corruption::MissingAuthorStorage),
    )?;
    let package = load_package(&source)?;
    let commit = package.set_table_cell_comment(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(1, 0),
        "new root with generated author storage",
    )?;
    assert_eq!(
        commit
            .package()
            .table_cell_comment(
                SheetSelector::index(0),
                TableSelector::index(0),
                CellPosition::new(1, 0),
            )?
            .as_ref()
            .map(|comment| comment.text()),
        Some("new root with generated author storage"),
    );
    let candidate = exact_bytes(commit.package())?;
    let archive = member_archive(&candidate, DOCUMENT_MEMBER)?;
    let author = archive
        .objects
        .iter()
        .find(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == ANNOTATION_AUTHOR_TYPE)
        })
        .ok_or_else(|| io::Error::other("generated author is missing"))?;
    let author_identifier = author.archive_info.identifier.unwrap_or_default();
    assert!(author_identifier > WATERMARK);
    let storage = archive
        .objects
        .iter()
        .find(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == ANNOTATION_AUTHOR_STORAGE_TYPE)
        })
        .ok_or_else(|| io::Error::other("generated author storage is missing"))?;
    let storage_identifier = storage.archive_info.identifier.unwrap_or_default();
    assert!(storage_identifier > WATERMARK);
    let storage_info = storage
        .archive_info
        .message_infos
        .first()
        .ok_or_else(|| io::Error::other("generated author-storage metadata is missing"))?;
    assert_eq!(storage_info.object_references, [author_identifier]);
    assert!(storage_info.field_infos.iter().any(|field| {
        field.path.as_slice() == [1] && field.object_references == [author_identifier]
    }));
    let fresh_ids = archive
        .objects
        .iter()
        .filter_map(|object| object.archive_info.identifier)
        .filter(|identifier| *identifier > WATERMARK)
        .collect::<Vec<_>>();
    assert_eq!(fresh_ids.len(), 3);
    let metadata_archive = member_archive(&candidate, METADATA_MEMBER)?;
    let metadata_payload = metadata_archive
        .objects
        .iter()
        .flat_map(|object| object.messages.iter())
        .find(|message| message.type_ == METADATA_TYPE)
        .ok_or_else(|| io::Error::other("package metadata is missing"))?;
    let metadata = tsp::PackageMetadata::decode(metadata_payload.data.as_slice())?;
    assert_eq!(metadata.last_object_identifier, WATERMARK + 3);
    let document = metadata
        .components
        .iter()
        .find(|component| component.identifier == 100)
        .ok_or_else(|| io::Error::other("Document metadata component is missing"))?;
    for identifier in &fresh_ids {
        assert!(
            document
                .object_uuid_map_entries
                .iter()
                .any(|entry| entry.identifier == *identifier),
            "fresh object {identifier} is missing current UUID ownership"
        );
    }
    let restored = commit
        .package()
        .apply_table_cell_comment(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn root_comment_creation_materializes_a_missing_cell_slot_atomically() -> TestResult {
    let source = fixture(FixtureMode::SparseCell, None)?;
    let package = load_package(&source)?;
    assert_eq!(
        package
            .table(SheetSelector::index(0), TableSelector::index(0))?
            .ok_or_else(|| io::Error::other("sparse table is missing"))?
            .dimensions()
            .columns(),
        2,
    );
    assert_eq!(
        package.table_cell_comment(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(1, 1),
        )?,
        None,
    );
    let commit = package.set_table_cell_comment(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(1, 1),
        "new root in sparse slot",
    )?;
    assert_eq!(
        commit
            .package()
            .table_cell_comment(
                SheetSelector::index(0),
                TableSelector::index(0),
                CellPosition::new(1, 1),
            )?
            .as_ref()
            .map(|comment| comment.text()),
        Some("new root in sparse slot"),
    );
    assert_eq!(
        commit.package().table_cell_comment(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(1, 0),
        )?,
        None,
    );
    let restored = commit
        .package()
        .apply_table_cell_comment(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn root_comment_creation_fails_closed_on_unsupported_graphs() -> TestResult {
    for (mode, corruption) in [
        (FixtureMode::SingleRoot, Some(Corruption::MissingAuthor)),
        (FixtureMode::SingleRoot, Some(Corruption::UnknownMetadata)),
        (FixtureMode::CrossComponent, None),
    ] {
        let source = fixture(mode, corruption)?;
        let package = load_package(&source)?;
        let result = package.set_table_cell_comment(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(1, 0),
            "must not publish",
        );
        assert!(
            result.is_err(),
            "unsupported root-creation graph was accepted"
        );
        assert_eq!(exact_bytes(&package)?, source);
    }
    Ok(())
}

fn member_data(source: &[u8], name: &str) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    Ok(catalog
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or_else(|| io::Error::other(format!("member {name} is missing")))?
        .data()
        .to_vec())
}

fn member_archive(source: &[u8], name: &str) -> TestResult<Archive> {
    Ok(Archive::parse(
        SnappyStream::decompress(&member_data(source, name)?)?.as_bytes(),
    )?)
}

fn reply_member_name(source: &[u8], identifier: u64) -> TestResult<String> {
    let catalog = Catalog::from_bytes(source)?;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
        if archive.object(identifier).is_some() {
            return Ok(entry.name().to_owned());
        }
    }
    Err(io::Error::other(format!("reply object {identifier} is missing")).into())
}

fn object_exists(source: &[u8], identifier: u64) -> TestResult<bool> {
    let catalog = Catalog::from_bytes(source)?;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
        if archive.object(identifier).is_some() {
            return Ok(true);
        }
    }
    Ok(false)
}

fn rewrite_member(
    source: &[u8],
    name: &str,
    mut rewrite: impl FnMut(&mut Archive) -> TestResult,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or_else(|| io::Error::other(format!("member {name} is missing")))?;
    let mut archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
    rewrite(&mut archive)?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            name,
            &compressed,
        )],
        Limits::default(),
    )?)
}

fn rewrite_reply_payload(
    source: &[u8],
    identifier: u64,
    mut rewrite: impl FnMut(&mut RawMessage) -> TestResult,
) -> TestResult<Vec<u8>> {
    let name = reply_member_name(source, identifier)?;
    rewrite_member(source, &name, |archive| {
        let object = archive
            .object_mut(identifier)
            .ok_or_else(|| io::Error::other("reply object is missing"))?;
        let message = object
            .messages
            .first_mut()
            .ok_or_else(|| io::Error::other("reply payload is missing"))?;
        rewrite(message)
    })
}
fn assert_read_rejected(source: &[u8], label: &str) -> TestResult {
    let original = source.to_vec();
    match Package::from_bytes(source) {
        Err(_) => {},
        Ok(package) => {
            assert!(
                package
                    .table_cell_comment_replies(
                        SheetSelector::index(0),
                        TableSelector::index(0),
                        CellPosition::new(0, 0),
                    )
                    .is_err(),
                "hostile reply source was accepted: {label}"
            );
            assert_eq!(exact_bytes(&package)?, original);
        },
    }
    Ok(())
}

fn assert_edit_rejected_atomically(source: &[u8], label: &str) -> TestResult {
    let original = source.to_vec();
    let package = match Package::from_bytes(source) {
        Err(_) => return Ok(()),
        Ok(package) => package,
    };
    let before = exact_bytes(&package)?;
    let result = package.set_table_cell_comment_reply(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(0, 0),
        CommentReplyIndex::new(0),
        "replacement",
    );
    assert!(
        result.is_err(),
        "hostile reply source was published: {label}"
    );
    assert_eq!(exact_bytes(&package)?, before);
    assert_eq!(source, original.as_slice());
    Ok(())
}

fn with_corruption(mode: FixtureMode, corruption: Corruption) -> TestResult<Vec<u8>> {
    let source = fixture(mode, Some(corruption))?;
    match corruption {
        Corruption::UnknownWire => rewrite_reply_payload(&source, FIRST_REPLY_ID, |message| {
            // Field 90 is an overlong unknown scalar.  The balanced group is
            // deliberately unknown to the selected CommentStorage schema.
            message.data.extend_from_slice(&[
                0xd0, 0x05, 0x80, 0x00, // unknown field 90, overlong zero
                0xdb, 0x05, 0x08, 0x07, 0xdc, 0x05, // balanced field-91 group
            ]);
            Ok(())
        }),
        _ => Ok(source),
    }
}

#[test]
fn malformed_reply_graphs_are_rejected_before_any_public_value() -> TestResult {
    for corruption in [
        Corruption::DuplicateReplyReference,
        Corruption::SelfReplyReference,
        Corruption::MissingReplyReference,
        Corruption::NestedReplyReference,
        Corruption::ExternalReplyReference,
        Corruption::TypedReplyReference,
        Corruption::DuplicateRootPayload,
        Corruption::WrongReplyType,
        Corruption::MissingReply,
        Corruption::SegmentedList,
        Corruption::DuplicateListKey,
        Corruption::WrongListType,
        Corruption::ZeroRefcount,
        Corruption::AggregateMissing,
        Corruption::AggregateExtra,
        Corruption::FieldInfoMissing,
        Corruption::FieldInfoDuplicate,
        Corruption::FieldInfoWrong,
        Corruption::OpaqueInbound,
        Corruption::UnknownMetadata,
        Corruption::MissingAuthor,
    ] {
        let source = with_corruption(FixtureMode::SingleRoot, corruption)?;
        assert_read_rejected(&source, &format!("{corruption:?}"))?;
        assert_edit_rejected_atomically(&source, &format!("{corruption:?}"))?;
    }
    Ok(())
}

#[test]
fn locked_tables_reject_changed_reply_edits_atomically() -> TestResult {
    let source = with_corruption(FixtureMode::SingleRoot, Corruption::Locked)?;
    assert_edit_rejected_atomically(&source, "locked table")
}

#[test]
fn stored_reply_refcounts_must_match_the_complete_bnc_census() -> TestResult {
    for corruption in [
        Corruption::RefcountUndercount,
        Corruption::RefcountOvercount,
    ] {
        let source = with_corruption(FixtureMode::SharedRoot, corruption)?;
        assert_read_rejected(&source, &format!("{corruption:?}"))?;
        assert_edit_rejected_atomically(&source, &format!("{corruption:?}"))?;
    }
    Ok(())
}

#[test]
fn metadata_ownership_and_missing_metadata_fail_closed_atomically() -> TestResult {
    for corruption in [
        Corruption::MissingMetadata,
        Corruption::MissingReplyUuid,
        Corruption::VersionedReplyUuid,
        Corruption::DuplicateReplyUuid,
        Corruption::AmbiguousReplyIdentifier,
        Corruption::DataOwnerReplyIdentifier,
        Corruption::RootDataMapReplyIdentifier,
    ] {
        let source = with_corruption(FixtureMode::SingleRoot, corruption)?;
        assert_edit_rejected_atomically(&source, &format!("{corruption:?}"))?;
    }
    Ok(())
}

#[test]
fn cross_component_and_unknown_inbound_edges_do_not_publish_partial_replies() -> TestResult {
    let split = fixture(FixtureMode::CrossComponent, None)?;
    assert_edit_rejected_atomically(&split, "cross-component reply graph")?;

    let inbound = with_corruption(FixtureMode::SingleRoot, Corruption::OpaqueInbound)?;
    assert_edit_rejected_atomically(&inbound, "opaque inbound reply edge")?;
    Ok(())
}

#[test]
fn unknown_scalar_and_group_are_retained_or_rejected_without_source_mutation() -> TestResult {
    let source = with_corruption(FixtureMode::SingleRoot, Corruption::UnknownWire)?;
    let original = source.clone();
    let package = match Package::from_bytes(&source) {
        Err(_) => return Ok(()),
        Ok(package) => package,
    };
    let before = exact_bytes(&package)?;
    let result = package.set_table_cell_comment_reply(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(0, 0),
        CommentReplyIndex::new(0),
        "unknown-preserving rewrite",
    );
    match result {
        Err(_) => assert_eq!(exact_bytes(&package)?, before),
        Ok(commit) => {
            let target = exact_bytes(commit.package())?;
            let member = reply_member_name(&target, FIRST_REPLY_ID)?;
            let archive = member_archive(&target, &member)?;
            let reply = archive
                .object(FIRST_REPLY_ID)
                .ok_or_else(|| io::Error::other("rewritten reply is missing"))?;
            let data = &reply
                .messages
                .first()
                .ok_or_else(|| io::Error::other("rewritten reply payload is missing"))?
                .data;
            assert!(data.windows(2).any(|window| window == [0xdb, 0x05]));
            assert!(data.windows(2).any(|window| window == [0xd0, 0x05]));
            let inverse = commit.patch().inverse();
            let restored = commit.package().apply_table_cell_comment_reply(&inverse)?;
            assert_eq!(exact_bytes(restored.package())?, original);
        },
    }
    Ok(())
}

#[test]
fn semantic_and_physical_limits_reject_without_mutating_the_source() -> TestResult {
    let source = fixture(FixtureMode::DuplicateText, None)?;
    let original = source.clone();
    let semantic = PackageSemanticLimits::new(
        PackageSemanticLimits::MAX_OBJECTS,
        1,
        PackageSemanticLimits::MAX_TABLES,
        1,
    )?;
    let options = PackageReadOptions::new(PackageLimits::default(), semantic);
    assert!(Package::from_bytes_with_options(&source, options).is_err());
    assert_eq!(source, original);

    let default_semantic = PackageSemanticLimits::default();
    let capped_semantic =
        default_semantic.with_projection_limits(default_semantic.max_materialized_cells(), 1024)?;
    let package = Package::from_bytes_with_options(
        &source,
        PackageReadOptions::new(PackageLimits::default(), capped_semantic),
    )?;
    let before = exact_bytes(&package)?;
    let result = package
        .edit_table_cell_comment_replies(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .append("x".repeat(16 * 1024))
        .commit();
    assert!(result.is_err());
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn reply_index_is_checked_and_duplicate_text_uses_source_ordinal() -> TestResult {
    let source = fixture(FixtureMode::DuplicateText, None)?;
    let original = source.clone();
    let package = load_package(&source)?;

    assert_eq!(
        read_all(&package, 0)?,
        vec![
            "duplicate".to_owned(),
            "duplicate".to_owned(),
            "duplicate".to_owned()
        ]
    );
    assert_eq!(read_one(&package, 0, 0)?, "duplicate");
    assert_eq!(read_one(&package, 0, 1)?, "duplicate");
    assert_eq!(read_one(&package, 0, 2)?, "duplicate");
    assert!(
        package
            .table_cell_comment_reply(
                SheetSelector::index(0),
                TableSelector::index(0),
                CellPosition::new(0, 0),
                CommentReplyIndex::new(3),
            )
            .is_err()
    );

    let index = CommentReplyIndex::new(1);
    assert_eq!(index.index(), 1);
    assert_eq!(index.get(), 1);
    assert!(CommentReplyIndex::try_from_usize(usize::MAX).is_err());
    let addressed = package.table_cell_comment_reply_a1(
        SheetSelector::index(0),
        TableSelector::index(0),
        "A1",
        index,
    )?;
    assert_eq!(addressed.text(), "duplicate");
    assert_eq!(exact_bytes(&package)?, original);
    Ok(())
}

#[test]
fn collection_noop_is_exact_and_debug_redacts_reply_text() -> TestResult {
    let source = fixture(FixtureMode::SingleRoot, None)?;
    let package = load_package(&source)?;
    let original = exact_bytes(&package)?;
    let before = read_one(&package, 0, 0)?;

    let edit = package.edit_table_cell_comment_replies(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(0, 0),
    )?;
    assert!(!format!("{edit:?}").contains(&before));
    let commit = edit
        .set(CommentReplyIndex::new(0), before.clone())
        .commit()?;
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(exact_bytes(commit.package())?, original);
    assert!(!format!("{:?}", commit.patch()).contains(&before));
    assert!(!format!("{commit:?}").contains(&before));
    Ok(())
}

#[test]
fn append_set_remove_and_a1_wrappers_are_reversible() -> TestResult {
    let source = fixture(FixtureMode::SingleRoot, None)?;
    let package = load_package(&source)?;
    let original = exact_bytes(&package)?;

    let appended = package
        .edit_table_cell_comment_replies(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .append("second reply")
        .commit()?;
    assert_eq!(
        read_all(appended.package(), 0)?,
        vec!["first reply".to_owned(), "second reply".to_owned()]
    );
    assert!(appended.diagnostics().changed());
    assert!(!appended.patch().is_noop());
    let appended_target = exact_bytes(appended.package())?;
    let reopened = load_package(&appended_target)?;
    assert_eq!(
        read_all(&reopened, 0)?,
        vec!["first reply".to_owned(), "second reply".to_owned()]
    );

    let inverse = appended.patch().inverse();
    let restored = appended
        .package()
        .apply_table_cell_comment_reply(&inverse)?;
    assert_eq!(exact_bytes(restored.package())?, original);

    let set = package.set_table_cell_comment_reply(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(0, 0),
        CommentReplyIndex::new(0),
        "replaced reply",
    )?;
    assert_eq!(read_one(set.package(), 0, 0)?, "replaced reply");

    let removed = package.remove_table_cell_comment_reply_a1(
        SheetSelector::index(0),
        TableSelector::index(0),
        "A1",
        CommentReplyIndex::new(0),
    )?;
    assert!(read_all(removed.package(), 0)?.is_empty());

    let directly_added = package.add_table_cell_comment_reply(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(0, 0),
        "directly appended reply",
    )?;
    assert_eq!(
        read_all(directly_added.package(), 0)?,
        vec![
            "first reply".to_owned(),
            "directly appended reply".to_owned()
        ]
    );
    let directly_added_a1 = package.add_table_cell_comment_reply_a1(
        SheetSelector::index(0),
        TableSelector::index(0),
        "A1",
        "direct A1 reply",
    )?;
    assert_eq!(
        read_all(directly_added_a1.package(), 0)?,
        vec!["first reply".to_owned(), "direct A1 reply".to_owned()]
    );
    Ok(())
}

#[test]
fn reply_text_edits_preserve_metadata_and_append_inherits_the_source_leaf() -> TestResult {
    let source = fixture(FixtureMode::SingleRoot, None)?;
    let package = load_package(&source)?;
    let original = package.table_cell_comment_replies(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(0, 0),
    )?;
    assert_eq!(original.len(), 1);
    let original_reply = &original[0];
    assert!(original_reply.timestamp().is_some());
    assert!(original_reply.author().is_some());

    let appended = package.add_table_cell_comment_reply(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(0, 0),
        "metadata-preserving append",
    )?;
    let appended_replies = appended.package().table_cell_comment_replies(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(0, 0),
    )?;
    assert_eq!(appended.patch().before(), original.as_ref());
    assert_eq!(appended.patch().after(), appended_replies.as_ref());
    assert_eq!(appended_replies[0], *original_reply);
    assert_eq!(
        appended_replies[1].timestamp(),
        original_reply.timestamp(),
        "an appended reply must inherit its source leaf timestamp"
    );
    assert_eq!(
        appended_replies[1].author(),
        original_reply.author(),
        "an appended reply must inherit its source leaf author"
    );

    let replaced = package.set_table_cell_comment_reply(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(0, 0),
        CommentReplyIndex::new(0),
        "metadata-preserving replacement",
    )?;
    let replaced_reply = replaced.package().table_cell_comment_reply(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(0, 0),
        CommentReplyIndex::new(0),
    )?;
    assert_eq!(replaced.patch().before(), original.as_ref());
    assert_eq!(replaced.patch().after()[0], replaced_reply);
    assert_eq!(replaced_reply.timestamp(), original_reply.timestamp());
    assert_eq!(replaced_reply.author(), original_reply.author());
    Ok(())
}

#[test]
fn direct_a1_read_and_collection_a1_edit_select_the_same_reply() -> TestResult {
    let source = fixture(FixtureMode::DuplicateText, None)?;
    let package = load_package(&source)?;
    let index = CommentReplyIndex::new(2);
    assert_eq!(
        package
            .table_cell_comment_reply_a1(
                SheetSelector::index(0),
                TableSelector::index(0),
                "A1",
                index,
            )?
            .text(),
        "duplicate"
    );
    let changed = package
        .edit_table_cell_comment_replies_a1(SheetSelector::index(0), TableSelector::index(0), "A1")?
        .set(index, "third changed")
        .commit()?;
    assert_eq!(read_one(changed.package(), 0, 2)?, "third changed");
    assert_eq!(read_one(changed.package(), 0, 0)?, "duplicate");
    Ok(())
}

#[test]
fn shared_root_edit_is_copy_on_write_and_preserves_sibling() -> TestResult {
    let source = fixture(FixtureMode::SharedRoot, None)?;
    let package = load_package(&source)?;
    assert_eq!(read_all(&package, 0)?, vec!["first reply".to_owned()]);
    assert_eq!(read_all(&package, 1)?, vec!["first reply".to_owned()]);

    let commit = package.set_table_cell_comment_reply(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(0, 0),
        CommentReplyIndex::new(0),
        "first-only",
    )?;
    assert_eq!(
        read_all(commit.package(), 0)?,
        vec!["first-only".to_owned()]
    );
    assert_eq!(
        read_all(commit.package(), 1)?,
        vec!["first reply".to_owned()]
    );
    assert!(object_exists(
        &exact_bytes(commit.package())?,
        ROOT_COMMENT_ID
    )?);
    Ok(())
}

#[test]
fn shared_reply_graph_is_rejected_by_the_direct_leaf_scope() -> TestResult {
    let source = fixture(FixtureMode::SharedReply, None)?;
    assert_read_rejected(&source, "shared reply archive")?;
    assert_edit_rejected_atomically(&source, "shared reply archive")
}

#[test]
fn removing_the_last_reply_culls_the_unshared_reply_archive() -> TestResult {
    let source = fixture(FixtureMode::SingleRoot, None)?;
    let package = load_package(&source)?;
    let commit = package.remove_table_cell_comment_reply(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(0, 0),
        CommentReplyIndex::new(0),
    )?;
    assert!(read_all(commit.package(), 0)?.is_empty());
    assert!(!object_exists(
        &exact_bytes(commit.package())?,
        FIRST_REPLY_ID
    )?);
    Ok(())
}

#[test]
fn inverse_conflict_and_locality_are_exact_source_operations() -> TestResult {
    let source = fixture(FixtureMode::SingleRoot, None)?;
    let package = load_package(&source)?;
    let source_view = member_data(&source, VIEW_STATE_MEMBER)?;
    let source_data = member_data(&source, "Data/sentinel.bin")?;

    let first = package.set_table_cell_comment_reply(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(0, 0),
        CommentReplyIndex::new(0),
        "first branch",
    )?;
    let second = package.set_table_cell_comment_reply(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(0, 0),
        CommentReplyIndex::new(0),
        "second branch",
    )?;
    assert!(
        second
            .package()
            .apply_table_cell_comment_reply(first.patch())
            .is_err()
    );

    let target = exact_bytes(first.package())?;
    assert_eq!(member_data(&target, VIEW_STATE_MEMBER)?, source_view);
    assert_eq!(member_data(&target, "Data/sentinel.bin")?, source_data);
    let reopened = load_package(&target)?;
    assert_eq!(read_one(&reopened, 0, 0)?, "first branch");
    let inverse = first.patch().inverse();
    let restored = first.package().apply_table_cell_comment_reply(&inverse)?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn missing_comment_root_reports_a_typed_error_without_source_mutation() -> TestResult {
    let source = fixture(FixtureMode::SingleRoot, None)?;
    let package = load_package(&source)?;
    let before = exact_bytes(&package)?;
    let error = package
        .table_cell_comment_reply(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(1, 0),
            CommentReplyIndex::new(0),
        )
        .expect_err("a cell without a root comment must reject an ordinal read");
    assert!(!format!("{error:?}").is_empty());
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn export_comment_metadata_fuzz_fixture() -> TestResult {
    let Some(path) = std::env::var_os("LITCHI_NUMBERS_COMMENT_METADATA_FUZZ_OUTPUT") else {
        return Ok(());
    };
    let source = fixture(FixtureMode::DuplicateText, None)?;
    let package = load_package(&source)?;
    assert_eq!(read_all(&package, 0)?, ["duplicate"; 3]);
    std::fs::write(path, source)?;
    Ok(())
}
