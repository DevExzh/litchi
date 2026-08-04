use super::*;
use crate::archive::Archive;

fn drawable(identifier: u64) -> DrawableObjectId {
    DrawableObjectId::new(identifier).unwrap()
}

fn storage(identifier: u64) -> CommentStorageId {
    CommentStorageId::new(identifier).unwrap()
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn fixture_comment_storage_uuid(identifier: u64) -> tsp::Uuid {
    tsp::Uuid {
        lower: identifier,
        upper: identifier.rotate_left(23) ^ 0x6c69_7463_6869_6977,
    }
}

fn object<T: Message>(identifier: u64, type_: u32, value: T) -> ArchiveObject {
    ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_,
            data: value.encode_to_vec(),
        }],
    )
    .unwrap()
}

fn placeholder(comment: Option<u64>) -> kn::PlaceholderArchive {
    kn::PlaceholderArchive {
        super_: tswp::ShapeInfoArchive {
            super_: tsd::ShapeArchive {
                super_: tsd::DrawableArchive {
                    comment: comment.map(reference),
                    ..Default::default()
                },
                ..Default::default()
            },
            ..Default::default()
        },
        ..Default::default()
    }
}

fn comment_storage(identifier: u64, text: &str, replies: Vec<u64>) -> ArchiveObject {
    let mut object = object(
        identifier,
        COMMENT_STORAGE_MESSAGE_TYPE,
        tsd::CommentStorageArchive {
            text: Some(text.to_owned()),
            creation_date: Some(tsp::Date { seconds: 42.5 }),
            author: Some(reference(30)),
            replies: replies.iter().copied().map(reference).collect(),
            storage_uuid: Some(fixture_comment_storage_uuid(identifier)),
        },
    );
    object.archive_info.message_infos[0]
        .object_references
        .extend([30].into_iter().chain(replies));
    object
}

fn keynote_package(shared_comment: bool) -> IWorkPackage {
    let document = kn::DocumentArchive {
        show: reference(2),
        ..Default::default()
    };
    let mut first = object(5, 7, placeholder(Some(20)));
    first.archive_info.message_infos[0]
        .object_references
        .push(20);
    let mut second = object(6, 7, placeholder(shared_comment.then_some(20)));
    if shared_comment {
        second.archive_info.message_infos[0]
            .object_references
            .push(20);
    }
    let mut package = IWorkPackage::new();
    package
        .replace_archive(
            "Index/Document.iwa",
            &Archive {
                objects: vec![
                    object(1, 1, document),
                    object(2, 2, kn::ShowArchive::default()),
                    first,
                    second,
                    comment_storage(20, "Original", vec![21]),
                    comment_storage(21, "Reply", Vec::new()),
                    object(
                        30,
                        ANNOTATION_AUTHOR_MESSAGE_TYPE,
                        tsk::AnnotationAuthorArchive {
                            name: Some("Native author".to_owned()),
                            ..Default::default()
                        },
                    ),
                ],
            },
        )
        .unwrap();
    package
}

fn keynote_package_with_empty_author_storage() -> IWorkPackage {
    let mut package = keynote_package(false);
    package
        .update_archive("Index/Document.iwa", |archive| {
            let drawable = archive.object_mut(5).unwrap();
            let mut value = kn::PlaceholderArchive::decode(drawable.messages[0].data.as_slice())?;
            value.super_.super_.super_.comment = None;
            drawable.replace_message(
                0,
                RawMessage {
                    type_: 7,
                    data: value.encode_to_vec(),
                },
            )?;
            drawable.archive_info.message_infos[0]
                .object_references
                .clear();
            Ok(())
        })
        .unwrap();
    package
        .replace_archive(
            "Index/AnnotationAuthorStorage.iwa",
            &Archive {
                objects: vec![object(
                    40,
                    ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE,
                    tsk::AnnotationAuthorStorageArchive::default(),
                )],
            },
        )
        .unwrap();
    package
        .replace_archive(
            "Index/Metadata.iwa",
            &Archive {
                objects: vec![object(
                    100,
                    crate::package_metadata::PACKAGE_METADATA_MESSAGE_TYPE,
                    tsp::PackageMetadata {
                        last_object_identifier: 100,
                        components: vec![
                            tsp::ComponentInfo {
                                identifier: 1,
                                preferred_locator: "Document".to_owned(),
                                save_token: Some(7),
                                ..Default::default()
                            },
                            tsp::ComponentInfo {
                                identifier: 40,
                                preferred_locator: "AnnotationAuthorStorage".to_owned(),
                                save_token: Some(6),
                                ..Default::default()
                            },
                        ],
                        save_token: Some(8),
                        ..Default::default()
                    },
                )],
            },
        )
        .unwrap();
    package
}

fn package_metadata(package: &IWorkPackage) -> tsp::PackageMetadata {
    let archive = package.archive("Index/Metadata.iwa").unwrap();
    let message = archive
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .find(|message| message.type_ == crate::package_metadata::PACKAGE_METADATA_MESSAGE_TYPE)
        .unwrap();
    tsp::PackageMetadata::decode(message.data.as_slice()).unwrap()
}

fn append_unknown_varint(data: &mut Vec<u8>, field_number: u32, value: u64) -> Vec<u8> {
    let mut field = litchi_iwa_common::varint::encode_varint(u64::from(field_number) << 3);
    field.extend(litchi_iwa_common::varint::encode_varint(value));
    data.extend_from_slice(&field);
    field
}

fn placeholder_bytes_with_unknown_fields() -> Vec<u8> {
    let mut drawable = tsd::DrawableArchive::default().encode_to_vec();
    append_unknown_varint(&mut drawable, 90, 900);

    let shape = tsd::ShapeArchive::default().encode_to_vec();
    let mut shape = patch_length_delimited_field(&shape, 1, true, Some(&drawable)).unwrap();
    append_unknown_varint(&mut shape, 91, 901);

    let shape_info = tswp::ShapeInfoArchive::default().encode_to_vec();
    let mut shape_info = patch_length_delimited_field(&shape_info, 1, true, Some(&shape)).unwrap();
    append_unknown_varint(&mut shape_info, 92, 902);

    let placeholder = kn::PlaceholderArchive::default().encode_to_vec();
    let mut placeholder =
        patch_length_delimited_field(&placeholder, 1, true, Some(&shape_info)).unwrap();
    append_unknown_varint(&mut placeholder, 93, 903);
    placeholder
}

fn object_payload(package: &IWorkPackage, identifier: u64) -> Vec<u8> {
    let locations = object_locations(package).unwrap();
    let archive = package.archive(&locations[&identifier]).unwrap();
    archive.object(identifier).unwrap().messages[0].data.clone()
}

#[test]
fn drawable_and_storage_ids_reject_zero_and_support_typed_editor_calls() {
    assert_eq!(DrawableObjectId::new(0), None);
    assert_eq!(CommentStorageId::new(0), None);

    let drawable_id = DrawableObjectId::try_from(5).unwrap();
    assert_eq!(drawable_id.object_id(), 5);
    assert_eq!(u64::from(drawable_id), 5);

    let mut editor =
        IWorkDrawableCommentEditor::from_package(keynote_package_with_empty_author_storage())
            .unwrap();
    editor.set_comment(drawable_id, "Typed root").unwrap();
    let root = editor.comment(drawable_id).unwrap().unwrap();
    assert_eq!(root.drawable_object_id, drawable_id);
    assert_eq!(root.storage_object_id.object_id(), 102);

    let reply_id = editor.add_reply(drawable_id, "Typed reply").unwrap();
    assert_eq!(reply_id.object_id(), 104);
    let updated_reply_id = editor
        .set_reply(drawable_id, reply_id, "Updated typed reply")
        .unwrap();
    assert_eq!(updated_reply_id.object_id(), 106);
    editor.remove_reply(drawable_id, updated_reply_id).unwrap();
    assert!(DrawableObjectId::from_object_id(0).is_err());
}

#[test]
fn creates_updates_and_clears_direct_drawable_comments() {
    let mut package = keynote_package(false);
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(5).unwrap();
            let mut value = kn::PlaceholderArchive::decode(object.messages[0].data.as_slice())?;
            value.super_.super_.super_.comment = None;
            object.replace_message(
                0,
                RawMessage {
                    type_: 7,
                    data: value.encode_to_vec(),
                },
            )?;
            object.archive_info.message_infos[0]
                .object_references
                .clear();
            Ok(())
        })
        .unwrap();

    let mut editor = IWorkDrawableCommentEditor::from_package(package).unwrap();
    assert_eq!(editor.application(), Application::Keynote);
    assert_eq!(editor.drawables().unwrap().len(), 2);
    assert!(editor.comment(drawable(5)).unwrap().is_none());

    editor.set_comment(drawable(5), "Created").unwrap();
    let created = editor.comment(drawable(5)).unwrap().unwrap();
    assert_eq!(created.comment.text, "Created");
    assert!(created.comment.creation_date_seconds.is_some());
    assert!(created.comment.storage_uuid.is_some());
    let bytes = editor.to_bytes().unwrap();
    editor.set_comment(drawable(5), "Created").unwrap();
    assert_eq!(editor.to_bytes().unwrap(), bytes);

    editor.set_comment(drawable(5), "Updated").unwrap();
    let updated = editor.comment(drawable(5)).unwrap().unwrap();
    assert_eq!(updated.storage_object_id, created.storage_object_id);
    assert_eq!(updated.comment.text, "Updated");
    editor.clear_comment(drawable(5)).unwrap();
    assert!(editor.comment(drawable(5)).unwrap().is_none());
    assert!(
        !object_locations(editor.package())
            .unwrap()
            .contains_key(&created.storage_object_id.object_id())
    );
}

#[test]
fn creates_reuses_and_cleans_native_author_graph() {
    let mut editor =
        IWorkDrawableCommentEditor::from_package(keynote_package_with_empty_author_storage())
            .unwrap();

    editor.set_comment(drawable(5), "First").unwrap();
    let first = editor.comment(drawable(5)).unwrap().unwrap();
    assert_eq!(first.storage_object_id.object_id(), 102);
    assert_eq!(first.comment.author_object_id, Some(101));
    assert!(first.comment.creation_date_seconds.is_some());
    assert!(first.comment.storage_uuid.is_some());

    let author_archive = editor
        .package()
        .archive("Index/AnnotationAuthorStorage.iwa")
        .unwrap();
    let storage = tsk::AnnotationAuthorStorageArchive::decode(
        author_archive.object(40).unwrap().messages[0]
            .data
            .as_slice(),
    )
    .unwrap();
    assert_eq!(
        storage
            .annotation_author
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>(),
        [101]
    );
    let author = author_archive.object(101).unwrap();
    assert_eq!(
        tsk::AnnotationAuthorArchive::decode(author.messages[0].data.as_slice()).unwrap(),
        generated_annotation_author()
    );
    assert_eq!(
        author.archive_info.message_infos[0]
            .field_infos
            .iter()
            .map(|field| field.path.path.as_slice())
            .collect::<Vec<_>>(),
        [vec![4].as_slice(), vec![3].as_slice()]
    );

    let metadata = package_metadata(editor.package());
    assert_eq!(metadata.last_object_identifier, 102);
    assert_eq!(metadata.save_token, Some(9));
    for identifier in [1, 40] {
        assert_eq!(
            metadata
                .components
                .iter()
                .find(|component| component.identifier == identifier)
                .unwrap()
                .save_token,
            Some(9)
        );
    }
    let document = metadata
        .components
        .iter()
        .find(|component| component.identifier == 1)
        .unwrap();
    assert_eq!(
        document
            .external_references
            .iter()
            .filter(|reference| {
                reference.component_identifier == 40 && reference.object_identifier == Some(101)
            })
            .count(),
        1
    );

    editor.set_comment(drawable(6), "Second").unwrap();
    let second = editor.comment(drawable(6)).unwrap().unwrap();
    assert_eq!(second.storage_object_id.object_id(), 103);
    assert_eq!(second.comment.author_object_id, Some(101));
    let metadata = package_metadata(editor.package());
    assert_eq!(metadata.last_object_identifier, 103);
    assert_eq!(metadata.save_token, Some(10));
    assert_eq!(
        metadata
            .components
            .iter()
            .find(|component| component.identifier == 1)
            .unwrap()
            .save_token,
        Some(10)
    );
    assert_eq!(
        metadata
            .components
            .iter()
            .find(|component| component.identifier == 40)
            .unwrap()
            .save_token,
        Some(9)
    );
    assert_eq!(
        metadata
            .components
            .iter()
            .flat_map(|component| &component.external_references)
            .filter(|reference| {
                reference.component_identifier == 40 && reference.object_identifier == Some(101)
            })
            .count(),
        1
    );

    editor.clear_comment(drawable(5)).unwrap();
    assert!(
        object_locations(editor.package())
            .unwrap()
            .contains_key(&101)
    );
    let metadata = package_metadata(editor.package());
    assert_eq!(metadata.save_token, Some(11));
    assert_eq!(
        metadata
            .components
            .iter()
            .find(|component| component.identifier == 1)
            .unwrap()
            .save_token,
        Some(11)
    );
    assert_eq!(
        metadata
            .components
            .iter()
            .find(|component| component.identifier == 40)
            .unwrap()
            .save_token,
        Some(9)
    );
    editor.clear_comment(drawable(6)).unwrap();
    let locations = object_locations(editor.package()).unwrap();
    assert!(!locations.contains_key(&101));
    assert!(!locations.contains_key(&102));
    assert!(!locations.contains_key(&103));
    let metadata = package_metadata(editor.package());
    assert_eq!(metadata.last_object_identifier, 100);
    assert_eq!(metadata.save_token, Some(12));
    for identifier in [1, 40] {
        assert_eq!(
            metadata
                .components
                .iter()
                .find(|component| component.identifier == identifier)
                .unwrap()
                .save_token,
            Some(12)
        );
    }
    assert!(
        metadata
            .components
            .iter()
            .flat_map(|component| &component.external_references)
            .all(|reference| reference.object_identifier != Some(101))
    );
    let author_archive = editor
        .package()
        .archive("Index/AnnotationAuthorStorage.iwa")
        .unwrap();
    let storage = tsk::AnnotationAuthorStorageArchive::decode(
        author_archive.object(40).unwrap().messages[0]
            .data
            .as_slice(),
    )
    .unwrap();
    assert!(storage.annotation_author.is_empty());
}

#[test]
fn duplicate_author_component_reference_fails_transactionally() {
    let mut editor =
        IWorkDrawableCommentEditor::from_package(keynote_package_with_empty_author_storage())
            .unwrap();
    editor.set_comment(drawable(5), "First").unwrap();
    let mut package = editor.into_package();
    package
        .update_archive("Index/Metadata.iwa", |archive| {
            let object = archive.object_mut(100).unwrap();
            let mut metadata = tsp::PackageMetadata::decode(object.messages[0].data.as_slice())?;
            let document = metadata
                .components
                .iter_mut()
                .find(|component| component.identifier == 1)
                .unwrap();
            document
                .external_references
                .push(document.external_references[0]);
            object.replace_message(
                0,
                RawMessage {
                    type_: crate::package_metadata::PACKAGE_METADATA_MESSAGE_TYPE,
                    data: metadata.encode_to_vec(),
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = IWorkDrawableCommentEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.set_comment(drawable(6), "Rejected").is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn creates_updates_and_removes_replies_with_native_copy_on_write() {
    let mut editor =
        IWorkDrawableCommentEditor::from_package(keynote_package_with_empty_author_storage())
            .unwrap();
    editor.set_comment(drawable(5), "Root").unwrap();
    let original_root = editor.comment(drawable(5)).unwrap().unwrap();
    assert_eq!(original_root.storage_object_id.object_id(), 102);

    let reply_id = editor.add_reply(drawable(5), "First reply").unwrap();
    assert_eq!(reply_id.object_id(), 104);
    let root = editor.comment(drawable(5)).unwrap().unwrap();
    assert_eq!(root.storage_object_id.object_id(), 103);
    assert_eq!(
        root.comment.storage_uuid,
        original_root.comment.storage_uuid
    );
    assert_eq!(root.comment.reply_object_ids, [104]);
    assert!(
        !object_locations(editor.package())
            .unwrap()
            .contains_key(&102)
    );
    let replies = editor.replies(drawable(5)).unwrap();
    assert_eq!(replies.len(), 1);
    assert_eq!(replies[0].storage_object_id.object_id(), 104);
    assert_eq!(replies[0].comment.text, "First reply");
    assert_eq!(replies[0].comment.author_object_id, Some(101));
    let original_reply_uuid = replies[0].comment.storage_uuid;

    let updated_reply_id = editor
        .set_reply(drawable(5), reply_id, "Updated reply")
        .unwrap();
    assert_eq!(updated_reply_id.object_id(), 106);
    let root = editor.comment(drawable(5)).unwrap().unwrap();
    assert_eq!(root.storage_object_id.object_id(), 105);
    assert_eq!(
        root.comment.storage_uuid,
        original_root.comment.storage_uuid
    );
    assert_eq!(root.comment.reply_object_ids, [106]);
    assert!(
        !object_locations(editor.package())
            .unwrap()
            .contains_key(&103)
    );
    assert!(
        !object_locations(editor.package())
            .unwrap()
            .contains_key(&104)
    );
    let updated = editor.replies(drawable(5)).unwrap().remove(0);
    assert_eq!(updated.comment.text, "Updated reply");
    assert_eq!(updated.comment.storage_uuid, original_reply_uuid);
    let unchanged = editor.to_bytes().unwrap();
    assert_eq!(
        editor
            .set_reply(drawable(5), updated_reply_id, "Updated reply")
            .unwrap(),
        updated_reply_id
    );
    assert_eq!(editor.to_bytes().unwrap(), unchanged);

    editor.remove_reply(drawable(5), updated_reply_id).unwrap();
    let root = editor.comment(drawable(5)).unwrap().unwrap();
    assert_eq!(root.storage_object_id.object_id(), 107);
    assert!(root.comment.reply_object_ids.is_empty());
    assert!(editor.replies(drawable(5)).unwrap().is_empty());
    let locations = object_locations(editor.package()).unwrap();
    assert!(!locations.contains_key(&105));
    assert!(!locations.contains_key(&106));
    assert!(locations.contains_key(&101));

    editor.clear_comment(drawable(5)).unwrap();
    let locations = object_locations(editor.package()).unwrap();
    assert!(!locations.contains_key(&101));
    assert!(!locations.contains_key(&107));
    let metadata = package_metadata(editor.package());
    assert_eq!(metadata.last_object_identifier, 100);
    assert_eq!(metadata.save_token, Some(13));
    for identifier in [1, 40] {
        assert_eq!(
            metadata
                .components
                .iter()
                .find(|component| component.identifier == identifier)
                .unwrap()
                .save_token,
            Some(13)
        );
    }
}

#[test]
fn reply_mutations_isolate_shared_comment_threads() {
    let mut editor = IWorkDrawableCommentEditor::from_package(keynote_package(true)).unwrap();
    let added_reply = editor.add_reply(drawable(5), "Added").unwrap();
    assert_eq!(added_reply.object_id(), 32);
    assert_eq!(
        editor
            .comment(drawable(5))
            .unwrap()
            .unwrap()
            .comment
            .reply_object_ids,
        [21, 32]
    );
    assert_eq!(
        editor
            .comment(drawable(6))
            .unwrap()
            .unwrap()
            .storage_object_id
            .object_id(),
        20
    );
    assert_eq!(
        editor
            .comment(drawable(6))
            .unwrap()
            .unwrap()
            .comment
            .reply_object_ids,
        [21]
    );

    let updated_reply = editor
        .set_reply(drawable(5), storage(21), "Isolated update")
        .unwrap();
    assert_eq!(updated_reply.object_id(), 34);
    let first = editor.replies(drawable(5)).unwrap();
    assert_eq!(
        first
            .iter()
            .map(|reply| (
                reply.storage_object_id.object_id(),
                reply.comment.text.as_str()
            ))
            .collect::<Vec<_>>(),
        [(34, "Isolated update"), (32, "Added")]
    );
    let second = editor.replies(drawable(6)).unwrap();
    assert_eq!(second[0].storage_object_id.object_id(), 21);
    assert_eq!(second[0].comment.text, "Reply");

    editor.remove_reply(drawable(5), added_reply).unwrap();
    let first = editor.replies(drawable(5)).unwrap();
    assert_eq!(first.len(), 1);
    assert_eq!(first[0].storage_object_id.object_id(), 34);
    let locations = object_locations(editor.package()).unwrap();
    assert!(!locations.contains_key(&32));
    assert!(locations.contains_key(&20));
    assert!(locations.contains_key(&21));
}

#[test]
fn malformed_reply_graph_fails_transactionally() {
    let mut package = keynote_package(true);
    package
        .update_archive("Index/Document.iwa", |archive| {
            let root = archive.object_mut(20).unwrap();
            let mut comment = tsd::CommentStorageArchive::decode(root.messages[0].data.as_slice())?;
            comment.replies.push(reference(21));
            root.replace_message(
                0,
                RawMessage {
                    type_: COMMENT_STORAGE_MESSAGE_TYPE,
                    data: comment.encode_to_vec(),
                },
            )?;
            root.archive_info.message_infos[0]
                .object_references
                .push(21);
            Ok(())
        })
        .unwrap();
    let mut editor = IWorkDrawableCommentEditor::from_package(package).unwrap();
    assert!(editor.replies(drawable(5)).is_err());
    let before = editor.to_bytes().unwrap();
    assert!(editor.add_reply(drawable(5), "Rejected").is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn clearing_malformed_reply_graph_fails_transactionally() {
    let mut package = keynote_package(true);
    package
        .update_archive("Index/Document.iwa", |archive| {
            let root = archive.object_mut(20).unwrap();
            let mut comment = tsd::CommentStorageArchive::decode(root.messages[0].data.as_slice())?;
            comment.replies.push(reference(21));
            root.replace_message(
                0,
                RawMessage {
                    type_: COMMENT_STORAGE_MESSAGE_TYPE,
                    data: comment.encode_to_vec(),
                },
            )?;
            root.archive_info.message_infos[0]
                .object_references
                .push(21);
            Ok(())
        })
        .unwrap();

    let mut editor = IWorkDrawableCommentEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.clear_comment(drawable(5)).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn generated_author_is_not_confused_with_existing_native_author() {
    let mut package = keynote_package_with_empty_author_storage();
    package
        .update_archive("Index/AnnotationAuthorStorage.iwa", |archive| {
            let storage = archive.object_mut(40).unwrap();
            storage.replace_message(
                0,
                RawMessage {
                    type_: ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE,
                    data: tsk::AnnotationAuthorStorageArchive {
                        annotation_author: vec![reference(41)],
                    }
                    .encode_to_vec(),
                },
            )?;
            Ok(archive.insert_object(object(
                41,
                ANNOTATION_AUTHOR_MESSAGE_TYPE,
                tsk::AnnotationAuthorArchive {
                    name: Some("Native author".to_owned()),
                    ..Default::default()
                },
            ))?)
        })
        .unwrap();
    let mut editor = IWorkDrawableCommentEditor::from_package(package).unwrap();
    editor.set_comment(drawable(5), "Generated author").unwrap();
    assert_eq!(
        editor
            .comment(drawable(5))
            .unwrap()
            .unwrap()
            .comment
            .author_object_id,
        Some(101)
    );
    let archive = editor
        .package()
        .archive("Index/AnnotationAuthorStorage.iwa")
        .unwrap();
    let storage = tsk::AnnotationAuthorStorageArchive::decode(
        archive.object(40).unwrap().messages[0].data.as_slice(),
    )
    .unwrap();
    assert_eq!(
        storage
            .annotation_author
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>(),
        [41, 101]
    );
    editor.clear_comment(drawable(5)).unwrap();
    let archive = editor
        .package()
        .archive("Index/AnnotationAuthorStorage.iwa")
        .unwrap();
    assert!(archive.object(41).is_some());
    assert!(archive.object(101).is_none());
}

#[test]
fn shared_comment_uses_copy_on_write_and_preserves_metadata() {
    let mut editor = IWorkDrawableCommentEditor::from_package(keynote_package(true)).unwrap();
    editor.set_comment(drawable(5), "Only first").unwrap();
    let first = editor.comment(drawable(5)).unwrap().unwrap();
    let second = editor.comment(drawable(6)).unwrap().unwrap();
    assert_ne!(first.storage_object_id, second.storage_object_id);
    assert_eq!(first.comment.text, "Only first");
    assert_eq!(second.comment.text, "Original");
    assert_eq!(first.comment.creation_date_seconds, Some(42.5));
    assert_eq!(first.comment.author_object_id, Some(30));
    assert_eq!(first.comment.reply_object_ids, vec![21]);
    assert_ne!(first.comment.storage_uuid, second.comment.storage_uuid);

    let reparsed = IWorkDrawableCommentEditor::from_bytes(&editor.to_bytes().unwrap()).unwrap();
    assert_eq!(reparsed.comment(drawable(5)).unwrap(), Some(first));
    assert_eq!(reparsed.comment(drawable(6)).unwrap(), Some(second));
}

#[test]
fn create_then_clear_restores_unknown_drawable_bytes_exactly() {
    let mut package = keynote_package(false);
    let original_payload = placeholder_bytes_with_unknown_fields();
    package
        .update_archive("Index/Document.iwa", |archive| {
            let object = archive.object_mut(5).unwrap();
            object.replace_message(
                0,
                RawMessage {
                    type_: 7,
                    data: original_payload.clone(),
                },
            )?;
            object.archive_info.message_infos[0]
                .object_references
                .clear();
            Ok(())
        })
        .unwrap();
    let before = package.to_bytes().unwrap();
    let mut editor = IWorkDrawableCommentEditor::from_package(package).unwrap();

    editor.set_comment(drawable(5), "Temporary").unwrap();
    assert_ne!(object_payload(editor.package(), 5), original_payload);
    editor.clear_comment(drawable(5)).unwrap();

    assert_eq!(object_payload(editor.package(), 5), original_payload);
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn storage_updates_and_copy_on_write_preserve_unknown_fields() {
    let mut package = keynote_package(false);
    let mut unknown = Vec::new();
    append_unknown_varint(&mut unknown, 99, 999);
    package
        .update_archive("Index/Document.iwa", |archive| {
            let storage = archive.object_mut(20).unwrap();
            let mut data = storage.messages[0].data.clone();
            data.extend_from_slice(&unknown);
            storage.replace_message(
                0,
                RawMessage {
                    type_: COMMENT_STORAGE_MESSAGE_TYPE,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = IWorkDrawableCommentEditor::from_package(package).unwrap();
    editor.set_comment(drawable(5), "In place").unwrap();
    assert!(object_payload(editor.package(), 20).ends_with(&unknown));

    let mut package = keynote_package(true);
    package
        .update_archive("Index/Document.iwa", |archive| {
            let storage = archive.object_mut(20).unwrap();
            let mut data = storage.messages[0].data.clone();
            data.extend_from_slice(&unknown);
            storage.replace_message(
                0,
                RawMessage {
                    type_: COMMENT_STORAGE_MESSAGE_TYPE,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = IWorkDrawableCommentEditor::from_package(package).unwrap();
    editor.set_comment(drawable(5), "Copy").unwrap();
    let clone_id = editor
        .comment(drawable(5))
        .unwrap()
        .unwrap()
        .storage_object_id;
    assert_ne!(clone_id.object_id(), 20);
    assert!(object_payload(editor.package(), clone_id.object_id()).ends_with(&unknown));
}

#[test]
fn wire_patcher_rejects_duplicate_and_truncated_fields() {
    let reference = reference(7).encode_to_vec();
    let mut duplicate = Vec::new();
    duplicate.extend([0x32, reference.len() as u8]);
    duplicate.extend_from_slice(&reference);
    duplicate.extend([0x32, reference.len() as u8]);
    duplicate.extend_from_slice(&reference);
    assert!(patch_length_delimited_field(&duplicate, 6, true, None).is_err());
    assert!(parse_wire_fields(&[0x32, 0x05, 0x08]).is_err());
    assert!(parse_wire_fields(&[0x33]).is_err());
}

#[test]
fn clearing_last_user_removes_orphan_reply_graph_but_keeps_author() {
    let mut editor = IWorkDrawableCommentEditor::from_package(keynote_package(true)).unwrap();
    editor.clear_comment(drawable(5)).unwrap();
    assert!(
        object_locations(editor.package())
            .unwrap()
            .contains_key(&20)
    );
    editor.clear_comment(drawable(6)).unwrap();
    let locations = object_locations(editor.package()).unwrap();
    assert!(!locations.contains_key(&20));
    assert!(!locations.contains_key(&21));
    assert!(locations.contains_key(&30));
}

#[test]
fn malformed_storage_fails_transactionally() {
    let mut package = keynote_package(false);
    package
        .update_archive("Index/Document.iwa", |archive| {
            let storage = archive.object_mut(20).unwrap();
            storage.replace_message(
                0,
                RawMessage {
                    type_: COMMENT_STORAGE_MESSAGE_TYPE,
                    data: vec![0x80],
                },
            )?;
            Ok(())
        })
        .unwrap();
    let mut editor = IWorkDrawableCommentEditor::from_package(package).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor.set_comment(drawable(5), "No").is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
    assert!(editor.clear_comment(drawable(5)).is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn every_supported_drawable_nesting_round_trips_comment_reference() {
    let cases = vec![
        (
            Application::Keynote,
            3002,
            DrawablePayload::Drawable(tsd::DrawableArchive::default()),
        ),
        (
            Application::Keynote,
            3004,
            DrawablePayload::Shape(tsd::ShapeArchive::default()),
        ),
        (
            Application::Keynote,
            3005,
            DrawablePayload::Image(tsd::ImageArchive::default()),
        ),
        (
            Application::Keynote,
            3006,
            DrawablePayload::Mask(tsd::MaskArchive::default()),
        ),
        (
            Application::Keynote,
            3007,
            DrawablePayload::Movie(tsd::MovieArchive::default()),
        ),
        (
            Application::Keynote,
            3008,
            DrawablePayload::Group(tsd::GroupArchive::default()),
        ),
        (
            Application::Keynote,
            3009,
            DrawablePayload::ConnectionLine(tsd::ConnectionLineArchive::default()),
        ),
        (
            Application::Keynote,
            2011,
            DrawablePayload::ShapeInfo(tswp::ShapeInfoArchive::default()),
        ),
        (
            Application::Keynote,
            2014,
            DrawablePayload::CommentInfo(tswp::CommentInfoArchive::default()),
        ),
        (
            Application::Pages,
            7,
            DrawablePayload::PagesPlaceholder(tp::PlaceholderArchive::default()),
        ),
        (
            Application::Keynote,
            7,
            DrawablePayload::KeynotePlaceholder(kn::PlaceholderArchive::default()),
        ),
        (
            Application::Numbers,
            7,
            DrawablePayload::NumbersPlaceholder(tn::PlaceholderArchive::default()),
        ),
        (
            Application::Keynote,
            12,
            DrawablePayload::KeynotePlaceholder(kn::PlaceholderArchive::default()),
        ),
        (
            Application::Keynote,
            5021,
            DrawablePayload::Chart(tsch::ChartDrawableArchive {
                super_: Some(tsd::DrawableArchive::default()),
            }),
        ),
        (
            Application::Pages,
            6000,
            DrawablePayload::Table(tst::TableInfoArchive::default()),
        ),
        (
            Application::Pages,
            6007,
            DrawablePayload::WpTable(tst::WpTableInfoArchive::default()),
        ),
    ];

    for (application, message_type, payload) in cases {
        let encoded = payload.encode_to_vec();
        let mut decoded = DrawablePayload::decode(application, message_type, &encoded)
            .unwrap()
            .unwrap();
        assert_eq!(decoded.comment_identifier(), None);
        decoded.set_comment_identifier(Some(123));
        let reparsed = DrawablePayload::decode(application, message_type, &decoded.encode_to_vec())
            .unwrap()
            .unwrap();
        assert_eq!(reparsed.comment_identifier(), Some(123));
    }
}
