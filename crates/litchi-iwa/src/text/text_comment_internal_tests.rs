use crate::archive::{ArchiveObject, RawMessage};
use crate::pages::PagesEditor;
use crate::shapes::{DrawablePoint, DrawableSize};
use crate::text::highlight_object::validate_annotation_graph;
use crate::text::highlight_storage::{HIGHLIGHT_TABLE_FIELD, TABLE_ENTRIES_FIELD, locate_storage};
use crate::text::storage_wire::STORAGE_MESSAGE_TYPES;
use crate::wire::{
    parse_wire_fields, patch_varint_field, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields, transform_length_delimited_field,
};
use litchi_iwa_text::comment::raw::{comment_id_value, reply_id_value};

use super::*;

const POSITION: DrawablePoint = DrawablePoint { x: 40.0, y: 80.0 };
const SIZE: DrawableSize = DrawableSize {
    width: 320.0,
    height: 140.0,
};

fn fixture() -> (super::super::IWorkTextEditor, u64, Vec<u8>) {
    let mut pages = PagesEditor::create_with_text("Body").unwrap();
    let text_box = pages.add_text_box(4, "Alpha Beta", POSITION, SIZE).unwrap();
    let baseline = pages.to_bytes().unwrap();
    (
        super::super::IWorkTextEditor::from_package(pages.into_package()),
        text_box.storage.object_id,
        baseline,
    )
}

fn range(start: usize, end: usize) -> TextRange {
    TextRange::from_utf16_indexes(start, end).unwrap()
}

fn body(value: &str) -> TextCommentBody {
    TextCommentBody::new(value).unwrap()
}

fn reply_body(value: &str) -> TextCommentReplyBody {
    TextCommentReplyBody::new(value).unwrap()
}

fn storage_message_index(object: &ArchiveObject) -> usize {
    object
        .messages
        .iter()
        .position(|message| STORAGE_MESSAGE_TYPES.contains(&message.type_))
        .unwrap()
}

#[test]
fn unknown_annotation_and_comment_fields_survive_updates() {
    let (mut editor, storage_id, _) = fixture();
    let comment = editor
        .add_text_comment(storage_id, range(0, 5), body("one"))
        .unwrap();
    let mut package = editor.into_package();
    let archive_name = locate_storage(&package, storage_id).unwrap().archive_name;
    let graph = {
        let archive = package.archive(&archive_name).unwrap();
        let annotation = archive.object(comment_id_value(comment.id())).unwrap();
        validate_annotation_graph(
            &package,
            &archive_name,
            comment_id_value(comment.id()),
            annotation,
        )
        .unwrap()
        .unwrap()
    };
    let comment_storage_id = graph.comment_storage_id;
    package
        .update_archive(&archive_name, |archive| {
            let storage = archive.object_mut(storage_id).unwrap();
            let index = storage_message_index(storage);
            let original = &storage.messages[index];
            let data =
                transform_length_delimited_field(&original.data, HIGHLIGHT_TABLE_FIELD, |table| {
                    let entries = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD)?;
                    let first = patch_varint_field(entries[0], 77, false, Some(42))?;
                    let table = rewrite_repeated_length_delimited_fields(
                        table,
                        TABLE_ENTRIES_FIELD,
                        &[first, entries[1].to_vec()],
                    )?;
                    patch_varint_field(&table, 99, false, Some(7))
                })?;
            storage.replace_message(
                index,
                RawMessage {
                    type_: original.type_,
                    data,
                },
            )?;

            for (object_id, field, value) in [
                (comment_id_value(comment.id()), 88, 9),
                (comment_storage_id, 89, 10),
            ] {
                let object = archive.object_mut(object_id).unwrap();
                let original = &object.messages[0];
                let data = patch_varint_field(&original.data, field, false, Some(value))?;
                object.replace_message(
                    0,
                    RawMessage {
                        type_: original.type_,
                        data,
                    },
                )?;
            }
            Ok(())
        })
        .unwrap();

    let mut editor = super::super::IWorkTextEditor::from_package(package);
    editor
        .update_text_comment(storage_id, comment.id(), range(6, 10), body("two"))
        .unwrap();
    let package = editor.into_package();
    let archive = package.archive(&archive_name).unwrap();
    let storage = archive.object(storage_id).unwrap();
    let message = &storage.messages[storage_message_index(storage)];
    let table =
        repeated_length_delimited_payloads(&message.data, HIGHLIGHT_TABLE_FIELD).unwrap()[0];
    assert!(
        parse_wire_fields(table)
            .unwrap()
            .iter()
            .any(|field| field.number() == 99)
    );
    assert!(
        repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD)
            .unwrap()
            .iter()
            .any(|entry| parse_wire_fields(entry)
                .unwrap()
                .iter()
                .any(|field| field.number() == 77))
    );
    for (object_id, field) in [
        (comment_id_value(comment.id()), 88),
        (comment_storage_id, 89),
    ] {
        assert!(
            parse_wire_fields(&archive.object(object_id).unwrap().messages[0].data)
                .unwrap()
                .iter()
                .any(|wire| wire.number() == field)
        );
    }
}

#[test]
fn replies_survive_root_updates_and_are_reclaimed_on_delete() {
    let (mut editor, storage_id, baseline) = fixture();
    let comment = editor
        .add_text_comment(storage_id, range(0, 5), body("root"))
        .unwrap();
    let reply = editor
        .add_text_comment_reply(storage_id, comment.id(), reply_body("reply"))
        .unwrap();

    assert_eq!(
        editor.text_comments(storage_id).unwrap()[0].reply_count(),
        1
    );
    assert_eq!(
        editor
            .text_comment_replies(storage_id, comment.id())
            .unwrap()[0],
        reply
    );
    let updated = editor
        .update_text_comment(storage_id, comment.id(), range(6, 10), body("updated"))
        .unwrap();
    assert_eq!(updated.reply_count(), 1);
    editor
        .remove_text_comment(storage_id, comment.id())
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
}

#[test]
fn reply_order_survives_middle_update_and_delete() {
    let (mut editor, storage_id, baseline) = fixture();
    let comment = editor
        .add_text_comment(storage_id, range(0, 5), body("root"))
        .unwrap();
    let first = editor
        .add_text_comment_reply(storage_id, comment.id(), reply_body("first"))
        .unwrap();
    let second = editor
        .add_text_comment_reply(storage_id, comment.id(), reply_body("second"))
        .unwrap();
    let third = editor
        .add_text_comment_reply(storage_id, comment.id(), reply_body("third"))
        .unwrap();
    assert_eq!(
        editor
            .text_comment_replies(storage_id, comment.id())
            .unwrap()
            .iter()
            .map(|reply| reply.id())
            .collect::<Vec<_>>(),
        vec![first.id(), second.id(), third.id()]
    );

    let updated = editor
        .update_text_comment_reply(
            storage_id,
            comment.id(),
            second.id(),
            reply_body("second updated"),
        )
        .unwrap();
    assert_eq!(updated.id(), second.id());
    editor
        .remove_text_comment_reply(storage_id, comment.id(), second.id())
        .unwrap();
    assert_eq!(
        editor
            .text_comment_replies(storage_id, comment.id())
            .unwrap()
            .iter()
            .map(|reply| reply.id())
            .collect::<Vec<_>>(),
        vec![first.id(), third.id()]
    );

    editor
        .remove_text_comment(storage_id, comment.id())
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
}

#[test]
fn reply_updates_preserve_unknown_fields_and_identity() {
    let (mut editor, storage_id, baseline) = fixture();
    let comment = editor
        .add_text_comment(storage_id, range(0, 5), body("root"))
        .unwrap();
    let reply = editor
        .add_text_comment_reply(storage_id, comment.id(), reply_body("reply"))
        .unwrap();
    let mut package = editor.into_package();
    let archive_name = locate_storage(&package, storage_id).unwrap().archive_name;
    package
        .update_archive(&archive_name, |archive| {
            let object = archive.object_mut(reply_id_value(reply.id())).unwrap();
            let original = &object.messages[0];
            let data = patch_varint_field(&original.data, 97, false, Some(31))?;
            object.replace_message(
                0,
                RawMessage {
                    type_: original.type_,
                    data,
                },
            )?;
            Ok(())
        })
        .unwrap();

    let mut editor = super::super::IWorkTextEditor::from_package(package);
    let updated = editor
        .update_text_comment_reply(storage_id, comment.id(), reply.id(), reply_body("updated"))
        .unwrap();
    assert_eq!(updated.id(), reply.id());
    assert_eq!(updated.creation_date(), reply.creation_date());
    assert_eq!(updated.author(), reply.author());
    assert_eq!(updated.storage_uuid(), reply.storage_uuid());
    let package = editor.package();
    assert!(
        parse_wire_fields(
            &package
                .archive(&archive_name)
                .unwrap()
                .object(reply_id_value(reply.id()))
                .unwrap()
                .messages[0]
                .data,
        )
        .unwrap()
        .iter()
        .any(|field| field.number() == 97)
    );

    editor
        .remove_text_comment_reply(storage_id, comment.id(), reply.id())
        .unwrap();
    editor
        .remove_text_comment(storage_id, comment.id())
        .unwrap();
    assert_eq!(editor.to_bytes().unwrap(), baseline);
}

#[test]
fn shared_or_wrong_parent_replies_are_rejected_transactionally() {
    let (mut editor, storage_id, _) = fixture();
    let first = editor
        .add_text_comment(storage_id, range(0, 5), body("first"))
        .unwrap();
    let second = editor
        .add_text_comment(storage_id, range(6, 10), body("second"))
        .unwrap();
    let reply = editor
        .add_text_comment_reply(storage_id, first.id(), reply_body("reply"))
        .unwrap();

    let before_wrong_parent = editor.to_bytes().unwrap();
    assert!(
        editor
            .update_text_comment_reply(
                storage_id,
                second.id(),
                reply.id(),
                reply_body("wrong parent"),
            )
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before_wrong_parent);

    let mut package = editor.into_package();
    let archive_name = locate_storage(&package, storage_id).unwrap().archive_name;
    package
        .update_archive(&archive_name, |archive| {
            let other = archive
                .objects
                .iter_mut()
                .find(|object| {
                    object.archive_info.identifier != Some(reply_id_value(reply.id()))
                        && !object.archive_info.message_infos.is_empty()
                })
                .unwrap();
            other.archive_info.message_infos[0]
                .object_references
                .push(reply_id_value(reply.id()));
            Ok(())
        })
        .unwrap();
    let mut editor = super::super::IWorkTextEditor::from_package(package);
    let before_shared = editor.to_bytes().unwrap();
    assert!(
        editor
            .update_text_comment_reply(storage_id, first.id(), reply.id(), reply_body("shared"),)
            .is_err()
    );
    assert!(
        editor
            .remove_text_comment_reply(storage_id, first.id(), reply.id())
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before_shared);
}

#[test]
fn shared_comment_storage_rejects_update_and_delete_transactionally() {
    let (mut editor, storage_id, _) = fixture();
    let comment = editor
        .add_text_comment(storage_id, range(0, 5), body("root"))
        .unwrap();
    let mut package = editor.into_package();
    let archive_name = locate_storage(&package, storage_id).unwrap().archive_name;
    let graph = {
        let archive = package.archive(&archive_name).unwrap();
        let annotation = archive.object(comment_id_value(comment.id())).unwrap();
        validate_annotation_graph(
            &package,
            &archive_name,
            comment_id_value(comment.id()),
            annotation,
        )
        .unwrap()
        .unwrap()
    };
    package
        .update_archive(&archive_name, |archive| {
            let other = archive
                .objects
                .iter_mut()
                .find(|object| {
                    object.archive_info.identifier != Some(comment_id_value(comment.id()))
                        && !object.archive_info.message_infos.is_empty()
                })
                .unwrap();
            other.archive_info.message_infos[0]
                .object_references
                .push(graph.comment_storage_id);
            Ok(())
        })
        .unwrap();
    let mut editor = super::super::IWorkTextEditor::from_package(package);
    let before = editor.to_bytes().unwrap();
    assert!(
        editor
            .update_text_comment(storage_id, comment.id(), range(6, 10), body("changed"))
            .is_err()
    );
    assert!(
        editor
            .remove_text_comment(storage_id, comment.id())
            .is_err()
    );
    assert_eq!(editor.to_bytes().unwrap(), before);
}
