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

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::wire::append_varint_field;
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, FieldType, RawMessage, SnappyStream};
use litchi_iwa_protos::{tn, tsce, tsd, tsk, tsp, tst};
use litchi_numbers::cell::comment::CommentReplyIndex;
use litchi_numbers::{
    CellPosition, Package, PackageLimits, PackageReadOptions, PackageSemanticLimits, SheetSelector,
    TableSelector,
};
use litchi_numbers_wire::BncCell;
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const REPLIES_MEMBER: &str = "Index/Replies.iwa";
const VIEW_STATE_MEMBER: &str = "Index/ViewState.iwa";
const METADATA_MEMBER: &str = "Index/Metadata.iwa";

const DOCUMENT_TYPE: u32 = 1;
const SHEET_TYPE: u32 = 2;
const TABLE_INFO_TYPE: u32 = 6_000;
const TABLE_MODEL_TYPE: u32 = 6_001;
const TILE_TYPE: u32 = 6_002;
const TABLE_DATA_LIST_TYPE: u32 = 6_005;
const COMMENT_STORAGE_TYPE: u32 = 3_056;
const METADATA_TYPE: u32 = 11_006;

const DOCUMENT_ID: u64 = 1;
const SHEET_ID: u64 = 2;
const TABLE_INFO_ID: u64 = 3;
const TABLE_MODEL_ID: u64 = 4;
const SIDECAR_ID: u64 = 5;
const TILE_ID: u64 = 6;
const ROOT_COMMENT_ID: u64 = 20;
const SECOND_ROOT_COMMENT_ID: u64 = 21;
const FIRST_REPLY_ID: u64 = 30;
const SECOND_REPLY_ID: u64 = 31;
const THIRD_REPLY_ID: u64 = 32;
const SHARED_REPLY_ID: u64 = 33;
const AUTHOR_ID: u64 = 40;
const AUTHOR_STORAGE_ID: u64 = 41;
const VIEW_STATE_ID: u64 = 300;
const METADATA_OBJECT_ID: u64 = 900;
const WATERMARK: u64 = 1_000;
const ANNOTATION_AUTHOR_TYPE: u32 = 212;
const ANNOTATION_AUTHOR_STORAGE_TYPE: u32 = 213;

const SHEET_NAME: &str = "Reply fixture sheet";
const TABLE_NAME: &str = "Reply fixture table";

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FixtureMode {
    /// An empty comment list plus an empty author registry.
    Rootless,
    /// One rooted comment with duplicate reply text for ordinal selection.
    DuplicateText,
    /// The same root comment is referenced by two cells.
    SharedRoot,
    /// Two roots share one reply archive.
    SharedReply,
    /// One root/reply pair is unshared and can be culled after removal.
    SingleRoot,
    /// The reply archive is moved to a second current component.
    CrossComponent,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Corruption {
    DuplicateReplyReference,
    SelfReplyReference,
    MissingReplyReference,
    NestedReplyReference,
    ExternalReplyReference,
    TypedReplyReference,
    DuplicateRootPayload,
    WrongReplyType,
    MissingReply,
    SegmentedList,
    DuplicateListKey,
    WrongListType,
    ZeroRefcount,
    RefcountUndercount,
    RefcountOvercount,
    AggregateMissing,
    AggregateExtra,
    FieldInfoMissing,
    FieldInfoDuplicate,
    FieldInfoWrong,
    MissingMetadata,
    MissingReplyUuid,
    VersionedReplyUuid,
    DuplicateReplyUuid,
    AmbiguousReplyIdentifier,
    DataOwnerReplyIdentifier,
    RootDataMapReplyIdentifier,
    OpaqueInbound,
    UnknownWire,
    UnknownMetadata,
    Locked,
    MissingAuthor,
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn external_reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        deprecated_is_external: Some(true),
        ..Default::default()
    }
}

fn typed_reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        deprecated_type: Some(1),
        ..Default::default()
    }
}

fn uuid_entry(identifier: u64) -> tsp::ObjectUuidMapEntry {
    tsp::ObjectUuidMapEntry {
        identifier,
        uuid: tsp::Uuid {
            lower: identifier,
            upper: identifier.saturating_add(0x1000),
        },
    }
}

fn object(identifier: u64, type_: u32, data: Vec<u8>) -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage { type_, data }],
    )?)
}

fn set_message_info(object: &mut ArchiveObject, references: &[u64]) -> TestResult {
    let info = object
        .archive_info
        .message_infos
        .first_mut()
        .ok_or_else(|| io::Error::other("synthetic object has no message info"))?;
    info.object_references = references.to_vec();
    Ok(())
}

fn set_field_info(
    object: &mut ArchiveObject,
    field_path: Vec<u32>,
    references: &[u64],
) -> TestResult {
    let info = object
        .archive_info
        .message_infos
        .first_mut()
        .ok_or_else(|| io::Error::other("synthetic object has no message info"))?;
    let mut field = FieldInfo::new(field_path);
    field.r#type = Some(FieldType::ObjectReference);
    field.object_references = references.to_vec();
    info.field_infos.push(field);
    Ok(())
}

fn comment_entry(key: u32, root_identifier: u64, refcount: u32) -> tst::table_data_list::ListEntry {
    tst::table_data_list::ListEntry {
        key,
        refcount,
        comment_storage: Some(reference(root_identifier)),
        ..Default::default()
    }
}

fn string_entry() -> tst::table_data_list::ListEntry {
    tst::table_data_list::ListEntry {
        key: 1,
        refcount: 1,
        string: Some("fixture seed".to_owned()),
        ..Default::default()
    }
}

fn formula_entry() -> tst::table_data_list::ListEntry {
    tst::table_data_list::ListEntry {
        key: 1,
        refcount: 1,
        formula: Some(tsce::FormulaArchive {
            ast_node_array: tsce::AstNodeArrayArchive {
                ast_node: vec![tsce::ast_node_array_archive::AstNodeArchive {
                    ast_node_type: tsce::ast_node_array_archive::AstNodeType::NumberNode as i32,
                    ast_number_node_number: Some(1.0),
                    ..Default::default()
                }],
            },
            ..Default::default()
        }),
        ..Default::default()
    }
}

fn formula_error_entry() -> tst::table_data_list::ListEntry {
    tst::table_data_list::ListEntry {
        key: 1,
        refcount: 1,
        string: Some("#VALUE!".to_owned()),
        ..Default::default()
    }
}

fn format_entry() -> tst::table_data_list::ListEntry {
    tst::table_data_list::ListEntry {
        key: 1,
        refcount: 2,
        format: Some(tsk::FormatStructArchive {
            format_type: Some(260),
            ..Default::default()
        }),
        ..Default::default()
    }
}

fn list_message(
    list_type: tst::table_data_list::ListType,
    entries: Vec<tst::table_data_list::ListEntry>,
    segments: Vec<tsp::Reference>,
) -> RawMessage {
    RawMessage {
        type_: TABLE_DATA_LIST_TYPE,
        data: tst::TableDataList {
            list_type: list_type as i32,
            next_list_id: 10,
            entries,
            segments,
            is_new_for_bnc: Some(true),
        }
        .encode_to_vec(),
    }
}

fn sidecar(mode: FixtureMode, corruption: Option<Corruption>) -> TestResult<ArchiveObject> {
    let mut comment_entries = match mode {
        FixtureMode::Rootless => Vec::new(),
        FixtureMode::SharedRoot | FixtureMode::DuplicateText | FixtureMode::SingleRoot => {
            vec![comment_entry(
                1,
                ROOT_COMMENT_ID,
                if matches!(mode, FixtureMode::SharedRoot) {
                    2
                } else {
                    1
                },
            )]
        },
        FixtureMode::SharedReply | FixtureMode::CrossComponent => vec![
            comment_entry(1, ROOT_COMMENT_ID, 1),
            comment_entry(2, SECOND_ROOT_COMMENT_ID, 1),
        ],
    };
    if !comment_entries.is_empty() && matches!(corruption, Some(Corruption::DuplicateListKey)) {
        comment_entries.push(comment_entries[0].clone());
    }
    if !comment_entries.is_empty() && matches!(corruption, Some(Corruption::ZeroRefcount)) {
        comment_entries[0].refcount = 0;
    }
    let actual_first_key_references: u32 = if matches!(mode, FixtureMode::SharedRoot) {
        2
    } else {
        1
    };
    if !comment_entries.is_empty() && matches!(corruption, Some(Corruption::RefcountUndercount)) {
        comment_entries[0].refcount = actual_first_key_references.saturating_sub(1);
    }
    if !comment_entries.is_empty() && matches!(corruption, Some(Corruption::RefcountOvercount)) {
        comment_entries[0].refcount = actual_first_key_references + 1;
    }
    if !comment_entries.is_empty() && matches!(corruption, Some(Corruption::WrongListType)) {
        comment_entries[0].comment_storage = None;
        comment_entries[0].string = Some("wrong list payload".to_owned());
    }

    let mut messages = vec![
        list_message(
            tst::table_data_list::ListType::String,
            vec![string_entry()],
            Vec::new(),
        ),
        list_message(
            tst::table_data_list::ListType::Formula,
            vec![formula_entry()],
            Vec::new(),
        ),
        list_message(
            tst::table_data_list::ListType::FormulaError,
            vec![formula_error_entry()],
            Vec::new(),
        ),
        list_message(
            tst::table_data_list::ListType::Format,
            vec![format_entry()],
            Vec::new(),
        ),
        list_message(
            tst::table_data_list::ListType::CommentStorage,
            comment_entries,
            if matches!(corruption, Some(Corruption::SegmentedList)) {
                vec![reference(700)]
            } else {
                Vec::new()
            },
        ),
    ];
    if matches!(corruption, Some(Corruption::DuplicateListKey)) {
        messages.push(messages[3].clone());
    }

    let mut result = ArchiveObject::new(SIDECAR_ID, messages)?;
    let list_index = result
        .messages
        .iter()
        .position(|message| {
            tst::TableDataList::decode(message.data.as_slice())
                .map(|list| list.list_type == tst::table_data_list::ListType::CommentStorage as i32)
                .unwrap_or(false)
        })
        .ok_or_else(|| io::Error::other("comment list is missing"))?;
    let info = result
        .archive_info
        .message_infos
        .get_mut(list_index)
        .ok_or_else(|| io::Error::other("comment list info is missing"))?;
    let roots: Vec<u64> = match mode {
        FixtureMode::Rootless => Vec::new(),
        FixtureMode::SharedReply | FixtureMode::CrossComponent => {
            vec![ROOT_COMMENT_ID, SECOND_ROOT_COMMENT_ID]
        },
        _ => vec![ROOT_COMMENT_ID],
    };
    info.object_references = roots.clone();
    for (key, root) in roots.iter().enumerate() {
        let mut field = FieldInfo::new(vec![3, u32::try_from(key + 1)?]);
        field.r#type = Some(FieldType::ObjectReference);
        field.object_references = vec![*root];
        info.field_infos.push(field);
    }
    match corruption {
        Some(Corruption::AggregateMissing) => info.object_references.clear(),
        Some(Corruption::AggregateExtra) => info.object_references.push(999_999),
        Some(Corruption::FieldInfoMissing) => info.field_infos.clear(),
        Some(Corruption::FieldInfoDuplicate) => {
            let field = info
                .field_infos
                .first()
                .cloned()
                .ok_or_else(|| io::Error::other("field info is missing"))?;
            info.field_infos.push(field);
        },
        Some(Corruption::FieldInfoWrong) => {
            let field = info
                .field_infos
                .first_mut()
                .ok_or_else(|| io::Error::other("field info is missing"))?;
            field.object_references = vec![999_999];
        },
        _ => {},
    }
    Ok(result)
}

fn table_model() -> tst::TableModelArchive {
    tst::TableModelArchive {
        table_id: "reply-fixture-table-id".to_owned(),
        table_name: TABLE_NAME.to_owned(),
        table_style: reference(SIDECAR_ID),
        body_text_style: reference(SIDECAR_ID),
        header_row_text_style: reference(SIDECAR_ID),
        header_column_text_style: reference(SIDECAR_ID),
        footer_row_text_style: reference(SIDECAR_ID),
        body_cell_style: reference(SIDECAR_ID),
        header_row_style: reference(SIDECAR_ID),
        header_column_style: reference(SIDECAR_ID),
        footer_row_style: reference(SIDECAR_ID),
        number_of_rows: 2,
        number_of_columns: 1,
        base_data_store: tst::DataStore {
            row_headers: tst::HeaderStorage {
                bucket_hash_function: 1,
                ..Default::default()
            },
            column_headers: reference(SIDECAR_ID),
            tiles: tst::TileStorage {
                tiles: vec![tst::tile_storage::Tile {
                    tileid: 0,
                    tile: reference(TILE_ID),
                }],
                tile_size: Some(256),
                ..Default::default()
            },
            string_table: reference(SIDECAR_ID),
            style_table: reference(SIDECAR_ID),
            formula_table: reference(SIDECAR_ID),
            formula_error_table: Some(reference(SIDECAR_ID)),
            format_table_pre_bnc: reference(SIDECAR_ID),
            format_table: Some(reference(SIDECAR_ID)),
            comment_storage_table: Some(reference(SIDECAR_ID)),
            next_row_strip_id: 1,
            next_column_strip_id: 1,
            row_tile_tree: tst::TableRbTree::default(),
            column_tile_tree: tst::TableRbTree::default(),
            ..Default::default()
        },
        ..Default::default()
    }
}

fn comment_cell(key: Option<u32>) -> TestResult<Vec<u8>> {
    let mut cell = BncCell::minimal();
    cell.set_number(42.0)?;
    cell.set_comment_identifier(key);
    Ok(cell.encode())
}

fn tile(mode: FixtureMode) -> TestResult<tst::Tile> {
    let keys = match mode {
        FixtureMode::Rootless => [None, None],
        FixtureMode::SharedRoot => [Some(1), Some(1)],
        FixtureMode::SharedReply | FixtureMode::CrossComponent => [Some(1), Some(2)],
        FixtureMode::DuplicateText | FixtureMode::SingleRoot => [Some(1), None],
    };
    let mut rows = Vec::new();
    for (row, key) in keys.into_iter().enumerate() {
        let bytes = comment_cell(key)?;
        rows.push(tst::TileRowInfo {
            tile_row_index: u32::try_from(row)?,
            cell_count: 1,
            storage_version: Some(5),
            cell_storage_buffer_pre_bnc: bytes.clone(),
            cell_offsets_pre_bnc: vec![0, 0],
            cell_storage_buffer: Some(bytes),
            cell_offsets: Some(vec![0, 0]),
            ..Default::default()
        });
    }
    Ok(tst::Tile {
        max_column: 0,
        max_row: 1,
        num_cells: 2,
        numrows: 2,
        row_infos: rows,
        storage_version: Some(5),
        last_saved_in_bnc: Some(true),
        ..Default::default()
    })
}

fn comment_archive(
    identifier: u64,
    text: &str,
    replies: &[tsp::Reference],
) -> TestResult<ArchiveObject> {
    let mut result = object(
        identifier,
        COMMENT_STORAGE_TYPE,
        tsd::CommentStorageArchive {
            text: Some(text.to_owned()),
            creation_date: Some(tsp::Date { seconds: 123.0 }),
            author: Some(reference(AUTHOR_ID)),
            replies: replies.to_vec(),
            storage_uuid: Some(tsp::Uuid {
                lower: identifier,
                upper: identifier.rotate_left(13),
            }),
        }
        .encode_to_vec(),
    )?;
    let reply_ids = replies
        .iter()
        .map(|reply| reply.identifier)
        .collect::<Vec<_>>();
    let mut references = vec![AUTHOR_ID];
    references.extend(reply_ids);
    set_message_info(&mut result, &references)?;
    set_field_info(&mut result, vec![3], &[AUTHOR_ID])?;
    for (ordinal, reply) in replies.iter().enumerate() {
        set_field_info(
            &mut result,
            vec![4, u32::try_from(ordinal)?],
            &[reply.identifier],
        )?;
    }
    Ok(result)
}

fn reply_objects(
    mode: FixtureMode,
    corruption: Option<Corruption>,
) -> TestResult<Vec<ArchiveObject>> {
    if matches!(mode, FixtureMode::Rootless) {
        return Ok(Vec::new());
    }
    let mut first_replies = match mode {
        FixtureMode::DuplicateText => vec![
            reference(FIRST_REPLY_ID),
            reference(SECOND_REPLY_ID),
            reference(THIRD_REPLY_ID),
        ],
        FixtureMode::SharedReply | FixtureMode::CrossComponent => {
            vec![reference(SHARED_REPLY_ID), reference(SECOND_REPLY_ID)]
        },
        _ => vec![reference(FIRST_REPLY_ID)],
    };
    if let Some(kind) = corruption {
        first_replies = match kind {
            Corruption::DuplicateReplyReference => {
                vec![reference(FIRST_REPLY_ID), reference(FIRST_REPLY_ID)]
            },
            Corruption::SelfReplyReference => vec![reference(ROOT_COMMENT_ID)],
            Corruption::MissingReplyReference => vec![reference(999_999)],
            Corruption::NestedReplyReference => vec![reference(SECOND_REPLY_ID)],
            Corruption::ExternalReplyReference => vec![external_reference(FIRST_REPLY_ID)],
            Corruption::TypedReplyReference => vec![typed_reference(FIRST_REPLY_ID)],
            _ => first_replies,
        };
    }
    let mut result = vec![comment_archive(
        ROOT_COMMENT_ID,
        "root comment",
        &first_replies,
    )?];
    if matches!(mode, FixtureMode::SharedReply | FixtureMode::CrossComponent) {
        result.push(comment_archive(
            SECOND_ROOT_COMMENT_ID,
            "second root",
            &[reference(SHARED_REPLY_ID)],
        )?);
    }
    let first_text = if matches!(mode, FixtureMode::DuplicateText) {
        "duplicate"
    } else {
        "first reply"
    };
    if !matches!(mode, FixtureMode::SharedReply | FixtureMode::CrossComponent) {
        result.push(comment_archive(FIRST_REPLY_ID, first_text, &[])?);
    }
    if matches!(mode, FixtureMode::DuplicateText) {
        result.push(comment_archive(SECOND_REPLY_ID, "duplicate", &[])?);
        result.push(comment_archive(THIRD_REPLY_ID, "duplicate", &[])?);
    }
    if matches!(mode, FixtureMode::SharedReply | FixtureMode::CrossComponent) {
        result.push(comment_archive(SECOND_REPLY_ID, "second reply", &[])?);
        result.push(comment_archive(SHARED_REPLY_ID, "shared reply", &[])?);
    }
    if matches!(corruption, Some(Corruption::MissingReply)) {
        result.retain(|object| object.archive_info.identifier != Some(FIRST_REPLY_ID));
    }
    if matches!(corruption, Some(Corruption::WrongReplyType)) {
        if let Some(object) = result
            .iter_mut()
            .find(|object| object.archive_info.identifier == Some(FIRST_REPLY_ID))
        {
            object.messages[0].type_ = COMMENT_STORAGE_TYPE + 1;
        }
    }
    if matches!(corruption, Some(Corruption::DuplicateRootPayload)) {
        let root = result
            .first_mut()
            .ok_or_else(|| io::Error::other("root comment is missing"))?;
        let message = root
            .messages
            .first()
            .cloned()
            .ok_or_else(|| io::Error::other("root payload is missing"))?;
        root.messages.push(message);
        root.archive_info
            .message_infos
            .push(root.archive_info.message_infos[0].clone());
    }
    Ok(result)
}

fn metadata_component(
    identifier: u64,
    locator: &str,
    save_token: u64,
    object_ids: &[u64],
) -> tsp::ComponentInfo {
    tsp::ComponentInfo {
        identifier,
        preferred_locator: locator.to_owned(),
        locator: Some(locator.to_owned()),
        save_token: Some(save_token),
        object_uuid_map_entries: object_ids.iter().copied().map(uuid_entry).collect(),
        ..Default::default()
    }
}

fn metadata(mode: FixtureMode, corruption: Option<Corruption>) -> TestResult<Vec<u8>> {
    let mut document_ids = vec![
        DOCUMENT_ID,
        SHEET_ID,
        TABLE_INFO_ID,
        TABLE_MODEL_ID,
        SIDECAR_ID,
        TILE_ID,
        ROOT_COMMENT_ID,
        SECOND_ROOT_COMMENT_ID,
        FIRST_REPLY_ID,
        SECOND_REPLY_ID,
        THIRD_REPLY_ID,
        SHARED_REPLY_ID,
        AUTHOR_ID,
        AUTHOR_STORAGE_ID,
    ];
    if matches!(mode, FixtureMode::Rootless) {
        document_ids.retain(|identifier| {
            !matches!(
                *identifier,
                ROOT_COMMENT_ID
                    | SECOND_ROOT_COMMENT_ID
                    | FIRST_REPLY_ID
                    | SECOND_REPLY_ID
                    | THIRD_REPLY_ID
                    | SHARED_REPLY_ID
                    | AUTHOR_ID
            )
        });
    }
    document_ids.retain(|identifier| {
        *identifier != SECOND_ROOT_COMMENT_ID
            || matches!(mode, FixtureMode::SharedReply | FixtureMode::CrossComponent)
    });
    if matches!(mode, FixtureMode::CrossComponent) {
        document_ids.retain(|identifier| {
            !matches!(
                *identifier,
                FIRST_REPLY_ID | SECOND_REPLY_ID | THIRD_REPLY_ID | SHARED_REPLY_ID
            )
        });
    }
    let reply_ids = if matches!(mode, FixtureMode::CrossComponent) {
        vec![SECOND_REPLY_ID, SHARED_REPLY_ID]
    } else {
        Vec::new()
    };
    let document = metadata_component(100, "Document", 9, &document_ids);
    let view = metadata_component(300, "ViewState", 7, &[VIEW_STATE_ID]);
    let replies = metadata_component(400, "Replies", 8, &reply_ids);
    let versioned = metadata_component(101, "Document", 3, &[999]);
    let mut package = tsp::PackageMetadata {
        last_object_identifier: WATERMARK,
        save_token: Some(10),
        components: vec![document, view],
        versioned_components: vec![versioned],
        ..Default::default()
    };
    if matches!(mode, FixtureMode::CrossComponent) {
        package.components.push(replies);
    }
    match corruption {
        Some(Corruption::MissingReplyUuid) => {
            if let Some(component) = package
                .components
                .iter_mut()
                .find(|component| component.identifier == 100)
            {
                component
                    .object_uuid_map_entries
                    .retain(|entry| entry.identifier != FIRST_REPLY_ID);
            }
        },
        Some(Corruption::VersionedReplyUuid) => {
            if let Some(component) = package
                .components
                .iter_mut()
                .find(|component| component.identifier == 100)
            {
                component
                    .object_uuid_map_entries
                    .retain(|entry| entry.identifier != FIRST_REPLY_ID);
            }
            package.versioned_components[0]
                .object_uuid_map_entries
                .push(uuid_entry(FIRST_REPLY_ID));
        },
        Some(Corruption::DuplicateReplyUuid) => {
            if let Some(component) = package
                .components
                .iter_mut()
                .find(|component| component.identifier == 100)
            {
                component
                    .object_uuid_map_entries
                    .push(uuid_entry(FIRST_REPLY_ID));
            }
        },
        Some(Corruption::AmbiguousReplyIdentifier) => {
            if let Some(component) = package
                .components
                .iter_mut()
                .find(|component| component.identifier == 100)
            {
                component.ambiguous_object_identifiers.push(FIRST_REPLY_ID);
            }
        },
        Some(Corruption::DataOwnerReplyIdentifier) => {
            if let Some(component) = package
                .components
                .iter_mut()
                .find(|component| component.identifier == 100)
            {
                component.data_references.push(tsp::ComponentDataReference {
                    data_identifier: 2_000,
                    object_reference_list: vec![tsp::component_data_reference::ObjectReference {
                        object_identifier: FIRST_REPLY_ID,
                        count: 1,
                    }],
                });
            }
        },
        Some(Corruption::RootDataMapReplyIdentifier) => {
            package.data_metadata_map = Some(reference(FIRST_REPLY_ID));
        },
        _ => {},
    }
    let mut data = package.encode_to_vec();
    if matches!(corruption, Some(Corruption::UnknownMetadata)) {
        append_varint_field(&mut data, 90, 0xfeed_beef)?;
    }
    Ok(data)
}

fn compressed(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

fn author_objects(empty: bool) -> TestResult<Vec<ArchiveObject>> {
    if empty {
        let mut storage = object(
            AUTHOR_STORAGE_ID,
            ANNOTATION_AUTHOR_STORAGE_TYPE,
            tsk::AnnotationAuthorStorageArchive::default().encode_to_vec(),
        )?;
        set_message_info(&mut storage, &[])?;
        return Ok(vec![storage]);
    }
    let author = object(
        AUTHOR_ID,
        ANNOTATION_AUTHOR_TYPE,
        tsk::AnnotationAuthorArchive {
            name: Some("Reply fixture author".to_owned()),
            public_id: Some("reply-fixture-author".to_owned()),
            public_ids: vec!["reply-fixture-author".to_owned()],
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    let mut storage = object(
        AUTHOR_STORAGE_ID,
        ANNOTATION_AUTHOR_STORAGE_TYPE,
        tsk::AnnotationAuthorStorageArchive {
            annotation_author: vec![reference(AUTHOR_ID)],
        }
        .encode_to_vec(),
    )?;
    set_message_info(&mut storage, &[AUTHOR_ID])?;
    Ok(vec![author, storage])
}

fn fixture(mode: FixtureMode, corruption: Option<Corruption>) -> TestResult<Vec<u8>> {
    let mut document = object(
        DOCUMENT_ID,
        DOCUMENT_TYPE,
        tn::DocumentArchive {
            sheets: vec![reference(SHEET_ID)],
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    set_message_info(&mut document, &[SHEET_ID])?;
    set_field_info(&mut document, vec![1], &[SHEET_ID])?;

    let mut sheet = object(
        SHEET_ID,
        SHEET_TYPE,
        tn::SheetArchive {
            name: SHEET_NAME.to_owned(),
            drawable_infos: vec![reference(TABLE_INFO_ID)],
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    set_message_info(&mut sheet, &[TABLE_INFO_ID])?;
    set_field_info(&mut sheet, vec![4], &[TABLE_INFO_ID])?;

    let mut info = object(
        TABLE_INFO_ID,
        TABLE_INFO_TYPE,
        tst::TableInfoArchive {
            super_: tsd::DrawableArchive {
                locked: matches!(corruption, Some(Corruption::Locked)).then_some(true),
                ..Default::default()
            },
            table_model: reference(TABLE_MODEL_ID),
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    set_message_info(&mut info, &[TABLE_MODEL_ID])?;
    set_field_info(&mut info, vec![4], &[TABLE_MODEL_ID])?;

    let mut model = object(
        TABLE_MODEL_ID,
        TABLE_MODEL_TYPE,
        table_model().encode_to_vec(),
    )?;
    set_message_info(&mut model, &[SIDECAR_ID, TILE_ID])?;
    set_field_info(&mut model, vec![25], &[SIDECAR_ID])?;
    set_field_info(&mut model, vec![26], &[TILE_ID])?;

    let mut tile = object(TILE_ID, TILE_TYPE, tile(mode)?.encode_to_vec())?;
    set_message_info(&mut tile, &[])?;
    let sidecar = sidecar(mode, corruption)?;
    let mut document_objects = vec![document, sheet, info, model, sidecar, tile];
    if !matches!(corruption, Some(Corruption::MissingAuthor)) {
        document_objects.extend(author_objects(matches!(
            (mode, corruption),
            (FixtureMode::Rootless, _)
        ))?);
    }
    let all_comments = reply_objects(mode, corruption)?;
    let mut split_replies = Vec::new();
    if matches!(mode, FixtureMode::CrossComponent) {
        for comment in all_comments {
            let identifier = comment.archive_info.identifier.unwrap_or_default();
            if matches!(identifier, ROOT_COMMENT_ID | SECOND_ROOT_COMMENT_ID) {
                document_objects.push(comment);
            } else {
                split_replies.push(comment);
            }
        }
    } else {
        document_objects.extend(all_comments);
    }

    let mut entries = vec![
        (DOCUMENT_MEMBER, compressed(document_objects)?),
        (
            VIEW_STATE_MEMBER,
            compressed(vec![object(
                VIEW_STATE_ID,
                210,
                b"unselected view state".to_vec(),
            )?])?,
        ),
        (
            METADATA_MEMBER,
            compressed(vec![object(
                METADATA_OBJECT_ID,
                METADATA_TYPE,
                metadata(mode, corruption)?,
            )?])?,
        ),
        ("preview.jpg", b"reply preview".to_vec()),
        ("preview-micro.jpg", b"reply micro".to_vec()),
        ("preview-web.jpg", b"reply web".to_vec()),
        ("Data/sentinel.bin", b"unselected data".to_vec()),
    ];
    if matches!(mode, FixtureMode::CrossComponent) {
        entries.insert(1, (REPLIES_MEMBER, compressed(split_replies)?));
    }
    if matches!(corruption, Some(Corruption::OpaqueInbound)) {
        let mut inbound = object(700, 99_999, b"opaque inbound".to_vec())?;
        set_message_info(&mut inbound, &[FIRST_REPLY_ID])?;
        let mut entries_data = entries
            .iter()
            .map(|(name, data)| (*name, data.clone()))
            .collect::<Vec<_>>();
        entries_data.push(("Index/OpaqueInbound.iwa", compressed(vec![inbound])?));
        entries = entries_data;
    }
    if matches!(corruption, Some(Corruption::MissingMetadata)) {
        entries.retain(|(name, _)| *name != METADATA_MEMBER);
    }
    Ok(litchi_iwa_archive::package::to_bytes(
        entries.iter().map(|(name, data)| (*name, data.as_slice())),
        Limits::default(),
    )?)
}

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
        Some("new root with generated author"),
    );
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
