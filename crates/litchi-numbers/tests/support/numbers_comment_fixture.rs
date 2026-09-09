//! Shared bounded Numbers comment graph fixtures for focused integration tests.
//!
//! The builder is kept in one test support module so the package and migration
//! host differential suites exercise the same native graph. It uses generated
//! protobuf values only to author test bytes; production readers remain the
//! source-authoritative lazy path.

use std::io;

use litchi_iwa_archive::{
    Limits,
    iwa::{Archive, ArchiveObject, FieldInfo, FieldType, RawMessage, SnappyStream},
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::wire::append_varint_field;
use litchi_iwa_protos::{tn, tsce, tsd, tsk, tsp, tst};
use litchi_numbers_wire::BncCell;
use prost::Message as _;

pub(crate) type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

pub(crate) const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
pub(crate) const REPLIES_MEMBER: &str = "Index/Replies.iwa";
pub(crate) const VIEW_STATE_MEMBER: &str = "Index/ViewState.iwa";
pub(crate) const METADATA_MEMBER: &str = "Index/Metadata.iwa";

pub(crate) const DOCUMENT_TYPE: u32 = 1;
pub(crate) const SHEET_TYPE: u32 = 2;
pub(crate) const TABLE_INFO_TYPE: u32 = 6_000;
pub(crate) const TABLE_MODEL_TYPE: u32 = 6_001;
pub(crate) const TILE_TYPE: u32 = 6_002;
pub(crate) const TABLE_DATA_LIST_TYPE: u32 = 6_005;
pub(crate) const COMMENT_STORAGE_TYPE: u32 = 3_056;
pub(crate) const METADATA_TYPE: u32 = 11_006;

pub(crate) const DOCUMENT_ID: u64 = 1;
pub(crate) const SHEET_ID: u64 = 2;
pub(crate) const TABLE_INFO_ID: u64 = 3;
pub(crate) const TABLE_MODEL_ID: u64 = 4;
pub(crate) const SIDECAR_ID: u64 = 5;
pub(crate) const TILE_ID: u64 = 6;
pub(crate) const ROOT_COMMENT_ID: u64 = 20;
pub(crate) const SECOND_ROOT_COMMENT_ID: u64 = 21;
pub(crate) const SEGMENT_ID: u64 = 700;
#[allow(
    dead_code,
    reason = "used by the focused cross-crate provenance fixture"
)]
pub(crate) const UNRELATED_SEGMENT_ID: u64 = 701;
pub(crate) const FIRST_REPLY_ID: u64 = 30;
pub(crate) const SECOND_REPLY_ID: u64 = 31;
pub(crate) const THIRD_REPLY_ID: u64 = 32;
pub(crate) const SHARED_REPLY_ID: u64 = 33;
pub(crate) const AUTHOR_ID: u64 = 40;
pub(crate) const AUTHOR_STORAGE_ID: u64 = 41;
pub(crate) const VIEW_STATE_ID: u64 = 300;
pub(crate) const METADATA_OBJECT_ID: u64 = 900;
pub(crate) const WATERMARK: u64 = 1_000;
pub(crate) const ANNOTATION_AUTHOR_TYPE: u32 = 212;
pub(crate) const ANNOTATION_AUTHOR_STORAGE_TYPE: u32 = 213;

pub(crate) const SHEET_NAME: &str = "Reply fixture sheet";
pub(crate) const TABLE_NAME: &str = "Reply fixture table";
#[allow(
    dead_code,
    reason = "used by the focused cross-crate multitable fixture"
)]
pub(crate) const SECOND_TABLE_NAME: &str = "Reply fixture table 2";
#[allow(
    dead_code,
    reason = "used by the focused cross-crate multitable fixture"
)]
pub(crate) const SECOND_TABLE_INFO_ID: u64 = 7;
#[allow(
    dead_code,
    reason = "used by the focused cross-crate multitable fixture"
)]
pub(crate) const SECOND_TABLE_MODEL_ID: u64 = 8;
#[allow(
    dead_code,
    reason = "used by the focused cross-crate multitable fixture"
)]
pub(crate) const SECOND_SIDECAR_ID: u64 = 9;
#[allow(
    dead_code,
    reason = "used by the focused cross-crate multitable fixture"
)]
pub(crate) const SECOND_TILE_ID: u64 = 10;

#[allow(
    dead_code,
    reason = "the shared fixture spans focused read cases and the broader host migration suite"
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum FixtureMode {
    /// An empty comment list plus an empty author registry.
    Rootless,
    /// An empty comment graph with a missing second-column slot.
    SparseCell,
    /// One rooted comment with duplicate reply text for ordinal selection.
    DuplicateText,
    /// The same root comment is referenced by two cells.
    SharedRoot,
    /// Two roots share one reply archive.
    SharedReply,
    /// One root/reply pair is unshared and can be culled after removal.
    SingleRoot,
    /// The selected root entry lives in a valid TableDataListSegment.
    SegmentedRoot,
    /// The reply archive is moved to a second current component.
    CrossComponent,
}

#[allow(
    dead_code,
    reason = "the shared fixture spans focused read cases and the broader host migration suite"
)]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum Corruption {
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
    MissingAuthorStorage,
    MissingCommentList,
}

pub(crate) fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

pub(crate) fn external_reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        deprecated_is_external: Some(true),
        ..Default::default()
    }
}

pub(crate) fn typed_reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        deprecated_type: Some(1),
        ..Default::default()
    }
}

pub(crate) fn uuid_entry(identifier: u64) -> tsp::ObjectUuidMapEntry {
    tsp::ObjectUuidMapEntry {
        identifier,
        uuid: tsp::Uuid {
            lower: identifier,
            upper: identifier.saturating_add(0x1000),
        },
    }
}

pub(crate) fn object(identifier: u64, type_: u32, data: Vec<u8>) -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage { type_, data }],
    )?)
}

pub(crate) fn set_message_info(object: &mut ArchiveObject, references: &[u64]) -> TestResult {
    let info = object
        .archive_info
        .message_infos
        .first_mut()
        .ok_or_else(|| io::Error::other("synthetic object has no message info"))?;
    info.object_references = references.to_vec();
    Ok(())
}

pub(crate) fn set_field_info(
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

pub(crate) fn comment_entry(
    key: u32,
    root_identifier: u64,
    refcount: u32,
) -> tst::table_data_list::ListEntry {
    tst::table_data_list::ListEntry {
        key,
        refcount,
        comment_storage: Some(reference(root_identifier)),
        ..Default::default()
    }
}

pub(crate) fn string_entry() -> tst::table_data_list::ListEntry {
    tst::table_data_list::ListEntry {
        key: 1,
        refcount: 1,
        string: Some("fixture seed".to_owned()),
        ..Default::default()
    }
}

pub(crate) fn formula_entry() -> tst::table_data_list::ListEntry {
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

pub(crate) fn formula_error_entry() -> tst::table_data_list::ListEntry {
    tst::table_data_list::ListEntry {
        key: 1,
        refcount: 1,
        string: Some("#VALUE!".to_owned()),
        ..Default::default()
    }
}

pub(crate) fn format_entry() -> tst::table_data_list::ListEntry {
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

pub(crate) fn list_message(
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

/// A valid comment-list segment used by the read-parity cases. The segment
/// owns the same semantic key as the root-list cases, but its entry and
/// storage edge are carried by the type-6011 payload.
pub(crate) fn comment_segment() -> TestResult<ArchiveObject> {
    let mut result = object(
        SEGMENT_ID,
        6_011,
        tst::TableDataListSegment {
            list_type: tst::table_data_list::ListType::CommentStorage as i32,
            key_range: tsp::Range {
                location: 1,
                length: 1,
            },
            entries: vec![comment_entry(1, ROOT_COMMENT_ID, 1)],
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    set_message_info(&mut result, &[ROOT_COMMENT_ID])?;
    set_field_info(&mut result, vec![3, 1], &[ROOT_COMMENT_ID])?;
    Ok(result)
}

#[allow(
    dead_code,
    reason = "used by the focused cross-crate provenance fixture"
)]
fn unrelated_comment_segment() -> TestResult<ArchiveObject> {
    let mut result = object(
        UNRELATED_SEGMENT_ID,
        6_011,
        tst::TableDataListSegment {
            list_type: tst::table_data_list::ListType::CommentStorage as i32,
            key_range: tsp::Range {
                location: 1,
                length: 1,
            },
            entries: vec![comment_entry(1, ROOT_COMMENT_ID, 1)],
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    set_message_info(&mut result, &[ROOT_COMMENT_ID])?;
    set_field_info(&mut result, vec![3, 1], &[ROOT_COMMENT_ID])?;
    Ok(result)
}

pub(crate) fn sidecar(
    mode: FixtureMode,
    corruption: Option<Corruption>,
) -> TestResult<ArchiveObject> {
    let mut comment_entries = match mode {
        FixtureMode::Rootless | FixtureMode::SparseCell | FixtureMode::SegmentedRoot => Vec::new(),
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
            if matches!(mode, FixtureMode::SegmentedRoot)
                || matches!(corruption, Some(Corruption::SegmentedList))
            {
                vec![reference(SEGMENT_ID)]
            } else {
                Vec::new()
            },
        ),
    ];
    if matches!(corruption, Some(Corruption::MissingCommentList)) {
        messages.pop();
        return Ok(ArchiveObject::new(SIDECAR_ID, messages)?);
    }
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
        FixtureMode::Rootless | FixtureMode::SparseCell => Vec::new(),
        FixtureMode::SegmentedRoot => vec![SEGMENT_ID],
        FixtureMode::SharedReply | FixtureMode::CrossComponent => {
            vec![ROOT_COMMENT_ID, SECOND_ROOT_COMMENT_ID]
        },
        _ => vec![ROOT_COMMENT_ID],
    };
    info.object_references = roots.clone();
    for (key, root) in roots.iter().enumerate() {
        let field_number = if matches!(mode, FixtureMode::SegmentedRoot) {
            4
        } else {
            3
        };
        let mut field = FieldInfo::new(vec![field_number, u32::try_from(key + 1)?]);
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

pub(crate) fn table_model(comment_list: bool, columns: u32) -> tst::TableModelArchive {
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
        number_of_columns: columns,
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
            comment_storage_table: comment_list.then(|| reference(SIDECAR_ID)),
            next_row_strip_id: 1,
            next_column_strip_id: 1,
            row_tile_tree: tst::TableRbTree::default(),
            column_tile_tree: tst::TableRbTree::default(),
            ..Default::default()
        },
        ..Default::default()
    }
}

pub(crate) fn comment_cell(key: Option<u32>) -> TestResult<Vec<u8>> {
    let mut cell = BncCell::minimal();
    cell.set_number(42.0)?;
    cell.set_comment_identifier(key);
    Ok(cell.encode())
}

pub(crate) fn tile(mode: FixtureMode) -> TestResult<tst::Tile> {
    let keys = match mode {
        FixtureMode::Rootless | FixtureMode::SparseCell => [None, None],
        FixtureMode::SharedRoot => [Some(1), Some(1)],
        FixtureMode::SharedReply | FixtureMode::CrossComponent => [Some(1), Some(2)],
        FixtureMode::DuplicateText | FixtureMode::SingleRoot | FixtureMode::SegmentedRoot => {
            [Some(1), None]
        },
    };
    let mut rows = Vec::new();
    for (row, key) in keys.into_iter().enumerate() {
        let bytes = comment_cell(key)?;
        let sparse = matches!(mode, FixtureMode::SparseCell);
        let offsets = if sparse {
            vec![0, 0, 0xff, 0xff]
        } else {
            vec![0, 0]
        };
        rows.push(tst::TileRowInfo {
            tile_row_index: u32::try_from(row)?,
            cell_count: 1,
            storage_version: Some(5),
            cell_storage_buffer_pre_bnc: bytes.clone(),
            cell_offsets_pre_bnc: vec![0, 0],
            cell_storage_buffer: Some(bytes),
            cell_offsets: Some(offsets),
            ..Default::default()
        });
    }
    Ok(tst::Tile {
        max_column: if matches!(mode, FixtureMode::SparseCell) {
            1
        } else {
            0
        },
        max_row: 1,
        num_cells: 2,
        numrows: 2,
        row_infos: rows,
        storage_version: Some(5),
        last_saved_in_bnc: Some(true),
        ..Default::default()
    })
}

pub(crate) fn comment_archive(
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

pub(crate) fn reply_objects(
    mode: FixtureMode,
    corruption: Option<Corruption>,
) -> TestResult<Vec<ArchiveObject>> {
    if matches!(mode, FixtureMode::Rootless | FixtureMode::SparseCell) {
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
            // Keep every node present so this is a true nested graph rather
            // than a missing-reference case; the focused reader must reject
            // the second-level reply before publishing the first reply.
            Corruption::NestedReplyReference => vec![reference(FIRST_REPLY_ID)],
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
        let nested = matches!(corruption, Some(Corruption::NestedReplyReference));
        let nested_replies = nested
            .then(|| vec![reference(SECOND_REPLY_ID)])
            .unwrap_or_default();
        result.push(comment_archive(
            FIRST_REPLY_ID,
            first_text,
            &nested_replies,
        )?);
        if nested {
            result.push(comment_archive(SECOND_REPLY_ID, "nested reply", &[])?);
        }
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

pub(crate) fn metadata_component(
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

pub(crate) fn metadata(mode: FixtureMode, corruption: Option<Corruption>) -> TestResult<Vec<u8>> {
    let mut document_ids = vec![
        DOCUMENT_ID,
        SHEET_ID,
        TABLE_INFO_ID,
        TABLE_MODEL_ID,
        SIDECAR_ID,
        TILE_ID,
        SEGMENT_ID,
        ROOT_COMMENT_ID,
        SECOND_ROOT_COMMENT_ID,
        FIRST_REPLY_ID,
        SECOND_REPLY_ID,
        THIRD_REPLY_ID,
        SHARED_REPLY_ID,
        AUTHOR_ID,
        AUTHOR_STORAGE_ID,
    ];
    if matches!(mode, FixtureMode::Rootless | FixtureMode::SparseCell) {
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
    if !matches!(mode, FixtureMode::SegmentedRoot) {
        document_ids.retain(|identifier| *identifier != SEGMENT_ID);
    }
    if matches!(corruption, Some(Corruption::MissingAuthorStorage)) {
        document_ids.retain(|identifier| !matches!(*identifier, AUTHOR_ID | AUTHOR_STORAGE_ID));
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

pub(crate) fn compressed(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

pub(crate) fn author_objects(empty: bool) -> TestResult<Vec<ArchiveObject>> {
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

pub(crate) fn fixture(mode: FixtureMode, corruption: Option<Corruption>) -> TestResult<Vec<u8>> {
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
        table_model(
            !matches!(corruption, Some(Corruption::MissingCommentList)),
            if matches!(mode, FixtureMode::SparseCell) {
                2
            } else {
                1
            },
        )
        .encode_to_vec(),
    )?;
    set_message_info(&mut model, &[SIDECAR_ID, TILE_ID])?;
    set_field_info(&mut model, vec![25], &[SIDECAR_ID])?;
    set_field_info(&mut model, vec![26], &[TILE_ID])?;
    if !matches!(corruption, Some(Corruption::MissingCommentList)) {
        set_field_info(&mut model, vec![4, 19], &[SIDECAR_ID])?;
    }

    let mut tile = object(TILE_ID, TILE_TYPE, tile(mode)?.encode_to_vec())?;
    set_message_info(&mut tile, &[])?;
    let sidecar = sidecar(mode, corruption)?;
    let mut document_objects = vec![document, sheet, info, model, sidecar, tile];
    if matches!(mode, FixtureMode::SegmentedRoot) {
        document_objects.push(comment_segment()?);
    }
    if !matches!(corruption, Some(Corruption::MissingAuthor)) {
        document_objects.extend(author_objects(matches!(
            (mode, corruption),
            (FixtureMode::Rootless | FixtureMode::SparseCell, _)
        ))?);
    }
    if matches!(corruption, Some(Corruption::MissingAuthorStorage)) {
        document_objects.retain(|object| {
            !object.messages.iter().any(|message| {
                message.type_ == ANNOTATION_AUTHOR_STORAGE_TYPE
                    || message.type_ == ANNOTATION_AUTHOR_TYPE
            })
        });
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

/// Extend the ordinary fixture with a second independent table whose first
/// cell deliberately reuses key `1`. Numbers list keys are local to a table's
/// comment-storage list; this source proves that a reader does not use the
/// package-wide number of cells carrying the same key as the selected list's
/// refcount.
#[allow(
    dead_code,
    reason = "used by the focused cross-crate multitable fixture"
)]
pub(crate) fn multitable_same_key_fixture() -> TestResult<Vec<u8>> {
    let source = fixture(FixtureMode::SingleRoot, None)?;
    let catalog = Catalog::from_bytes(&source)?;

    let document_entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("comment fixture document member is missing"))?;
    let document_stream = SnappyStream::decompress(document_entry.data())?;
    let mut document = Archive::parse(document_stream.as_bytes())?;

    {
        let sheet = document
            .object_mut(SHEET_ID)
            .ok_or_else(|| io::Error::other("comment fixture sheet is missing"))?;
        let message = sheet
            .messages
            .first_mut()
            .ok_or_else(|| io::Error::other("comment fixture sheet payload is missing"))?;
        let mut archive = tn::SheetArchive::decode(message.data.as_slice())?;
        archive.drawable_infos.push(reference(SECOND_TABLE_INFO_ID));
        message.data = archive.encode_to_vec();
        let info = sheet
            .archive_info
            .message_infos
            .first_mut()
            .ok_or_else(|| io::Error::other("comment fixture sheet metadata is missing"))?;
        info.object_references.push(SECOND_TABLE_INFO_ID);
        let mut field = FieldInfo::new(vec![4, 2]);
        field.r#type = Some(FieldType::ObjectReference);
        field.object_references.push(SECOND_TABLE_INFO_ID);
        info.field_infos.push(field);
    }

    let mut second_info = object(
        SECOND_TABLE_INFO_ID,
        TABLE_INFO_TYPE,
        tst::TableInfoArchive {
            super_: tsd::DrawableArchive::default(),
            table_model: reference(SECOND_TABLE_MODEL_ID),
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    set_message_info(&mut second_info, &[SECOND_TABLE_MODEL_ID])?;
    set_field_info(&mut second_info, vec![4], &[SECOND_TABLE_MODEL_ID])?;

    let mut second_model_archive = table_model(true, 1);
    second_model_archive.table_id = "reply-fixture-table-2-id".to_owned();
    second_model_archive.table_name = SECOND_TABLE_NAME.to_owned();
    second_model_archive.table_style = reference(SECOND_SIDECAR_ID);
    second_model_archive.body_text_style = reference(SECOND_SIDECAR_ID);
    second_model_archive.header_row_text_style = reference(SECOND_SIDECAR_ID);
    second_model_archive.header_column_text_style = reference(SECOND_SIDECAR_ID);
    second_model_archive.footer_row_text_style = reference(SECOND_SIDECAR_ID);
    second_model_archive.body_cell_style = reference(SECOND_SIDECAR_ID);
    second_model_archive.header_row_style = reference(SECOND_SIDECAR_ID);
    second_model_archive.header_column_style = reference(SECOND_SIDECAR_ID);
    second_model_archive.footer_row_style = reference(SECOND_SIDECAR_ID);
    second_model_archive.base_data_store.column_headers = reference(SECOND_SIDECAR_ID);
    second_model_archive.base_data_store.tiles.tiles[0].tile = reference(SECOND_TILE_ID);
    second_model_archive.base_data_store.string_table = reference(SECOND_SIDECAR_ID);
    second_model_archive.base_data_store.style_table = reference(SECOND_SIDECAR_ID);
    second_model_archive.base_data_store.formula_table = reference(SECOND_SIDECAR_ID);
    second_model_archive.base_data_store.formula_error_table = Some(reference(SECOND_SIDECAR_ID));
    second_model_archive.base_data_store.format_table_pre_bnc = reference(SECOND_SIDECAR_ID);
    second_model_archive.base_data_store.format_table = Some(reference(SECOND_SIDECAR_ID));
    second_model_archive.base_data_store.comment_storage_table = Some(reference(SECOND_SIDECAR_ID));
    let mut second_model = object(
        SECOND_TABLE_MODEL_ID,
        TABLE_MODEL_TYPE,
        second_model_archive.encode_to_vec(),
    )?;
    set_message_info(&mut second_model, &[SECOND_SIDECAR_ID, SECOND_TILE_ID])?;
    set_field_info(&mut second_model, vec![25], &[SECOND_SIDECAR_ID])?;
    set_field_info(&mut second_model, vec![26], &[SECOND_TILE_ID])?;
    set_field_info(&mut second_model, vec![4, 19], &[SECOND_SIDECAR_ID])?;

    let mut second_sidecar = sidecar(FixtureMode::SingleRoot, None)?;
    second_sidecar.archive_info.identifier = Some(SECOND_SIDECAR_ID);
    let comment_index = second_sidecar
        .messages
        .iter()
        .position(|message| {
            tst::TableDataList::decode(message.data.as_slice())
                .map(|list| list.list_type == tst::table_data_list::ListType::CommentStorage as i32)
                .unwrap_or(false)
        })
        .ok_or_else(|| io::Error::other("second comment list is missing"))?;
    let list_message = second_sidecar
        .messages
        .get_mut(comment_index)
        .ok_or_else(|| io::Error::other("second comment list payload is missing"))?;
    let mut list = tst::TableDataList::decode(list_message.data.as_slice())?;
    let entry = list
        .entries
        .first_mut()
        .ok_or_else(|| io::Error::other("second comment list entry is missing"))?;
    entry.comment_storage = Some(reference(SECOND_ROOT_COMMENT_ID));
    list_message.data = list.encode_to_vec();
    let info = second_sidecar
        .archive_info
        .message_infos
        .get_mut(comment_index)
        .ok_or_else(|| io::Error::other("second comment list metadata is missing"))?;
    info.object_references = vec![SECOND_ROOT_COMMENT_ID];
    for field in &mut info.field_infos {
        field.object_references = vec![SECOND_ROOT_COMMENT_ID];
    }

    let mut second_tile = object(
        SECOND_TILE_ID,
        TILE_TYPE,
        tile(FixtureMode::SingleRoot)?.encode_to_vec(),
    )?;
    set_message_info(&mut second_tile, &[])?;
    let second_root = comment_archive(
        SECOND_ROOT_COMMENT_ID,
        "second table root",
        &[reference(SECOND_REPLY_ID)],
    )?;
    let second_reply = comment_archive(SECOND_REPLY_ID, "second table reply", &[])?;
    document.objects.extend([
        second_info,
        second_model,
        second_sidecar,
        second_tile,
        second_root,
        second_reply,
    ]);

    let compressed_document = SnappyStream::compress(&document.to_bytes()?)?;

    let metadata_entry = catalog
        .iter()
        .find(|entry| entry.name() == METADATA_MEMBER)
        .ok_or_else(|| io::Error::other("comment fixture metadata member is missing"))?;
    let metadata_stream = SnappyStream::decompress(metadata_entry.data())?;
    let mut metadata_archive = Archive::parse(metadata_stream.as_bytes())?;
    let metadata = metadata_archive
        .object_mut(METADATA_OBJECT_ID)
        .ok_or_else(|| io::Error::other("comment fixture metadata object is missing"))?;
    let metadata_message = metadata
        .messages
        .first_mut()
        .ok_or_else(|| io::Error::other("comment fixture metadata payload is missing"))?;
    let mut package_metadata = tsp::PackageMetadata::decode(metadata_message.data.as_slice())?;
    let document_component = package_metadata
        .components
        .iter_mut()
        .find(|component| component.preferred_locator == "Document")
        .ok_or_else(|| io::Error::other("comment fixture Document component is missing"))?;
    for identifier in [
        SECOND_TABLE_INFO_ID,
        SECOND_TABLE_MODEL_ID,
        SECOND_SIDECAR_ID,
        SECOND_TILE_ID,
        SECOND_ROOT_COMMENT_ID,
        SECOND_REPLY_ID,
    ] {
        if !document_component
            .object_uuid_map_entries
            .iter()
            .any(|entry| entry.identifier == identifier)
        {
            document_component
                .object_uuid_map_entries
                .push(uuid_entry(identifier));
        }
    }
    metadata_message.data = package_metadata.encode_to_vec();
    let compressed_metadata = SnappyStream::compress(&metadata_archive.to_bytes()?)?;

    Ok(catalog.reassemble_to_bytes(
        &[
            EntryEdit::new(DOCUMENT_MEMBER, compressed_document.as_slice()),
            EntryEdit::new(METADATA_MEMBER, compressed_metadata.as_slice()),
        ],
        Limits::default(),
    )?)
}

/// Add an unreferenced segment carrying the same key and storage edge as the
/// selected segment. The root list still points only to `SEGMENT_ID`; a
/// reader must use that parent edge when resolving the selected entry rather
/// than accepting an unrelated segment by key/storage coincidence.
#[allow(
    dead_code,
    reason = "used by the focused cross-crate provenance fixture"
)]
pub(crate) fn segmented_with_unrelated_segment_fixture() -> TestResult<Vec<u8>> {
    let source = fixture(FixtureMode::SegmentedRoot, None)?;
    let catalog = Catalog::from_bytes(&source)?;

    let document_entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("comment fixture document member is missing"))?;
    let document_stream = SnappyStream::decompress(document_entry.data())?;
    let mut document = Archive::parse(document_stream.as_bytes())?;
    document.objects.push(unrelated_comment_segment()?);
    let compressed_document = SnappyStream::compress(&document.to_bytes()?)?;

    let metadata_entry = catalog
        .iter()
        .find(|entry| entry.name() == METADATA_MEMBER)
        .ok_or_else(|| io::Error::other("comment fixture metadata member is missing"))?;
    let metadata_stream = SnappyStream::decompress(metadata_entry.data())?;
    let mut metadata_archive = Archive::parse(metadata_stream.as_bytes())?;
    let metadata = metadata_archive
        .object_mut(METADATA_OBJECT_ID)
        .ok_or_else(|| io::Error::other("comment fixture metadata object is missing"))?;
    let metadata_message = metadata
        .messages
        .first_mut()
        .ok_or_else(|| io::Error::other("comment fixture metadata payload is missing"))?;
    let mut package_metadata = tsp::PackageMetadata::decode(metadata_message.data.as_slice())?;
    let document_component = package_metadata
        .components
        .iter_mut()
        .find(|component| component.preferred_locator == "Document")
        .ok_or_else(|| io::Error::other("comment fixture Document component is missing"))?;
    if !document_component
        .object_uuid_map_entries
        .iter()
        .any(|entry| entry.identifier == UNRELATED_SEGMENT_ID)
    {
        document_component
            .object_uuid_map_entries
            .push(uuid_entry(UNRELATED_SEGMENT_ID));
    }
    metadata_message.data = package_metadata.encode_to_vec();
    let compressed_metadata = SnappyStream::compress(&metadata_archive.to_bytes()?)?;

    Ok(catalog.reassemble_to_bytes(
        &[
            EntryEdit::new(DOCUMENT_MEMBER, compressed_document.as_slice()),
            EntryEdit::new(METADATA_MEMBER, compressed_metadata.as_slice()),
        ],
        Limits::default(),
    )?)
}
