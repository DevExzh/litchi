#![no_main]

//! Strict source-backed integration coverage for Pages body-table hidden axes.
//!
//! The package fixture below is intentionally assembled as a native IWA graph:
//! a body table attachment owns a table-info record, the table-info record
//! points at a model and a column/row UID map, and the model owns hidden-state
//! extents through formula-owner records.  The semantic value is therefore
//! read through the same ownership proof used by a real package.  Several
//! fields are appended with literal wire helpers after protobuf construction;
//! this keeps unknown-field and malformed-wire checks independent of the
//! focused hidden-axis writer.

use std::error::Error as StdError;
use std::sync::Arc;
use std::thread;

use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_common::{
    decode_varint_from_bytes,
    wire::{RawWireField, RawWireFields, WireView},
};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, FieldPath, RawMessage, SnappyStream};
use litchi_iwa_protos::{
    pages_hidden_state_codec as hidden_codec, tp, tsa, tsce, tsd, tsp, tst, tswp,
};
use litchi_pages::table::hidden_axes::{AxisIndex, HiddenAxes};
use litchi_pages::{
    BodyTableHiddenAxesError as Error, BodyTableHiddenAxesLimitKind, BodyTableSelector, Limits,
    Package, PackageError,
};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const ROOT_IDENTIFIER: u64 = 1;
const BODY_IDENTIFIER: u64 = 42;
const TABLE_COUNT: usize = 2;
const ATTACHMENT_BASE: u64 = 100;
const DRAWABLE_BASE: u64 = 200;
const MODEL_BASE: u64 = 300;
const UID_MAP_BASE: u64 = 400;
const FORMULA_OWNER_BASE: u64 = 500;
const FILTER_SET_BASE: u64 = 600;
const FORMULA_OBJECT_BASE: u64 = 700;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const TABLE_INFO_VIEW_UIDS_FIELD: u32 = 6;
const TABLE_MODEL_UID_MAP_FIELD: u32 = 46;
// Canonical TST.ColumnRowUidMapArchive message type in Pages components.
const UID_MAP_MESSAGE_TYPE: u32 = 6_267;
const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;
const HIDDEN_STATE_FORMULA_OWNER_MESSAGE_TYPE: u32 = 6_204;
const FILTER_SET_MESSAGE_TYPE: u32 = 6_220;
const ROOT_MESSAGE_TYPE: u32 = 10_000;
const BODY_MESSAGE_TYPE: u32 = 2_001;
const ATTACHMENT_MESSAGE_TYPE: u32 = 2_003;
const TABLE_ROWS: u32 = 4;
const TABLE_COLUMNS: u32 = 4;
const PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

// Unknown fields deliberately sit outside every selected native projection.
const UNKNOWN_MODEL_FIELD: u32 = 99;
const UNKNOWN_MODEL_VALUE: u64 = 0xfeed_beef;
const UNKNOWN_INFO_FIELD: u32 = 98;
const UNKNOWN_INFO_VALUE: u64 = 0xcafe_babe;
const UNKNOWN_OWNER_FIELD: u32 = 97;
const UNKNOWN_OWNER_VALUE: u64 = 0xabad_1dea;
const UNKNOWN_FIXED32_FIELD: u32 = 100;
const UNKNOWN_FIXED32_VALUE: [u8; 4] = [0x78, 0x56, 0x34, 0x12];
const UNKNOWN_FIXED64_FIELD: u32 = 101;
const UNKNOWN_FIXED64_VALUE: [u8; 8] = [0xef, 0xcd, 0xab, 0x89, 0x67, 0x45, 0x23, 0x01];
const UNKNOWN_LENGTH_FIELD: u32 = 102;
const UNKNOWN_LENGTH_VALUE: &[u8] = b"opaque-unknown";
const UNKNOWN_GROUP_FIELD: u32 = 103;
const UNKNOWN_GROUP_VALUE: &[u8] = &[0x08, 0x01];
const UNKNOWN_NESTED_REFERENCE_FIELD: u32 = 90;
const UNKNOWN_NESTED_REFERENCE_VALUE: u64 = 0x1234_5678;
const UNKNOWN_NESTED_UUID_FIELD: u32 = 91;
const UNKNOWN_NESTED_UUID_VALUE: u64 = 0x8765_4321;

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

fn append_unknown_scalar_fields(output: &mut Vec<u8>, varint_field: u32, value: u64) -> TestResult {
    litchi_iwa_common::wire::append_varint_field(output, varint_field, value)?;
    // Canonical keys for fields 100 (fixed32), 101 (fixed64), and 102
    // (length-delimited). Keeping these records literal makes preservation
    // independent of generated protobuf structs while exercising scalar wire
    // kinds that are accepted by the top-level discovery projections.
    output.extend_from_slice(&[0xa5, 0x06]);
    output.extend_from_slice(&UNKNOWN_FIXED32_VALUE);
    output.extend_from_slice(&[0xa9, 0x06]);
    output.extend_from_slice(&UNKNOWN_FIXED64_VALUE);
    litchi_iwa_common::wire::append_length_delimited_field(
        output,
        UNKNOWN_LENGTH_FIELD,
        UNKNOWN_LENGTH_VALUE,
    )?;
    Ok(())
}

fn append_unknown_wire_fields(output: &mut Vec<u8>, varint_field: u32, value: u64) -> TestResult {
    append_unknown_scalar_fields(output, varint_field, value)?;
    // Field 103's start/end keys are 0xbb 0x06 / 0xbc 0x06.
    output.extend_from_slice(&[0xbb, 0x06]);
    output.extend_from_slice(UNKNOWN_GROUP_VALUE);
    output.extend_from_slice(&[0xbc, 0x06]);
    Ok(())
}

trait ExactBytes {
    fn exact_bytes(&self) -> Vec<u8>;
}

impl ExactBytes for Package {
    fn exact_bytes(&self) -> Vec<u8> {
        let mut bytes = Vec::new();
        self.write_to(&mut bytes)
            .expect("an in-memory Vec accepts package bytes");
        bytes
    }
}

#[derive(Clone, Copy, Debug, Default)]
struct TableOptions {
    /// Include a user-hidden row and column in the active state.
    user_hidden: bool,
    /// Add non-user filtered/pivot markers that must not be discarded.
    filtered_or_pivot: bool,
    /// Mark the table as a pivot topology unsupported by this transaction.
    pivot_table: bool,
    /// Mark the table's drawable as locked.
    locked: bool,
    /// Use a valid non-identity physical-to-stable UID permutation.  This
    /// keeps the semantic resolver honest: positional edits must go through
    /// the map rather than relying on sorted payload order.
    non_identity_uid_map: bool,
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn uuid(lower: u64, upper: u64) -> tsp::Uuid {
    tsp::Uuid { lower, upper }
}

fn table_identifier(base: u64, index: usize) -> u64 {
    base + u64::try_from(index).expect("fixture index fits")
}

fn table_attachment(index: usize) -> u64 {
    table_identifier(ATTACHMENT_BASE, index)
}

fn table_drawable(index: usize) -> u64 {
    table_identifier(DRAWABLE_BASE, index)
}

fn table_model(index: usize) -> u64 {
    table_identifier(MODEL_BASE, index)
}

fn table_uid_map(index: usize) -> u64 {
    table_identifier(UID_MAP_BASE, index)
}

fn table_formula_owner(index: usize) -> u64 {
    table_identifier(FORMULA_OWNER_BASE, index)
}

fn table_filter_set(index: usize, column: bool) -> u64 {
    FILTER_SET_BASE + u64::try_from(index * 2 + usize::from(!column)).expect("fixture index fits")
}

fn table_formula_object(index: usize, column: bool) -> u64 {
    FORMULA_OBJECT_BASE
        + u64::try_from(index * 2 + usize::from(!column)).expect("fixture index fits")
}

fn field_reference(path: impl Into<FieldPath>, identifier: u64) -> FieldInfo {
    let mut field = FieldInfo::new(path);
    field.object_references.push(identifier);
    field
}

fn object(
    identifier: u64,
    type_: u32,
    data: Vec<u8>,
    references: &[u64],
) -> TestResult<ArchiveObject> {
    let mut object = ArchiveObject::new(identifier, vec![RawMessage { type_, data }])?;
    object.archive_info.message_infos[0]
        .object_references
        .extend_from_slice(references);
    Ok(object)
}

fn table_uids(index: usize) -> (Vec<tsp::Uuid>, Vec<tsp::Uuid>) {
    let upper = 0x100 + u64::try_from(index).expect("fixture index fits");
    let columns = (0..TABLE_COLUMNS)
        .map(|offset| uuid(0x1000 + u64::from(offset), upper))
        .collect();
    let rows = (0..TABLE_ROWS)
        .map(|offset| uuid(0x2000 + u64::from(offset), upper))
        .collect();
    (rows, columns)
}

fn uid_map_payload(index: usize, non_identity: bool) -> Vec<u8> {
    let (rows, columns) = table_uids(index);
    let (column_index_for_uid, column_uid_for_index, row_index_for_uid, row_uid_for_index) =
        if non_identity {
            (
                vec![1, 3, 0, 2],
                vec![2, 0, 3, 1],
                vec![2, 1, 3, 0],
                vec![3, 1, 0, 2],
            )
        } else {
            (
                (0..TABLE_COLUMNS).collect(),
                (0..TABLE_COLUMNS).collect(),
                (0..TABLE_ROWS).collect(),
                (0..TABLE_ROWS).collect(),
            )
        };
    tst::ColumnRowUidMapArchive {
        sorted_column_uids: columns,
        column_index_for_uid,
        column_uid_for_index,
        sorted_row_uids: rows,
        row_index_for_uid,
        row_uid_for_index,
    }
    .encode_to_vec()
}

fn extent_state(
    uid: tsp::Uuid,
    user_hidden: Option<bool>,
    filtered: Option<bool>,
    pivot_hidden: Option<bool>,
) -> tst::hidden_state_extent_archive::RowOrColumnState {
    tst::hidden_state_extent_archive::RowOrColumnState {
        row_or_column_uid: uid,
        user_hidden,
        filtered,
        pivot_hidden,
    }
}

fn hidden_owner_payload(
    index: usize,
    options: TableOptions,
) -> (tst::HiddenStatesOwnerArchive, tsp::Uuid) {
    let (rows, columns) = table_uids(index);
    let formula_owner_uid = uuid(
        0x3000 + u64::try_from(index).expect("fixture index fits"),
        0x100 + u64::try_from(index).expect("fixture index fits"),
    );
    let hidden_states_uid = uuid(formula_owner_uid.lower + 4, formula_owner_uid.upper);
    let column_extent_uid = uuid(hidden_states_uid.lower + 7, hidden_states_uid.upper);
    let row_states = if options.user_hidden {
        let mut states = vec![
            // Native Pages may retain an explicit visible marker.  It is not
            // part of the public semantic value, but an existing-owner edit
            // must preserve it while the axis remains visible.
            extent_state(rows[0], Some(false), None, None),
            extent_state(rows[1], Some(true), None, None),
        ];
        if options.filtered_or_pivot {
            // This marker is not user-owned and must survive a rewrite.
            states.push(extent_state(rows[2], None, Some(true), None));
        }
        states
    } else if options.filtered_or_pivot {
        vec![extent_state(rows[2], None, Some(true), None)]
    } else {
        Vec::new()
    };
    let column_states = if options.user_hidden {
        let mut states = vec![
            extent_state(columns[0], Some(false), None, None),
            extent_state(columns[2], Some(true), None, None),
        ];
        if options.filtered_or_pivot {
            states.push(extent_state(columns[1], None, None, Some(true)));
        }
        states
    } else if options.filtered_or_pivot {
        vec![extent_state(columns[1], None, None, Some(true))]
    } else {
        Vec::new()
    };
    let column_extent = tst::HiddenStateExtentArchive {
        hidden_state_extent_uid: column_extent_uid,
        base_hidden_states: column_states,
        row_or_column_direction:
            tst::hidden_state_extent_archive::RowOrColumnDirection::ColumnDirection as i32,
        filter_set: Some(reference(table_filter_set(index, true))),
        needs_to_update_filter_set_for_import: Some(false),
        ..tst::HiddenStateExtentArchive::default()
    };
    let row_extent = tst::HiddenStateExtentArchive {
        hidden_state_extent_uid: hidden_states_uid,
        base_hidden_states: row_states,
        row_or_column_direction:
            tst::hidden_state_extent_archive::RowOrColumnDirection::RowDirection as i32,
        filter_set: Some(reference(table_filter_set(index, false))),
        needs_to_update_filter_set_for_import: Some(false),
        ..tst::HiddenStateExtentArchive::default()
    };
    let owner = tst::HiddenStatesOwnerArchive {
        owner_uid: hidden_states_uid,
        hidden_states: vec![tst::HiddenStatesArchive {
            hidden_states_uid,
            column_hidden_state_extent: column_extent,
            row_hidden_state_extent: row_extent,
        }],
    };
    (owner, formula_owner_uid)
}

fn model_payload(index: usize, name: &str, options: TableOptions) -> TestResult<Vec<u8>> {
    let model_id = table_model(index);
    let map_id = table_uid_map(index);
    let (owner, _formula_owner_uid) = hidden_owner_payload(index, options);
    let model = tst::TableModelArchive {
        table_id: format!("table-{name}"),
        table_style: reference(model_id + 10_000),
        body_text_style: reference(model_id + 10_001),
        header_row_text_style: reference(model_id + 10_002),
        header_column_text_style: reference(model_id + 10_003),
        footer_row_text_style: reference(model_id + 10_004),
        body_cell_style: reference(model_id + 10_005),
        header_row_style: reference(model_id + 10_006),
        header_column_style: reference(model_id + 10_007),
        footer_row_style: reference(model_id + 10_008),
        // `base_data_store` is a required nested field in the Pages archive
        // schema.  Prost omits a default nested message, so keep one required
        // scalar nonzero to ensure the outer field is emitted.
        base_data_store: tst::DataStore {
            next_row_strip_id: 1,
            ..tst::DataStore::default()
        },
        base_column_row_uids: Some(reference(map_id)),
        number_of_rows: TABLE_ROWS,
        number_of_columns: TABLE_COLUMNS,
        table_name: name.to_owned(),
        default_row_height: 18.0,
        default_column_width: 64.0,
        hidden_state_formula_owner_for_columns: options
            .user_hidden
            .then(|| reference(table_formula_object(index, true))),
        hidden_state_formula_owner_for_rows: options
            .user_hidden
            .then(|| reference(table_formula_object(index, false))),
        // The owner is appended below as a raw length-delimited field after
        // adding an unknown nested field.  Protobuf's generated encoder would
        // otherwise discard that field before the package can test retention.
        hidden_states_owner: None,
        number_of_hidden_rows: (options.user_hidden || options.filtered_or_pivot)
            .then_some(u32::from(options.user_hidden) + u32::from(options.filtered_or_pivot)),
        number_of_hidden_columns: (options.user_hidden || options.filtered_or_pivot)
            .then_some(u32::from(options.user_hidden) + u32::from(options.filtered_or_pivot)),
        number_of_filtered_rows: options.filtered_or_pivot.then_some(1),
        number_of_user_hidden_rows: options.user_hidden.then_some(1),
        number_of_user_hidden_columns: options.user_hidden.then_some(1),
        pivot_owner: options.pivot_table.then(|| reference(model_id + 50_000)),
        ..tst::TableModelArchive::default()
    };
    let mut payload = model.encode_to_vec();
    let mut map_reference = reference(map_id).encode_to_vec();
    litchi_iwa_common::wire::append_varint_field(
        &mut map_reference,
        UNKNOWN_NESTED_REFERENCE_FIELD,
        UNKNOWN_NESTED_REFERENCE_VALUE,
    )?;
    payload = replace_length_field(&payload, TABLE_MODEL_UID_MAP_FIELD, &map_reference)?;
    if options.user_hidden {
        let mut owner_payload = owner.encode_to_vec();
        append_unknown_wire_fields(&mut owner_payload, UNKNOWN_OWNER_FIELD, UNKNOWN_OWNER_VALUE)?;
        litchi_iwa_common::wire::append_length_delimited_field(&mut payload, 70, &owner_payload)?;
    }
    // Literal unknown model field: no focused codec can manufacture this
    // field, so source preservation is tested against an independent shape.
    append_unknown_scalar_fields(&mut payload, UNKNOWN_MODEL_FIELD, UNKNOWN_MODEL_VALUE)?;
    Ok(payload)
}

fn table_info_payload(
    index: usize,
    options: TableOptions,
    hidden_uuid: Option<tsp::Uuid>,
) -> Vec<u8> {
    let info = tst::TableInfoArchive {
        super_: tsd::DrawableArchive {
            parent: Some(reference(BODY_IDENTIFIER)),
            locked: options.locked.then_some(true),
            ..tsd::DrawableArchive::default()
        },
        table_model: reference(table_model(index)),
        // The first table deliberately exercises the model's native
        // field-46 fallback.  The second table retains table-info field 6 so
        // both accepted native ownership shapes remain reachable.
        view_column_row_uids: (index != 0).then(|| reference(table_uid_map(index))),
        hidden_states_uuid: hidden_uuid.clone(),
        is_a_pivot_table: options.pivot_table.then_some(true),
        ..tst::TableInfoArchive::default()
    };
    let mut payload = info.encode_to_vec();
    let mut model_reference = reference(table_model(index)).encode_to_vec();
    litchi_iwa_common::wire::append_varint_field(
        &mut model_reference,
        UNKNOWN_NESTED_REFERENCE_FIELD,
        UNKNOWN_NESTED_REFERENCE_VALUE,
    )
    .expect("nested unknown reference field fits");
    payload = replace_length_field(&payload, 2, &model_reference)
        .expect("table-info model reference is present");
    if let Some(hidden_uuid) = hidden_uuid {
        let mut hidden_uuid_payload = hidden_uuid.encode_to_vec();
        litchi_iwa_common::wire::append_varint_field(
            &mut hidden_uuid_payload,
            UNKNOWN_NESTED_UUID_FIELD,
            UNKNOWN_NESTED_UUID_VALUE,
        )
        .expect("nested unknown UUID field fits");
        payload = replace_length_field(&payload, 8, &hidden_uuid_payload)
            .expect("table-info hidden UUID field is present");
    }
    append_unknown_scalar_fields(&mut payload, UNKNOWN_INFO_FIELD, UNKNOWN_INFO_VALUE)
        .expect("unknown info fields fit");
    payload
}

fn formula_owner_payload(index: usize, hidden_uuid: tsp::Uuid) -> Vec<u8> {
    let formula_owner_uid = uuid(hidden_uuid.lower - 4, hidden_uuid.upper);
    tsce::FormulaOwnerDependenciesArchive {
        formula_owner_uid,
        internal_formula_owner_id: 1,
        formula_owner: Some(reference(table_drawable(index))),
        ..tsce::FormulaOwnerDependenciesArchive::default()
    }
    .encode_to_vec()
}

fn hidden_formula_owner_payload(extent_uid: tsp::Uuid) -> Vec<u8> {
    tst::HiddenStateFormulaOwnerArchive {
        owner_id: Some(tsp::CfuuidArchive {
            uuid_bytes: None,
            uuid_w0: Some(extent_uid.lower as u32),
            uuid_w1: Some((extent_uid.lower >> 32) as u32),
            uuid_w2: Some(extent_uid.upper as u32),
            uuid_w3: Some((extent_uid.upper >> 32) as u32),
        }),
        needs_to_update_filter_set_for_import: Some(false),
        ..tst::HiddenStateFormulaOwnerArchive::default()
    }
    .encode_to_vec()
}

fn filter_set_payload() -> Vec<u8> {
    tst::FilterSetArchive {
        r#type: Some(tst::filter_set_archive::FilterSetType::FilterSetArchiveTypeAll as i32),
        is_enabled: Some(false),
        needs_formula_rewrite_for_import: Some(false),
        filter_offsets: vec![0],
        ..tst::FilterSetArchive::default()
    }
    .encode_to_vec()
}

fn hidden_state_uuid(index: usize) -> tsp::Uuid {
    uuid(
        0x3000 + u64::try_from(index).expect("fixture index fits") + 4,
        0x100 + u64::try_from(index).expect("fixture index fits"),
    )
}

fn synthetic_package(
    options: [TableOptions; TABLE_COUNT],
    names: [&str; TABLE_COUNT],
) -> TestResult<Vec<u8>> {
    let root = tp::DocumentArchive {
        super_: tsa::DocumentArchive::default(),
        body_storage: Some(reference(BODY_IDENTIFIER)),
        ..tp::DocumentArchive::default()
    };
    let body = tswp::StorageArchive {
        kind: Some(tswp::storage_archive::KindType::Body as i32),
        text: vec!["\u{fffc}\u{fffc}".to_owned()],
        table_attachment: Some(tswp::ObjectAttributeTable {
            entries: (0..TABLE_COUNT)
                .map(|index| tswp::object_attribute_table::ObjectAttribute {
                    character_index: u32::try_from(index).expect("fixture index fits"),
                    object: Some(reference(table_attachment(index))),
                })
                .collect(),
        }),
        ..tswp::StorageArchive::default()
    };
    let mut root_object = object(
        ROOT_IDENTIFIER,
        ROOT_MESSAGE_TYPE,
        root.encode_to_vec(),
        &[BODY_IDENTIFIER],
    )?;
    root_object.archive_info.message_infos[0]
        .field_infos
        .push(field_reference(vec![4], BODY_IDENTIFIER));
    let attachment_ids = (0..TABLE_COUNT).map(table_attachment).collect::<Vec<_>>();
    let mut body_object = object(
        BODY_IDENTIFIER,
        BODY_MESSAGE_TYPE,
        body.encode_to_vec(),
        &attachment_ids,
    )?;
    body_object.archive_info.message_infos[0].field_infos = (0..TABLE_COUNT)
        .map(|index| field_reference(vec![9], table_attachment(index)))
        .collect();

    let mut objects = vec![root_object, body_object];
    for index in 0..TABLE_COUNT {
        let model_id = table_model(index);
        let info_id = table_drawable(index);
        let map_id = table_uid_map(index);
        let owner_uuid = hidden_state_uuid(index);
        let attachment_payload = tswp::DrawableAttachmentArchive {
            drawable: Some(reference(info_id)),
            ..tswp::DrawableAttachmentArchive::default()
        }
        .encode_to_vec();
        let mut attachment_object = object(
            table_attachment(index),
            ATTACHMENT_MESSAGE_TYPE,
            attachment_payload,
            &[info_id],
        )?;
        attachment_object.archive_info.message_infos[0]
            .field_infos
            .push(field_reference(vec![1], info_id));

        let info_references = if index == 0 {
            vec![BODY_IDENTIFIER, model_id]
        } else {
            vec![BODY_IDENTIFIER, model_id, map_id]
        };
        let mut info_object = object(
            info_id,
            TABLE_INFO_MESSAGE_TYPE,
            table_info_payload(
                index,
                options[index],
                options[index].user_hidden.then_some(owner_uuid),
            ),
            &info_references,
        )?;
        info_object.archive_info.message_infos[0]
            .field_infos
            .extend([
                field_reference(vec![1, 2], BODY_IDENTIFIER),
                field_reference(vec![2], model_id),
            ]);
        if index != 0 {
            info_object.archive_info.message_infos[0]
                .field_infos
                .push(field_reference(vec![6], map_id));
        }

        let mut model_references = vec![map_id];
        if options[index].user_hidden {
            model_references.extend([
                table_formula_object(index, true),
                table_formula_object(index, false),
                table_filter_set(index, true),
                table_filter_set(index, false),
            ]);
        }
        let mut model_object = object(
            model_id,
            TABLE_MODEL_MESSAGE_TYPE,
            model_payload(index, names[index], options[index])?,
            &model_references,
        )?;
        model_object.archive_info.message_infos[0]
            .field_infos
            .push(field_reference(vec![46], map_id));
        if options[index].user_hidden {
            model_object.archive_info.message_infos[0]
                .field_infos
                .extend([
                    field_reference(vec![34], table_formula_object(index, true)),
                    field_reference(vec![35], table_formula_object(index, false)),
                ]);
        }

        let uid_map_object = object(
            map_id,
            UID_MAP_MESSAGE_TYPE,
            uid_map_payload(index, options[index].non_identity_uid_map),
            &[],
        )?;
        objects.extend([attachment_object, info_object, model_object, uid_map_object]);

        // Every native table has the formula-owner dependency record.  The
        // hidden-state owner itself is optional; absent-owner reads are empty
        // and non-empty creation requests must fail closed.
        let formula_owner_object = object(
            table_formula_owner(index),
            FORMULA_OWNER_MESSAGE_TYPE,
            formula_owner_payload(index, owner_uuid),
            &[info_id],
        )?;
        if options[index].user_hidden {
            let (rows, columns) = table_uids(index);
            let column_extent_uid = uuid(owner_uuid.lower + 7, owner_uuid.upper);
            let formula_columns = object(
                table_formula_object(index, true),
                HIDDEN_STATE_FORMULA_OWNER_MESSAGE_TYPE,
                hidden_formula_owner_payload(column_extent_uid),
                &[],
            )?;
            let formula_rows = object(
                table_formula_object(index, false),
                HIDDEN_STATE_FORMULA_OWNER_MESSAGE_TYPE,
                hidden_formula_owner_payload(owner_uuid),
                &[],
            )?;
            let filter_columns = object(
                table_filter_set(index, true),
                FILTER_SET_MESSAGE_TYPE,
                filter_set_payload(),
                &[],
            )?;
            let filter_rows = object(
                table_filter_set(index, false),
                FILTER_SET_MESSAGE_TYPE,
                filter_set_payload(),
                &[],
            )?;
            // Keep the row/column UUIDs live in the archive metadata too; the
            // semantic resolver must not infer a positional map from payload
            // ordering alone.
            let _ = (rows, columns);
            objects.extend([
                formula_owner_object,
                formula_columns,
                formula_rows,
                filter_columns,
                filter_rows,
            ]);
        } else {
            objects.push(formula_owner_object);
        }
    }

    let archive = Archive { objects };
    let component = SnappyStream::compress(&archive.to_bytes()?)?;
    let mut members: Vec<(&str, &[u8])> = vec![
        ("Data/sentinel.bin", b"untouched-sentinel"),
        (DOCUMENT_MEMBER, component.as_slice()),
    ];
    members.extend(
        PREVIEWS
            .into_iter()
            .map(|name| (name, b"preview".as_slice())),
    );
    Ok(litchi_iwa_archive::package::to_bytes(
        members,
        Limits::default(),
    )?)
}

fn document_archive(package: &[u8]) -> TestResult<Archive> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or("missing document member")?;
    Ok(Archive::parse(
        SnappyStream::decompress(entry.data())?.as_bytes(),
    )?)
}

fn preview_count(package: &[u8]) -> usize {
    Catalog::from_bytes(package)
        .unwrap_or_else(|error| panic!("package preview catalog is valid: {error}"))
        .iter()
        .filter(|entry| PREVIEWS.contains(&entry.name()))
        .count()
}

fn rewrite_document_archive(
    package: &[u8],
    mutate: impl FnOnce(&mut Archive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or("missing document member")?;
    let mut archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
    mutate(&mut archive)?;
    let component = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(DOCUMENT_MEMBER, &component)],
        Limits::default(),
    )?)
}

fn rewrite_model_raw(package: &[u8], index: usize, raw: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let model = archive
            .object_mut(table_model(index))
            .ok_or("missing table model")?;
        let message = model
            .messages
            .iter_mut()
            .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or("missing table-model message")?;
        message.data.extend_from_slice(raw);
        Ok(())
    })
}

fn replace_model_raw_field(
    package: &[u8],
    index: usize,
    field_number: u32,
    replacement: &[u8],
) -> TestResult<Vec<u8>> {
    let payload = model_payload_from_package(package, index)?;
    let rewritten = replace_raw_field(&payload, field_number, replacement)?;
    rewrite_document_archive(package, |archive| {
        let model = archive
            .object_mut(table_model(index))
            .ok_or("missing table model")?;
        let message = model
            .messages
            .iter_mut()
            .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or("missing table-model message")?;
        message.data = rewritten;
        Ok(())
    })
}

fn replace_length_field(
    source: &[u8],
    field_number: u32,
    replacement: &[u8],
) -> TestResult<Vec<u8>> {
    let view = WireView::parse(source)?;
    let mut output = Vec::with_capacity(source.len() + replacement.len());
    let mut replaced = false;
    for field in view.fields() {
        if field.number() == field_number && !replaced {
            litchi_iwa_common::wire::append_length_delimited_field(
                &mut output,
                field_number,
                replacement,
            )?;
            replaced = true;
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    if !replaced {
        return Err(format!("missing field {field_number}").into());
    }
    Ok(output)
}

fn replace_raw_field(source: &[u8], field_number: u32, replacement: &[u8]) -> TestResult<Vec<u8>> {
    let view = WireView::parse(source)?;
    let mut output = Vec::with_capacity(source.len() + replacement.len());
    let mut replaced = false;
    for field in view.fields() {
        if field.number() == field_number && !replaced {
            output.extend_from_slice(replacement);
            replaced = true;
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    if !replaced {
        return Err(format!("missing field {field_number}").into());
    }
    Ok(output)
}

fn append_duplicate_model_message(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let model = archive
            .object_mut(table_model(index))
            .ok_or("missing table model")?;
        let payload = model
            .messages
            .iter()
            .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or("missing table-model message")?
            .data
            .clone();
        model.push_message(RawMessage {
            type_: TABLE_MODEL_MESSAGE_TYPE,
            data: payload,
        })?;
        Ok(())
    })
}

fn append_duplicate_info_message(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let info = archive
            .object_mut(table_drawable(index))
            .ok_or("missing table info")?;
        let payload = info
            .messages
            .iter()
            .find(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .ok_or("missing table-info message")?
            .data
            .clone();
        info.push_message(RawMessage {
            type_: TABLE_INFO_MESSAGE_TYPE,
            data: payload,
        })?;
        Ok(())
    })
}

fn remove_object(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        archive
            .objects
            .retain(|object| object.archive_info.identifier != Some(identifier));
        Ok(())
    })
}

fn append_duplicate_formula_owner(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let owner = archive
            .object(table_formula_owner(index))
            .ok_or("missing formula-owner object")?
            .clone();
        let mut duplicate = owner;
        duplicate.archive_info.identifier = Some(table_formula_owner(index) + 10_000);
        archive.insert_object(duplicate)?;
        Ok(())
    })
}

fn duplicate_table_name(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    let first =
        tst::TableModelArchive::decode(model_payload_from_package(package, index)?.as_slice())?;
    let second_index = index.checked_add(1).ok_or("table index overflow")?;
    let second_payload = model_payload_from_package(package, second_index)?;
    let rewritten = replace_length_field(&second_payload, 8, first.table_name.as_bytes())?;
    rewrite_document_archive(package, |archive| {
        let model = archive
            .object_mut(table_model(second_index))
            .ok_or("missing second table model")?;
        let message = model
            .messages
            .iter_mut()
            .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or("missing second table-model message")?;
        message.data = rewritten;
        Ok(())
    })
}

fn shared_formula_owner(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    let second_index = index.checked_add(1).ok_or("table index overflow")?;
    rewrite_document_archive(package, |archive| {
        let owner = archive
            .object_mut(table_formula_owner(second_index))
            .ok_or("missing second formula-owner object")?;
        let message = owner
            .messages
            .iter_mut()
            .find(|message| message.type_ == FORMULA_OWNER_MESSAGE_TYPE)
            .ok_or("missing second formula-owner message")?;
        let mut decoded = tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())?;
        decoded.formula_owner = Some(reference(table_drawable(index)));
        message.data = decoded.encode_to_vec();
        Ok(())
    })
}

fn duplicate_active_hidden_state(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model(package, index, |model| {
        if let Some(owner) = model.hidden_states_owner.as_mut() {
            if let Some(active) = owner.hidden_states.first().cloned() {
                owner.hidden_states.push(active);
            }
        }
    })
}

fn wrong_hidden_owner_uuid(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model(package, index, |model| {
        if let Some(owner) = model.hidden_states_owner.as_mut() {
            owner.owner_uid.lower = owner.owner_uid.lower.saturating_add(1);
        }
    })
}

fn zero_owner_uuid(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model(package, index, |model| {
        if let Some(owner) = model.hidden_states_owner.as_mut() {
            owner.owner_uid = uuid(0, 0);
        }
    })
}

fn wrong_active_uuid(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    let payload = info_payload(package, index)?;
    let mut info = tst::TableInfoArchive::decode(payload.as_slice())?;
    info.hidden_states_uuid = Some(uuid(0xdead, 0xbeef));
    let mut replacement = info.encode_to_vec();
    append_unknown_scalar_fields(&mut replacement, UNKNOWN_INFO_FIELD, UNKNOWN_INFO_VALUE)?;
    rewrite_document_archive(package, |archive| {
        let object = archive
            .object_mut(table_drawable(index))
            .ok_or("missing table info")?;
        let message = object
            .messages
            .iter_mut()
            .find(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .ok_or("missing table-info message")?;
        message.data = replacement;
        Ok(())
    })
}

fn wrong_uid_map_message_type(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let map = archive
            .object_mut(table_uid_map(index))
            .ok_or("missing UID map")?;
        let message_index = map
            .messages
            .iter()
            .position(|message| message.type_ == UID_MAP_MESSAGE_TYPE)
            .ok_or("missing canonical UID-map message")?;
        map.messages[message_index].type_ = 6_005;
        map.archive_info.message_infos[message_index].type_ = 6_005;
        Ok(())
    })
}

fn dangling_formula_owner(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let owner = archive
            .object_mut(table_formula_owner(index))
            .ok_or("missing formula-owner object")?;
        let message = owner
            .messages
            .iter_mut()
            .find(|message| message.type_ == FORMULA_OWNER_MESSAGE_TYPE)
            .ok_or("missing formula-owner message")?;
        let mut decoded = tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())?;
        decoded.formula_owner = Some(reference(0xdead_beef));
        message.data = decoded.encode_to_vec();
        Ok(())
    })
}

fn wrong_uid_map_lengths(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    let mut map = tst::ColumnRowUidMapArchive::decode(
        document_archive(package)?
            .object(table_uid_map(index))
            .ok_or("missing UID map")?
            .messages
            .first()
            .ok_or("missing UID-map message")?
            .data
            .as_slice(),
    )?;
    map.row_uid_for_index.pop();
    rewrite_uid_map(package, index, map)
}

fn duplicate_uid_map_entry(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    let mut map = tst::ColumnRowUidMapArchive::decode(
        document_archive(package)?
            .object(table_uid_map(index))
            .ok_or("missing UID map")?
            .messages
            .first()
            .ok_or("missing UID-map message")?
            .data
            .as_slice(),
    )?;
    map.row_uid_for_index[1] = map.row_uid_for_index[0];
    rewrite_uid_map(package, index, map)
}

fn duplicate_hidden_owner_field(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    let payload = model_payload_from_package(package, index)?;
    let owner = WireView::parse(&payload)?
        .fields()
        .find(|field| field.number() == 70)
        .ok_or("missing hidden-state owner")?
        .payload()
        .to_vec();
    let mut duplicate = Vec::new();
    litchi_iwa_common::wire::append_length_delimited_field(&mut duplicate, 70, &owner)?;
    rewrite_model_raw(package, index, &duplicate)
}

fn wrong_hidden_extent_direction(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model(package, index, |model| {
        if let Some(owner) = model.hidden_states_owner.as_mut()
            && let Some(active) = owner.hidden_states.first_mut()
        {
            active.row_hidden_state_extent.row_or_column_direction =
                tst::hidden_state_extent_archive::RowOrColumnDirection::ColumnDirection as i32;
        }
    })
}

fn rewrite_uid_map(
    package: &[u8],
    index: usize,
    map: tst::ColumnRowUidMapArchive,
) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let map_object = archive
            .object_mut(table_uid_map(index))
            .ok_or("missing UID map")?;
        let message = map_object
            .messages
            .iter_mut()
            .find(|message| message.type_ == UID_MAP_MESSAGE_TYPE)
            .ok_or("missing UID-map message")?;
        message.data = map.encode_to_vec();
        Ok(())
    })
}

fn rewrite_model(
    package: &[u8],
    index: usize,
    mutate: impl FnOnce(&mut tst::TableModelArchive),
) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let model = archive
            .object_mut(table_model(index))
            .ok_or("missing table model")?;
        let message = model
            .messages
            .iter_mut()
            .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or("missing table-model message")?;
        let mut decoded = tst::TableModelArchive::decode(message.data.as_slice())?;
        mutate(&mut decoded);
        let mut payload = decoded.encode_to_vec();
        append_unknown_scalar_fields(&mut payload, UNKNOWN_MODEL_FIELD, UNKNOWN_MODEL_VALUE)?;
        message.data = payload;
        Ok(())
    })
}

fn rewrite_model_metadata(
    package: &[u8],
    index: usize,
    mutate: impl FnOnce(&mut ArchiveObject, usize) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let model = archive
            .object_mut(table_model(index))
            .ok_or("missing table model")?;
        let message_index = model
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or("missing table-model message")?;
        mutate(model, message_index)
    })
}

fn rewrite_info_metadata(
    package: &[u8],
    index: usize,
    mutate: impl FnOnce(&mut ArchiveObject, usize) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let info = archive
            .object_mut(table_drawable(index))
            .ok_or("missing table info")?;
        let message_index = info
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .ok_or("missing table-info message")?;
        mutate(info, message_index)
    })
}

fn remove_model_map_metadata(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model_metadata(package, index, |model, message_index| {
        model.archive_info.message_infos[message_index]
            .field_infos
            .retain(|field| field.path.as_slice() != [TABLE_MODEL_UID_MAP_FIELD]);
        Ok(())
    })
}

fn wrong_model_map_metadata_path(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model_metadata(package, index, |model, message_index| {
        let field = model.archive_info.message_infos[message_index]
            .field_infos
            .iter_mut()
            .find(|field| field.path.as_slice() == [TABLE_MODEL_UID_MAP_FIELD])
            .ok_or("missing model UID-map field metadata")?;
        field.path = FieldPath::new(vec![TABLE_MODEL_UID_MAP_FIELD + 1]);
        Ok(())
    })
}

fn remove_info_map_metadata(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_info_metadata(package, index, |info, message_index| {
        info.archive_info.message_infos[message_index]
            .field_infos
            .retain(|field| field.path.as_slice() != [TABLE_INFO_VIEW_UIDS_FIELD]);
        Ok(())
    })
}

fn invalid_filter_metadata(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let filter = archive
            .object_mut(table_filter_set(index, true))
            .ok_or("missing column filter set")?;
        let message = filter
            .messages
            .iter_mut()
            .find(|message| message.type_ == FILTER_SET_MESSAGE_TYPE)
            .ok_or("missing filter-set message")?;
        let mut decoded = tst::FilterSetArchive::decode(message.data.as_slice())?;
        decoded.needs_formula_rewrite_for_import = Some(true);
        message.data = decoded.encode_to_vec();
        Ok(())
    })
}

fn model_payload_from_package(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    Ok(document_archive(package)?
        .object(table_model(index))
        .ok_or("missing table model")?
        .messages
        .iter()
        .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .ok_or("missing table-model message")?
        .data
        .clone())
}

fn info_payload(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    Ok(document_archive(package)?
        .object(table_drawable(index))
        .ok_or("missing table info")?
        .messages
        .iter()
        .find(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
        .ok_or("missing table-info message")?
        .data
        .clone())
}

fn hidden_owner_payload_from_package(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    let payload = model_payload_from_package(package, index)?;
    Ok(WireView::parse(&payload)?
        .fields()
        .find(|field| field.number() == 70)
        .ok_or("missing hidden-state owner")?
        .payload()
        .to_vec())
}

// The fuzz input is a small descriptor.  It always builds and exercises a
// valid rooted package first; malformed graph and wire variants are additional
// probes selected by the descriptor.  This keeps the operation reachable for
// empty/random input while retaining deterministic native-shape coverage.

use libfuzzer_sys::fuzz_target;
use std::hint::black_box;
use std::sync::OnceLock;

const MAX_INPUT_BYTES: usize = 64 * 1024;
const MAX_DESCRIPTOR_BYTES: usize = 256;
const MAX_TABLE_ROWS: usize = TABLE_ROWS as usize;
const MAX_TABLE_COLUMNS: usize = TABLE_COLUMNS as usize;
// The fixture is intentionally tiny, but its Snappy streams expand beyond
// the ZIP byte count. Keep these subordinate ceilings finite and independent
// from the package-byte boundary under test; deriving them from compressed
// input size would turn compression ratio into an accidental test condition.
const FIXTURE_ARCHIVE_CEILING: u64 = 64 * 1024;
const FIXTURE_STREAM_CEILING: usize = 64 * 1024;

const COMMAND_SET: u8 = 0;
const COMMAND_CLEAR: u8 = 1;
const COMMAND_RESET: u8 = 2;

const MODE_VALID: u8 = 0;
const MODE_MISSING_ROOT: u8 = 1;
const MODE_MISSING_BODY: u8 = 2;
const MODE_MISSING_ATTACHMENT: u8 = 3;
const MODE_MISSING_DRAWABLE: u8 = 4;
const MODE_MISSING_MODEL: u8 = 5;
const MODE_MISSING_UID_MAP: u8 = 6;
const MODE_MISSING_FORMULA_OWNER: u8 = 7;
const MODE_DUPLICATE_MODEL: u8 = 8;
const MODE_DUPLICATE_INFO: u8 = 9;
const MODE_DUPLICATE_FORMULA_OWNER: u8 = 10;
const MODE_DUPLICATE_ACTIVE_STATE: u8 = 11;
const MODE_WRONG_OWNER_UUID: u8 = 12;
const MODE_WRONG_UID_MAP_LENGTH: u8 = 13;
const MODE_DUPLICATE_UID: u8 = 14;
const MODE_UNKNOWN_AXIS_UUID: u8 = 15;
const MODE_WRONG_DIRECTION: u8 = 16;
const MODE_WRONG_MESSAGE_TYPE: u8 = 17;
const MODE_DANGLING_REFERENCE: u8 = 18;
const MODE_DUPLICATE_NAME: u8 = 19;
const MODE_SHARED_OWNER: u8 = 20;
const MODE_WIRE_DUPLICATE: u8 = 21;
const MODE_WIRE_WRONG_KIND: u8 = 22;
const MODE_WIRE_TRUNCATED: u8 = 23;
const MODE_WIRE_GROUP: u8 = 24;
const MODE_WIRE_NONCANONICAL: u8 = 25;
const MODE_WIRE_INVALID_UTF8: u8 = 26;
// Dependency/topology and arithmetic-bound probes stay separate from the
// wire modes above so each corpus descriptor identifies one failure class.
const MODE_MISSING_FILTER_OWNER: u8 = 27;
const MODE_CROSS_COMPONENT_ALIAS: u8 = 28;
const MODE_AXIS_BOUNDS: u8 = 29;
const MODE_COUNT_OVERFLOW: u8 = 30;
const MODE_ZERO_UUID: u8 = 31;
const MODE_WRONG_ACTIVE_UUID: u8 = 32;
const MODE_UID_MAP_WRONG_MESSAGE_TYPE: u8 = 33;
const MODE_DANGLING_FORMULA_OWNER: u8 = 34;
const MODE_MODEL_MAP_METADATA_MISSING: u8 = 35;
const MODE_MODEL_MAP_METADATA_PATH: u8 = 36;
const MODE_INFO_MAP_METADATA_MISSING: u8 = 37;
const MODE_FILTER_METADATA: u8 = 38;
const MODE_COUNT: u8 = 39;

fuzz_target!(|data: &[u8]| {
    let descriptor = normalize_descriptor(data);

    // This is the primary path for every input, including malformed-mode
    // bytes.  The descriptor controls the selector, command, topology,
    // ownership state, and requested semantic axes.
    let source = base_fixture(&descriptor)
        .unwrap_or_else(|error| panic!("bounded valid fixture must build: {error}"));
    exercise_primary(&source, &descriptor);

    let mode = descriptor.get(3).copied().unwrap_or_default() % MODE_COUNT;
    if mode != MODE_VALID {
        let malformed = malformed_fixture(&source, mode)
            .unwrap_or_else(|error| panic!("malformed fixture mode {mode} must build: {error}"));
        exercise_malformed(&malformed, &descriptor);
    }

    // Keep one data-dependent resource probe in every iteration.  A
    // process-wide sweep below is only supplemental.
    exercise_data_dependent_limits(&source, &descriptor);
    black_box(source.len());

    static LIMIT_SWEEP: OnceLock<()> = OnceLock::new();
    LIMIT_SWEEP.get_or_init(|| {
        let source = base_fixture(&[1, COMMAND_SET, 3, MODE_VALID, 0, 0, 0, 0])
            .unwrap_or_else(|error| panic!("limit sweep fixture must build: {error}"));
        exercise_limit_boundaries(&source);
    });

    static CODEC_SWEEP: OnceLock<()> = OnceLock::new();
    CODEC_SWEEP.get_or_init(exercise_nested_codec_limits);
});

fn normalize_descriptor(data: &[u8]) -> Vec<u8> {
    let bounded = &data[..data.len().min(MAX_INPUT_BYTES)];
    if let Some(encoded) = bounded.strip_prefix(b"hex:") {
        if let Some(decoded) = decode_hex(encoded) {
            return decoded;
        }
        // Invalid hex is still a descriptor.  Falling back to the bounded
        // suffix prevents a malformed textual prefix from making the target
        // vacuous.
        return encoded[..encoded.len().min(MAX_DESCRIPTOR_BYTES)].to_vec();
    }
    bounded[..bounded.len().min(MAX_DESCRIPTOR_BYTES)].to_vec()
}

fn decode_hex(encoded: &[u8]) -> Option<Vec<u8>> {
    let mut output = Vec::with_capacity(encoded.len() / 2);
    let mut high = None;
    for byte in encoded.iter().copied() {
        if byte.is_ascii_whitespace() {
            continue;
        }
        let nibble = match byte {
            b'0'..=b'9' => byte - b'0',
            b'a'..=b'f' => byte - b'a' + 10,
            b'A'..=b'F' => byte - b'A' + 10,
            _ => return None,
        };
        if let Some(value) = high.take() {
            output.push((value << 4) | nibble);
            if output.len() > MAX_DESCRIPTOR_BYTES {
                return None;
            }
        } else {
            high = Some(nibble);
        }
    }
    high.is_none().then_some(output)
}

fn limit_maximum(source_len: usize, descriptor: &[u8]) -> u64 {
    let source_len = u64::try_from(source_len).unwrap_or(u64::MAX);
    match descriptor.get(7).copied().unwrap_or_default() & 0x0f {
        // Keep the deterministic corpus spelling explicit: 0 is exactly the
        // source size and 1 is source size + 1, not source size - 1.
        0 => source_len,
        1 => source_len.saturating_add(1),
        offset => source_len
            .saturating_sub(u64::from(offset.saturating_sub(1)))
            .max(1),
    }
}

fn base_fixture(descriptor: &[u8]) -> TestResult<Vec<u8>> {
    let command = descriptor.get(1).copied().unwrap_or(COMMAND_SET) % 4;
    let mode = descriptor.get(3).copied().unwrap_or_default() % MODE_COUNT;
    let flags = descriptor.first().copied().unwrap_or_default();

    // A changed clear/reset/no-op needs an existing owner; a nonempty set with
    // the owner bit clear deliberately reaches the terminal unsupported-
    // creation path.
    let existing_owner = flags & 1 != 0
        || command != COMMAND_SET
        || matches!(
            mode,
            MODE_MISSING_FORMULA_OWNER
                | MODE_DUPLICATE_FORMULA_OWNER
                | MODE_DUPLICATE_NAME
                | MODE_DUPLICATE_ACTIVE_STATE
                | MODE_WRONG_OWNER_UUID
                | MODE_WRONG_UID_MAP_LENGTH
                | MODE_DUPLICATE_UID
                | MODE_UNKNOWN_AXIS_UUID
                | MODE_WRONG_DIRECTION
                | MODE_WRONG_MESSAGE_TYPE
                | MODE_DANGLING_REFERENCE
                | MODE_SHARED_OWNER
                | MODE_WIRE_DUPLICATE
                | MODE_WIRE_WRONG_KIND
                | MODE_WIRE_TRUNCATED
                | MODE_WIRE_GROUP
                | MODE_WIRE_NONCANONICAL
                | MODE_WIRE_INVALID_UTF8
                | MODE_MISSING_FILTER_OWNER
                | MODE_CROSS_COMPONENT_ALIAS
                | MODE_AXIS_BOUNDS
                | MODE_COUNT_OVERFLOW
                | MODE_ZERO_UUID
                | MODE_WRONG_ACTIVE_UUID
                | MODE_UID_MAP_WRONG_MESSAGE_TYPE
                | MODE_DANGLING_FORMULA_OWNER
                | MODE_MODEL_MAP_METADATA_MISSING
                | MODE_MODEL_MAP_METADATA_PATH
                | MODE_INFO_MAP_METADATA_MISSING
                | MODE_FILTER_METADATA
        );

    let filtered = flags & 2 != 0 || mode == MODE_WIRE_NONCANONICAL;
    let pivot = flags & 4 != 0;
    let locked = flags & 8 != 0;
    let topology = mode == MODE_WRONG_DIRECTION || mode == MODE_SHARED_OWNER;
    let first = TableOptions {
        user_hidden: existing_owner || filtered || pivot || locked || topology,
        filtered_or_pivot: filtered,
        pivot_table: pivot,
        locked,
        non_identity_uid_map: flags & 0x40 != 0,
    };
    let second = TableOptions {
        // The shared-owner probe needs two independently rooted tables so it
        // can redirect only the second dependency record to the first info
        // object.  Other descriptors retain input-controlled second-table
        // ownership.
        user_hidden: flags & 0x10 != 0 || mode == MODE_SHARED_OWNER,
        filtered_or_pivot: flags & 0x20 != 0,
        pivot_table: false,
        locked: false,
        non_identity_uid_map: flags & 0x80 != 0,
    };
    synthetic_package([first, second], ["Revenue", "Costs"])
}

fn malformed_fixture(source: &[u8], mode: u8) -> TestResult<Vec<u8>> {
    match mode {
        MODE_MISSING_ROOT => remove_object(source, ROOT_IDENTIFIER),
        MODE_MISSING_BODY => remove_object(source, BODY_IDENTIFIER),
        MODE_MISSING_ATTACHMENT => remove_object(source, table_attachment(0)),
        MODE_MISSING_DRAWABLE => remove_object(source, table_drawable(0)),
        MODE_MISSING_MODEL => remove_object(source, table_model(0)),
        MODE_MISSING_UID_MAP => remove_object(source, table_uid_map(0)),
        MODE_MISSING_FORMULA_OWNER => remove_object(source, table_formula_owner(0)),
        MODE_DUPLICATE_MODEL => append_duplicate_model_message(source, 0),
        MODE_DUPLICATE_INFO => append_duplicate_info_message(source, 0),
        MODE_DUPLICATE_FORMULA_OWNER => append_duplicate_formula_owner(source, 0),
        MODE_DUPLICATE_ACTIVE_STATE => duplicate_active_hidden_state(source, 0),
        MODE_WRONG_OWNER_UUID => wrong_hidden_owner_uuid(source, 0),
        MODE_WRONG_UID_MAP_LENGTH => wrong_uid_map_lengths(source, 0),
        MODE_DUPLICATE_UID => duplicate_uid_map_entry(source, 0),
        MODE_UNKNOWN_AXIS_UUID => unknown_axis_uuid(source, 0),
        MODE_WRONG_DIRECTION => wrong_hidden_extent_direction(source, 0),
        MODE_WRONG_MESSAGE_TYPE => wrong_message_type(source, 0),
        MODE_DANGLING_REFERENCE => dangling_reference(source, 0),
        MODE_DUPLICATE_NAME => duplicate_table_name(source, 0),
        MODE_SHARED_OWNER => shared_formula_owner(source, 0),
        MODE_WIRE_DUPLICATE => duplicate_hidden_owner_field(source, 0),
        // Field 70 is a length-delimited hidden-state owner; encode the same
        // selected field with a varint wire kind.
        MODE_WIRE_WRONG_KIND => replace_model_raw_field(source, 0, 70, &[0xb0, 0x04, 0x01]),
        // Field 70's length-delimited key is b2 04.  Preserve that field
        // number while varying its wire framing so the owner scanner sees
        // each malformed case.
        MODE_WIRE_TRUNCATED => replace_model_raw_field(source, 0, 70, &[0xb2, 0x04, 0x80]),
        MODE_WIRE_GROUP => {
            replace_model_raw_field(source, 0, 70, &[0xb3, 0x04, 0x08, 0x01, 0xb4, 0x04])
        },
        MODE_WIRE_NONCANONICAL => replace_model_raw_field(source, 0, 70, &[0xb2, 0x04, 0x80, 0x00]),
        MODE_WIRE_INVALID_UTF8 => replace_model_field(source, 0, 8, &[0xff]),
        MODE_MISSING_FILTER_OWNER => remove_object(source, table_filter_set(0, true)),
        MODE_CROSS_COMPONENT_ALIAS => cross_component_alias(source, 0),
        MODE_AXIS_BOUNDS => axis_bounds(source, 0),
        MODE_COUNT_OVERFLOW => count_overflow(source, 0),
        MODE_ZERO_UUID => zero_owner_uuid(source, 0),
        MODE_WRONG_ACTIVE_UUID => wrong_active_uuid(source, 0),
        MODE_UID_MAP_WRONG_MESSAGE_TYPE => wrong_uid_map_message_type(source, 0),
        MODE_DANGLING_FORMULA_OWNER => dangling_formula_owner(source, 0),
        MODE_MODEL_MAP_METADATA_MISSING => remove_model_map_metadata(source, 0),
        MODE_MODEL_MAP_METADATA_PATH => wrong_model_map_metadata_path(source, 0),
        MODE_INFO_MAP_METADATA_MISSING => remove_info_map_metadata(source, 1),
        MODE_FILTER_METADATA => invalid_filter_metadata(source, 0),
        _ => Ok(source.to_vec()),
    }
}

fn primary_selector(descriptor: &[u8]) -> BodyTableSelector<'static> {
    let flags = descriptor.first().copied().unwrap_or_default();
    let mode = descriptor.get(3).copied().unwrap_or_default() % MODE_COUNT;
    // Keep every malformed-mode command on the mutated rooted table.  Name /
    // index variation remains input-dependent for the ordinary valid graph,
    // while lock/filter/pivot flags also stay rooted on table zero.
    if mode != MODE_VALID || flags & 0x0e != 0 {
        return BodyTableSelector::index(0);
    }
    match descriptor.get(4).copied().unwrap_or_default() % 4 {
        0 => BodyTableSelector::index(0),
        1 => BodyTableSelector::name("Revenue"),
        2 => BodyTableSelector::index(1),
        _ => BodyTableSelector::name("Costs"),
    }
}

fn probe_selectors(descriptor: &[u8]) -> [BodyTableSelector<'static>; 4] {
    let selected = primary_selector(descriptor);
    [
        selected,
        BodyTableSelector::index(usize::from(descriptor.get(6).copied().unwrap_or_default()) % 2),
        BodyTableSelector::name("Missing"),
        BodyTableSelector::index(usize::MAX),
    ]
}

fn requested_axes(descriptor: &[u8]) -> HiddenAxes {
    let row = usize::from(descriptor.get(5).copied().unwrap_or_default()) % MAX_TABLE_ROWS;
    let column = usize::from(descriptor.get(6).copied().unwrap_or_default()) % MAX_TABLE_COLUMNS;
    let axes = match descriptor.get(2).copied().unwrap_or_default() % 4 {
        0 => Vec::new(),
        1 => vec![AxisIndex::row(row)],
        2 => vec![AxisIndex::column(column)],
        _ => vec![AxisIndex::row(row), AxisIndex::column(column)],
    };
    HiddenAxes::new(axes).unwrap_or_else(|_| HiddenAxes::empty())
}

fn exercise_primary(source: &[u8], descriptor: &[u8]) {
    let package = Package::from_bytes(source)
        .unwrap_or_else(|error| panic!("valid descriptor fixture must parse: {error}"));
    let selector = primary_selector(descriptor);
    let before = package
        .body_table_hidden_axes(selector)
        .unwrap_or_else(|error| panic!("valid descriptor selector must read: {error}"));
    assert_canonical_and_bounded(&before);
    let source_bytes = package.exact_bytes();
    assert_eq!(source_bytes, source);
    let source_archive = document_archive(&source_bytes)
        .unwrap_or_else(|error| panic!("valid document archive must reopen: {error}"));
    for index in 0..TABLE_COUNT {
        assert_canonical_uid_map_message(&source_archive, index);
        assert_uid_map_route(&source_bytes, index);
        let non_identity = match index {
            0 => descriptor.first().copied().unwrap_or_default() & 0x40 != 0,
            1 => descriptor.first().copied().unwrap_or_default() & 0x80 != 0,
            _ => false,
        };
        if non_identity {
            assert_non_identity_uid_map(&source_archive, index);
        }
    }

    // Read paths are observational.  A second read and every selector probe
    // must leave the exact source untouched.
    for probe in probe_selectors(descriptor) {
        if let Ok(value) = package.body_table_hidden_axes(probe) {
            assert_canonical_and_bounded(&value);
        }
    }
    assert_eq!(package.exact_bytes(), source_bytes);
    exercise_shared_ownership(&package, selector, descriptor, &before, &source_bytes);

    let command = descriptor.get(1).copied().unwrap_or(COMMAND_SET) % 4;
    let requested = requested_axes(descriptor);
    let edit = package
        .edit_body_table_hidden_axes(selector)
        .unwrap_or_else(|error| panic!("valid descriptor edit must start: {error}"));
    let staged = match command {
        COMMAND_SET => edit.set(requested.clone()),
        COMMAND_CLEAR => edit.clear(),
        COMMAND_RESET => edit.reset(),
        _ => edit.set(before.clone()),
    };
    let expected_noop = match command {
        COMMAND_SET => requested == before,
        COMMAND_CLEAR | COMMAND_RESET => before.is_empty(),
        _ => true,
    };
    let ownerless_nonempty_set =
        command == COMMAND_SET && before.is_empty() && !requested.is_empty();
    let commit = match staged.commit() {
        Ok(commit) => commit,
        Err(error) => {
            // A valid graph may refuse a changed lock/filter/pivot topology or
            // an absent-owner nonempty set, but every such refusal must be
            // atomic. Ordinary valid graphs must publish (or publish an exact
            // no-op), so this assertion keeps each input on a substantive
            // success or explicitly expected error branch.
            assert_eq!(package.exact_bytes(), source_bytes);
            let flags = descriptor.first().copied().unwrap_or_default();
            assert!(
                !expected_noop,
                "valid no-op edit unexpectedly rejected: {error}"
            );
            assert!(
                ownerless_nonempty_set || flags & 0x0c != 0,
                "valid descriptor edit unexpectedly rejected: {error}"
            );
            if ownerless_nonempty_set {
                assert_eq!(error, Error::UnsupportedDependency);
            } else if flags & 0x08 != 0 {
                assert_eq!(error, Error::TableLocked);
            } else {
                assert_eq!(error, Error::UnsupportedDependency);
            }
            black_box(error);
            return;
        },
    };

    let expected = match command {
        COMMAND_SET => requested,
        COMMAND_CLEAR | COMMAND_RESET => HiddenAxes::empty(),
        _ => before.clone(),
    };
    let target_bytes = commit.package().exact_bytes();
    let after = commit
        .package()
        .body_table_hidden_axes(selector)
        .unwrap_or_else(|error| panic!("committed hidden-axis read failed: {error}"));
    assert_eq!(after, expected);
    assert_canonical_and_bounded(&after);

    assert_eq!(commit.patch().is_noop(), expected_noop);
    if commit.patch().is_noop() {
        assert_eq!(target_bytes, source_bytes);
        assert!(!commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().touched_components(), 0);
        assert_eq!(commit.diagnostics().deleted_previews(), 0);
        assert!(!commit.diagnostics().full_reparse_performed());
    } else {
        assert_ne!(target_bytes, source_bytes);
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().touched_components(), 1);
        assert!(commit.diagnostics().full_reparse_performed());
        assert_preview_diagnostic(
            &source_bytes,
            &target_bytes,
            commit.diagnostics().deleted_previews(),
        );
    }

    assert_member_locality(
        &source_bytes,
        &target_bytes,
        selected_table_index(descriptor),
    );
    if !commit.patch().is_noop() {
        assert_existing_owner_wire_state_preserved(
            &source_bytes,
            &target_bytes,
            selected_table_index(descriptor),
            &after,
        );
    }
    assert_reopen_and_patch_invariants(
        &package,
        &commit,
        selector,
        &before,
        &after,
        &source_bytes,
        &target_bytes,
    );
}

fn exercise_shared_ownership(
    package: &Package,
    selector: BodyTableSelector<'static>,
    descriptor: &[u8],
    before: &HiddenAxes,
    source_bytes: &[u8],
) {
    // Keep an immutable Arc observer active while a COW fork stages an edit.
    // The observer must continue to see the exact source regardless of
    // whether the fork publishes a candidate or refuses the operation.
    let shared = Arc::new(package.clone());
    let observer = Arc::clone(&shared);
    let reader = thread::spawn(move || observer.body_table_hidden_axes(selector));

    let fork = package.clone();
    let command = descriptor.get(1).copied().unwrap_or(COMMAND_SET) % 4;
    let requested = requested_axes(descriptor);
    let result = fork
        .edit_body_table_hidden_axes(selector)
        .unwrap_or_else(|error| panic!("shared-owner edit must start: {error}"));
    let staged = match command {
        COMMAND_SET => result.set(requested.clone()),
        COMMAND_CLEAR => result.clear(),
        COMMAND_RESET => result.reset(),
        _ => result.set(before.clone()),
    };
    match staged.commit() {
        Ok(commit) => {
            let expected = match command {
                COMMAND_SET => requested,
                COMMAND_CLEAR | COMMAND_RESET => HiddenAxes::empty(),
                _ => before.clone(),
            };
            assert_eq!(
                commit
                    .package()
                    .body_table_hidden_axes(selector)
                    .unwrap_or_else(|error| panic!("shared candidate read failed: {error}")),
                expected
            );
            let candidate_bytes = commit.package().exact_bytes();
            if commit.patch().is_noop() {
                assert_eq!(candidate_bytes, source_bytes);
            } else {
                assert_ne!(candidate_bytes, source_bytes);
            }
            assert_eq!(fork.exact_bytes(), source_bytes);
            assert_eq!(commit.patch().inverse().inverse(), *commit.patch());
        },
        Err(error) => {
            assert_eq!(fork.exact_bytes(), source_bytes);
            black_box(error);
        },
    }
    let observed = reader
        .join()
        .expect("shared source reader must not panic")
        .unwrap_or_else(|error| panic!("shared source read failed: {error}"));
    assert_eq!(observed, *before);
    assert_eq!(shared.exact_bytes(), source_bytes);
}

fn assert_canonical_and_bounded(hidden: &HiddenAxes) {
    for pair in hidden.as_slice().windows(2) {
        assert!(pair[0] < pair[1], "hidden-axis positions are not canonical");
    }
    for axis in hidden.iter() {
        match axis {
            AxisIndex::Row(index) => assert!(index < MAX_TABLE_ROWS),
            AxisIndex::Column(index) => assert!(index < MAX_TABLE_COLUMNS),
        }
    }
}

fn assert_preview_diagnostic(source: &[u8], target: &[u8], deleted: usize) {
    let source_count = preview_count(source);
    let target_count = preview_count(target);
    assert_eq!(deleted, source_count.saturating_sub(target_count));
    assert!(deleted <= PREVIEWS.len());
}

fn selected_table_index(descriptor: &[u8]) -> usize {
    let flags = descriptor.first().copied().unwrap_or_default();
    let mode = descriptor.get(3).copied().unwrap_or_default() % MODE_COUNT;
    if mode != MODE_VALID || flags & 0x0e != 0 {
        return 0;
    }
    usize::from(descriptor.get(4).copied().unwrap_or_default() % 4 >= 2)
}

fn assert_member_locality(source: &[u8], target: &[u8], selected_index: usize) {
    let source_catalog = Catalog::from_bytes(source).expect("source catalog is valid");
    let target_catalog = Catalog::from_bytes(target).expect("target catalog is valid");

    for source_entry in source_catalog.iter() {
        let name = source_entry.name();
        if name == DOCUMENT_MEMBER || PREVIEWS.contains(&name) {
            continue;
        }
        let target_entry = target_catalog
            .iter()
            .find(|entry| entry.name() == name)
            .unwrap_or_else(|| panic!("member {name} disappeared during local rewrite"));
        assert_eq!(
            source_entry.data(),
            target_entry.data(),
            "member {name} changed"
        );
    }

    let source_archive = document_archive(source).expect("source document archive");
    let target_archive = document_archive(target).expect("target document archive");
    let selected_ids = [
        table_attachment(selected_index),
        table_drawable(selected_index),
        table_model(selected_index),
        table_uid_map(selected_index),
        table_formula_owner(selected_index),
        table_filter_set(selected_index, true),
        table_filter_set(selected_index, false),
        table_formula_object(selected_index, true),
        table_formula_object(selected_index, false),
    ];
    let source_unselected = source_archive
        .objects
        .iter()
        .filter(|object| {
            object
                .archive_info
                .identifier
                .is_some_and(|identifier| !selected_ids.contains(&identifier))
        })
        .count();
    let target_unselected = target_archive
        .objects
        .iter()
        .filter(|object| {
            object
                .archive_info
                .identifier
                .is_some_and(|identifier| !selected_ids.contains(&identifier))
        })
        .count();
    assert_eq!(
        source_unselected, target_unselected,
        "local rewrite added or removed an unselected archive object"
    );
    for source_object in source_archive.objects.iter().filter(|object| {
        object
            .archive_info
            .identifier
            .is_some_and(|identifier| !selected_ids.contains(&identifier))
    }) {
        let identifier = source_object
            .archive_info
            .identifier
            .expect("unselected archive object has an identifier");
        let target_object = target_archive
            .object(identifier)
            .unwrap_or_else(|| panic!("unselected object {identifier} disappeared"));
        assert!(
            source_object.same_content_ignoring_offsets(target_object),
            "unselected object {identifier} or its archive metadata changed"
        );
    }

    // The UID map, attachment, and formula/filter helpers are selected
    // dependencies of the focused graph, but this writer owns only the
    // model/info payloads.  Compare them explicitly so a successful edit
    // cannot smuggle a dependency mutation through the selected-object
    // allowance.
    for identifier in [
        table_attachment(selected_index),
        table_uid_map(selected_index),
        table_formula_owner(selected_index),
        table_filter_set(selected_index, true),
        table_filter_set(selected_index, false),
        table_formula_object(selected_index, true),
        table_formula_object(selected_index, false),
    ] {
        let source_object = source_archive.object(identifier);
        let target_object = target_archive.object(identifier);
        assert_eq!(
            source_object.is_some(),
            target_object.is_some(),
            "selected helper {identifier} presence changed"
        );
        if let (Some(source_object), Some(target_object)) = (source_object, target_object) {
            assert!(
                source_object.same_content_ignoring_offsets(target_object),
                "selected helper {identifier} or its archive metadata changed"
            );
        }
    }
}

fn assert_canonical_uid_map_message(archive: &Archive, index: usize) {
    let map = archive
        .object(table_uid_map(index))
        .unwrap_or_else(|| panic!("missing UID-map object for table {index}"));
    let canonical = map
        .messages
        .iter()
        .filter(|message| message.type_ == UID_MAP_MESSAGE_TYPE)
        .count();
    assert_eq!(
        canonical, 1,
        "table {index} must have exactly one canonical 6267 UID-map message"
    );
    assert!(
        map.messages.iter().all(|message| message.type_ != 6_005),
        "table {index} must not use the unrelated 6005 table-data message"
    );
}

fn assert_non_identity_uid_map(archive: &Archive, index: usize) {
    let map = archive
        .object(table_uid_map(index))
        .unwrap_or_else(|| panic!("missing non-identity UID-map object for table {index}"));
    let message = map
        .messages
        .iter()
        .find(|message| message.type_ == UID_MAP_MESSAGE_TYPE)
        .unwrap_or_else(|| panic!("missing canonical UID-map message for table {index}"));
    let decoded = tst::ColumnRowUidMapArchive::decode(message.data.as_slice())
        .unwrap_or_else(|error| panic!("table {index} UID-map payload is valid: {error}"));
    assert_ne!(
        decoded.column_uid_for_index,
        (0..TABLE_COLUMNS).collect::<Vec<_>>(),
        "table {index} non-identity column map collapsed to identity"
    );
    assert_ne!(
        decoded.row_uid_for_index,
        (0..TABLE_ROWS).collect::<Vec<_>>(),
        "table {index} non-identity row map collapsed to identity"
    );
    assert_eq!(decoded.column_uid_for_index, vec![2, 0, 3, 1]);
    assert_eq!(decoded.column_index_for_uid, vec![1, 3, 0, 2]);
    assert_eq!(decoded.row_uid_for_index, vec![3, 1, 0, 2]);
    assert_eq!(decoded.row_index_for_uid, vec![2, 1, 3, 0]);
    for (physical, stable) in decoded.column_uid_for_index.iter().copied().enumerate() {
        assert_eq!(
            decoded.column_index_for_uid[stable as usize],
            u32::try_from(physical).expect("fixture position fits")
        );
    }
    for (physical, stable) in decoded.row_uid_for_index.iter().copied().enumerate() {
        assert_eq!(
            decoded.row_index_for_uid[stable as usize],
            u32::try_from(physical).expect("fixture position fits")
        );
    }
}

fn assert_uid_map_route(package: &[u8], index: usize) {
    let model = model_payload_from_package(package, index)
        .unwrap_or_else(|error| panic!("table {index} model payload is present: {error}"));
    let model_fields = WireView::parse(&model)
        .unwrap_or_else(|error| panic!("table {index} model wire is valid: {error}"))
        .fields()
        .filter(|field| field.number() == TABLE_MODEL_UID_MAP_FIELD)
        .count();
    assert_eq!(
        model_fields, 1,
        "table {index} must retain exactly one model field-46 UID-map route"
    );

    let info = info_payload(package, index)
        .unwrap_or_else(|error| panic!("table {index} info payload is present: {error}"));
    let info_fields = WireView::parse(&info)
        .unwrap_or_else(|error| panic!("table {index} info wire is valid: {error}"))
        .fields()
        .filter(|field| field.number() == TABLE_INFO_VIEW_UIDS_FIELD)
        .count();
    if index == 0 {
        assert_eq!(
            info_fields, 0,
            "the first table must exercise model field-46 fallback"
        );
    } else {
        assert_eq!(
            info_fields, 1,
            "the second table must exercise table-info field-6 route"
        );
    }
}

fn raw_field(payload: &[u8], field_number: u32) -> RawWireField<'_> {
    let mut fields = RawWireFields::new(payload);
    loop {
        let field = fields
            .next()
            .unwrap_or_else(|error| panic!("unknown-field payload became malformed: {error}"));
        let Some(field) = field else {
            panic!("unknown field {field_number} was dropped");
        };
        if field.number() == field_number {
            return field;
        }
    }
}

fn assert_unknown_varint_field(payload: &[u8], field_number: u32, expected: u64) {
    let field = raw_field(payload, field_number);
    let value = decode_varint_from_bytes(field.payload())
        .unwrap_or_else(|error| panic!("unknown field {field_number} is not a varint: {error}"))
        .0;
    assert_eq!(value, expected, "unknown field {field_number} changed");
}

fn assert_unknown_scalar_fields(payload: &[u8], varint_field: u32, varint_value: u64) {
    assert_unknown_varint_field(payload, varint_field, varint_value);
    let fixed32 = raw_field(payload, UNKNOWN_FIXED32_FIELD);
    assert_eq!(fixed32.wire_type(), 5);
    assert_eq!(fixed32.key(), &[0xa5, 0x06]);
    assert_eq!(fixed32.payload(), UNKNOWN_FIXED32_VALUE);
    let fixed64 = raw_field(payload, UNKNOWN_FIXED64_FIELD);
    assert_eq!(fixed64.wire_type(), 1);
    assert_eq!(fixed64.key(), &[0xa9, 0x06]);
    assert_eq!(fixed64.payload(), UNKNOWN_FIXED64_VALUE);
    let length = raw_field(payload, UNKNOWN_LENGTH_FIELD);
    assert_eq!(length.wire_type(), 2);
    assert_eq!(length.payload(), UNKNOWN_LENGTH_VALUE);
}

fn assert_unknown_wire_fields(payload: &[u8], varint_field: u32, varint_value: u64) {
    assert_unknown_scalar_fields(payload, varint_field, varint_value);
    let group = raw_field(payload, UNKNOWN_GROUP_FIELD);
    assert!(group.is_group_start());
    assert_eq!(group.key(), &[0xbb, 0x06]);
    assert_eq!(group.group_payload(), Some(UNKNOWN_GROUP_VALUE));
}

fn unknown_raw_fields(payload: &[u8], field_numbers: &[u32]) -> Vec<Vec<u8>> {
    let mut fields = RawWireFields::new(payload);
    let mut result = Vec::new();
    loop {
        let field = fields
            .next()
            .unwrap_or_else(|error| panic!("unknown-field payload became malformed: {error}"));
        let Some(field) = field else {
            break;
        };
        if field_numbers.contains(&field.number()) {
            result.push(field.raw().to_vec());
        }
    }
    result
}

fn assert_unknown_wire_fields_exact(source: &[u8], target: &[u8], field_numbers: &[u32]) {
    assert_eq!(
        unknown_raw_fields(source, field_numbers),
        unknown_raw_fields(target, field_numbers),
        "unknown wire records changed during the focused rewrite"
    );
}

fn assert_nested_unknown_varint_field(
    payload: &[u8],
    outer_field: u32,
    nested_field: u32,
    expected: u64,
) {
    let nested = raw_field(payload, outer_field).payload().to_vec();
    assert_unknown_varint_field(&nested, nested_field, expected);
}

fn nested_unknown_raw_fields(
    payload: &[u8],
    outer_field: u32,
    field_numbers: &[u32],
) -> Vec<Vec<u8>> {
    let nested = raw_field(payload, outer_field).payload().to_vec();
    unknown_raw_fields(&nested, field_numbers)
}

fn assert_nested_unknown_wire_fields_exact(
    source: &[u8],
    target: &[u8],
    outer_field: u32,
    field_numbers: &[u32],
) {
    assert_eq!(
        nested_unknown_raw_fields(source, outer_field, field_numbers),
        nested_unknown_raw_fields(target, outer_field, field_numbers),
        "nested unknown wire records changed during the focused rewrite"
    );
}

fn assert_existing_owner_wire_state_preserved(
    source: &[u8],
    target: &[u8],
    index: usize,
    after: &HiddenAxes,
) {
    let source_archive = document_archive(source)
        .unwrap_or_else(|error| panic!("source document archive is present: {error}"));
    let uid_map = source_archive
        .object(table_uid_map(index))
        .and_then(|object| {
            object
                .messages
                .iter()
                .find(|message| message.type_ == UID_MAP_MESSAGE_TYPE)
        })
        .and_then(|message| tst::ColumnRowUidMapArchive::decode(message.data.as_slice()).ok())
        .unwrap_or_else(|| panic!("source table {index} UID map is present and decodable"));
    let source_model = model_payload_from_package(source, index)
        .unwrap_or_else(|error| panic!("source model payload is present: {error}"));
    let target_model = model_payload_from_package(target, index)
        .unwrap_or_else(|error| panic!("target model payload is present: {error}"));
    assert_unknown_scalar_fields(&source_model, UNKNOWN_MODEL_FIELD, UNKNOWN_MODEL_VALUE);
    assert_unknown_scalar_fields(&target_model, UNKNOWN_MODEL_FIELD, UNKNOWN_MODEL_VALUE);
    let top_level_unknown_fields = [
        UNKNOWN_MODEL_FIELD,
        UNKNOWN_FIXED32_FIELD,
        UNKNOWN_FIXED64_FIELD,
        UNKNOWN_LENGTH_FIELD,
    ];
    assert_unknown_wire_fields_exact(&source_model, &target_model, &top_level_unknown_fields);
    assert_nested_unknown_varint_field(
        &source_model,
        TABLE_MODEL_UID_MAP_FIELD,
        UNKNOWN_NESTED_REFERENCE_FIELD,
        UNKNOWN_NESTED_REFERENCE_VALUE,
    );
    assert_nested_unknown_wire_fields_exact(
        &source_model,
        &target_model,
        TABLE_MODEL_UID_MAP_FIELD,
        &[UNKNOWN_NESTED_REFERENCE_FIELD],
    );
    assert_nested_unknown_varint_field(
        &target_model,
        TABLE_MODEL_UID_MAP_FIELD,
        UNKNOWN_NESTED_REFERENCE_FIELD,
        UNKNOWN_NESTED_REFERENCE_VALUE,
    );

    let source_info = info_payload(source, index)
        .unwrap_or_else(|error| panic!("source table-info payload is present: {error}"));
    let target_info = info_payload(target, index)
        .unwrap_or_else(|error| panic!("target table-info payload is present: {error}"));
    assert_unknown_scalar_fields(&source_info, UNKNOWN_INFO_FIELD, UNKNOWN_INFO_VALUE);
    assert_unknown_scalar_fields(&target_info, UNKNOWN_INFO_FIELD, UNKNOWN_INFO_VALUE);
    let top_level_unknown_fields = [
        UNKNOWN_INFO_FIELD,
        UNKNOWN_FIXED32_FIELD,
        UNKNOWN_FIXED64_FIELD,
        UNKNOWN_LENGTH_FIELD,
    ];
    assert_unknown_wire_fields_exact(&source_info, &target_info, &top_level_unknown_fields);
    assert_nested_unknown_varint_field(
        &source_info,
        2,
        UNKNOWN_NESTED_REFERENCE_FIELD,
        UNKNOWN_NESTED_REFERENCE_VALUE,
    );
    assert_nested_unknown_wire_fields_exact(
        &source_info,
        &target_info,
        2,
        &[UNKNOWN_NESTED_REFERENCE_FIELD],
    );
    assert_nested_unknown_varint_field(
        &target_info,
        2,
        UNKNOWN_NESTED_REFERENCE_FIELD,
        UNKNOWN_NESTED_REFERENCE_VALUE,
    );
    assert_nested_unknown_varint_field(
        &source_info,
        8,
        UNKNOWN_NESTED_UUID_FIELD,
        UNKNOWN_NESTED_UUID_VALUE,
    );
    assert_nested_unknown_wire_fields_exact(
        &source_info,
        &target_info,
        8,
        &[UNKNOWN_NESTED_UUID_FIELD],
    );
    assert_nested_unknown_varint_field(
        &target_info,
        8,
        UNKNOWN_NESTED_UUID_FIELD,
        UNKNOWN_NESTED_UUID_VALUE,
    );

    let source_owner = hidden_owner_payload_from_package(source, index)
        .unwrap_or_else(|error| panic!("changed source must have an existing owner: {error}"));
    let target_owner = hidden_owner_payload_from_package(target, index)
        .unwrap_or_else(|error| panic!("changed target must retain the owner: {error}"));
    assert_unknown_wire_fields(&source_owner, UNKNOWN_OWNER_FIELD, UNKNOWN_OWNER_VALUE);
    assert_unknown_wire_fields(&target_owner, UNKNOWN_OWNER_FIELD, UNKNOWN_OWNER_VALUE);
    let top_level_unknown_fields = [
        UNKNOWN_OWNER_FIELD,
        UNKNOWN_FIXED32_FIELD,
        UNKNOWN_FIXED64_FIELD,
        UNKNOWN_LENGTH_FIELD,
        UNKNOWN_GROUP_FIELD,
    ];
    assert_unknown_wire_fields_exact(&source_owner, &target_owner, &top_level_unknown_fields);

    let source_owner = tst::HiddenStatesOwnerArchive::decode(source_owner.as_slice())
        .unwrap_or_else(|error| panic!("source hidden owner is decodable: {error}"));
    let target_owner = tst::HiddenStatesOwnerArchive::decode(target_owner.as_slice())
        .unwrap_or_else(|error| panic!("target hidden owner is decodable: {error}"));
    let source_active = source_owner
        .hidden_states
        .iter()
        .find(|state| state.hidden_states_uid == source_owner.owner_uid)
        .or_else(|| source_owner.hidden_states.first())
        .unwrap_or_else(|| panic!("source hidden owner has no active state"));
    let target_active = target_owner
        .hidden_states
        .iter()
        .find(|state| state.hidden_states_uid == source_active.hidden_states_uid)
        .unwrap_or_else(|| panic!("target hidden owner lost the active state"));

    for (is_row, (source_extent, target_extent)) in [
        (
            true,
            (
                &source_active.row_hidden_state_extent,
                &target_active.row_hidden_state_extent,
            ),
        ),
        (
            false,
            (
                &source_active.column_hidden_state_extent,
                &target_active.column_hidden_state_extent,
            ),
        ),
    ] {
        for source_state in &source_extent.base_hidden_states {
            // User-hidden entries are intentionally replaced by the request.
            // Filtered and pivot markers are independent native state and
            // must survive, while an explicit visible marker is preserved
            // only when that position remains visible after the edit.
            // The UID map stores stable UUID order separately from physical
            // axis order. Resolve through its inverse permutation so a
            // non-identity map cannot make a visible marker look selected.
            let physical_index = if is_row {
                uid_map
                    .sorted_row_uids
                    .iter()
                    .position(|uid| *uid == source_state.row_or_column_uid)
                    .and_then(|stable| uid_map.row_index_for_uid.get(stable))
                    .and_then(|physical| usize::try_from(*physical).ok())
            } else {
                uid_map
                    .sorted_column_uids
                    .iter()
                    .position(|uid| *uid == source_state.row_or_column_uid)
                    .and_then(|stable| uid_map.column_index_for_uid.get(stable))
                    .and_then(|physical| usize::try_from(*physical).ok())
            };
            let requested_hidden = physical_index.is_some_and(|physical| {
                after.contains(if is_row {
                    AxisIndex::row(physical)
                } else {
                    AxisIndex::column(physical)
                })
            });
            let target_state = target_extent
                .base_hidden_states
                .iter()
                .find(|target_state| {
                    target_state.row_or_column_uid == source_state.row_or_column_uid
                });
            let has_non_user_state =
                source_state.filtered.is_some() || source_state.pivot_hidden.is_some();
            if has_non_user_state {
                let target_state = target_state.unwrap_or_else(|| {
                    panic!(
                        "non-user native state {:?} was dropped",
                        source_state.row_or_column_uid
                    )
                });
                assert_eq!(target_state.filtered, source_state.filtered);
                assert_eq!(target_state.pivot_hidden, source_state.pivot_hidden);
                if source_state.user_hidden == Some(false) && !requested_hidden {
                    assert_eq!(target_state.user_hidden, Some(false));
                }
            } else if source_state.user_hidden == Some(false) && !requested_hidden {
                let target_state = target_state.unwrap_or_else(|| {
                    panic!(
                        "explicit visible native state {:?} was dropped",
                        source_state.row_or_column_uid
                    )
                });
                assert_eq!(target_state.user_hidden, Some(false));
            }
        }
    }
}

fn assert_reopen_and_patch_invariants(
    source: &Package,
    commit: &litchi_pages::BodyTableHiddenAxesCommit,
    selector: BodyTableSelector<'static>,
    before: &HiddenAxes,
    after: &HiddenAxes,
    source_bytes: &[u8],
    target_bytes: &[u8],
) {
    let reopened = Package::from_bytes(target_bytes)
        .unwrap_or_else(|error| panic!("candidate did not reopen: {error}"));
    assert_eq!(
        reopened
            .body_table_hidden_axes(selector)
            .unwrap_or_else(|error| panic!("reopened candidate read failed: {error}")),
        *after
    );

    assert_eq!(commit.patch().inverse().inverse(), *commit.patch());

    let applied = source
        .apply_body_table_hidden_axes(commit.patch())
        .unwrap_or_else(|error| panic!("fresh patch did not apply: {error}"));
    assert_eq!(applied.package().exact_bytes(), target_bytes);
    assert_eq!(
        applied.package().body_table_hidden_axes(selector).unwrap(),
        *after
    );

    let target_package = commit.package();
    let target_before_conflict = target_package.exact_bytes();
    if commit.patch().is_noop() {
        // A no-op patch is intentionally replayable as an exact identity;
        // conflict fencing applies only to a changed source-bound patch.
        let replayed = target_package
            .apply_body_table_hidden_axes(commit.patch())
            .unwrap_or_else(|error| panic!("no-op patch replay failed: {error}"));
        assert_eq!(replayed.package().exact_bytes(), target_before_conflict);
        assert_eq!(
            replayed
                .package()
                .body_table_hidden_axes(selector)
                .unwrap_or_else(|error| panic!("no-op replay read failed: {error}")),
            *after
        );
        let restored = target_package
            .apply_body_table_hidden_axes(&commit.patch().inverse())
            .unwrap_or_else(|error| panic!("no-op inverse patch did not apply: {error}"));
        assert_eq!(restored.package().exact_bytes(), source_bytes);
        assert_eq!(
            restored
                .package()
                .body_table_hidden_axes(selector)
                .unwrap_or_else(|error| panic!("no-op inverse read failed: {error}")),
            *before
        );
    } else {
        assert!(
            target_package
                .apply_body_table_hidden_axes(commit.patch())
                .is_err(),
            "reapplying a source-bound patch must conflict"
        );
        assert_eq!(target_package.exact_bytes(), target_before_conflict);

        let restored = target_package
            .apply_body_table_hidden_axes(&commit.patch().inverse())
            .unwrap_or_else(|error| panic!("inverse patch did not apply: {error}"));
        assert_eq!(restored.package().exact_bytes(), source_bytes);
        assert_eq!(
            restored
                .package()
                .body_table_hidden_axes(selector)
                .unwrap_or_else(|error| panic!("inverse semantic read failed: {error}")),
            *before
        );
    }
}

fn exercise_malformed(source: &[u8], descriptor: &[u8]) {
    let package = match Package::from_bytes(source) {
        Ok(package) => package,
        Err(error) => {
            // Physical ingress is allowed to reject a malformed graph/wire
            // before the focused owner; this is an intentional error branch.
            assert_ne!(
                descriptor.get(3).copied().unwrap_or_default() % MODE_COUNT,
                MODE_VALID
            );
            black_box(error);
            return;
        },
    };
    let before = package.exact_bytes();
    let selector = malformed_selector(descriptor);
    let read = package.body_table_hidden_axes(selector);
    assert!(
        read.is_err(),
        "malformed fixture mode {} was readable",
        descriptor.get(3).copied().unwrap_or_default() % MODE_COUNT
    );
    let read_error = read.expect_err("malformed read error was lost");
    assert_precise_malformed_error(
        descriptor.get(3).copied().unwrap_or_default() % MODE_COUNT,
        read_error,
    );
    black_box(read_error);

    if descriptor.get(3).copied().unwrap_or_default() % MODE_COUNT == MODE_DUPLICATE_NAME {
        let before_name_probe = package.exact_bytes();
        let name_selector = BodyTableSelector::name("Revenue");
        assert!(
            package.body_table_hidden_axes(name_selector).is_err(),
            "duplicate table names must not resolve by name"
        );
        let name_edit = package
            .edit_body_table_hidden_axes(name_selector)
            .and_then(|edit| edit.set(HiddenAxes::empty()).commit());
        assert!(
            name_edit.is_err(),
            "duplicate table names must reject edits"
        );
        assert_eq!(package.exact_bytes(), before_name_probe);
    }
    let requested = requested_axes(descriptor);
    let result = package
        .edit_body_table_hidden_axes(selector)
        .and_then(|edit| {
            let command = descriptor.get(1).copied().unwrap_or_default() % 4;
            let edit = match command {
                COMMAND_CLEAR => edit.clear(),
                COMMAND_RESET => edit.reset(),
                _ => edit.set(requested),
            };
            edit.commit()
        });
    assert!(
        result.is_err(),
        "malformed fixture mode {} was accepted",
        descriptor.get(3).copied().unwrap_or_default() % MODE_COUNT
    );
    let edit_error = result.expect_err("malformed edit error was lost");
    assert_precise_malformed_error(
        descriptor.get(3).copied().unwrap_or_default() % MODE_COUNT,
        edit_error,
    );
    black_box(edit_error);
    assert_eq!(package.exact_bytes(), before);
}

fn assert_precise_malformed_error(mode: u8, error: Error) {
    if mode == MODE_COUNT_OVERFLOW {
        assert!(
            matches!(
                error,
                Error::LimitExceeded {
                    kind: BodyTableHiddenAxesLimitKind::WireFields,
                    ..
                }
            ),
            "count-overflow fixture returned an imprecise error: {error:?}"
        );
        return;
    }
    let expected = match mode {
        MODE_DUPLICATE_NAME => Error::AmbiguousTableName,
        MODE_FILTER_METADATA => Error::UnsupportedDependency,
        _ => Error::InvalidSource,
    };
    assert_eq!(
        error, expected,
        "malformed fixture mode {mode} returned an imprecise error"
    );
}

fn malformed_selector(descriptor: &[u8]) -> BodyTableSelector<'static> {
    match descriptor.get(3).copied().unwrap_or_default() % MODE_COUNT {
        // The duplicate-name mutation targets the second model, so a name
        // lookup is the intended ambiguous-selector error path.
        MODE_DUPLICATE_NAME => BodyTableSelector::name("Revenue"),
        // The shared-owner mutation redirects table two's dependency into
        // table one; select that table so the inbound alias is exercised.
        MODE_SHARED_OWNER => BodyTableSelector::index(1),
        // The info-field route exists only on the second table.  The other
        // metadata probes deliberately target the selected first table.
        MODE_INFO_MAP_METADATA_MISSING => BodyTableSelector::index(1),
        _ => primary_selector(descriptor),
    }
}

fn exercise_data_dependent_limits(source: &[u8], descriptor: &[u8]) {
    let maximum = limit_maximum(source.len(), descriptor);
    let source_length = u64::try_from(source.len()).unwrap_or(u64::MAX);
    assert_fixture_resource_ceiling(source_length);
    let limits = Limits::new(
        maximum,
        64,
        FIXTURE_ARCHIVE_CEILING,
        FIXTURE_ARCHIVE_CEILING,
        FIXTURE_STREAM_CEILING,
    )
    .unwrap_or_else(|error| panic!("bounded limit profile must be valid: {error}"));
    let package = match Package::from_bytes_with_limits(source, limits) {
        Ok(package) => {
            assert!(
                maximum >= source_length,
                "a package admitted above its physical byte ceiling"
            );
            package
        },
        Err(error) => {
            assert!(
                maximum < source_length,
                "a package rejected despite an admitted physical byte ceiling: {error}"
            );
            assert_input_limit_error(error, maximum, source_length);
            return;
        },
    };
    let before = package.exact_bytes();
    let selector = primary_selector(descriptor);
    match package.body_table_hidden_axes(selector) {
        Ok(value) => assert_canonical_and_bounded(&value),
        Err(error) => assert_resource_error(error),
    }
    let result = package
        .edit_body_table_hidden_axes(selector)
        .and_then(|edit| edit.set(requested_axes(descriptor)).commit());
    match result {
        Ok(commit) => {
            let target = commit.package().exact_bytes();
            assert!(
                u64::try_from(target.len()).unwrap_or(u64::MAX) <= maximum,
                "successful bounded edit exceeded its package-byte ceiling"
            );
            if commit.patch().is_noop() {
                assert_eq!(target, before);
            } else {
                assert_ne!(target, before);
            }
        },
        Err(error) => {
            assert_eq!(package.exact_bytes(), before);
            assert_resource_or_expected_edit_error(error);
        },
    }
}

fn assert_input_limit_error(error: PackageError, maximum: u64, source_length: u64) {
    let is_expected = matches!(
        &error,
        PackageError::Archive(litchi_iwa_archive::Error::Limit {
            kind: litchi_iwa_archive::LimitKind::InputBytes,
            observed,
            maximum: actual_maximum,
        }) if *observed == source_length && *actual_maximum == maximum
    );
    assert!(
        is_expected,
        "a tight physical ingress returned an unexpected error: {error:?}"
    );
}

fn exercise_limit_boundaries(source: &[u8]) {
    let length = u64::try_from(source.len()).unwrap_or(u64::MAX);
    assert_fixture_resource_ceiling(length);
    let max_plus_one = length.saturating_add(1);
    let mut saw_max_plus_one = false;
    for maximum in [length.saturating_sub(1).max(1), length, max_plus_one] {
        let limits = Limits::new(
            maximum,
            64,
            FIXTURE_ARCHIVE_CEILING,
            FIXTURE_ARCHIVE_CEILING,
            FIXTURE_STREAM_CEILING,
        )
        .unwrap_or_else(|error| panic!("limit-boundary profile must be valid: {error}"));
        let package = match Package::from_bytes_with_limits(source, limits) {
            Ok(package) => {
                assert!(
                    maximum >= length,
                    "the package crossed a limit below its source length"
                );
                if maximum == max_plus_one {
                    saw_max_plus_one = true;
                }
                package
            },
            Err(error) => {
                assert!(
                    maximum < length,
                    "the package rejected a limit at or above its source length: {error}"
                );
                assert_input_limit_error(error, maximum, length);
                continue;
            },
        };
        let before = package.exact_bytes();
        match package.body_table_hidden_axes(0usize) {
            Ok(value) => assert_canonical_and_bounded(&value),
            Err(error) => assert_resource_error(error),
        }
        let result = package
            .edit_body_table_hidden_axes(0usize)
            .and_then(|edit| edit.set(HiddenAxes::empty()).commit());
        match result {
            Ok(commit) => {
                let target = commit.package().exact_bytes();
                assert!(
                    target.len() <= usize::try_from(maximum).unwrap_or(usize::MAX),
                    "successful limit-boundary edit exceeded its package-byte ceiling"
                );
                if commit.patch().is_noop() {
                    assert_eq!(target, before);
                } else {
                    assert_ne!(target, before);
                }
            },
            Err(error) => {
                assert_eq!(package.exact_bytes(), before);
                assert_resource_or_expected_edit_error(error);
            },
        }
    }
    assert!(
        saw_max_plus_one,
        "the exact source-length-plus-one package limit must admit ingress"
    );
}

fn assert_fixture_resource_ceiling(source_length: u64) {
    assert!(
        source_length < FIXTURE_ARCHIVE_CEILING,
        "the bounded fixture unexpectedly exceeds its subordinate archive ceiling"
    );
    assert!(
        FIXTURE_STREAM_CEILING <= Limits::MAX_IWA_STREAM_BYTES,
        "the bounded fixture stream ceiling exceeds the format hard limit"
    );
}

fn exercise_nested_codec_limits() {
    let (owner, _) = hidden_owner_payload(
        0,
        TableOptions {
            user_hidden: true,
            filtered_or_pivot: true,
            ..TableOptions::default()
        },
    );
    let payload = owner.encode_to_vec();
    // `for_source` is intentionally conservative for a caller that does not
    // know the shape of a payload.  This fixture is a trusted bounded owner
    // graph with nested extent/state records, so size the baseline from the
    // same finite accounting profile used by the codec's own tests.  The
    // individual one-unit ceilings below still prove that every accounting
    // dimension is enforced.
    let options = nested_codec_options(&payload);
    let (decoded, report) = hidden_codec::decode_hidden_states_owner_with_report(&payload, options)
        .unwrap_or_else(|error| panic!("bounded hidden-state codec baseline failed: {error}"));
    assert!(report.fields() > 0, "codec field accounting was vacuous");
    assert!(report.work_bytes() > 0, "codec work accounting was vacuous");
    assert!(
        report.allocations() > 0,
        "codec allocation accounting was vacuous"
    );
    assert!(
        report.retained_bytes() > 0,
        "codec retained accounting was vacuous"
    );
    assert!(
        report.scratch_bytes() > 0,
        "codec scratch accounting was vacuous"
    );

    let prepared = hidden_codec::prepare_hidden_states_owner_rewrite(&payload, &decoded, options)
        .unwrap_or_else(|error| panic!("bounded hidden-state rewrite preparation failed: {error}"));
    let requirements = prepared.execution_requirements();
    let output = prepared
        .execute(hidden_codec::RewriteExecutionLimits::exact(requirements))
        .unwrap_or_else(|error| panic!("exact hidden-state rewrite budget failed: {error}"));
    let (round_trip, _) = hidden_codec::decode_hidden_states_owner_with_report(
        output.bytes(),
        nested_codec_options(output.bytes()),
    )
    .unwrap_or_else(|error| panic!("rewritten hidden-state payload is not decodable: {error}"));
    assert_eq!(round_trip, decoded, "codec rewrite changed semantic state");

    for (label, limit, expected) in [
        (
            "fields",
            report.fields().saturating_sub(1),
            hidden_codec::DecodeLimit::Fields {
                observed: report.fields(),
                maximum: report.fields().saturating_sub(1),
            },
        ),
        (
            "work",
            report.work_bytes().saturating_sub(1),
            hidden_codec::DecodeLimit::WorkBytes {
                observed: report.work_bytes(),
                maximum: report.work_bytes().saturating_sub(1),
            },
        ),
        (
            "allocations",
            report.allocations().saturating_sub(1),
            hidden_codec::DecodeLimit::Allocations {
                observed: report.allocations(),
                maximum: report.allocations().saturating_sub(1),
            },
        ),
        (
            "retained",
            report.retained_bytes().saturating_sub(1),
            hidden_codec::DecodeLimit::RetainedBytes {
                observed: report.retained_bytes(),
                maximum: report.retained_bytes().saturating_sub(1),
            },
        ),
        (
            "scratch",
            report.scratch_bytes().saturating_sub(1),
            hidden_codec::DecodeLimit::ScratchBytes {
                observed: report.scratch_bytes(),
                maximum: report.scratch_bytes().saturating_sub(1),
            },
        ),
    ] {
        let constrained = match label {
            "fields" => options.with_max_fields(limit),
            "work" => options.with_max_work_bytes(limit),
            "allocations" => options.with_max_allocations(limit),
            "retained" => options.with_max_retained_bytes(limit),
            "scratch" => options.with_max_scratch_bytes(limit),
            _ => unreachable!("codec limit label is exhaustive"),
        };
        let error = hidden_codec::decode_hidden_states_owner(&payload, constrained)
            .expect_err("a one-unit nested codec ceiling must be observable");
        assert_eq!(
            error.resource_limit(),
            Some(expected),
            "nested codec {label} ceiling was not enforced"
        );
    }

    for (label, limit, expected) in [
        (
            "fields",
            requirements.fields().saturating_sub(1),
            hidden_codec::DecodeLimit::Fields {
                observed: requirements.fields(),
                maximum: requirements.fields().saturating_sub(1),
            },
        ),
        (
            "work",
            requirements.work_bytes().saturating_sub(1),
            hidden_codec::DecodeLimit::WorkBytes {
                observed: requirements.work_bytes(),
                maximum: requirements.work_bytes().saturating_sub(1),
            },
        ),
        (
            "allocations",
            requirements.allocations().saturating_sub(1),
            hidden_codec::DecodeLimit::Allocations {
                observed: requirements.allocations(),
                maximum: requirements.allocations().saturating_sub(1),
            },
        ),
        (
            "retained",
            requirements.retained_bytes().saturating_sub(1),
            hidden_codec::DecodeLimit::RetainedBytes {
                observed: requirements.retained_bytes(),
                maximum: requirements.retained_bytes().saturating_sub(1),
            },
        ),
        (
            "scratch",
            requirements.scratch_bytes().saturating_sub(1),
            hidden_codec::DecodeLimit::ScratchBytes {
                observed: requirements.scratch_bytes(),
                maximum: requirements.scratch_bytes().saturating_sub(1),
            },
        ),
    ] {
        let prepared =
            hidden_codec::prepare_hidden_states_owner_rewrite(&payload, &decoded, options)
                .unwrap_or_else(|error| {
                    panic!("nested codec {label} rewrite preparation failed: {error}")
                });
        let constrained = match label {
            "fields" => hidden_codec::RewriteExecutionLimits::unrestricted().with_fields(limit),
            "work" => hidden_codec::RewriteExecutionLimits::unrestricted().with_work_bytes(limit),
            "allocations" => {
                hidden_codec::RewriteExecutionLimits::unrestricted().with_allocations(limit)
            },
            "retained" => {
                hidden_codec::RewriteExecutionLimits::unrestricted().with_retained_bytes(limit)
            },
            "scratch" => {
                hidden_codec::RewriteExecutionLimits::unrestricted().with_scratch_bytes(limit)
            },
            _ => unreachable!("codec limit label is exhaustive"),
        };
        let error = prepared
            .execute(constrained)
            .expect_err("a one-unit rewrite ceiling must be observable");
        assert_eq!(
            error.resource_limit(),
            Some(expected),
            "nested codec rewrite {label} ceiling was not enforced"
        );
    }
}

fn nested_codec_options(source: &[u8]) -> hidden_codec::DecodeOptions {
    hidden_codec::DecodeOptions::new(
        source.len().max(1),
        source.len().saturating_mul(2).max(1),
        source.len().saturating_mul(32).max(64),
        source.len().saturating_mul(128).max(128),
        16,
        128,
    )
    .with_max_allocations(4_096)
    .with_max_retained_bytes(source.len().saturating_mul(8).max(1))
    .with_max_scratch_bytes(source.len().saturating_mul(128).max(1))
}

fn assert_resource_error(error: Error) {
    assert!(
        matches!(
            error,
            Error::LimitExceeded { .. } | Error::Allocation { .. }
        ),
        "valid bounded read returned an unexpected error: {error:?}"
    );
}

fn assert_resource_or_expected_edit_error(error: Error) {
    assert!(
        matches!(
            error,
            Error::LimitExceeded { .. }
                | Error::Allocation { .. }
                | Error::TableLocked
                | Error::UnsupportedDependency
                | Error::UnsupportedSource
        ),
        "valid bounded edit returned an unexpected error: {error:?}"
    );
}

fn unknown_axis_uuid(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model(package, index, |model| {
        if let Some(owner) = model.hidden_states_owner.as_mut()
            && let Some(active) = owner.hidden_states.first_mut()
            && let Some(state) = active
                .row_hidden_state_extent
                .base_hidden_states
                .first_mut()
        {
            state.row_or_column_uid = uuid(0xffff, 0xffff);
        }
    })
}

fn wrong_message_type(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let model = archive
            .object_mut(table_model(index))
            .ok_or("missing model")?;
        let message = model
            .messages
            .iter_mut()
            .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or("missing model message")?;
        message.type_ = TABLE_MODEL_MESSAGE_TYPE + 100;
        Ok(())
    })
}

fn dangling_reference(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model(package, index, |model| {
        model.base_column_row_uids = Some(reference(0xdead_beef));
    })
}

fn cross_component_alias(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model(package, index, |model| {
        // Point the UID-map slot at a hidden-state formula owner.  The
        // identifier exists, but its component/type is not a UID map.
        model.base_column_row_uids = Some(reference(table_formula_object(index, true)));
    })
}

fn axis_bounds(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model(package, index, |model| {
        // Leave the UID map at its original four-row cardinality while
        // shrinking the model to three rows.  The graph must reject the
        // inconsistent bound before an edit can address row three.
        model.number_of_rows = TABLE_ROWS.saturating_sub(1);
        model.number_of_hidden_rows = Some(TABLE_ROWS.saturating_add(1));
        model.number_of_hidden_columns = Some(TABLE_COLUMNS.saturating_add(1));
    })
}

fn count_overflow(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model(package, index, |model| {
        model.number_of_rows = u32::MAX;
        model.number_of_columns = u32::MAX;
        model.number_of_hidden_rows = Some(u32::MAX);
        model.number_of_hidden_columns = Some(u32::MAX);
    })
}

fn replace_model_field(
    package: &[u8],
    index: usize,
    field: u32,
    replacement: &[u8],
) -> TestResult<Vec<u8>> {
    let payload = model_payload_from_package(package, index)?;
    let rewritten = replace_length_field(&payload, field, replacement)?;
    rewrite_document_archive(package, |archive| {
        let model = archive
            .object_mut(table_model(index))
            .ok_or("missing model")?;
        let message = model
            .messages
            .iter_mut()
            .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or("missing model message")?;
        message.data = rewritten;
        Ok(())
    })
}
