// Strict source-backed integration coverage for Pages body-table hidden axes.
//
// The package fixture below is intentionally assembled as a native IWA graph:
// a body table attachment owns a table-info record, the table-info record
// points at a model and a column/row UID map, and the model owns hidden-state
// extents through formula-owner records.  The semantic value is therefore
// read through the same ownership proof used by a real package.  Several
// fields are appended with literal wire helpers after protobuf construction;
// this keeps unknown-field and malformed-wire checks independent of the
// focused hidden-axis writer.

use std::error::Error as StdError;
use std::sync::Arc;
use std::thread;

use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_common::{
    decode_varint_from_bytes,
    wire::{RawWireField, RawWireFields, WireView},
};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, FieldPath, RawMessage, SnappyStream};
use litchi_iwa_protos::{tp, tsa, tsce, tsd, tsp, tst, tswp};
use litchi_pages::table::hidden_axes::{AxisIndex, HiddenAxes};
use litchi_pages::{
    BodyTableHiddenAxesError as Error, BodyTableHiddenAxesLimitKind, BodyTableSelector, Limits,
    Package,
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
// Native Pages table UID maps are carried by the 6267 archive message.  The
// nearby 6005 family is a Numbers table-data list, not this map.
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
// These names resemble preview assets but are not canonical lifecycle
// authorities.  A focused hidden-axis rewrite must retain them byte-for-byte.
const NONCANONICAL_PREVIEWS: [&str; 2] = ["Preview/preview.jpg", "preview.jpeg"];

// Unknown fields deliberately sit outside every selected native projection.
const UNKNOWN_MODEL_FIELD: u32 = 99;
const UNKNOWN_MODEL_VALUE: u64 = 0xfeed_beef;
const UNKNOWN_INFO_FIELD: u32 = 98;
const UNKNOWN_INFO_VALUE: u64 = 0xcafe_babe;
const UNKNOWN_OWNER_FIELD: u32 = 97;
const UNKNOWN_OWNER_VALUE: u64 = 0xabad_1dea;
const UNKNOWN_FORMULA_OWNER_FIELD: u32 = 96;
const UNKNOWN_FORMULA_OWNER_VALUE: u64 = 0xfeed_cafe;
const UNKNOWN_FILTER_FIELD: u32 = 95;
const UNKNOWN_FILTER_VALUE: u64 = 0xface_b00c;
const UNKNOWN_EXTENT_FIELD: u32 = 94;
const UNKNOWN_EXTENT_VALUE: u64 = 0xbeef_cafe;
const UNKNOWN_HIDDEN_FORMULA_FIELD: u32 = 93;
const UNKNOWN_HIDDEN_FORMULA_VALUE: u64 = 0xabcd_1234;
const UNKNOWN_FIXED64_FIELD: u32 = 91;
const UNKNOWN_FIXED32_FIELD: u32 = 92;
const UNKNOWN_GROUP_FIELD: u32 = 90;
// Keep this outside the known fields of every nested TST fixture message;
// unlike a generated prost round-trip, the source-owned wire helper retains
// its opaque length-delimited payload.
const UNKNOWN_LENGTH_FIELD: u32 = 120;

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

fn is_package_limit_error(error: &litchi_pages::PackageError) -> bool {
    matches!(
        error,
        litchi_pages::PackageError::Archive(litchi_iwa_archive::Error::Limit { .. })
            | litchi_pages::PackageError::Archive(litchi_iwa_archive::Error::Iwa(
                litchi_iwa_core::Error::Limit { .. },
            ))
            | litchi_pages::PackageError::ObjectLimit { .. }
            | litchi_pages::PackageError::PayloadLimit { .. }
    )
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
    /// Emit optional count scalars with an explicit zero value.  Presence is
    /// meaningful in the native model and must survive a rewrite.
    explicit_zero_counts: bool,
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn external_reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        deprecated_is_external: Some(true),
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

// Frame a test field independently: the generic append helper scans existing
// output and rejects groups, while these fixtures intentionally retain them.
fn append_message_field(output: &mut Vec<u8>, number: u32, payload: &[u8]) -> TestResult {
    let mut field = Vec::new();
    litchi_iwa_common::wire::append_length_delimited_field(&mut field, number, payload)?;
    output.extend_from_slice(&field);
    Ok(())
}

fn raw_wire_fields(source: &[u8]) -> TestResult<Vec<RawWireField<'_>>> {
    let mut scanner = RawWireFields::new(source);
    let mut fields = Vec::new();
    while let Some(field) = scanner.next()? {
        fields.push(field);
    }
    Ok(fields)
}

fn insert_raw_after_field(source: &[u8], after_field: u32, unknown: &[u8]) -> TestResult<Vec<u8>> {
    let fields = raw_wire_fields(source)?;
    let mut offset = 0usize;
    for field in fields {
        let end = offset
            .checked_add(field.raw().len())
            .ok_or("wire field offset overflow")?;
        if field.number() == after_field {
            let mut result = Vec::with_capacity(
                source
                    .len()
                    .checked_add(unknown.len())
                    .ok_or("wire insertion length overflow")?,
            );
            result.extend_from_slice(&source[..end]);
            result.extend_from_slice(unknown);
            result.extend_from_slice(&source[end..]);
            return Ok(result);
        }
        offset = end;
    }
    Err(format!("missing field {after_field} for unknown insertion").into())
}

fn insert_unknown_varint_after_field(
    source: &[u8],
    after_field: u32,
    field_number: u32,
    value: u64,
) -> TestResult<Vec<u8>> {
    let mut unknown = Vec::new();
    litchi_iwa_common::wire::append_varint_field(&mut unknown, field_number, value)?;
    insert_raw_after_field(source, after_field, &unknown)
}

fn insert_raw_after_nested_path(
    source: &[u8],
    path: &[u32],
    after_field: u32,
    unknown: &[u8],
) -> TestResult<Vec<u8>> {
    if path.is_empty() {
        return insert_raw_after_field(source, after_field, unknown);
    }
    let fields = raw_wire_fields(source)?;
    let mut offset = 0usize;
    let mut result = Vec::with_capacity(source.len());
    let mut changed = false;
    for field in fields {
        let end = offset
            .checked_add(field.raw().len())
            .ok_or("wire field offset overflow")?;
        if !changed && field.number() == path[0] && field.wire_type() == 2 {
            let nested =
                insert_raw_after_nested_path(field.payload(), &path[1..], after_field, unknown)?;
            append_message_field(&mut result, field.number(), &nested)?;
            changed = true;
        } else {
            result.extend_from_slice(field.raw());
        }
        offset = end;
    }
    if !changed {
        return Err(format!("missing nested field path {path:?}").into());
    }
    Ok(result)
}

fn insert_unknown_after_nested_path(
    source: &[u8],
    path: &[u32],
    after_field: u32,
    field_number: u32,
    value: u64,
) -> TestResult<Vec<u8>> {
    let mut unknown = Vec::new();
    litchi_iwa_common::wire::append_varint_field(&mut unknown, field_number, value)?;
    insert_raw_after_nested_path(source, path, after_field, &unknown)
}

fn unknown_fixed32_wire() -> [u8; 6] {
    // field 92, fixed32 wire type (5), followed by four arbitrary bytes
    [0xe5, 0x05, 0x01, 0x23, 0x45, 0x67]
}

fn unknown_fixed64_wire() -> [u8; 10] {
    // field 91, fixed64 wire type (1), followed by eight arbitrary bytes
    [0xd9, 0x05, 0x89, 0xab, 0xcd, 0xef, 0x01, 0x23, 0x45, 0x67]
}

fn unknown_group_wire() -> [u8; 7] {
    // A balanced unknown group (field 90) with one unknown varint member.
    [0xd3, 0x05, 0xc8, 0x05, 0x01, 0xd4, 0x05]
}

fn unknown_length_delimited_wire() -> TestResult<Vec<u8>> {
    let mut raw = Vec::new();
    append_message_field(&mut raw, UNKNOWN_LENGTH_FIELD, b"opaque-source-owned-bytes")?;
    Ok(raw)
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

fn uid_map_payload(index: usize) -> Vec<u8> {
    let (rows, columns) = table_uids(index);
    tst::ColumnRowUidMapArchive {
        sorted_column_uids: columns,
        column_index_for_uid: (0..TABLE_COLUMNS).collect(),
        column_uid_for_index: (0..TABLE_COLUMNS).collect(),
        sorted_row_uids: rows,
        row_index_for_uid: (0..TABLE_ROWS).collect(),
        row_uid_for_index: (0..TABLE_ROWS).collect(),
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
        vec![
            // Native Pages may retain an explicit visible marker.  It is not
            // part of the public semantic value, but an existing-owner edit
            // must preserve it while the axis remains visible.
            extent_state(rows[0], Some(false), None, None),
            extent_state(rows[1], Some(true), None, None),
            // This marker is not user-owned and must survive a rewrite.
            extent_state(
                rows[2],
                None,
                options.filtered_or_pivot.then_some(true),
                None,
            ),
        ]
    } else if options.filtered_or_pivot {
        vec![extent_state(rows[2], None, Some(true), None)]
    } else {
        Vec::new()
    };
    let column_states = if options.user_hidden {
        vec![
            extent_state(columns[0], Some(false), None, None),
            extent_state(columns[2], Some(true), None, None),
            extent_state(
                columns[1],
                None,
                None,
                options.filtered_or_pivot.then_some(true),
            ),
        ]
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
        number_of_filtered_rows: if options.filtered_or_pivot {
            Some(1)
        } else {
            options.explicit_zero_counts.then_some(0)
        },
        number_of_user_hidden_rows: if options.user_hidden {
            Some(1)
        } else {
            options.explicit_zero_counts.then_some(0)
        },
        number_of_user_hidden_columns: if options.user_hidden {
            Some(1)
        } else {
            options.explicit_zero_counts.then_some(0)
        },
        pivot_owner: options.pivot_table.then(|| reference(model_id + 50_000)),
        ..tst::TableModelArchive::default()
    };
    let mut payload = model.encode_to_vec();
    if options.user_hidden {
        let mut owner_payload = owner.encode_to_vec();
        owner_payload = insert_unknown_after_nested_path(
            &owner_payload,
            &[2, 2],
            1,
            UNKNOWN_EXTENT_FIELD,
            UNKNOWN_EXTENT_VALUE,
        )?;
        owner_payload = insert_unknown_after_nested_path(
            &owner_payload,
            &[2, 3],
            1,
            UNKNOWN_EXTENT_FIELD,
            UNKNOWN_EXTENT_VALUE,
        )?;
        owner_payload =
            insert_raw_after_nested_path(&owner_payload, &[2, 2], 1, &unknown_fixed32_wire())?;
        owner_payload = insert_raw_after_nested_path(
            &owner_payload,
            &[2, 2],
            1,
            &unknown_length_delimited_wire()?,
        )?;
        owner_payload = insert_raw_after_nested_path(
            &owner_payload,
            &[2, 3],
            1,
            &unknown_length_delimited_wire()?,
        )?;
        owner_payload = insert_unknown_varint_after_field(
            &owner_payload,
            1,
            UNKNOWN_OWNER_FIELD,
            UNKNOWN_OWNER_VALUE,
        )?;
        owner_payload =
            insert_raw_after_field(&owner_payload, 1, &unknown_length_delimited_wire()?)?;
        owner_payload =
            insert_raw_after_nested_path(&owner_payload, &[2, 3], 1, &unknown_group_wire())?;
        owner_payload = insert_raw_after_field(&owner_payload, 1, &unknown_group_wire())?;
        append_message_field(&mut payload, 70, &owner_payload)?;
    }
    // Literal unknown model field: no focused codec can manufacture this
    // field, so source preservation is tested against an independent shape.
    litchi_iwa_common::wire::append_varint_field(
        &mut payload,
        UNKNOWN_MODEL_FIELD,
        UNKNOWN_MODEL_VALUE,
    )
    .expect("unknown model field fits");
    payload.extend_from_slice(&unknown_length_delimited_wire()?);
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
        // The selected/user-hidden table deliberately exercises the model's
        // native field-46 fallback.  The second table retains the table-info
        // field-6 form so both accepted native ownership shapes are covered.
        view_column_row_uids: (index != 0).then(|| reference(table_uid_map(index))),
        hidden_states_uuid: hidden_uuid,
        is_a_pivot_table: options.pivot_table.then_some(true),
        ..tst::TableInfoArchive::default()
    };
    let mut payload = info.encode_to_vec();
    litchi_iwa_common::wire::append_varint_field(
        &mut payload,
        UNKNOWN_INFO_FIELD,
        UNKNOWN_INFO_VALUE,
    )
    .expect("unknown info field fits");
    payload.extend_from_slice(&unknown_length_delimited_wire().expect("unknown info field fits"));
    payload
}

fn formula_owner_payload(index: usize, hidden_uuid: tsp::Uuid) -> TestResult<Vec<u8>> {
    let formula_owner_uid = uuid(hidden_uuid.lower - 4, hidden_uuid.upper);
    let payload = tsce::FormulaOwnerDependenciesArchive {
        formula_owner_uid,
        internal_formula_owner_id: 1,
        formula_owner: Some(reference(table_drawable(index))),
        ..tsce::FormulaOwnerDependenciesArchive::default()
    }
    .encode_to_vec();
    let payload = insert_unknown_varint_after_field(
        &payload,
        1,
        UNKNOWN_FORMULA_OWNER_FIELD,
        UNKNOWN_FORMULA_OWNER_VALUE,
    )?;
    let payload = insert_raw_after_field(&payload, 1, &unknown_fixed64_wire())?;
    insert_raw_after_field(&payload, 1, &unknown_length_delimited_wire()?)
}

fn hidden_formula_owner_payload(extent_uid: tsp::Uuid) -> TestResult<Vec<u8>> {
    let payload = tst::HiddenStateFormulaOwnerArchive {
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
    .encode_to_vec();
    let payload = insert_unknown_varint_after_field(
        &payload,
        1,
        UNKNOWN_HIDDEN_FORMULA_FIELD,
        UNKNOWN_HIDDEN_FORMULA_VALUE,
    )?;
    let payload = insert_raw_after_field(&payload, 1, &unknown_length_delimited_wire()?)?;
    // Keep the group last for the same reason as the owner fixture above.
    insert_raw_after_field(&payload, 1, &unknown_group_wire())
}

fn filter_set_payload() -> TestResult<Vec<u8>> {
    let payload = tst::FilterSetArchive {
        r#type: Some(tst::filter_set_archive::FilterSetType::FilterSetArchiveTypeAll as i32),
        is_enabled: Some(false),
        needs_formula_rewrite_for_import: Some(false),
        filter_offsets: vec![0],
        ..tst::FilterSetArchive::default()
    }
    .encode_to_vec();
    let payload =
        insert_unknown_varint_after_field(&payload, 1, UNKNOWN_FILTER_FIELD, UNKNOWN_FILTER_VALUE)?;
    let payload = insert_raw_after_field(&payload, 1, &unknown_fixed32_wire())?;
    insert_raw_after_field(&payload, 1, &unknown_length_delimited_wire()?)
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

        let uid_map_object = object(map_id, UID_MAP_MESSAGE_TYPE, uid_map_payload(index), &[])?;
        objects.extend([attachment_object, info_object, model_object, uid_map_object]);

        // Every native table has the formula-owner dependency record.  The
        // hidden-state owner itself is optional, but the dependency UUID is
        // still the stable identity used by the absence/read path.
        let formula_owner_object = object(
            table_formula_owner(index),
            FORMULA_OWNER_MESSAGE_TYPE,
            formula_owner_payload(index, owner_uuid)?,
            &[info_id],
        )?;
        if options[index].user_hidden {
            let (rows, columns) = table_uids(index);
            let column_extent_uid = uuid(owner_uuid.lower + 7, owner_uuid.upper);
            let formula_columns = object(
                table_formula_object(index, true),
                HIDDEN_STATE_FORMULA_OWNER_MESSAGE_TYPE,
                hidden_formula_owner_payload(column_extent_uid)?,
                &[],
            )?;
            let formula_rows = object(
                table_formula_object(index, false),
                HIDDEN_STATE_FORMULA_OWNER_MESSAGE_TYPE,
                hidden_formula_owner_payload(owner_uuid)?,
                &[],
            )?;
            let filter_columns = object(
                table_filter_set(index, true),
                FILTER_SET_MESSAGE_TYPE,
                filter_set_payload()?,
                &[],
            )?;
            let filter_rows = object(
                table_filter_set(index, false),
                FILTER_SET_MESSAGE_TYPE,
                filter_set_payload()?,
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
    members.extend(
        NONCANONICAL_PREVIEWS
            .into_iter()
            .map(|name| (name, b"noncanonical-preview".as_slice())),
    );
    Ok(litchi_iwa_archive::package::to_bytes(
        members,
        Limits::default(),
    )?)
}

fn normal_package() -> TestResult<Vec<u8>> {
    let source = synthetic_package(
        [
            TableOptions {
                user_hidden: true,
                ..TableOptions::default()
            },
            TableOptions::default(),
        ],
        ["Revenue", "Costs"],
    )?;
    Ok(source)
}

fn filtered_package() -> TestResult<Vec<u8>> {
    synthetic_package(
        [
            TableOptions {
                user_hidden: true,
                filtered_or_pivot: true,
                ..TableOptions::default()
            },
            TableOptions::default(),
        ],
        ["Revenue", "Costs"],
    )
}

fn pivot_package() -> TestResult<Vec<u8>> {
    synthetic_package(
        [
            TableOptions {
                user_hidden: true,
                pivot_table: true,
                ..TableOptions::default()
            },
            TableOptions::default(),
        ],
        ["Revenue", "Costs"],
    )
}

fn locked_package() -> TestResult<Vec<u8>> {
    synthetic_package(
        [
            TableOptions {
                user_hidden: true,
                locked: true,
                ..TableOptions::default()
            },
            TableOptions::default(),
        ],
        ["Revenue", "Costs"],
    )
}

fn explicit_zero_count_package() -> TestResult<Vec<u8>> {
    synthetic_package(
        [
            TableOptions {
                user_hidden: true,
                explicit_zero_counts: true,
                ..TableOptions::default()
            },
            TableOptions::default(),
        ],
        ["Revenue", "Costs"],
    )
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

fn member_bytes(package: &[u8], name: &str) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    Ok(catalog
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or_else(|| format!("missing package member {name}"))?
        .data()
        .to_vec())
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

fn object_header_bytes(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or("missing document member")?;
    let decompressed = SnappyStream::decompress(entry.data())?.into_bytes();
    let archive = Archive::parse(&decompressed)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| format!("missing object {identifier}"))?;
    let object_start = usize::try_from(object.header_offset)?;
    let (header_length, prefix_length) = decode_varint_from_bytes(&decompressed[object_start..])?;
    let header_length = usize::try_from(header_length)?;
    let header_start = object_start
        .checked_add(prefix_length)
        .ok_or("object header offset overflow")?;
    let header_end = header_start
        .checked_add(header_length)
        .ok_or("object header range overflow")?;
    Ok(decompressed
        .get(header_start..header_end)
        .ok_or("object header is truncated")?
        .to_vec())
}

fn rewrite_raw_object_header(
    package: &[u8],
    identifier: u64,
    mutate: impl FnOnce(&[u8]) -> TestResult<Vec<u8>>,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or("missing document member")?;
    let decompressed = SnappyStream::decompress(entry.data())?.into_bytes();
    let archive = Archive::parse(&decompressed)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| format!("missing object {identifier}"))?;
    let object_start = usize::try_from(object.header_offset)?;
    let (old_header_length, prefix_length) =
        decode_varint_from_bytes(&decompressed[object_start..])?;
    let old_header_length = usize::try_from(old_header_length)?;
    let header_start = object_start
        .checked_add(prefix_length)
        .ok_or("object header offset overflow")?;
    let header_end = header_start
        .checked_add(old_header_length)
        .ok_or("object header range overflow")?;
    let header = decompressed
        .get(header_start..header_end)
        .ok_or("object header is truncated")?
        .to_vec();
    let header = mutate(&header)?;

    let mut rewritten = Vec::with_capacity(
        decompressed
            .len()
            .checked_add(header.len().saturating_sub(old_header_length))
            .ok_or("component length overflow")?,
    );
    rewritten.extend_from_slice(&decompressed[..object_start]);
    litchi_iwa_common::varint::encode_varint_into(&mut rewritten, u64::try_from(header.len())?);
    rewritten.extend_from_slice(&header);
    rewritten.extend_from_slice(
        decompressed
            .get(header_end..)
            .ok_or("object payload is truncated")?,
    );
    let component = SnappyStream::compress(&rewritten)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(DOCUMENT_MEMBER, &component)],
        Limits::default(),
    )?)
}

fn with_noncanonical_model_header(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_raw_object_header(package, table_model(index), |source| {
        let mut header = source.to_vec();
        litchi_iwa_common::wire::append_varint_field(
            &mut header,
            UNKNOWN_LENGTH_FIELD,
            0xdecafbad,
        )?;
        Ok(header)
    })
}

fn mismatched_uid_map_message_info_type(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_raw_object_header(package, table_uid_map(index), |source| {
        let mut replacement = Vec::new();
        litchi_iwa_common::wire::append_varint_field(
            &mut replacement,
            1,
            u64::from(UID_MAP_MESSAGE_TYPE + 1),
        )?;
        replace_first_raw_field_at_nested_path(source, &[2], 1, &replacement)
    })
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

fn rewrite_model_payload(package: &[u8], index: usize, payload: Vec<u8>) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let model = archive
            .object_mut(table_model(index))
            .ok_or("missing table model")?;
        let message = model
            .messages
            .iter_mut()
            .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or("missing table-model message")?;
        message.data = payload;
        Ok(())
    })
}

fn replace_first_raw_field(source: &[u8], number: u32, replacement: &[u8]) -> TestResult<Vec<u8>> {
    let fields = raw_wire_fields(source)?;
    let mut result = Vec::with_capacity(source.len());
    let mut replaced = false;
    for field in fields {
        if !replaced && field.number() == number {
            result.extend_from_slice(replacement);
            replaced = true;
        } else {
            result.extend_from_slice(field.raw());
        }
    }
    replaced
        .then_some(result)
        .ok_or_else(|| format!("missing field {number} for raw replacement").into())
}

fn remove_first_raw_field(source: &[u8], number: u32) -> TestResult<Vec<u8>> {
    let fields = raw_wire_fields(source)?;
    let mut result = Vec::with_capacity(source.len());
    let mut removed = false;
    for field in fields {
        if !removed && field.number() == number {
            removed = true;
        } else {
            result.extend_from_slice(field.raw());
        }
    }
    removed
        .then_some(result)
        .ok_or_else(|| format!("missing field {number} for raw removal").into())
}

fn replace_first_raw_field_at_nested_path(
    source: &[u8],
    path: &[u32],
    number: u32,
    replacement: &[u8],
) -> TestResult<Vec<u8>> {
    if path.is_empty() {
        return replace_first_raw_field(source, number, replacement);
    }
    let fields = raw_wire_fields(source)?;
    let mut result = Vec::with_capacity(source.len());
    let mut changed = false;
    for field in fields {
        if !changed && field.number() == path[0] && field.wire_type() == 2 {
            let nested = replace_first_raw_field_at_nested_path(
                field.payload(),
                &path[1..],
                number,
                replacement,
            )?;
            append_message_field(&mut result, field.number(), &nested)?;
            changed = true;
        } else {
            result.extend_from_slice(field.raw());
        }
    }
    changed
        .then_some(result)
        .ok_or_else(|| format!("missing nested field path {path:?}").into())
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

fn append_ambiguous_body_attachment_owner(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let body = archive
            .object_mut(BODY_IDENTIFIER)
            .ok_or("missing body storage")?;
        let message = body.messages.first_mut().ok_or("missing body message")?;
        let mut storage = tswp::StorageArchive::decode(message.data.as_slice())?;
        storage.text = vec!["\u{fffc}\u{fffc}\u{fffc}".to_owned()];
        storage
            .table_attachment
            .as_mut()
            .ok_or("missing table-attachment inventory")?
            .entries
            .push(tswp::object_attribute_table::ObjectAttribute {
                character_index: 2,
                object: Some(reference(table_attachment(index))),
            });
        message.data = storage.encode_to_vec();
        let message_info = body
            .archive_info
            .message_infos
            .first_mut()
            .ok_or("missing body message metadata")?;
        // The same attachment is now declared twice at the rooted field-9
        // ownership edge.  Keep the bytes parseable, then require the native
        // ownership proof to reject the ambiguity before any rewrite.
        message_info.object_references.push(table_attachment(index));
        message_info
            .field_infos
            .push(field_reference(vec![9], table_attachment(index)));
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

fn duplicate_active_hidden_state(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model(package, index, |model| {
        if let Some(owner) = model.hidden_states_owner.as_mut() {
            if let Some(active) = owner.hidden_states.first().cloned() {
                owner.hidden_states.push(active);
            }
        }
    })
}

fn append_inactive_hidden_state(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model(package, index, |model| {
        if let Some(owner) = model.hidden_states_owner.as_mut()
            && let Some(mut inactive) = owner.hidden_states.first().cloned()
        {
            let (rows, columns) = table_uids(index);
            // The table-info UUID still selects the original active view.  A
            // second, distinct stored view is therefore either preserved
            // byte-for-byte by an edit of the active view or conservatively
            // refused; it must never be silently rewritten as if active.
            inactive.hidden_states_uid.lower =
                inactive.hidden_states_uid.lower.saturating_add(0x400);
            inactive
                .column_hidden_state_extent
                .hidden_state_extent_uid
                .lower = inactive
                .column_hidden_state_extent
                .hidden_state_extent_uid
                .lower
                .saturating_add(0x400);
            inactive
                .row_hidden_state_extent
                .hidden_state_extent_uid
                .lower = inactive
                .row_hidden_state_extent
                .hidden_state_extent_uid
                .lower
                .saturating_add(0x400);
            // Give the inactive view markers that cannot be mistaken for the
            // active view.  Both UIDs remain members of the same physical
            // row/column map so a resolver must use the selected view rather
            // than blindly taking the first state.
            inactive.row_hidden_state_extent.base_hidden_states =
                vec![extent_state(rows[3], Some(true), None, None)];
            inactive.column_hidden_state_extent.base_hidden_states =
                vec![extent_state(columns[0], Some(true), None, None)];
            // The admitted read profile requires uniquely owned filter sets.
            // An inactive view without filters retains independent axis state.
            inactive.row_hidden_state_extent.filter_set = None;
            inactive.column_hidden_state_extent.filter_set = None;
            owner.hidden_states.push(inactive);
        }
    })
}

fn missing_active_hidden_state_uuid(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    let payload = info_payload(package, index)?;
    let mut info = tst::TableInfoArchive::decode(payload.as_slice())?;
    info.hidden_states_uuid = None;
    let mut replacement = info.encode_to_vec();
    litchi_iwa_common::wire::append_varint_field(
        &mut replacement,
        UNKNOWN_INFO_FIELD,
        UNKNOWN_INFO_VALUE,
    )?;
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

fn select_second_hidden_state_as_active(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    let source = append_inactive_hidden_state(package, index)?;
    let owner = tst::HiddenStatesOwnerArchive::decode(
        hidden_owner_payload_from_package(&source, index)?.as_slice(),
    )?;
    let second = owner
        .hidden_states
        .get(1)
        .ok_or("missing inactive hidden-state view")?;
    let second_uuid = second.hidden_states_uid;
    let second_column_extent_uid = second.column_hidden_state_extent.hidden_state_extent_uid;
    let second_row_extent_uid = second.row_hidden_state_extent.hidden_state_extent_uid;
    let payload = info_payload(&source, index)?;
    let mut info = tst::TableInfoArchive::decode(payload.as_slice())?;
    info.hidden_states_uuid = Some(second_uuid);
    let mut replacement = info.encode_to_vec();
    litchi_iwa_common::wire::append_varint_field(
        &mut replacement,
        UNKNOWN_INFO_FIELD,
        UNKNOWN_INFO_VALUE,
    )?;
    rewrite_document_archive(&source, |archive| {
        let object = archive
            .object_mut(table_drawable(index))
            .ok_or("missing table info")?;
        let message = object
            .messages
            .iter_mut()
            .find(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .ok_or("missing table-info message")?;
        message.data = replacement;

        let model = archive
            .object_mut(table_model(index))
            .ok_or("missing model")?;
        let model_message = model
            .messages
            .iter_mut()
            .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or("missing model message")?;
        let mut model_value = tst::TableModelArchive::decode(model_message.data.as_slice())?;
        model_value
            .hidden_states_owner
            .as_mut()
            .ok_or("missing hidden owner")?
            .owner_uid = second_uuid;
        model_message.data = model_value.encode_to_vec();

        // The 4008 formula-owner UUID is the owner-level predecessor of the
        // selected 6204 hidden-state UUID.  When the table-info selector
        // moves to the second view, update that identity as well; leaving it
        // tied to the first view would make a valid non-first active view
        // look like a dangling dependency.
        let formula = archive
            .object_mut(table_formula_owner(index))
            .ok_or("missing formula-owner object")?;
        let formula_message = formula
            .messages
            .iter_mut()
            .find(|message| message.type_ == FORMULA_OWNER_MESSAGE_TYPE)
            .ok_or("missing formula-owner message")?;
        let mut formula_owner =
            tsce::FormulaOwnerDependenciesArchive::decode(formula_message.data.as_slice())?;
        formula_owner.formula_owner_uid =
            uuid(second_uuid.lower.saturating_sub(4), second_uuid.upper);
        formula_message.data = formula_owner.encode_to_vec();

        for (identifier, extent_uid) in [
            (table_formula_object(index, true), second_column_extent_uid),
            (table_formula_object(index, false), second_row_extent_uid),
        ] {
            let formula = archive
                .object_mut(identifier)
                .ok_or("missing hidden-state formula owner")?;
            let formula_message = formula
                .messages
                .iter_mut()
                .find(|message| message.type_ == HIDDEN_STATE_FORMULA_OWNER_MESSAGE_TYPE)
                .ok_or("missing hidden-state formula-owner message")?;
            formula_message.data = hidden_formula_owner_payload(extent_uid)?;
        }
        Ok(())
    })
}

fn duplicate_hidden_row_state(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model(package, index, |model| {
        if let Some(owner) = model.hidden_states_owner.as_mut()
            && let Some(active) = owner.hidden_states.first_mut()
            && let Some(state) = active
                .row_hidden_state_extent
                .base_hidden_states
                .first()
                .cloned()
        {
            active
                .row_hidden_state_extent
                .base_hidden_states
                .push(state);
        }
    })
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

fn dangling_uid_map_reference(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model(package, index, |model| {
        model.base_column_row_uids = Some(reference(0xdead_beef));
    })
}

fn cross_type_uid_map_reference(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model(package, index, |model| {
        // The identifier exists, but its native payload is a hidden-state
        // formula owner rather than a UID map.
        model.base_column_row_uids = Some(reference(table_formula_object(index, true)));
    })
}

fn inconsistent_row_count(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model(package, index, |model| {
        model.number_of_rows = TABLE_ROWS.saturating_sub(1);
    })
}

fn out_of_range_row_count(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model(package, index, |model| {
        model.number_of_rows = TABLE_ROWS.saturating_add(1);
    })
}

fn overlong_known_row_count(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    // Field 6 is the known u32 row count.  This is the canonical key followed
    // by a varint for 2^32, which must not be truncated into a valid count.
    let payload = model_payload_from_package(package, index)?;
    rewrite_model_payload(
        package,
        index,
        replace_first_raw_field(&payload, 6, &[0x30, 0x80, 0x80, 0x80, 0x80, 0x10])?,
    )
}

fn stale_hidden_count(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model(package, index, |model| {
        model.number_of_hidden_rows = Some(99);
    })
}

fn stale_filtered_count(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model(package, index, |model| {
        model.number_of_filtered_rows = Some(1);
    })
}

fn replace_model_field_wire(
    package: &[u8],
    index: usize,
    field: u32,
    replacement: &[u8],
) -> TestResult<Vec<u8>> {
    let payload = model_payload_from_package(package, index)?;
    rewrite_model_payload(
        package,
        index,
        replace_first_raw_field(&payload, field, replacement)?,
    )
}

fn wrong_model_message_type(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let model = archive
            .object_mut(table_model(index))
            .ok_or("missing table model")?;
        let message = model
            .messages
            .iter_mut()
            .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or("missing table-model message")?;
        message.type_ = TABLE_MODEL_MESSAGE_TYPE + 100;
        Ok(())
    })
}

fn wrong_hidden_owner_uuid(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model(package, index, |model| {
        if let Some(owner) = model.hidden_states_owner.as_mut() {
            owner.owner_uid.lower = owner.owner_uid.lower.saturating_add(1);
        }
    })
}

fn zero_hidden_owner_uuid(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model(package, index, |model| {
        if let Some(owner) = model.hidden_states_owner.as_mut() {
            owner.owner_uid = uuid(0, 0);
        }
    })
}

fn cross_axis_extent_uid_alias(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model(package, index, |model| {
        if let Some(owner) = model.hidden_states_owner.as_mut()
            && let Some(active) = owner.hidden_states.first_mut()
        {
            active.column_hidden_state_extent.hidden_state_extent_uid =
                active.row_hidden_state_extent.hidden_state_extent_uid;
        }
    })
}

fn invalid_extent_boolean(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    let payload = model_payload_from_package(package, index)?;
    // `needs_to_update_filter_set_for_import` is a known bool at extent field
    // 6.  A protobuf bool is not an arbitrary integer: value 2 must be
    // rejected instead of being silently coerced to true.
    let payload = replace_first_raw_field_at_nested_path(&payload, &[70, 2, 2], 6, &[0x30, 0x02])?;
    rewrite_model_payload(package, index, payload)
}

fn pivot_owner_only(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model(package, index, |model| {
        model.pivot_owner = Some(reference(table_model(index) + 50_000));
    })
}

fn merge_owner_only(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model(package, index, |model| {
        model.merge_owner = Some(tst::MergeOwnerArchive {
            owner_id: tsp::CfuuidArchive {
                uuid_bytes: None,
                uuid_w0: Some(1),
                uuid_w1: Some(2),
                uuid_w2: Some(3),
                uuid_w3: Some(4),
            },
            formula_store: None,
        });
    })
}

fn role_sharing_object_reference(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    // A hidden-state formula-owner object is already occupied by another
    // native role.  Reusing it as the model's UID-map edge must not make the
    // graph appear valid merely because the identifier exists.
    cross_type_uid_map_reference(package, index)
}

fn rewrite_uid_map_metadata(
    package: &[u8],
    index: usize,
    mutate: impl FnOnce(&mut ArchiveObject, usize) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let map = archive
            .object_mut(table_uid_map(index))
            .ok_or("missing UID map")?;
        let message_index = map
            .messages
            .iter()
            .position(|message| message.type_ == UID_MAP_MESSAGE_TYPE || message.type_ == 6_200)
            .ok_or("missing UID-map message")?;
        mutate(map, message_index)
    })
}

fn qualified_legacy_uid_map(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_uid_map_metadata(package, index, |map, message_index| {
        map.messages[message_index].type_ = 6_200;
        let info = &mut map.archive_info.message_infos[message_index];
        info.type_ = 6_200;
        info.versions = vec![1, 0, 5];
        Ok(())
    })
}

fn unqualified_legacy_uid_map(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_uid_map_metadata(package, index, |map, message_index| {
        map.messages[message_index].type_ = 6_200;
        let info = &mut map.archive_info.message_infos[message_index];
        info.type_ = 6_200;
        // A legacy type without the native archive version tuple is not a
        // qualified Pages UID-map route.
        info.versions = vec![1, 0, 4];
        Ok(())
    })
}

fn mismatched_uid_map_header_type(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_uid_map_metadata(package, index, |map, message_index| {
        // Archive::to_bytes synchronizes MessageInfo.type_ from the raw
        // message.  Change the raw message type so both projections agree on
        // the wrong canonical type instead of accidentally normalizing this
        // case back to 6267.
        map.messages[message_index].type_ = UID_MAP_MESSAGE_TYPE + 1;
        Ok(())
    })
}

fn uid_map_should_merge(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_uid_map_metadata(package, index, |map, _| {
        map.archive_info.should_merge = Some(true);
        Ok(())
    })
}

fn uid_map_base_message(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_uid_map_metadata(package, index, |map, message_index| {
        map.archive_info.message_infos[message_index].base_message_index = Some(0);
        Ok(())
    })
}

fn uid_map_diff_merge_version(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_uid_map_metadata(package, index, |map, message_index| {
        map.archive_info.message_infos[message_index].diff_merge_version = vec![1];
        Ok(())
    })
}

fn uid_map_diff_field_path(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_uid_map_metadata(package, index, |map, message_index| {
        map.archive_info.message_infos[message_index].diff_field_path =
            Some(FieldPath::new(vec![1]));
        Ok(())
    })
}

fn uid_map_fields_to_remove(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_uid_map_metadata(package, index, |map, message_index| {
        map.archive_info.message_infos[message_index].fields_to_remove =
            vec![FieldPath::new(vec![1])];
        Ok(())
    })
}

fn uid_map_diff_read_version(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_uid_map_metadata(package, index, |map, message_index| {
        map.archive_info.message_infos[message_index].diff_read_version = vec![1];
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

fn duplicate_sorted_uid(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
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
    map.sorted_row_uids[1] = map.sorted_row_uids[0];
    rewrite_uid_map(package, index, map)
}

fn malformed_uid_map_inverse(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
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
    map.column_index_for_uid = vec![0, 0, 2, 3];
    rewrite_uid_map(package, index, map)
}

fn non_identity_uid_permutation(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
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
    // `*_uid_for_index` maps physical positions to stable sorted entries;
    // the other arrays are their exact inverses.  Keep both pairs valid while
    // making positional projection depend on the map rather than list order.
    map.column_uid_for_index = vec![2, 0, 3, 1];
    map.column_index_for_uid = vec![1, 3, 0, 2];
    map.row_uid_for_index = vec![3, 1, 0, 2];
    map.row_index_for_uid = vec![2, 1, 3, 0];
    rewrite_uid_map(package, index, map)
}

fn wrong_active_uuid(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    let payload = info_payload(package, index)?;
    let mut info = tst::TableInfoArchive::decode(payload.as_slice())?;
    info.hidden_states_uuid = Some(uuid(0xdead, 0xbeef));
    let mut replacement = info.encode_to_vec();
    litchi_iwa_common::wire::append_varint_field(
        &mut replacement,
        UNKNOWN_INFO_FIELD,
        UNKNOWN_INFO_VALUE,
    )?;
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

fn duplicate_hidden_owner_field(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    let payload = model_payload_from_package(package, index)?;
    let owner = WireView::parse(&payload)?
        .fields()
        .find(|field| field.number() == 70)
        .ok_or("missing hidden-state owner")?
        .payload()
        .to_vec();
    let mut duplicate = Vec::new();
    append_message_field(&mut duplicate, 70, &owner)?;
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

#[test]
fn malformed_and_ambiguous_native_ownership_fails_closed_atomically() -> TestResult {
    let source = normal_package()?;
    let malformed_sources = [
        append_duplicate_model_message(&source, 0)?,
        append_duplicate_info_message(&source, 0)?,
        append_ambiguous_body_attachment_owner(&source, 0)?,
        remove_object(&source, table_uid_map(0))?,
        remove_object(&source, table_formula_owner(0))?,
        append_duplicate_formula_owner(&source, 0)?,
        duplicate_active_hidden_state(&source, 0)?,
        duplicate_hidden_row_state(&source, 0)?,
        unknown_axis_uuid(&source, 0)?,
        zero_hidden_owner_uuid(&source, 0)?,
        cross_axis_extent_uid_alias(&source, 0)?,
        invalid_extent_boolean(&source, 0)?,
        pivot_owner_only(&source, 0)?,
        role_sharing_object_reference(&source, 0)?,
        formula_owner_external_reference(&source, 0)?,
        hidden_extent_external_filter_reference(&source, 0)?,
        dangling_uid_map_reference(&source, 0)?,
        cross_type_uid_map_reference(&source, 0)?,
        inconsistent_row_count(&source, 0)?,
        out_of_range_row_count(&source, 0)?,
        stale_hidden_count(&source, 0)?,
        stale_filtered_count(&source, 0)?,
        overlong_known_row_count(&source, 0)?,
        wrong_model_message_type(&source, 0)?,
        wrong_hidden_owner_uuid(&source, 0)?,
        unqualified_legacy_uid_map(&source, 0)?,
        mismatched_uid_map_header_type(&source, 0)?,
        mismatched_uid_map_message_info_type(&source, 0)?,
        uid_map_should_merge(&source, 0)?,
        uid_map_base_message(&source, 0)?,
        uid_map_diff_merge_version(&source, 0)?,
        uid_map_diff_field_path(&source, 0)?,
        uid_map_fields_to_remove(&source, 0)?,
        uid_map_diff_read_version(&source, 0)?,
        wrong_uid_map_lengths(&source, 0)?,
        duplicate_uid_map_entry(&source, 0)?,
        duplicate_sorted_uid(&source, 0)?,
        malformed_uid_map_inverse(&source, 0)?,
        remove_model_map_field(&source, 0)?,
        remove_model_map_payload_field(&source, 0)?,
        wrong_model_map_field_path(&source, 0)?,
        model_map_data_reference(&source, 0)?,
        model_map_extra_aggregate_reference(&source, 0)?,
        model_external_map_reference(&source, 0)?,
        wrong_active_uuid(&source, 0)?,
        duplicate_hidden_owner_field(&source, 0)?,
        wrong_hidden_extent_direction(&source, 0)?,
    ];
    for (case_index, bytes) in malformed_sources.into_iter().enumerate() {
        let Ok(package) = Package::from_bytes(&bytes) else {
            // A source-authoritative parser may reject malformed metadata
            // before the semantic transaction is opened.  That is still a
            // fail-closed result.
            continue;
        };
        let before = package.exact_bytes();
        assert_rejected_without_mutation(
            &package,
            package
                .edit_body_table_hidden_axes(0usize)
                .and_then(|edit| edit.set(HiddenAxes::empty()).commit()),
            &before,
            case_index,
        );
    }
    Ok(())
}

#[test]
fn uid_map_legacy_qualification_and_table_info_model_reference_edges_are_strict() -> TestResult {
    let source = normal_package()?;
    let qualified_source = qualified_legacy_uid_map(&source, 0)?;
    let qualified = Package::from_bytes(&qualified_source)?;
    assert_eq!(
        qualified.body_table_hidden_axes(0usize)?,
        HiddenAxes::new([AxisIndex::row(1), AxisIndex::column(2)])?
    );
    let qualified_commit = qualified
        .edit_body_table_hidden_axes(0usize)?
        .set(HiddenAxes::new([AxisIndex::row(0)])?)
        .commit()?;
    assert_eq!(
        Package::from_bytes(&qualified_commit.package().exact_bytes())?
            .body_table_hidden_axes(0usize)?,
        HiddenAxes::new([AxisIndex::row(0)])?
    );

    let permutation_source = non_identity_uid_permutation(&source, 0)?;
    let permutation = Package::from_bytes(&permutation_source)?;
    // The map's physical-position permutation moves the stable column UID
    // that was hidden at sorted index 2 to physical column 0.  A valid map
    // must be followed, not mistaken for malformed ordering.
    assert_eq!(
        permutation.body_table_hidden_axes(0usize)?,
        HiddenAxes::new([AxisIndex::row(1), AxisIndex::column(0)])?
    );
    let permutation_commit = permutation
        .edit_body_table_hidden_axes(0usize)?
        .set(HiddenAxes::new([AxisIndex::row(2), AxisIndex::column(3)])?)
        .commit()?;
    assert_eq!(
        Package::from_bytes(&permutation_commit.package().exact_bytes())?
            .body_table_hidden_axes(0usize)?,
        HiddenAxes::new([AxisIndex::row(2), AxisIndex::column(3)])?
    );

    // Exercise the optional table-info map edge against a table that has an
    // existing owner, so each case reaches graph validation instead of being
    // short-circuited by the existing-owner-only refusal.
    let both_owner_source = synthetic_package(
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
    let table_one_cases = [
        disagreeing_info_map_reference(&both_owner_source, 1)?,
        disagreeing_model_map_reference(&both_owner_source, 1)?,
        remove_model_map_field(&both_owner_source, 1)?,
        remove_model_map_payload_field(&both_owner_source, 1)?,
        wrong_model_map_field_path(&both_owner_source, 1)?,
        model_map_data_reference(&both_owner_source, 1)?,
        model_map_extra_aggregate_reference(&both_owner_source, 1)?,
        model_external_map_reference(&both_owner_source, 1)?,
        remove_info_map_field(&both_owner_source, 1)?,
        wrong_info_map_field_path(&both_owner_source, 1)?,
        info_map_data_reference(&both_owner_source, 1)?,
        info_map_extra_aggregate_reference(&both_owner_source, 1)?,
        info_external_map_reference(&both_owner_source, 1)?,
    ];
    let rejected_axes = HiddenAxes::new([AxisIndex::row(0)])?;
    for (case_index, bytes) in table_one_cases.into_iter().enumerate() {
        let Ok(package) = Package::from_bytes(&bytes) else {
            continue;
        };
        let before = package.exact_bytes();
        let result = package
            .edit_body_table_hidden_axes(1usize)
            .and_then(|edit| edit.set(rejected_axes.clone()).commit());
        assert!(
            result.is_err(),
            "invalid table-info/model edge was accepted (case {case_index})"
        );
        assert_eq!(package.exact_bytes(), before);
    }
    Ok(())
}

#[test]
fn active_view_selection_and_multiview_writes_fail_closed() -> TestResult {
    let source = append_inactive_hidden_state(&normal_package()?, 0)?;
    let package = Package::from_bytes(&source)?;
    let owner_payload = hidden_owner_payload_from_package(&source, 0)?;
    let owner = tst::HiddenStatesOwnerArchive::decode(owner_payload.as_slice())?;
    assert_eq!(owner.hidden_states.len(), 2);
    let (rows, columns) = table_uids(0);
    let inactive = owner.hidden_states.get(1).ok_or("missing inactive view")?;
    assert!(
        inactive
            .row_hidden_state_extent
            .base_hidden_states
            .iter()
            .any(|state| state.row_or_column_uid == rows[3] && state.user_hidden == Some(true))
    );
    assert!(
        inactive
            .column_hidden_state_extent
            .base_hidden_states
            .iter()
            .any(|state| state.row_or_column_uid == columns[0] && state.user_hidden == Some(true))
    );
    let before_read = package.exact_bytes();
    assert_eq!(
        package.body_table_hidden_axes(0usize)?,
        HiddenAxes::new([AxisIndex::row(1), AxisIndex::column(2)])?
    );
    assert_eq!(package.exact_bytes(), before_read);

    let before = before_read;
    let result = package
        .edit_body_table_hidden_axes(0usize)?
        .set(HiddenAxes::new([AxisIndex::row(0)])?)
        .commit();
    assert!(matches!(
        result,
        Err(Error::UnsupportedDependency | Error::InvalidSource | Error::UnsupportedSource)
    ));
    assert_eq!(package.exact_bytes(), before);

    let second_active_source = select_second_hidden_state_as_active(&normal_package()?, 0)?;
    let second_active = Package::from_bytes(&second_active_source)?;
    let second_before = second_active.exact_bytes();
    // A read must honor the table-info-selected non-first view.  The write
    // remains conservative because the owner stores more than one view.
    assert_eq!(
        second_active.body_table_hidden_axes(0usize)?,
        HiddenAxes::new([AxisIndex::row(3), AxisIndex::column(0)])?
    );
    let second_owner = tst::HiddenStatesOwnerArchive::decode(
        hidden_owner_payload_from_package(&second_active_source, 0)?.as_slice(),
    )?;
    assert!(
        second_owner
            .hidden_states
            .first()
            .ok_or("missing original hidden-state view")?
            .row_hidden_state_extent
            .base_hidden_states
            .iter()
            .any(|state| state.row_or_column_uid == rows[1] && state.user_hidden == Some(true))
    );
    assert!(
        second_owner
            .hidden_states
            .get(1)
            .ok_or("missing selected hidden-state view")?
            .row_hidden_state_extent
            .base_hidden_states
            .iter()
            .any(|state| state.row_or_column_uid == rows[3] && state.user_hidden == Some(true))
    );
    let formula_owner = tsce::FormulaOwnerDependenciesArchive::decode(
        object_message_payload(
            &second_active_source,
            table_formula_owner(0),
            FORMULA_OWNER_MESSAGE_TYPE,
        )?
        .as_slice(),
    )?;
    let selected_uuid = second_owner
        .hidden_states
        .get(1)
        .ok_or("missing selected hidden-state view")?
        .hidden_states_uid;
    assert_eq!(
        formula_owner.formula_owner_uid,
        uuid(selected_uuid.lower.saturating_sub(4), selected_uuid.upper)
    );
    let second_result = second_active
        .edit_body_table_hidden_axes(0usize)?
        .clear()
        .commit();
    assert!(matches!(
        second_result,
        Err(Error::UnsupportedDependency | Error::InvalidSource | Error::UnsupportedSource)
    ));
    assert_eq!(second_active.exact_bytes(), second_before);

    let missing_uuid_source = missing_active_hidden_state_uuid(&source, 0)?;
    let missing_uuid = Package::from_bytes(&missing_uuid_source)?;
    let missing_before = missing_uuid.exact_bytes();
    assert!(matches!(
        missing_uuid.body_table_hidden_axes(0usize),
        Err(Error::InvalidSource | Error::UnsupportedDependency | Error::UnsupportedSource)
    ));
    assert!(matches!(
        missing_uuid.edit_body_table_hidden_axes(0usize),
        Err(Error::InvalidSource | Error::UnsupportedDependency | Error::UnsupportedSource)
    ));
    assert_eq!(missing_uuid.exact_bytes(), missing_before);
    Ok(())
}

#[test]
fn literal_malformed_wire_fields_are_rejected_without_source_mutation() -> TestResult {
    let source = normal_package()?;
    // Field 70 is a length-delimited nested owner.  These are deliberately
    // literal wire shapes rather than protobuf-generated values.
    for raw in [
        &[0xb3, 0x04, 0x08, 0x01][..], // unbalanced group
        &[0xb0, 0x04, 0x01][..],       // owner with wrong wire type
        &[0xb2, 0x04, 0x80, 0x00][..], // non-canonical length varint
        &[0xb2, 0x04, 0x80][..],       // truncated length-delimited field
    ] {
        let malformed = replace_model_field_wire(&source, 0, 70, raw)?;
        let Ok(package) = Package::from_bytes(&malformed) else {
            // Rejection at package ingress is also a valid fail-closed result.
            continue;
        };
        let before = package.exact_bytes();
        let result = package.body_table_hidden_axes(0usize);
        assert!(
            result.is_err(),
            "literal malformed wire was accepted: {raw:?}"
        );
        assert_eq!(package.exact_bytes(), before);
    }
    Ok(())
}

#[test]
fn filtered_and_pivot_states_are_preserved_or_refused_fail_closed() -> TestResult {
    let source = filtered_package()?;
    let package = Package::from_bytes(&source)?;
    let helper_payloads = [
        (table_formula_owner(0), FORMULA_OWNER_MESSAGE_TYPE),
        (
            table_formula_object(0, true),
            HIDDEN_STATE_FORMULA_OWNER_MESSAGE_TYPE,
        ),
        (
            table_formula_object(0, false),
            HIDDEN_STATE_FORMULA_OWNER_MESSAGE_TYPE,
        ),
        (table_filter_set(0, true), FILTER_SET_MESSAGE_TYPE),
        (table_filter_set(0, false), FILTER_SET_MESSAGE_TYPE),
    ]
    .into_iter()
    .map(|(identifier, type_)| {
        Ok::<_, Box<dyn StdError>>((
            identifier,
            type_,
            object_message_payload(&source, identifier, type_)?,
        ))
    })
    .collect::<TestResult<Vec<_>>>()?;
    let before = package.body_table_hidden_axes(0usize)?;
    let source_model_payload = model_payload_from_package(&source, 0)?;
    let source_info_payload = info_payload(&source, 0)?;
    let source_owner_payload = hidden_owner_payload_from_package(&source, 0)?;
    let filtered_rows = WireView::parse(&source_model_payload)?
        .fields()
        .find(|field| field.number() == 40)
        .ok_or("filtered-row count field 40 is missing")?;
    assert_eq!(decode_varint_from_bytes(filtered_rows.payload())?.0, 1);
    assert!(before.contains(AxisIndex::row(1)));
    assert!(before.contains(AxisIndex::column(2)));

    let requested = HiddenAxes::new([AxisIndex::row(0), AxisIndex::column(3)])?;
    let commit = package
        .edit_body_table_hidden_axes(0usize)?
        .set(requested.clone())
        .commit()?;
    // A user edit may coexist with non-user filtered/pivot markers, but those
    // markers must remain present in the native extent.
    assert_eq!(commit.package().body_table_hidden_axes(0usize)?, requested);
    let reopened_filtered = Package::from_bytes(&commit.package().exact_bytes())?;
    assert_eq!(reopened_filtered.body_table_hidden_axes(0usize)?, requested);
    for (field, expected) in [(14, 2), (15, 2), (40, 1), (41, 1), (42, 1)] {
        assert_eq!(
            model_count_field(&commit.package().exact_bytes(), 0, field)?,
            Some(expected),
            "mixed filtered/user count field {field}"
        );
    }
    let owner = hidden_owner_payload_from_package(&commit.package().exact_bytes(), 0)?;
    let owner_view = raw_wire_fields(&owner)?;
    assert!(
        owner_view
            .iter()
            .any(|field| field.number() == UNKNOWN_OWNER_FIELD)
    );
    assert!(contains_field_at_path(
        &owner,
        &[2, 2],
        UNKNOWN_EXTENT_FIELD
    )?);
    assert!(contains_field_at_path(
        &owner,
        &[2, 3],
        UNKNOWN_EXTENT_FIELD
    )?);
    assert!(contains_field_at_path(
        &owner,
        &[2, 2],
        UNKNOWN_FIXED32_FIELD
    )?);
    assert!(contains_field_at_path(
        &owner,
        &[2, 2],
        UNKNOWN_LENGTH_FIELD
    )?);
    assert!(contains_field_at_path(
        &owner,
        &[2, 3],
        UNKNOWN_GROUP_FIELD
    )?);
    assert!(contains_field_at_path(
        &owner,
        &[2, 3],
        UNKNOWN_LENGTH_FIELD
    )?);
    assert!(contains_field_at_path(&owner, &[], UNKNOWN_GROUP_FIELD)?);
    assert!(contains_field_at_path(&owner, &[], UNKNOWN_LENGTH_FIELD)?);
    assert_eq!(
        raw_fields_at_path(
            &source_model_payload,
            &[],
            &[UNKNOWN_MODEL_FIELD, UNKNOWN_LENGTH_FIELD],
        )?,
        raw_fields_at_path(
            &model_payload_from_package(&commit.package().exact_bytes(), 0)?,
            &[],
            &[UNKNOWN_MODEL_FIELD, UNKNOWN_LENGTH_FIELD],
        )?
    );
    assert_eq!(
        raw_fields_at_path(
            &source_info_payload,
            &[],
            &[UNKNOWN_INFO_FIELD, UNKNOWN_LENGTH_FIELD],
        )?,
        raw_fields_at_path(
            &info_payload(&commit.package().exact_bytes(), 0)?,
            &[],
            &[UNKNOWN_INFO_FIELD, UNKNOWN_LENGTH_FIELD],
        )?
    );
    for (path, numbers) in [
        (
            &[][..],
            &[
                UNKNOWN_OWNER_FIELD,
                UNKNOWN_GROUP_FIELD,
                UNKNOWN_LENGTH_FIELD,
            ][..],
        ),
        (
            &[2, 2][..],
            &[
                UNKNOWN_EXTENT_FIELD,
                UNKNOWN_FIXED32_FIELD,
                UNKNOWN_LENGTH_FIELD,
            ][..],
        ),
        (
            &[2, 3][..],
            &[
                UNKNOWN_EXTENT_FIELD,
                UNKNOWN_GROUP_FIELD,
                UNKNOWN_LENGTH_FIELD,
            ][..],
        ),
    ] {
        assert_eq!(
            raw_fields_at_path(&source_owner_payload, path, numbers)?,
            raw_fields_at_path(&owner, path, numbers)?,
            "interleaved unknown fields changed at path {path:?}"
        );
    }
    let owner = tst::HiddenStatesOwnerArchive::decode(owner.as_slice())?;
    let active = owner
        .hidden_states
        .first()
        .ok_or("rewritten hidden-state owner lost its active state")?;
    assert!(
        active
            .row_hidden_state_extent
            .base_hidden_states
            .iter()
            .any(|state| state.filtered == Some(true))
    );
    assert!(
        active
            .column_hidden_state_extent
            .base_hidden_states
            .iter()
            .any(|state| state.pivot_hidden == Some(true))
    );
    assert_eq!(
        active
            .column_hidden_state_extent
            .needs_to_update_filter_set_for_import,
        Some(false)
    );
    assert_eq!(
        active
            .row_hidden_state_extent
            .needs_to_update_filter_set_for_import,
        Some(false)
    );
    for (identifier, type_, before_payload) in helper_payloads {
        let after_payload =
            object_message_payload(&commit.package().exact_bytes(), identifier, type_)?;
        assert_eq!(
            &after_payload, &before_payload,
            "helper object {identifier} changed during a hidden-axis rewrite"
        );
        let expected_unknown = match identifier {
            id if id == table_formula_owner(0) => UNKNOWN_FIXED64_FIELD,
            id if id == table_formula_object(0, true) || id == table_formula_object(0, false) => {
                UNKNOWN_GROUP_FIELD
            },
            _ => UNKNOWN_FIXED32_FIELD,
        };
        assert!(contains_field_at_path(
            &before_payload,
            &[],
            expected_unknown
        )?);
        assert!(contains_field_at_path(
            &before_payload,
            &[],
            UNKNOWN_LENGTH_FIELD
        )?);
        if identifier == table_filter_set(0, true) || identifier == table_filter_set(0, false) {
            let filter = tst::FilterSetArchive::decode(after_payload.as_slice())?;
            assert_eq!(filter.is_enabled, Some(false));
            assert_eq!(filter.needs_formula_rewrite_for_import, Some(false));
        }
        if identifier == table_formula_object(0, true)
            || identifier == table_formula_object(0, false)
        {
            let formula = tst::HiddenStateFormulaOwnerArchive::decode(after_payload.as_slice())?;
            assert_eq!(formula.needs_to_update_filter_set_for_import, Some(false));
        }
    }

    let pivot = Package::from_bytes(&pivot_package()?)?;
    let before = pivot.exact_bytes();
    let pivot_noop = pivot
        .edit_body_table_hidden_axes(0usize)?
        .set(pivot.body_table_hidden_axes(0usize)?)
        .commit()?;
    assert!(pivot_noop.patch().is_noop());
    assert_eq!(pivot_noop.package().exact_bytes(), before);
    let result = pivot
        .edit_body_table_hidden_axes(0usize)?
        .set(HiddenAxes::new([AxisIndex::row(0)])?)
        .commit();
    assert!(
        matches!(
            &result,
            Err(Error::UnsupportedDependency | Error::InvalidSource)
        ),
        "pivot topology must be refused: {result:?}"
    );
    assert_eq!(pivot.exact_bytes(), before);
    Ok(())
}

#[test]
fn explicit_visible_markers_survive_existing_owner_rewrites() -> TestResult {
    let source = normal_package()?;
    let package = Package::from_bytes(&source)?;
    let requested = HiddenAxes::new([AxisIndex::row(3), AxisIndex::column(2)])?;
    let commit = package
        .edit_body_table_hidden_axes(0usize)?
        .set(requested.clone())
        .commit()?;
    assert_eq!(commit.package().body_table_hidden_axes(0usize)?, requested);

    let (rows, columns) = table_uids(0);
    let owner = tst::HiddenStatesOwnerArchive::decode(
        hidden_owner_payload_from_package(&commit.package().exact_bytes(), 0)?.as_slice(),
    )?;
    let active = owner
        .hidden_states
        .first()
        .ok_or("rewritten hidden-state owner lost its active state")?;
    let row_states = &active.row_hidden_state_extent.base_hidden_states;
    let column_states = &active.column_hidden_state_extent.base_hidden_states;

    assert_eq!(
        row_states
            .iter()
            .find(|state| state.row_or_column_uid == rows[0])
            .and_then(|state| state.user_hidden),
        Some(false),
        "an explicit visible row marker must survive when the row remains visible"
    );
    assert_eq!(
        column_states
            .iter()
            .find(|state| state.row_or_column_uid == columns[0])
            .and_then(|state| state.user_hidden),
        Some(false),
        "an explicit visible column marker must survive when the column remains visible"
    );
    assert!(
        row_states
            .iter()
            .find(|state| state.row_or_column_uid == rows[1])
            .is_some_and(|state| state.user_hidden.is_none()),
        "clearing visibility preserves the source state identity and removes the hidden flag"
    );
    assert_eq!(
        column_states
            .iter()
            .find(|state| state.row_or_column_uid == columns[2])
            .and_then(|state| state.user_hidden),
        Some(true)
    );
    Ok(())
}

#[test]
fn descendant_formula_and_filter_metadata_is_not_silently_accepted() -> TestResult {
    let source = normal_package()?;
    let malformed_sources = [
        formula_owner_owner_kind(&source, 0)?,
        formula_owner_base_uid(&source, 0)?,
        formula_owner_cell_dependencies(&source, 0)?,
        formula_owner_range_dependencies(&source, 0)?,
        formula_owner_uuid_references(&source, 0)?,
        formula_owner_bad_message_type(&source, table_formula_owner(0))?,
        formula_owner_message_info_type_mismatch(&source, 0)?,
        formula_owner_message_info_length_mismatch(&source, 0)?,
        formula_owner_bad_versions(&source, table_formula_owner(0))?,
        formula_owner_bad_merge(&source, table_formula_owner(0))?,
        formula_owner_bad_base(&source, table_formula_owner(0))?,
        formula_owner_bad_diff(&source, table_formula_owner(0))?,
        hidden_formula_owner_conflicting_cfuuid(&source, table_formula_object(0, true))?,
        hidden_formula_owner_bad_metadata(&source, table_formula_object(0, true))?,
        filter_set_prepivot_rules(&source, table_filter_set(0, true))?,
        filter_set_enabled_vector(&source, table_filter_set(0, true))?,
        filter_set_rules(&source, table_filter_set(0, true))?,
        filter_set_out_of_range_offset(&source, table_filter_set(0, true))?,
    ];
    let rejected_axes = HiddenAxes::new([AxisIndex::row(0)])?;
    for (case_index, bytes) in malformed_sources.into_iter().enumerate() {
        let Ok(package) = Package::from_bytes(&bytes) else {
            continue;
        };
        let before = package.exact_bytes();
        let result = package
            .edit_body_table_hidden_axes(0usize)
            .and_then(|edit| edit.set(rejected_axes.clone()).commit());
        assert!(
            result.is_err(),
            "descendant metadata was accepted (case {case_index})"
        );
        assert_eq!(package.exact_bytes(), before);
    }
    Ok(())
}

#[test]
fn pivot_owner_is_distinct_from_merge_owner_and_bool_values_are_strict() -> TestResult {
    let source = normal_package()?;
    let pivot_source = pivot_owner_only(&source, 0)?;
    let pivot = Package::from_bytes(&pivot_source)?;
    let pivot_before = pivot.exact_bytes();
    let pivot_result = pivot
        .edit_body_table_hidden_axes(0usize)?
        .set(HiddenAxes::new([AxisIndex::row(0)])?)
        .commit();
    assert!(matches!(pivot_result, Err(Error::UnsupportedDependency)));
    assert_eq!(pivot.exact_bytes(), pivot_before);

    let merge_source = merge_owner_only(&source, 0)?;
    let merge = Package::from_bytes(&merge_source)?;
    let merge_commit = merge
        .edit_body_table_hidden_axes(0usize)?
        .set(HiddenAxes::new([AxisIndex::row(0)])?)
        .commit()?;
    assert_eq!(
        merge_commit.package().body_table_hidden_axes(0usize)?,
        HiddenAxes::new([AxisIndex::row(0)])?
    );

    let bool_source = invalid_extent_boolean(&source, 0)?;
    let bool_package = Package::from_bytes(&bool_source)?;
    let bool_before = bool_package.exact_bytes();
    let bool_result = bool_package.body_table_hidden_axes(0usize);
    assert!(matches!(bool_result, Err(Error::InvalidSource)));
    assert_eq!(bool_package.exact_bytes(), bool_before);
    Ok(())
}

fn model_count_field(package: &[u8], index: usize, number: u32) -> TestResult<Option<u64>> {
    let payload = model_payload_from_package(package, index)?;
    WireView::parse(&payload)?
        .fields()
        .find(|field| field.number() == number)
        .map(|field| decode_varint_from_bytes(field.payload()).map(|(value, _)| value))
        .transpose()
        .map_err(Into::into)
}

#[test]
fn model_hidden_count_presence_tracks_user_and_non_user_state() -> TestResult {
    let filtered_source = filtered_package()?;
    for (field, expected) in [(14, 2), (15, 2), (40, 1), (41, 1), (42, 1)] {
        assert_eq!(
            model_count_field(&filtered_source, 0, field)?,
            Some(expected),
            "filtered source count field {field}"
        );
    }
    let filtered = Package::from_bytes(&filtered_source)?;
    let filtered_clear = filtered
        .edit_body_table_hidden_axes(0usize)?
        .clear()
        .commit()?;
    let filtered_target = filtered_clear.package().exact_bytes();
    for (field, expected) in [(14, 1), (15, 1), (40, 1), (41, 0), (42, 0)] {
        assert_eq!(
            model_count_field(&filtered_target, 0, field)?,
            Some(expected),
            "filtered clear count field {field}"
        );
    }
    assert_eq!(
        Package::from_bytes(&filtered_target)?.body_table_hidden_axes(0usize)?,
        HiddenAxes::empty()
    );

    let normal_source = normal_package()?;
    let normal = Package::from_bytes(&normal_source)?;
    let normal_clear = normal
        .edit_body_table_hidden_axes(0usize)?
        .clear()
        .commit()?;
    let normal_target = normal_clear.package().exact_bytes();
    for (field, expected) in [(14, 0), (15, 0), (41, 0), (42, 0)] {
        assert_eq!(
            model_count_field(&normal_target, 0, field)?,
            Some(expected),
            "normal clear count field {field}"
        );
    }
    assert_eq!(
        model_count_field(&normal_target, 0, 40)?,
        None,
        "an originally absent filtered-row count must remain absent"
    );

    let zero_source = explicit_zero_count_package()?;
    assert_eq!(model_count_field(&zero_source, 0, 40)?, Some(0));
    let zero = Package::from_bytes(&zero_source)?;
    let zero_clear = zero.edit_body_table_hidden_axes(0usize)?.clear().commit()?;
    let zero_target = zero_clear.package().exact_bytes();
    assert_eq!(
        model_count_field(&zero_target, 0, 40)?,
        Some(0),
        "an explicit zero filtered-row count must remain present"
    );
    for field in [41, 42] {
        assert_eq!(
            model_count_field(&zero_target, 0, field)?,
            Some(0),
            "an explicit user-hidden count field {field} must remain present"
        );
    }
    Ok(())
}

#[test]
fn unknown_model_info_and_nested_owner_wire_fields_survive_rewrite() -> TestResult {
    let source = normal_package()?;
    let package = Package::from_bytes(&source)?;
    let requested = HiddenAxes::new([AxisIndex::row(0), AxisIndex::column(3)])?;
    let commit = package
        .edit_body_table_hidden_axes(0usize)?
        .set(requested)
        .commit()?;
    let target = commit.package().exact_bytes();

    let model_payload = model_payload_from_package(&target, 0)?;
    let model = WireView::parse(&model_payload)?;
    let model_unknown = model
        .fields()
        .find(|field| field.number() == UNKNOWN_MODEL_FIELD)
        .ok_or("unknown model field was dropped")?;
    assert_eq!(
        decode_varint_from_bytes(model_unknown.payload())?.0,
        UNKNOWN_MODEL_VALUE
    );
    assert!(
        model
            .fields()
            .any(|field| field.number() == UNKNOWN_LENGTH_FIELD)
    );
    let info_payload = info_payload(&target, 0)?;
    let info = WireView::parse(&info_payload)?;
    let info_unknown = info
        .fields()
        .find(|field| field.number() == UNKNOWN_INFO_FIELD)
        .ok_or("unknown table-info field was dropped")?;
    assert_eq!(
        decode_varint_from_bytes(info_unknown.payload())?.0,
        UNKNOWN_INFO_VALUE
    );
    assert!(
        info.fields()
            .any(|field| field.number() == UNKNOWN_LENGTH_FIELD)
    );
    let owner = hidden_owner_payload_from_package(&target, 0)?;
    let owner_unknown = raw_wire_fields(&owner)?
        .into_iter()
        .find(|field| field.number() == UNKNOWN_OWNER_FIELD)
        .ok_or("unknown hidden-owner field was dropped")?;
    assert_eq!(
        decode_varint_from_bytes(owner_unknown.payload())?.0,
        UNKNOWN_OWNER_VALUE
    );
    Ok(())
}

#[test]
fn rewrite_is_local_to_selected_table_and_invalidates_only_root_previews() -> TestResult {
    let source = normal_package()?;
    let package = Package::from_bytes(&source)?;
    let second_model_before = model_payload_from_package(&source, 1)?;
    let second_info_before = info_payload(&source, 1)?;
    let sentinel_before = member_bytes(&source, "Data/sentinel.bin")?;
    let unselected_before = unselected_object_snapshot(&source, 0)?;
    let selected_non_target_before = selected_non_target_object_snapshot(&source, 0)?;
    let object_order_before = archive_object_order(&document_archive(&source)?)?;
    let commit = package
        .edit_body_table_hidden_axes(BodyTableSelector::name("Revenue"))?
        .set(HiddenAxes::new([AxisIndex::row(0), AxisIndex::column(3)])?)
        .commit()?;
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert_eq!(commit.diagnostics().deleted_previews(), PREVIEWS.len());
    assert!(commit.diagnostics().full_reparse_performed());
    let target = commit.package().exact_bytes();
    assert_eq!(model_payload_from_package(&target, 1)?, second_model_before);
    assert_eq!(info_payload(&target, 1)?, second_info_before);
    assert_eq!(member_bytes(&target, "Data/sentinel.bin")?, sentinel_before);
    let target_archive = document_archive(&target)?;
    assert_eq!(archive_object_order(&target_archive)?, object_order_before);
    for (identifier, before) in selected_non_target_before {
        let after = target_archive
            .object(identifier)
            .ok_or_else(|| format!("selected helper object {identifier} disappeared"))?;
        assert!(
            before.same_content_ignoring_offsets(after),
            "selected non-target object {identifier} or its metadata changed"
        );
    }
    for (identifier, before) in unselected_before {
        let after = target_archive
            .object(identifier)
            .ok_or_else(|| format!("unselected object {identifier} disappeared"))?;
        assert!(
            before.same_content_ignoring_offsets(after),
            "unselected object {identifier} or its archive metadata changed"
        );
    }
    let catalog = Catalog::from_bytes(&target)?;
    for preview in PREVIEWS {
        assert!(catalog.iter().all(|entry| entry.name() != preview));
    }
    for preview in NONCANONICAL_PREVIEWS {
        assert!(
            catalog.iter().any(|entry| entry.name() == preview),
            "non-canonical preview-like member {preview} was removed"
        );
    }
    Ok(())
}

#[test]
fn selected_object_retains_noncanonical_header_while_payload_framing_changes_locally() -> TestResult
{
    let source = with_noncanonical_model_header(&normal_package()?, 0)?;
    let source_header = object_header_bytes(&source, table_model(0))?;
    let source_header_view = WireView::parse(&source_header)?;
    assert!(
        source_header_view
            .fields()
            .any(|field| field.number() == UNKNOWN_LENGTH_FIELD)
    );
    let source_object = document_archive(&source)?
        .object(table_model(0))
        .ok_or("missing selected model")?
        .clone();

    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_body_table_hidden_axes(0usize)?
        .set(HiddenAxes::new([
            AxisIndex::row(0),
            AxisIndex::row(1),
            AxisIndex::row(2),
            AxisIndex::row(3),
            AxisIndex::column(0),
            AxisIndex::column(1),
            AxisIndex::column(2),
            AxisIndex::column(3),
        ])?)
        .commit()?;
    let target = commit.package().exact_bytes();
    let target_header = object_header_bytes(&target, table_model(0))?;
    let target_header_view = WireView::parse(&target_header)?;
    let retained = target_header_view
        .fields()
        .find(|field| field.number() == UNKNOWN_LENGTH_FIELD)
        .ok_or("selected object's retained header field was dropped")?;
    assert_eq!(decode_varint_from_bytes(retained.payload())?.0, 0xdecafbad,);
    let target_archive = document_archive(&target)?;
    let target_object = target_archive
        .object(table_model(0))
        .ok_or("missing rewritten selected model")?;
    assert_eq!(
        target_object.archive_info.identifier,
        source_object.archive_info.identifier
    );
    assert_eq!(target_object.header_length, source_object.header_length);
    assert_ne!(
        target_object.data_length, source_object.data_length,
        "the selected payload framing should reflect the changed owner"
    );
    Ok(())
}

#[test]
fn output_and_wire_limits_reject_before_publication() -> TestResult {
    let source = without_previews(&normal_package()?)?;
    let package = Package::from_bytes(&source)?;
    // Native Pages does not synthesize a hidden-state owner for an otherwise
    // valid table.  Use the existing owner on table 0 when exercising output
    // growth and the output-byte limit.
    let requested = HiddenAxes::new([
        AxisIndex::row(0),
        AxisIndex::row(1),
        AxisIndex::row(2),
        AxisIndex::row(3),
        AxisIndex::column(0),
        AxisIndex::column(1),
        AxisIndex::column(2),
        AxisIndex::column(3),
    ])?;
    let unrestricted = package
        .edit_body_table_hidden_axes(0usize)?
        .set(requested.clone())
        .commit()?;
    let target = unrestricted.package().exact_bytes();
    assert!(
        target.len() > source.len(),
        "rewriting existing hidden-state ownership should grow the package"
    );
    assert_eq!(
        unrestricted.diagnostics().deleted_previews(),
        0,
        "a source with no canonical previews has nothing lifecycle-owned to delete"
    );

    let partial_source = without_preview_subset(&normal_package()?, &PREVIEWS[1..])?;
    let partial = Package::from_bytes(&partial_source)?;
    let partial_commit = partial
        .edit_body_table_hidden_axes(0usize)?
        .set(requested.clone())
        .commit()?;
    assert_eq!(partial_commit.diagnostics().deleted_previews(), 1);
    let partial_catalog = Catalog::from_bytes(&partial_commit.package().exact_bytes())?;
    assert!(
        partial_catalog
            .iter()
            .all(|entry| entry.name() != PREVIEWS[0])
    );
    for preview in NONCANONICAL_PREVIEWS {
        assert!(partial_catalog.iter().any(|entry| entry.name() == preview));
    }
    let bounded = Limits::new(
        u64::try_from(source.len())?,
        32,
        1024 * 1024,
        1024 * 1024,
        1024 * 1024,
    )?;
    let bounded_package = Package::from_bytes_with_limits(&source, bounded)?;
    let before = bounded_package.exact_bytes();
    let result = bounded_package
        .edit_body_table_hidden_axes(0usize)?
        .set(requested.clone())
        .commit();
    assert!(matches!(result, Err(Error::LimitExceeded { .. })));
    assert_eq!(bounded_package.exact_bytes(), before);

    let normal_source = normal_package()?;
    let mut successful_fields = None;
    for fields in 1..=4_096 {
        let archive_limits = litchi_iwa_core::Limits::default().with_header_fields(fields)?;
        let limits = Limits::default().with_archive_limits(archive_limits)?;
        match Package::from_bytes_with_limits(&normal_source, limits) {
            Ok(package) => {
                let result = package
                    .edit_body_table_hidden_axes(0usize)
                    .map(|edit| edit.set(requested.clone()))
                    .and_then(|edit| edit.commit());
                if result.is_ok() {
                    successful_fields = Some(fields);
                    break;
                }
            },
            Err(error) => {
                assert!(
                    is_package_limit_error(&error),
                    "unexpected header-field preflight error: {error:?}"
                );
            },
        }
    }
    let exact_fields =
        successful_fields.ok_or("no finite hidden-axis wire limit boundary was observed")?;
    assert!(exact_fields > 1);
    let archive_limits = litchi_iwa_core::Limits::default().with_header_fields(exact_fields - 1)?;
    let limits = Limits::default().with_archive_limits(archive_limits)?;
    match Package::from_bytes_with_limits(&normal_source, limits) {
        Ok(package) => {
            let before = package.exact_bytes();
            let result = package
                .edit_body_table_hidden_axes(0usize)
                .map(|edit| edit.set(requested.clone()))
                .and_then(|edit| edit.commit());
            assert!(matches!(result, Err(Error::LimitExceeded { .. })));
            assert_eq!(package.exact_bytes(), before);
        },
        Err(error) => {
            assert!(
                is_package_limit_error(&error),
                "unexpected below-boundary preflight error: {error:?}"
            );
        },
    }
    Ok(())
}

#[test]
fn public_hidden_axis_transactions_are_send_sync_debug_and_redacted() -> TestResult {
    fn assert_send_sync_debug<T: Send + Sync + std::fmt::Debug>() {}

    fn assert_redacted(output: &str) {
        // Public Debug/Display output may report semantic state counts and
        // booleans, but it must not become a second package-inspection API.
        for forbidden in [
            DOCUMENT_MEMBER,
            "Data/sentinel.bin",
            "preview.jpg",
            "untouched-sentinel",
            "Revenue",
            "Costs",
            "fingerprint",
            "Row(0)",
            "Row(1)",
            "Column(2)",
            "Column(3)",
            "row 0",
            "row 1",
            "column 2",
            "column 3",
            // Native fixture identifiers and UUID words.  These are
            // intentionally distinct from the small count diagnostics.
            "300",
            "400",
            "500",
            "600",
            "700",
            "12292",
            "12296",
            "12299",
            "256",
        ] {
            assert!(
                !output.contains(forbidden),
                "public hidden-axis output leaked {forbidden:?}: {output}"
            );
        }
    }

    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<AxisIndex>();
    assert_send_sync_debug::<HiddenAxes>();
    assert_send_sync_debug::<litchi_pages::BodyTableHiddenAxesEdit<'static>>();
    assert_send_sync_debug::<litchi_pages::BodyTableHiddenAxesPatch>();
    assert_send_sync_debug::<litchi_pages::BodyTableHiddenAxesCommit>();
    assert_send_sync_debug::<litchi_pages::BodyTableHiddenAxesDiagnostics>();
    assert_send_sync_debug::<Error>();
    assert_send_sync_debug::<BodyTableHiddenAxesLimitKind>();

    let package = Package::from_bytes(&normal_package()?)?;
    let edit = package.edit_body_table_hidden_axes(0usize)?;
    assert_redacted(&format!("{:?}", edit.path()));
    assert_redacted(&format!("{edit:?}"));
    let commit = edit.set(HiddenAxes::new([AxisIndex::row(0)])?).commit()?;
    assert_redacted(&format!("{:?}", commit.patch().path()));
    assert_redacted(&format!("{:?}", commit.patch()));
    assert_redacted(&format!("{:?}", commit.patch().inverse()));
    assert_redacted(&format!("{commit:?}"));
    assert_redacted(&format!("{:?}", commit.diagnostics()));
    for error in [
        Error::TableNotFound,
        Error::TableLocked,
        Error::PatchConflict,
        Error::LimitExceeded {
            kind: BodyTableHiddenAxesLimitKind::WireBytes,
            observed: 17,
            maximum: 23,
        },
        Error::Allocation { amount: 5 },
    ] {
        assert_redacted(&format!("{error:?}"));
        assert_redacted(&error.to_string());
    }
    Ok(())
}

#[test]
fn cloned_packages_are_copy_on_write_and_safe_for_concurrent_reads_and_edits() -> TestResult {
    let source = normal_package()?;
    let package = Arc::new(Package::from_bytes(&source)?);
    let source_before = package.exact_bytes();
    let mut handles = Vec::new();
    for (thread_index, requested) in [
        HiddenAxes::new([AxisIndex::row(0)])?,
        HiddenAxes::new([AxisIndex::column(0)])?,
        HiddenAxes::new([AxisIndex::row(3), AxisIndex::column(3)])?,
        HiddenAxes::empty(),
    ]
    .into_iter()
    .enumerate()
    {
        let package = Arc::clone(&package);
        let requested_is_empty = requested.is_empty();
        handles.push(thread::spawn(move || {
            let selector = if thread_index % 2 == 0 {
                0usize
            } else {
                1usize
            };
            let _ = package
                .body_table_hidden_axes(selector)
                .expect("concurrent hidden-axis read succeeds");
            let edit = package
                .edit_body_table_hidden_axes(selector)
                .expect("concurrent hidden-axis edit opens")
                .set(requested);
            if selector == 1 {
                let result = edit.commit();
                if requested_is_empty {
                    assert!(
                        matches!(&result, Ok(commit) if commit.patch().is_noop()),
                        "an unchanged absent-owner edit must be a no-op: {result:?}"
                    );
                } else {
                    assert!(
                        matches!(
                            &result,
                            Err(Error::UnsupportedDependency | Error::UnsupportedSource)
                        ),
                        "absent-owner creation must remain unsupported: {result:?}"
                    );
                }
                package.exact_bytes()
            } else {
                edit.commit()
                    .expect("existing-owner edit publishes")
                    .package()
                    .exact_bytes()
            }
        }));
    }
    for handle in handles {
        let result = handle.join().expect("concurrent hidden-axis edit panicked");
        assert!(!result.is_empty());
    }
    assert_eq!(package.exact_bytes(), source_before);

    // A second immutable owner observes the same source after another owner
    // publishes a candidate, proving that edits use COW package state.
    let clone = (*package).clone();
    let candidate = clone
        .edit_body_table_hidden_axes(0usize)?
        .set(HiddenAxes::new([AxisIndex::row(2)])?)
        .commit()?;
    assert_eq!(package.exact_bytes(), source_before);
    assert_ne!(candidate.package().exact_bytes(), source_before);
    Ok(())
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
        litchi_iwa_common::wire::append_varint_field(
            &mut payload,
            UNKNOWN_MODEL_FIELD,
            UNKNOWN_MODEL_VALUE,
        )?;
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

fn remove_model_map_field(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model_metadata(package, index, |model, message_index| {
        let info = &mut model.archive_info.message_infos[message_index];
        info.field_infos
            .retain(|field| field.path.as_slice() != [46]);
        Ok(())
    })
}

fn remove_model_map_payload_field(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    let payload = model_payload_from_package(package, index)?;
    rewrite_model_payload(package, index, remove_first_raw_field(&payload, 46)?)
}

fn wrong_model_map_field_path(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model_metadata(package, index, |model, message_index| {
        let field = model.archive_info.message_infos[message_index]
            .field_infos
            .iter_mut()
            .find(|field| field.path.as_slice() == [46])
            .ok_or("missing model map field declaration")?;
        field.path = FieldPath::new(vec![47]);
        Ok(())
    })
}

fn model_map_data_reference(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model_metadata(package, index, |model, message_index| {
        model.archive_info.message_infos[message_index]
            .data_references
            .push(table_uid_map(index));
        Ok(())
    })
}

fn model_map_extra_aggregate_reference(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model_metadata(package, index, |model, message_index| {
        model.archive_info.message_infos[message_index]
            .object_references
            .push(table_uid_map(index));
        Ok(())
    })
}

fn model_external_map_reference(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model(package, index, |model| {
        model.base_column_row_uids = Some(external_reference(table_uid_map(index)));
    })
}

fn remove_info_map_field(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_info_metadata(package, index, |info, message_index| {
        let metadata = &mut info.archive_info.message_infos[message_index];
        metadata
            .field_infos
            .retain(|field| field.path.as_slice() != [6]);
        Ok(())
    })
}

fn wrong_info_map_field_path(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_info_metadata(package, index, |info, message_index| {
        let field = info.archive_info.message_infos[message_index]
            .field_infos
            .iter_mut()
            .find(|field| field.path.as_slice() == [6])
            .ok_or("missing info map field declaration")?;
        field.path = FieldPath::new(vec![7]);
        Ok(())
    })
}

fn info_map_data_reference(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_info_metadata(package, index, |info, message_index| {
        info.archive_info.message_infos[message_index]
            .data_references
            .push(table_uid_map(index));
        Ok(())
    })
}

fn info_map_extra_aggregate_reference(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_info_metadata(package, index, |info, message_index| {
        info.archive_info.message_infos[message_index]
            .object_references
            .push(table_uid_map(index));
        Ok(())
    })
}

fn info_external_map_reference(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_info_metadata(package, index, |info, message_index| {
        let payload = info.messages[message_index].data.as_slice();
        let mut decoded = tst::TableInfoArchive::decode(payload)?;
        decoded.view_column_row_uids = Some(external_reference(table_uid_map(index)));
        info.messages[message_index].data = decoded.encode_to_vec();
        Ok(())
    })
}

fn disagreeing_info_map_reference(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    debug_assert_eq!(index, 1);
    rewrite_info_metadata(package, index, |info, message_index| {
        let mut decoded =
            tst::TableInfoArchive::decode(info.messages[message_index].data.as_slice())?;
        decoded.view_column_row_uids = Some(reference(table_uid_map(0)));
        info.messages[message_index].data = decoded.encode_to_vec();
        let metadata = &mut info.archive_info.message_infos[message_index];
        for identifier in &mut metadata.object_references {
            if *identifier == table_uid_map(index) {
                *identifier = table_uid_map(0);
            }
        }
        for field in &mut metadata.field_infos {
            if field.path.as_slice() == [6] {
                field.object_references = vec![table_uid_map(0)];
            }
        }
        Ok(())
    })
}

fn disagreeing_model_map_reference(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    debug_assert_eq!(index, 1);
    rewrite_model_metadata(package, index, |model, message_index| {
        let payload = model.messages[message_index].data.as_slice();
        let mut decoded = tst::TableModelArchive::decode(payload)?;
        decoded.base_column_row_uids = Some(reference(table_uid_map(0)));
        model.messages[message_index].data = decoded.encode_to_vec();
        let metadata = &mut model.archive_info.message_infos[message_index];
        for identifier in &mut metadata.object_references {
            if *identifier == table_uid_map(index) {
                *identifier = table_uid_map(0);
            }
        }
        for field in &mut metadata.field_infos {
            if field.path.as_slice() == [46] {
                field.object_references = vec![table_uid_map(0)];
            }
        }
        Ok(())
    })
}

fn rewrite_object_metadata(
    package: &[u8],
    identifier: u64,
    message_type: u32,
    mutate: impl FnOnce(&mut ArchiveObject, usize) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let object = archive
            .object_mut(identifier)
            .ok_or_else(|| format!("missing object {identifier}"))?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == message_type)
            .ok_or_else(|| format!("missing message type {message_type}"))?;
        mutate(object, message_index)
    })
}

fn rewrite_formula_owner_payload(
    package: &[u8],
    index: usize,
    mutate: impl FnOnce(&mut tsce::FormulaOwnerDependenciesArchive),
) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let object = archive
            .object_mut(table_formula_owner(index))
            .ok_or("missing formula-owner object")?;
        let message = object
            .messages
            .iter_mut()
            .find(|message| message.type_ == FORMULA_OWNER_MESSAGE_TYPE)
            .ok_or("missing formula-owner message")?;
        let mut decoded = tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())?;
        mutate(&mut decoded);
        message.data = decoded.encode_to_vec();
        Ok(())
    })
}

fn rewrite_hidden_formula_owner_payload(
    package: &[u8],
    identifier: u64,
    mutate: impl FnOnce(&mut tst::HiddenStateFormulaOwnerArchive),
) -> TestResult<Vec<u8>> {
    rewrite_object_payload(
        package,
        identifier,
        HIDDEN_STATE_FORMULA_OWNER_MESSAGE_TYPE,
        |data| {
            let mut decoded = tst::HiddenStateFormulaOwnerArchive::decode(data)?;
            mutate(&mut decoded);
            Ok(decoded.encode_to_vec())
        },
    )
}

fn rewrite_filter_set_payload(
    package: &[u8],
    identifier: u64,
    mutate: impl FnOnce(&mut tst::FilterSetArchive),
) -> TestResult<Vec<u8>> {
    rewrite_object_payload(package, identifier, FILTER_SET_MESSAGE_TYPE, |data| {
        let mut decoded = tst::FilterSetArchive::decode(data)?;
        mutate(&mut decoded);
        Ok(decoded.encode_to_vec())
    })
}

fn rewrite_object_payload(
    package: &[u8],
    identifier: u64,
    message_type: u32,
    mutate: impl FnOnce(&[u8]) -> TestResult<Vec<u8>>,
) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let object = archive
            .object_mut(identifier)
            .ok_or_else(|| format!("missing object {identifier}"))?;
        let message = object
            .messages
            .iter_mut()
            .find(|message| message.type_ == message_type)
            .ok_or_else(|| format!("missing message type {message_type}"))?;
        message.data = mutate(&message.data)?;
        Ok(())
    })
}

fn formula_owner_owner_kind(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_formula_owner_payload(package, index, |owner| owner.owner_kind = Some(1))
}

fn formula_owner_base_uid(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_formula_owner_payload(package, index, |owner| {
        owner.base_owner_uid = Some(uuid(0x9000, 0x9001));
    })
}

fn formula_owner_cell_dependencies(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_formula_owner_payload(package, index, |owner| {
        owner.cell_dependencies = Some(tsce::CellDependenciesExpandedArchive::default());
    })
}

fn formula_owner_range_dependencies(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_formula_owner_payload(package, index, |owner| {
        owner.range_dependencies = Some(tsce::RangeDependenciesArchive::default());
    })
}

fn formula_owner_uuid_references(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_formula_owner_payload(package, index, |owner| {
        owner.uuid_references = Some(tsce::UuidReferencesArchive::default());
    })
}

fn formula_owner_external_reference(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_formula_owner_payload(package, index, |owner| {
        owner.formula_owner = Some(external_reference(table_drawable(index)));
    })
}

fn hidden_extent_external_filter_reference(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_model(package, index, |model| {
        if let Some(owner) = model.hidden_states_owner.as_mut()
            && let Some(active) = owner.hidden_states.first_mut()
        {
            active.row_hidden_state_extent.filter_set =
                Some(external_reference(table_filter_set(index, false)));
        }
    })
}

fn hidden_formula_owner_conflicting_cfuuid(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    rewrite_hidden_formula_owner_payload(package, identifier, |owner| {
        if let Some(owner_id) = owner.owner_id.as_mut() {
            owner_id.uuid_bytes = Some(vec![0; 16]);
        }
    })
}

fn hidden_formula_owner_bad_metadata(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    rewrite_object_metadata(
        package,
        identifier,
        HIDDEN_STATE_FORMULA_OWNER_MESSAGE_TYPE,
        |object, message_index| {
            object.archive_info.message_infos[message_index].diff_read_version = vec![1];
            Ok(())
        },
    )
}

fn filter_set_prepivot_rules(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    rewrite_filter_set_payload(package, identifier, |filter| {
        filter
            .filter_rules_prepivot
            .push(tst::FilterRulePrePivotArchive::default());
    })
}

fn filter_set_enabled_vector(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    rewrite_filter_set_payload(package, identifier, |filter| {
        filter.filter_enabled.push(true);
    })
}

fn filter_set_rules(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    rewrite_filter_set_payload(package, identifier, |filter| {
        filter.filter_rules.push(tst::FilterRuleArchive::default());
    })
}

fn filter_set_out_of_range_offset(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    rewrite_filter_set_payload(package, identifier, |filter| {
        filter.filter_offsets = vec![u32::MAX];
    })
}

fn formula_owner_bad_message_type(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    rewrite_object_metadata(
        package,
        identifier,
        FORMULA_OWNER_MESSAGE_TYPE,
        |object, message_index| {
            object.messages[message_index].type_ = FORMULA_OWNER_MESSAGE_TYPE + 1;
            Ok(())
        },
    )
}

fn formula_owner_message_info_type_mismatch(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    rewrite_raw_object_header(package, table_formula_owner(index), |source| {
        let mut replacement = Vec::new();
        litchi_iwa_common::wire::append_varint_field(
            &mut replacement,
            1,
            u64::from(FORMULA_OWNER_MESSAGE_TYPE + 1),
        )?;
        replace_first_raw_field_at_nested_path(source, &[2], 1, &replacement)
    })
}

fn formula_owner_message_info_length_mismatch(package: &[u8], index: usize) -> TestResult<Vec<u8>> {
    let payload_length = object_message_payload(
        package,
        table_formula_owner(index),
        FORMULA_OWNER_MESSAGE_TYPE,
    )?
    .len();
    rewrite_raw_object_header(package, table_formula_owner(index), |source| {
        let mut replacement = Vec::new();
        litchi_iwa_common::wire::append_varint_field(
            &mut replacement,
            3,
            u64::try_from(payload_length.saturating_add(1))?,
        )?;
        replace_first_raw_field_at_nested_path(source, &[2], 3, &replacement)
    })
}

fn formula_owner_bad_versions(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    rewrite_object_metadata(
        package,
        identifier,
        FORMULA_OWNER_MESSAGE_TYPE,
        |object, message_index| {
            object.archive_info.message_infos[message_index].versions = vec![2, 0, 0];
            Ok(())
        },
    )
}

fn formula_owner_bad_merge(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    rewrite_object_metadata(
        package,
        identifier,
        FORMULA_OWNER_MESSAGE_TYPE,
        |object, _| {
            object.archive_info.should_merge = Some(true);
            Ok(())
        },
    )
}

fn formula_owner_bad_base(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    rewrite_object_metadata(
        package,
        identifier,
        FORMULA_OWNER_MESSAGE_TYPE,
        |object, message_index| {
            object.archive_info.message_infos[message_index].base_message_index = Some(0);
            Ok(())
        },
    )
}

fn formula_owner_bad_diff(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    rewrite_object_metadata(
        package,
        identifier,
        FORMULA_OWNER_MESSAGE_TYPE,
        |object, message_index| {
            object.archive_info.message_infos[message_index].diff_field_path =
                Some(FieldPath::new(vec![1]));
            Ok(())
        },
    )
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

fn object_message_payload(package: &[u8], identifier: u64, type_: u32) -> TestResult<Vec<u8>> {
    Ok(document_archive(package)?
        .object(identifier)
        .ok_or_else(|| format!("missing object {identifier}"))?
        .messages
        .iter()
        .find(|message| message.type_ == type_)
        .ok_or_else(|| format!("missing message type {type_} in object {identifier}"))?
        .data
        .clone())
}

fn contains_field_at_path(source: &[u8], path: &[u32], number: u32) -> TestResult<bool> {
    let fields = raw_wire_fields(source)?;
    if path.is_empty() {
        return Ok(fields.iter().any(|field| field.number() == number));
    }
    let field = fields
        .iter()
        .find(|field| field.number() == path[0] && field.wire_type() == 2);
    field
        .map(|field| contains_field_at_path(field.payload(), &path[1..], number))
        .transpose()
        .map(|value| value.unwrap_or(false))
}

fn raw_fields_at_path(source: &[u8], path: &[u32], numbers: &[u32]) -> TestResult<Vec<Vec<u8>>> {
    if let Some((&head, tail)) = path.split_first() {
        let fields = raw_wire_fields(source)?;
        let field = fields
            .iter()
            .find(|field| field.number() == head && field.wire_type() == 2)
            .ok_or_else(|| format!("missing nested field path {path:?}"))?;
        return raw_fields_at_path(field.payload(), tail, numbers);
    }
    Ok(raw_wire_fields(source)?
        .iter()
        .filter(|field| numbers.contains(&field.number()))
        .map(|field| field.raw().to_vec())
        .collect())
}

fn unselected_object_snapshot(
    package: &[u8],
    index: usize,
) -> TestResult<Vec<(u64, ArchiveObject)>> {
    let selected = [
        table_attachment(index),
        table_drawable(index),
        table_model(index),
        table_uid_map(index),
        table_formula_owner(index),
        table_filter_set(index, true),
        table_filter_set(index, false),
        table_formula_object(index, true),
        table_formula_object(index, false),
    ];
    let mut objects = document_archive(package)?
        .objects
        .into_iter()
        .filter_map(|object| {
            let identifier = object.archive_info.identifier?;
            (!selected.contains(&identifier)).then_some((identifier, object))
        })
        .collect::<Vec<_>>();
    objects.sort_unstable_by_key(|(identifier, _)| *identifier);
    Ok(objects)
}

fn selected_non_target_object_snapshot(
    package: &[u8],
    index: usize,
) -> TestResult<Vec<(u64, ArchiveObject)>> {
    let selected = [
        table_attachment(index),
        table_uid_map(index),
        table_formula_owner(index),
        table_filter_set(index, true),
        table_filter_set(index, false),
        table_formula_object(index, true),
        table_formula_object(index, false),
    ];
    Ok(document_archive(package)?
        .objects
        .into_iter()
        .filter_map(|object| {
            let identifier = object.archive_info.identifier?;
            selected
                .contains(&identifier)
                .then_some((identifier, object))
        })
        .collect())
}

fn archive_object_order(archive: &Archive) -> TestResult<Vec<u64>> {
    archive
        .objects
        .iter()
        .map(|object| {
            object
                .archive_info
                .identifier
                .ok_or_else(|| "archive object is missing its identifier".into())
        })
        .collect()
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

fn without_previews(package: &[u8]) -> TestResult<Vec<u8>> {
    without_preview_subset(package, &PREVIEWS)
}

fn without_preview_subset(package: &[u8], previews: &[&str]) -> TestResult<Vec<u8>> {
    Ok(
        Catalog::from_bytes(package)?.reassemble_with_deletions_to_bytes(
            &[],
            previews,
            Limits::default(),
        )?,
    )
}

fn assert_rejected_without_mutation(
    package: &Package,
    result: Result<impl std::fmt::Debug, Error>,
    before: &[u8],
    case_index: usize,
) {
    assert!(
        result.is_err(),
        "malformed or unsupported graph was accepted (case {case_index}): {result:?}"
    );
    assert_eq!(package.exact_bytes(), before);
}

#[test]
fn selectors_read_rows_columns_and_absence_without_collapsing_names() -> TestResult {
    let source = normal_package()?;
    let package = Package::from_bytes(&source)?;
    let expected = HiddenAxes::new([AxisIndex::row(1), AxisIndex::column(2)])?;

    // The first table resolves its UID map through TST.TableModelArchive's
    // base-column/row UID reference (field 46), not the optional table-info
    // view map (field 6).  Keep this assertion independent of the focused
    // hidden-axis reader so a fixture typo cannot mask an ownership bug.
    assert!(
        !WireView::parse(&info_payload(&source, 0)?)?
            .fields()
            .any(|field| field.number() == 6)
    );
    let model_payload = model_payload_from_package(&source, 0)?;
    let model_wire = WireView::parse(&model_payload)?;
    assert!(model_wire.fields().any(|field| field.number() == 46));
    let owner_field = model_wire
        .fields()
        .find(|field| field.number() == 70)
        .ok_or("missing native hidden-state owner")?;
    assert_eq!(owner_field.wire_type(), 2);
    assert_eq!(&owner_field.raw()[..2], &[0xb2, 0x04]);

    assert_eq!(package.body_table_hidden_axes(0usize)?, expected);
    assert_eq!(
        package.body_table_hidden_axes(BodyTableSelector::index(0))?,
        expected
    );
    assert_eq!(
        package.body_table_hidden_axes(BodyTableSelector::name("Revenue"))?,
        expected
    );
    assert_eq!(
        package.body_table_hidden_axes(BodyTableSelector::name("Costs"))?,
        HiddenAxes::empty()
    );
    assert!(matches!(
        package.body_table_hidden_axes(BodyTableSelector::name("Missing")),
        Err(Error::TableNotFound)
    ));
    assert!(matches!(
        package.body_table_hidden_axes(usize::MAX),
        Err(Error::TableNotFound | Error::AmbiguousSelector)
    ));

    let ambiguous_source = synthetic_package(
        [
            TableOptions {
                user_hidden: true,
                ..TableOptions::default()
            },
            TableOptions::default(),
        ],
        ["Revenue", "Revenue"],
    )?;
    let ambiguous = Package::from_bytes(&ambiguous_source)?;
    assert!(matches!(
        ambiguous.body_table_hidden_axes(BodyTableSelector::name("Revenue")),
        Err(Error::AmbiguousTableName | Error::AmbiguousSelector)
    ));
    Ok(())
}

#[test]
fn absent_state_reads_empty_and_changed_creation_is_refused_atomically() -> TestResult {
    let source = normal_package()?;
    let package = Package::from_bytes(&source)?;
    let requested = HiddenAxes::new([AxisIndex::row(0), AxisIndex::column(3)])?;
    assert_eq!(package.body_table_hidden_axes(1usize)?, HiddenAxes::empty());

    let before = package.exact_bytes();
    let result = package
        .edit_body_table_hidden_axes(BodyTableSelector::name("Costs"))?
        .set(requested.clone())
        .commit();
    assert!(
        matches!(
            result,
            Err(Error::UnsupportedDependency | Error::UnsupportedSource)
        ),
        "native absent-owner creation must be refused: {result:?}"
    );
    assert_eq!(package.exact_bytes(), before);

    let absent_noop = package
        .edit_body_table_hidden_axes(1usize)?
        .clear()
        .commit()?;
    assert!(absent_noop.patch().is_noop());
    assert_eq!(absent_noop.package().exact_bytes(), source);
    Ok(())
}

#[test]
fn set_clear_reset_and_noop_preserve_exact_source_and_presence() -> TestResult {
    let source = normal_package()?;
    let package = Package::from_bytes(&source)?;
    let before = package.body_table_hidden_axes(0usize)?;

    let noop = package
        .edit_body_table_hidden_axes(BodyTableSelector::name("Revenue"))?
        .set(before.clone())
        .commit()?;
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert_eq!(noop.diagnostics().deleted_previews(), 0);
    assert_eq!(noop.package().exact_bytes(), source);
    assert_eq!(
        noop.package()
            .apply_body_table_hidden_axes(noop.patch())?
            .package()
            .exact_bytes(),
        source
    );
    let replay = noop.package().apply_body_table_hidden_axes(noop.patch())?;
    let replay_again = replay
        .package()
        .apply_body_table_hidden_axes(replay.patch())?;
    assert!(replay_again.patch().is_noop());
    assert_eq!(replay_again.package().exact_bytes(), source);

    let tampered_source = Catalog::from_bytes(&source)?.reassemble_to_bytes(
        &[EntryEdit::new("Data/sentinel.bin", b"tampered-noop")],
        Limits::default(),
    )?;
    let tampered = Package::from_bytes(&tampered_source)?;
    assert!(matches!(
        tampered.apply_body_table_hidden_axes(noop.patch()),
        Err(Error::PatchConflict)
    ));

    let clear = package
        .edit_body_table_hidden_axes(0usize)?
        .clear()
        .commit()?;
    assert_eq!(
        clear.package().body_table_hidden_axes(0usize)?,
        HiddenAxes::empty()
    );
    assert!(
        WireView::parse(&model_payload_from_package(
            &clear.package().exact_bytes(),
            0
        )?)?
        .fields()
        .any(|field| field.number() == 70),
        "clearing user-hidden axes must retain the native owner presence"
    );
    assert!(
        WireView::parse(&info_payload(&clear.package().exact_bytes(), 0)?)?
            .fields()
            .any(|field| field.number() == 8),
        "clearing user-hidden axes must retain the active-state UUID presence"
    );

    let reset = Package::from_bytes(&clear.package().exact_bytes())?
        .edit_body_table_hidden_axes(0usize)?
        .reset()
        .commit()?;
    assert_eq!(
        reset.package().body_table_hidden_axes(0usize)?,
        HiddenAxes::empty()
    );
    assert_eq!(reset.patch().before().clone(), HiddenAxes::empty());
    Ok(())
}

#[test]
fn changed_axes_have_exact_apply_inverse_and_conflict_fences() -> TestResult {
    let source = normal_package()?;
    let package = Package::from_bytes(&source)?;
    let requested = HiddenAxes::new([AxisIndex::row(0), AxisIndex::row(3), AxisIndex::column(1)])?;
    let commit = package
        .edit_body_table_hidden_axes(0usize)?
        .set(requested.clone())
        .commit()?;
    let target = commit.package().exact_bytes();
    assert_eq!(
        commit.patch().before().clone(),
        HiddenAxes::new([AxisIndex::row(1), AxisIndex::column(2)])?
    );
    assert_eq!(commit.patch().after().clone(), requested);
    assert!(!commit.patch().is_noop());
    assert_eq!(
        package
            .apply_body_table_hidden_axes(commit.patch())?
            .package()
            .exact_bytes(),
        target
    );
    let reopened = Package::from_bytes(&target)?;
    assert_eq!(
        reopened.body_table_hidden_axes(0usize)?,
        commit.patch().after().clone()
    );
    let restored = commit
        .package()
        .apply_body_table_hidden_axes(&commit.patch().inverse())?;
    assert_eq!(restored.package().exact_bytes(), source);
    assert_eq!(
        restored.package().body_table_hidden_axes(0usize)?,
        commit.patch().before().clone()
    );

    let tampered_source = Catalog::from_bytes(&source)?.reassemble_to_bytes(
        &[EntryEdit::new("Data/sentinel.bin", b"tampered")],
        Limits::default(),
    )?;
    let tampered = Package::from_bytes(&tampered_source)?;
    assert!(matches!(
        tampered.apply_body_table_hidden_axes(commit.patch()),
        Err(Error::PatchConflict)
    ));
    let tampered_target_source = Catalog::from_bytes(&target)?.reassemble_to_bytes(
        &[EntryEdit::new("Data/sentinel.bin", b"tampered-target")],
        Limits::default(),
    )?;
    let tampered_target = Package::from_bytes(&tampered_target_source)?;
    assert!(matches!(
        tampered_target.apply_body_table_hidden_axes(&commit.patch().inverse()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(tampered_target.exact_bytes(), tampered_target_source);
    let foreign = Package::from_bytes(&synthetic_package(
        [
            TableOptions {
                user_hidden: true,
                ..TableOptions::default()
            },
            TableOptions::default(),
        ],
        ["Other", "Costs"],
    )?)?;
    assert!(matches!(
        foreign.apply_body_table_hidden_axes(commit.patch()),
        Err(Error::PatchConflict)
    ));
    Ok(())
}

#[test]
fn axis_bounds_and_semantic_duplicates_fail_before_publication() -> TestResult {
    assert!(matches!(
        HiddenAxes::new([AxisIndex::row(2), AxisIndex::row(2)]),
        Err(litchi_pages::table::hidden_axes::Error::Duplicate { .. })
    ));
    assert!(matches!(
        HiddenAxes::new([AxisIndex::column(1), AxisIndex::column(1)]),
        Err(litchi_pages::table::hidden_axes::Error::Duplicate { .. })
    ));

    let package = Package::from_bytes(&normal_package()?)?;
    let before = package.exact_bytes();
    let invalid_row = HiddenAxes::new([AxisIndex::row(TABLE_ROWS as usize)])?;
    let invalid_column = HiddenAxes::new([AxisIndex::column(TABLE_COLUMNS as usize)])?;
    assert!(
        package
            .edit_body_table_hidden_axes(0usize)?
            .set(invalid_row)
            .commit()
            .is_err()
    );
    assert!(
        package
            .edit_body_table_hidden_axes(0usize)?
            .set(invalid_column)
            .commit()
            .is_err()
    );
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

#[test]
fn locked_tables_refuse_changed_axis_publication_but_allow_reads() -> TestResult {
    let source = locked_package()?;
    let package = Package::from_bytes(&source)?;
    let current = package.body_table_hidden_axes(0usize)?;
    assert_eq!(
        current,
        HiddenAxes::new([AxisIndex::row(1), AxisIndex::column(2)])?
    );
    let noop = package
        .edit_body_table_hidden_axes(0usize)?
        .set(current.clone())
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package().exact_bytes(), source);
    let before = package.exact_bytes();
    let result = package
        .edit_body_table_hidden_axes(0usize)?
        .set(HiddenAxes::new([AxisIndex::row(0)])?)
        .commit();
    assert!(matches!(result, Err(Error::TableLocked)));
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}
