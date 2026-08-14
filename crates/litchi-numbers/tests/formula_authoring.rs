//! Public contract coverage for bounded, selector-first Numbers formulas.

use std::{fmt::Debug, path::PathBuf};

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, SnappyStream};
use litchi_iwa_protos::{tn, tsce, tsp, tst};
use litchi_numbers::{
    Package,
    cell::Value,
    formula::{
        AxisReference, BinaryOperator, CachedValue, CellReference, Error as FormulaError,
        Expression, LimitKind as FormulaLimitKind, MAX_DEPTH, MAX_NODES, MAX_OWNED_BYTES,
    },
    table::{
        CellPosition,
        cells::{Change, DependencyKind, Error as CellError, Input, Storage},
    },
};
use litchi_numbers_wire::{BncCell, CachedScalar};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/numbers/basic.numbers")
}

fn assert_send_sync_debug<T: Send + Sync + Debug>() {}

fn exact_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .expect("an in-memory Vec accepts package bytes");
    bytes
}

struct SyntheticComponent {
    name: String,
    archive: Archive,
    changed: bool,
}

fn object_route(components: &[SyntheticComponent], identifier: u64) -> TestResult<(usize, usize)> {
    let mut routes = components
        .iter()
        .enumerate()
        .flat_map(|(component, entry)| {
            entry
                .archive
                .objects
                .iter()
                .enumerate()
                .filter_map(move |(object, value)| {
                    (value.archive_info.identifier == Some(identifier))
                        .then_some((component, object))
                })
        });
    let route = routes
        .next()
        .ok_or_else(|| std::io::Error::other("synthetic object reference is unresolved"))?;
    if routes.next().is_some() {
        return Err(std::io::Error::other("synthetic object reference is ambiguous").into());
    }
    Ok(route)
}

fn object_message_index(object: &ArchiveObject, message_type: u32) -> TestResult<usize> {
    let mut matches = object
        .messages
        .iter()
        .enumerate()
        .filter_map(|(index, message)| (message.type_ == message_type).then_some(index));
    let index = matches
        .next()
        .ok_or_else(|| std::io::Error::other("synthetic object message is missing"))?;
    if matches.next().is_some() {
        return Err(std::io::Error::other("synthetic object message is ambiguous").into());
    }
    Ok(index)
}

fn replace_declared_reference(object: &mut ArchiveObject, old: u64, new: u64) -> TestResult {
    let mut aggregate = 0usize;
    for info in &mut object.archive_info.message_infos {
        for reference in &mut info.object_references {
            if *reference == old {
                *reference = new;
                aggregate += 1;
            }
        }
        for field in &mut info.field_infos {
            for reference in &mut field.object_references {
                if *reference == old {
                    *reference = new;
                }
            }
        }
    }
    if aggregate != 1 {
        return Err(std::io::Error::other(format!(
            "expected one aggregate reference, found {aggregate}"
        ))
        .into());
    }
    Ok(())
}

fn append_declared_reference(
    object: &mut ArchiveObject,
    message_index: usize,
    path: &[u32],
    identifier: u64,
) -> TestResult {
    let info = object
        .archive_info
        .message_infos
        .get_mut(message_index)
        .ok_or_else(|| std::io::Error::other("synthetic message metadata is missing"))?;
    if info.object_references.contains(&identifier) {
        return Err(std::io::Error::other("synthetic aggregate reference already exists").into());
    }
    info.object_references.push(identifier);
    if let Some(field) = info
        .field_infos
        .iter_mut()
        .find(|field| field.path.as_slice() == path)
    {
        field.object_references.push(identifier);
    } else {
        let mut field = FieldInfo::new(path.to_vec());
        field.object_references.push(identifier);
        info.field_infos.push(field);
    }
    Ok(())
}

fn replace_declared_field_reference(
    object: &mut ArchiveObject,
    message_index: usize,
    path: &[u32],
    old: u64,
    new: u64,
) -> TestResult {
    let info = object
        .archive_info
        .message_infos
        .get_mut(message_index)
        .ok_or_else(|| std::io::Error::other("synthetic message metadata is missing"))?;
    if info.object_references.contains(&new) {
        return Err(std::io::Error::other("synthetic replacement reference already exists").into());
    }
    let mut aggregate_replacements = 0usize;
    for reference in &mut info.object_references {
        if *reference == old {
            *reference = new;
            aggregate_replacements += 1;
        }
    }
    if aggregate_replacements != 1 {
        return Err(std::io::Error::other(format!(
            "expected one aggregate replacement, found {aggregate_replacements}"
        ))
        .into());
    }
    let mut matching = info
        .field_infos
        .iter_mut()
        .filter(|field| field.path.as_slice() == path);
    let Some(field) = matching.next() else {
        // Some native ArchiveInfo headers retain only the aggregate object
        // list. The semantic resolver accepts that envelope.
        return Ok(());
    };
    if matching.next().is_some() {
        return Err(std::io::Error::other("synthetic reference field is ambiguous").into());
    }
    let mut replacements = 0usize;
    for reference in &mut field.object_references {
        if *reference == old {
            *reference = new;
            replacements += 1;
        }
    }
    if replacements != 1 {
        return Err(std::io::Error::other(format!(
            "expected one field-local replacement, found {replacements}"
        ))
        .into());
    }
    Ok(())
}

fn next_object_identifier(components: &[SyntheticComponent]) -> TestResult<u64> {
    components
        .iter()
        .flat_map(|component| component.archive.objects.iter())
        .filter_map(|object| object.archive_info.identifier)
        .max()
        .ok_or_else(|| std::io::Error::other("synthetic package has no object identifiers").into())
}

fn successor_uid(value: tsp::Uuid) -> TestResult<tsp::Uuid> {
    if value.lower != u64::MAX {
        return Ok(tsp::Uuid {
            lower: value.lower + 1,
            upper: value.upper,
        });
    }
    Ok(tsp::Uuid {
        lower: 1,
        upper: value
            .upper
            .checked_add(1)
            .ok_or_else(|| std::io::Error::other("formula owner UID space exhausted"))?,
    })
}

fn cfuuid(value: tsp::Uuid) -> tsp::CfuuidArchive {
    tsp::CfuuidArchive {
        uuid_w0: Some(value.lower as u32),
        uuid_w1: Some((value.lower >> 32) as u32),
        uuid_w2: Some(value.upper as u32),
        uuid_w3: Some((value.upper >> 32) as u32),
        ..Default::default()
    }
}

/// Clone the native table graph into a distinct second semantic owner while
/// retaining the app-authored storage/list/engine envelopes. This fixture is
/// synthetic and deterministic; it never becomes a checked host file.
fn two_table_formula_package() -> TestResult<Package> {
    const DOCUMENT_MESSAGE_TYPE: u32 = 1;
    const SHEET_MESSAGE_TYPE: u32 = 2;
    const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
    const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
    const TILE_MESSAGE_TYPE: u32 = 6_002;
    const ENGINE_MESSAGE_TYPE: u32 = 4_000;
    const OWNER_MESSAGE_TYPE: u32 = 4_008;

    let source = std::fs::read(fixture_path())?;
    let catalog = Catalog::from_bytes(&source)?;
    let mut components = Vec::new();
    for member in catalog
        .iter()
        .filter(|member| member.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(member.data())?;
        components.push(SyntheticComponent {
            name: member.name().to_owned(),
            archive: Archive::parse(stream.as_bytes())?,
            changed: false,
        });
    }

    let mut root_route = None;
    for (component_index, component) in components.iter().enumerate() {
        for (object_index, object) in component.archive.objects.iter().enumerate() {
            if object.archive_info.identifier == Some(1)
                && object
                    .messages
                    .iter()
                    .any(|message| message.type_ == DOCUMENT_MESSAGE_TYPE)
            {
                if root_route
                    .replace((component_index, object_index))
                    .is_some()
                {
                    return Err(
                        std::io::Error::other("synthetic document root is ambiguous").into(),
                    );
                }
            }
        }
    }
    let (root_component, root_object) =
        root_route.ok_or_else(|| std::io::Error::other("synthetic document root is missing"))?;
    let root_message = object_message_index(
        &components[root_component].archive.objects[root_object],
        DOCUMENT_MESSAGE_TYPE,
    )?;
    let document = tn::DocumentArchive::decode(
        components[root_component].archive.objects[root_object].messages[root_message]
            .data
            .as_slice(),
    )?;
    let sheet_identifier = document
        .sheets
        .first()
        .ok_or_else(|| std::io::Error::other("native fixture has no sheet"))?
        .identifier;
    let engine_identifier = document
        .super_
        .calculation_engine
        .as_ref()
        .ok_or_else(|| std::io::Error::other("native fixture has no calculation engine"))?
        .identifier;

    let (sheet_component, sheet_object) = object_route(&components, sheet_identifier)?;
    let sheet_message = object_message_index(
        &components[sheet_component].archive.objects[sheet_object],
        SHEET_MESSAGE_TYPE,
    )?;
    let mut sheet = tn::SheetArchive::decode(
        components[sheet_component].archive.objects[sheet_object].messages[sheet_message]
            .data
            .as_slice(),
    )?;
    if sheet.drawable_infos.len() != 1 {
        return Err(
            std::io::Error::other("native formula seed must have exactly one table").into(),
        );
    }
    let table_info_identifier = sheet.drawable_infos[0].identifier;
    let (info_component, info_object) = object_route(&components, table_info_identifier)?;
    let info_message = object_message_index(
        &components[info_component].archive.objects[info_object],
        TABLE_INFO_MESSAGE_TYPE,
    )?;
    let table_info = tst::TableInfoArchive::decode(
        components[info_component].archive.objects[info_object].messages[info_message]
            .data
            .as_slice(),
    )?;
    let model_identifier = table_info.table_model.identifier;
    let (model_component, model_object) = object_route(&components, model_identifier)?;
    let model_message = object_message_index(
        &components[model_component].archive.objects[model_object],
        TABLE_MODEL_MESSAGE_TYPE,
    )?;
    let mut model = tst::TableModelArchive::decode(
        components[model_component].archive.objects[model_object].messages[model_message]
            .data
            .as_slice(),
    )?;
    if model.base_data_store.tiles.tiles.len() != 1 {
        return Err(std::io::Error::other("native formula seed must have one tile").into());
    }
    let tile_identifier = model.base_data_store.tiles.tiles[0].tile.identifier;
    let (tile_component, tile_object) = object_route(&components, tile_identifier)?;
    object_message_index(
        &components[tile_component].archive.objects[tile_object],
        TILE_MESSAGE_TYPE,
    )?;
    let formula_list_identifier = model.base_data_store.formula_table.identifier;
    let (formula_list_component, formula_list_object) =
        object_route(&components, formula_list_identifier)?;

    let (engine_component, engine_object) = object_route(&components, engine_identifier)?;
    let engine_message = object_message_index(
        &components[engine_component].archive.objects[engine_object],
        ENGINE_MESSAGE_TYPE,
    )?;
    let mut engine = tsce::CalculationEngineArchive::decode(
        components[engine_component].archive.objects[engine_object].messages[engine_message]
            .data
            .as_slice(),
    )?;
    let mut owner_route = None;
    let mut maximum_internal_owner = 0u32;
    let mut maximum_uid = None;
    for reference in &engine.dependency_tracker.formula_owner_dependencies {
        let route = object_route(&components, reference.identifier)?;
        let message = object_message_index(
            &components[route.0].archive.objects[route.1],
            OWNER_MESSAGE_TYPE,
        )?;
        let owner = tsce::FormulaOwnerDependenciesArchive::decode(
            components[route.0].archive.objects[route.1].messages[message]
                .data
                .as_slice(),
        )?;
        maximum_internal_owner = maximum_internal_owner.max(owner.internal_formula_owner_id);
        let uid = owner.formula_owner_uid;
        if maximum_uid.is_none_or(|current: tsp::Uuid| {
            (uid.lower, uid.upper) > (current.lower, current.upper)
        }) {
            maximum_uid = Some(uid);
        }
        if owner
            .formula_owner
            .as_ref()
            .is_some_and(|reference| reference.identifier == table_info_identifier)
        {
            if owner_route
                .replace((route.0, route.1, message, owner))
                .is_some()
            {
                return Err(std::io::Error::other("selected formula owner is ambiguous").into());
            }
        }
    }
    let (owner_component, owner_object, owner_message, mut owner) =
        owner_route.ok_or_else(|| std::io::Error::other("selected formula owner is missing"))?;
    if !components[owner_component].archive.objects[owner_object]
        .archive_info
        .message_infos[owner_message]
        .object_references
        .contains(&table_info_identifier)
    {
        append_declared_reference(
            &mut components[owner_component].archive.objects[owner_object],
            owner_message,
            &[11],
            table_info_identifier,
        )?;
        components[owner_component].changed = true;
    }

    let first_new = next_object_identifier(&components)?
        .checked_add(1)
        .ok_or_else(|| std::io::Error::other("synthetic identifier space exhausted"))?;
    let new_info_identifier = first_new;
    let new_model_identifier = first_new + 1;
    let new_tile_identifier = first_new + 2;
    let new_formula_list_identifier = first_new + 3;
    let new_owner_identifier = first_new + 4;
    let new_internal_owner = maximum_internal_owner
        .checked_add(1)
        .ok_or_else(|| std::io::Error::other("internal formula owner space exhausted"))?;
    let new_uid = successor_uid(
        maximum_uid.ok_or_else(|| std::io::Error::other("formula owner UID set is empty"))?,
    )?;

    let mut cloned_info = components[info_component].archive.objects[info_object].clone();
    cloned_info.archive_info.identifier = Some(new_info_identifier);
    replace_declared_reference(&mut cloned_info, model_identifier, new_model_identifier)?;
    let mut cloned_info_value = table_info;
    cloned_info_value.table_model.identifier = new_model_identifier;
    cloned_info.replace_message_preserving_header(
        info_message,
        litchi_iwa_core::RawMessage {
            type_: TABLE_INFO_MESSAGE_TYPE,
            data: cloned_info_value.encode_to_vec(),
        },
    )?;

    let mut cloned_model = components[model_component].archive.objects[model_object].clone();
    cloned_model.archive_info.identifier = Some(new_model_identifier);
    replace_declared_reference(&mut cloned_model, tile_identifier, new_tile_identifier)?;
    replace_declared_field_reference(
        &mut cloned_model,
        model_message,
        &[4, 6],
        formula_list_identifier,
        new_formula_list_identifier,
    )?;
    model.table_name = "External".to_owned();
    model.table_id = "litchi-formula-external-owner".to_owned();
    model.base_data_store.tiles.tiles[0].tile.identifier = new_tile_identifier;
    model.base_data_store.formula_table.identifier = new_formula_list_identifier;
    cloned_model.replace_message_preserving_header(
        model_message,
        litchi_iwa_core::RawMessage {
            type_: TABLE_MODEL_MESSAGE_TYPE,
            data: model.encode_to_vec(),
        },
    )?;

    let mut cloned_tile = components[tile_component].archive.objects[tile_object].clone();
    cloned_tile.archive_info.identifier = Some(new_tile_identifier);

    let mut cloned_formula_list =
        components[formula_list_component].archive.objects[formula_list_object].clone();
    cloned_formula_list.archive_info.identifier = Some(new_formula_list_identifier);

    let mut cloned_owner = components[owner_component].archive.objects[owner_object].clone();
    cloned_owner.archive_info.identifier = Some(new_owner_identifier);
    replace_declared_reference(
        &mut cloned_owner,
        table_info_identifier,
        new_info_identifier,
    )?;
    owner.formula_owner_uid = new_uid;
    owner.internal_formula_owner_id = new_internal_owner;
    owner
        .formula_owner
        .as_mut()
        .ok_or_else(|| std::io::Error::other("selected formula owner lost its table"))?
        .identifier = new_info_identifier;
    cloned_owner.replace_message_preserving_header(
        owner_message,
        litchi_iwa_core::RawMessage {
            type_: OWNER_MESSAGE_TYPE,
            data: owner.encode_to_vec(),
        },
    )?;

    sheet.drawable_infos.push(tsp::Reference {
        identifier: new_info_identifier,
        ..Default::default()
    });
    {
        let sheet_object = &mut components[sheet_component].archive.objects[sheet_object];
        append_declared_reference(sheet_object, sheet_message, &[2], new_info_identifier)?;
        sheet_object.replace_message_preserving_header(
            sheet_message,
            litchi_iwa_core::RawMessage {
                type_: SHEET_MESSAGE_TYPE,
                data: sheet.encode_to_vec(),
            },
        )?;
        components[sheet_component].changed = true;
    }

    engine
        .dependency_tracker
        .formula_owner_dependencies
        .push(tsp::Reference {
            identifier: new_owner_identifier,
            ..Default::default()
        });
    engine
        .dependency_tracker
        .owner_id_map
        .as_mut()
        .ok_or_else(|| std::io::Error::other("native engine owner map is missing"))?
        .map_entry
        .push(tsce::owner_id_map_archive::OwnerIdMapArchiveEntry {
            internal_owner_id: new_internal_owner,
            owner_id: cfuuid(new_uid),
        });
    {
        let engine_object = &mut components[engine_component].archive.objects[engine_object];
        append_declared_reference(engine_object, engine_message, &[2, 6], new_owner_identifier)?;
        engine_object.replace_message_preserving_header(
            engine_message,
            litchi_iwa_core::RawMessage {
                type_: ENGINE_MESSAGE_TYPE,
                data: engine.encode_to_vec(),
            },
        )?;
        components[engine_component].changed = true;
    }

    components[info_component].archive.objects.push(cloned_info);
    components[info_component].changed = true;
    components[model_component]
        .archive
        .objects
        .push(cloned_model);
    components[model_component].changed = true;
    components[tile_component].archive.objects.push(cloned_tile);
    components[tile_component].changed = true;
    components[formula_list_component]
        .archive
        .objects
        .push(cloned_formula_list);
    components[formula_list_component].changed = true;
    components[owner_component]
        .archive
        .objects
        .push(cloned_owner);
    components[owner_component].changed = true;

    let mut rewritten = Vec::new();
    for component in components.iter().filter(|component| component.changed) {
        rewritten.push((
            component.name.as_str(),
            SnappyStream::compress(&component.archive.to_bytes()?)?,
        ));
    }
    let edits = rewritten
        .iter()
        .map(|(name, bytes)| EntryEdit::new(name, bytes))
        .collect::<Vec<_>>();
    let bytes = catalog.reassemble_to_bytes(&edits, Limits::default())?;
    let package = Package::from_bytes(&bytes)?;
    assert_eq!(package.sheets()[0].tables().count(), 2);
    Ok(package)
}

fn data_list_entries(
    package: &[u8],
    kind: tst::table_data_list::ListType,
) -> TestResult<Vec<tst::table_data_list::ListEntry>> {
    let catalog = Catalog::from_bytes(package)?;
    let mut entries = Vec::new();
    for member in catalog
        .iter()
        .filter(|member| member.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(member.data())?;
        let archive = Archive::parse(stream.as_bytes())?;
        for message in archive
            .objects
            .iter()
            .flat_map(|object| object.messages.iter())
            .filter(|message| matches!(message.type_, 6_005 | 6_201))
        {
            let Ok(list) = tst::TableDataList::decode(message.data.as_slice()) else {
                continue;
            };
            if list.list_type == kind as i32 {
                entries.extend(list.entries);
            }
        }
    }
    Ok(entries)
}

fn formula_entries(package: &[u8]) -> TestResult<Vec<tst::table_data_list::ListEntry>> {
    data_list_entries(package, tst::table_data_list::ListType::Formula)
}

fn decode_offset(bytes: &[u8], width: usize) -> Option<usize> {
    match width {
        2 => {
            let value = u16::from_le_bytes(bytes.try_into().ok()?);
            (value != u16::MAX).then_some(usize::from(value))
        },
        4 => {
            let value = u32::from_le_bytes(bytes.try_into().ok()?);
            (value != u32::MAX)
                .then(|| usize::try_from(value).ok())
                .flatten()
        },
        _ => None,
    }
}

fn formula_cell(package: &[u8], position: CellPosition) -> TestResult<BncCell> {
    let catalog = Catalog::from_bytes(package)?;
    let mut matches = Vec::new();
    for member in catalog
        .iter()
        .filter(|member| member.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(member.data())?;
        let archive = Archive::parse(stream.as_bytes())?;
        for message in archive
            .objects
            .iter()
            .flat_map(|object| object.messages.iter())
            .filter(|message| message.type_ == 6_002)
        {
            let tile = tst::Tile::decode(message.data.as_slice())?;
            let Some(row) = tile
                .row_infos
                .iter()
                .find(|row| row.tile_row_index == position.row())
            else {
                continue;
            };
            let Some(offsets) = row.cell_offsets.as_deref() else {
                continue;
            };
            let Some(storage) = row.cell_storage_buffer.as_deref() else {
                continue;
            };
            let width = if row.has_wide_offsets.unwrap_or(false) {
                4
            } else {
                2
            };
            let column = usize::try_from(position.column())?;
            let start_index = column
                .checked_mul(width)
                .ok_or_else(|| std::io::Error::other("formula offset overflow"))?;
            let Some(start_bytes) = offsets.get(start_index..start_index + width) else {
                continue;
            };
            let Some(start) = decode_offset(start_bytes, width) else {
                continue;
            };
            let end = offsets[start_index + width..]
                .chunks_exact(width)
                .find_map(|bytes| decode_offset(bytes, width))
                .unwrap_or(storage.len());
            let encoded = storage
                .get(start..end)
                .ok_or_else(|| std::io::Error::other("formula cell storage range is malformed"))?;
            let cell = BncCell::parse(encoded)?;
            if matches!(
                cell.stored_value(),
                litchi_numbers_wire::StoredValue::Formula(_)
            ) {
                matches.push(cell);
            }
        }
    }
    if matches.len() != 1 {
        return Err(std::io::Error::other(format!(
            "expected one formula cell at {position:?}, found {}",
            matches.len()
        ))
        .into());
    }
    Ok(matches.remove(0))
}

fn assert_formula_source(package: &Package, position: CellPosition) -> TestResult {
    assert!(matches!(
        package.table_cell(0usize, 0usize, position)?.storage(),
        Storage::Stored(Value::Formula(source)) if source.starts_with('=') && source.len() > 1
    ));
    Ok(())
}

#[test]
fn public_formula_values_are_bounded_typed_and_content_free() -> TestResult {
    assert_send_sync_debug::<Expression>();
    assert_send_sync_debug::<CachedValue>();
    assert_send_sync_debug::<CellReference>();
    assert_send_sync_debug::<AxisReference>();
    assert_send_sync_debug::<BinaryOperator>();
    // Pivot categories intentionally have no forgeable public constructor;
    // they remain an opaque future Package-resolved selector.
    assert_send_sync_debug::<FormulaError>();
    assert_send_sync_debug::<FormulaLimitKind>();
    assert_send_sync_debug::<Input>();
    assert_send_sync_debug::<Change>();

    let secret = "private formula literal";
    let expression = Expression::text(secret)?;
    let cached = CachedValue::text(secret)?;
    let input = Input::formula_cached(expression.clone(), cached.clone());
    let change = Change::set_formula_cached(CellPosition::new(2, 2), expression, cached);
    for rendered in [
        format!("{input:?}"),
        format!("{change:?}"),
        format!("{:?}", Expression::text(secret)?),
        format!("{:?}", CachedValue::text(secret)?),
    ] {
        assert!(!rendered.contains(secret));
    }
    Ok(())
}

#[test]
fn literals_and_caches_reject_nonfinite_and_oversized_values() -> TestResult {
    for value in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
        assert_eq!(Expression::number(value), Err(FormulaError::NonFinite));
        assert_eq!(CachedValue::number(value), Err(FormulaError::NonFinite));
        assert_eq!(CachedValue::date(value), Err(FormulaError::NonFinite));
        assert_eq!(CachedValue::duration(value), Err(FormulaError::NonFinite));
    }

    let maximum = "x".repeat(MAX_OWNED_BYTES);
    Expression::text(&maximum)?;
    CachedValue::text(&maximum)?;
    let oversized = format!("{maximum}x");
    for error in [
        Expression::text(&oversized).expect_err("oversized expression text must fail"),
        CachedValue::text(&oversized).expect_err("oversized cached text must fail"),
    ] {
        assert!(matches!(
            error,
            FormulaError::LimitExceeded {
                kind: FormulaLimitKind::OwnedBytes,
                observed,
                maximum: actual,
            } if observed == MAX_OWNED_BYTES + 1 && actual == MAX_OWNED_BYTES
        ));
    }

    assert_eq!(Expression::boolean(true), Expression::boolean(true));
    assert_eq!(CachedValue::boolean(false), CachedValue::boolean(false));
    Ok(())
}

#[test]
fn composition_charges_literal_bytes_and_depth_aggregately() -> TestResult {
    let half = "x".repeat(MAX_OWNED_BYTES / 2);
    let left = Expression::text(&half)?;
    let right = Expression::text(&half)?;
    let error = Expression::function("SUM", [left, right])
        .expect_err("function name plus literals must exceed the aggregate byte limit");
    assert!(matches!(
        error,
        FormulaError::LimitExceeded {
            kind: FormulaLimitKind::OwnedBytes,
            observed,
            maximum,
        } if observed == MAX_OWNED_BYTES + 3 && maximum == MAX_OWNED_BYTES
    ));

    let mut deepest = Expression::number(1.0)?;
    for _ in 1..MAX_DEPTH {
        deepest = Expression::negate(deepest)?;
    }
    let error = Expression::percent(deepest)
        .expect_err("one node beyond the maximum formula depth must fail");
    assert!(matches!(
        error,
        FormulaError::LimitExceeded {
            kind: FormulaLimitKind::Depth,
            observed,
            maximum,
        } if observed == MAX_DEPTH + 1 && maximum == MAX_DEPTH
    ));
    Ok(())
}

#[test]
fn node_limit_accepts_the_exact_maximum_and_refuses_the_next_parent() -> TestResult {
    // Each inner SUM has 257 nodes. The outer SUM therefore has exactly
    // 1 + 255 * 257 = 65_536 nodes, without approaching the depth limit.
    let mut children = Vec::with_capacity(255);
    for _ in 0..255 {
        children.push(Expression::function(
            "SUM",
            (0..256).map(|_| Expression::boolean(true)),
        )?);
    }
    let maximum = Expression::function("SUM", children)?;
    assert!(format!("{maximum:?}").contains(&format!("nodes: {MAX_NODES}")));
    let error = Expression::negate(maximum)
        .expect_err("a unary parent above the exact node maximum must fail");
    assert!(matches!(
        error,
        FormulaError::LimitExceeded {
            kind: FormulaLimitKind::Nodes,
            observed,
            maximum,
        } if observed == MAX_NODES + 1 && maximum == MAX_NODES
    ));
    Ok(())
}

#[test]
fn functions_unary_binary_ranges_and_whole_axes_are_checked() -> TestResult {
    for (name, arguments) in [
        ("sum", vec![Expression::number(1.0)?]),
        ("AVERAGE", vec![Expression::number(1.0)?]),
        ("MIN", vec![Expression::number(1.0)?]),
        ("MAX", vec![Expression::number(1.0)?]),
        ("COUNT", vec![Expression::number(1.0)?]),
        ("COUNTA", vec![Expression::text("x")?]),
        ("AND", vec![Expression::boolean(true)]),
        ("OR", vec![Expression::boolean(false)]),
        (
            "IF",
            vec![Expression::boolean(true), Expression::number(1.0)?],
        ),
        (
            "IFERROR",
            vec![Expression::number(1.0)?, Expression::number(0.0)?],
        ),
        ("NOT", vec![Expression::boolean(true)]),
        ("ABS", vec![Expression::number(-1.0)?]),
        (
            "ROUND",
            vec![Expression::number(1.25)?, Expression::number(1.0)?],
        ),
    ] {
        Expression::function(name, arguments)?;
    }
    assert_eq!(
        Expression::function("MISSING", [Expression::number(1.0)?]),
        Err(FormulaError::UnsupportedFunction)
    );
    assert!(matches!(
        Expression::function("TOO-LONG", [Expression::number(1.0)?]),
        Err(FormulaError::LimitExceeded {
            kind: FormulaLimitKind::OwnedBytes,
            observed: 8,
            maximum: 7,
        })
    ));
    assert_eq!(
        Expression::function("IF", [Expression::boolean(true)]),
        Err(FormulaError::InvalidArity)
    );

    for operator in [
        BinaryOperator::Add,
        BinaryOperator::Subtract,
        BinaryOperator::Multiply,
        BinaryOperator::Divide,
        BinaryOperator::Power,
        BinaryOperator::Concatenate,
        BinaryOperator::GreaterThan,
        BinaryOperator::GreaterThanOrEqual,
        BinaryOperator::LessThan,
        BinaryOperator::LessThanOrEqual,
        BinaryOperator::Equal,
        BinaryOperator::NotEqual,
    ] {
        Expression::binary(operator, Expression::number(1.0)?, Expression::number(2.0)?)?;
    }
    Expression::negate(Expression::number(1.0)?)?;
    Expression::percent(Expression::number(50.0)?)?;

    let a1 = CellReference::relative(0, 0);
    let b2 = CellReference::mixed(1, 1, true, false);
    let _ = Expression::cell(a1);
    Expression::range(a1, b2)?;
    Expression::rows(AxisReference::relative(0), AxisReference::absolute(1))?;
    Expression::columns(AxisReference::absolute(0), AxisReference::relative(1))?;
    assert_eq!(Expression::range(b2, a1), Err(FormulaError::ReversedRange));
    assert_eq!(
        Expression::rows(AxisReference::relative(2), AxisReference::relative(1)),
        Err(FormulaError::ReversedRange)
    );
    assert!(matches!(
        Expression::function(
            "SUM",
            (0..=litchi_numbers::formula::MAX_FUNCTION_ARGUMENTS)
                .map(|_| Expression::boolean(true)),
        ),
        Err(FormulaError::LimitExceeded {
            kind: FormulaLimitKind::FunctionArguments,
            observed,
            maximum,
        }) if observed == litchi_numbers::formula::MAX_FUNCTION_ARGUMENTS + 1
            && maximum == litchi_numbers::formula::MAX_FUNCTION_ARGUMENTS
    ));
    Ok(())
}

#[test]
fn selector_derived_table_handles_are_source_bound() -> TestResult {
    let package = Package::open(fixture_path())?;
    let edit = package.edit_table_cells("Sheet 1", "Table 1")?;
    let by_name = edit.formula_table("Sheet 1", "Table 1")?;
    let by_index = edit.formula_table(0usize, 0usize)?;
    assert_eq!(by_name, by_index);
    assert!(!format!("{by_name:?}").contains("Table 1"));
    assert!(matches!(
        edit.formula_table("missing sheet", 0usize),
        Err(CellError::SheetNotFound)
    ));
    assert!(matches!(
        edit.formula_table(0usize, "missing table"),
        Err(CellError::TableNotFound)
    ));

    let local = Expression::function(
        "SUM",
        [
            Expression::table_cell(&by_name, CellReference::relative(2, 1)),
            Expression::table_range(
                &by_name,
                CellReference::relative(1, 1),
                CellReference::absolute(2, 2),
            )?,
            Expression::table_rows(
                &by_name,
                AxisReference::relative(1),
                AxisReference::relative(2),
            )?,
            Expression::table_columns(
                &by_name,
                AxisReference::relative(1),
                AxisReference::relative(2),
            )?,
        ],
    )?;
    let staged = edit.set_formula_a1("C3", local)?;
    assert_eq!(staged.len(), 1);

    let independent = Package::open(fixture_path())?;
    let independent_edit = independent.edit_table_cells(0usize, 0usize)?;
    let foreign = Expression::table_cell(&by_name, CellReference::relative(2, 1));
    assert!(matches!(
        independent_edit.set_formula_a1("C3", foreign),
        Err(CellError::PatchConflict)
    ));
    Ok(())
}

#[test]
fn formula_staging_checks_addresses_bounds_and_all_cache_kinds() -> TestResult {
    let package = Package::open(fixture_path())?;
    let source = exact_bytes(&package);
    let expression = || {
        Expression::function(
            "SUM",
            [
                Expression::cell(CellReference::relative(2, 1)),
                Expression::number(8.0).expect("finite literal"),
            ],
        )
        .expect("valid SUM")
    };
    let mut edit = package.edit_table_cells(0usize, 0usize)?;
    for (column, cached) in [
        (2, CachedValue::number(50.0)?),
        (3, CachedValue::text("fifty")?),
        (4, CachedValue::boolean(true)),
        (5, CachedValue::date(123_456.0)?),
        (6, CachedValue::duration(90.0)?),
    ] {
        edit = edit.set_formula_cached(CellPosition::new(2, column), expression(), cached)?;
    }
    assert_eq!(edit.len(), 5);

    assert!(matches!(
        package
            .edit_table_cells(0usize, 0usize)?
            .set_formula_a1("private invalid address", expression()),
        Err(CellError::InvalidAddress)
    ));
    assert!(matches!(
        package
            .edit_table_cells(0usize, 0usize)?
            .set_formula(CellPosition::new(22, 0), Expression::number(1.0)?,),
        Err(CellError::OutOfBounds { .. })
    ));
    assert!(matches!(
        package
            .edit_table_cells(0usize, 0usize)?
            .set_formula_a1("C3", Expression::cell(CellReference::relative(22, 0)),),
        Err(CellError::OutOfBounds { .. })
    ));
    assert_eq!(exact_bytes(&package), source);
    Ok(())
}

#[test]
fn cache_poison_cycles_and_unsupported_descendants_refuse_atomically() -> TestResult {
    let package = Package::open(fixture_path())?;
    let source = exact_bytes(&package);

    let poisoned = package
        .edit_table_cells(0usize, 0usize)?
        .set_formula_cached_a1(
            "C3",
            Expression::function(
                "SUM",
                [
                    Expression::cell(CellReference::relative(2, 1)),
                    Expression::number(8.0)?,
                ],
            )?,
            CachedValue::number(999.0)?,
        )?
        .set_formula_a1(
            "D3",
            Expression::binary(
                BinaryOperator::Multiply,
                Expression::cell(CellReference::relative(2, 2)),
                Expression::number(2.0)?,
            )?,
        )?
        .commit()
        .expect_err("a supplied poison cannot feed an evaluator-supported descendant");
    assert!(
        matches!(
            poisoned,
            CellError::UnsupportedDependency {
                kind: DependencyKind::FormulaCache,
                ..
            }
        ),
        "unexpected poisoned-cache error: {poisoned:?}"
    );
    assert_eq!(exact_bytes(&package), source);

    let cycle = package
        .edit_table_cells(0usize, 0usize)?
        .set_formula_a1("C3", Expression::cell(CellReference::relative(2, 3)))?
        .set_formula_a1("D3", Expression::cell(CellReference::relative(2, 2)))?
        .commit()
        .expect_err("an authored dependency cycle must refuse atomically");
    assert!(matches!(
        cycle,
        CellError::UnsupportedDependency {
            kind: DependencyKind::FormulaCache,
            ..
        }
    ));
    assert_eq!(exact_bytes(&package), source);

    let unsupported_expression = || {
        Expression::function(
            "IF",
            [
                Expression::boolean(true),
                Expression::number(1.0).expect("finite literal"),
                Expression::number(2.0).expect("finite literal"),
            ],
        )
        .expect("structurally supported function")
    };
    let unsupported_descendant = package
        .edit_table_cells(0usize, 0usize)?
        .set_formula_a1("C3", unsupported_expression())?
        .set_formula_a1(
            "D3",
            Expression::binary(
                BinaryOperator::Multiply,
                Expression::cell(CellReference::relative(2, 2)),
                Expression::number(2.0)?,
            )?,
        )?
        .commit()
        .expect_err("an impacted descendant cannot consume an unevaluated authored leaf");
    assert!(matches!(
        unsupported_descendant,
        CellError::UnsupportedDependency {
            kind: DependencyKind::FormulaCache,
            ..
        }
    ));
    assert_eq!(exact_bytes(&package), source);

    let unsupported_leaf = package
        .edit_table_cells(0usize, 0usize)?
        .set_formula_a1("C3", unsupported_expression())?
        .commit()?;
    assert_eq!(unsupported_leaf.diagnostics().changed_cells(), 1);
    assert_eq!(unsupported_leaf.diagnostics().refreshed_formula_caches(), 0);
    let target = exact_bytes(unsupported_leaf.package());
    assert_formula_source(unsupported_leaf.package(), CellPosition::from_a1("C3")?)?;
    assert_eq!(formula_entries(&target)?.len(), 1);
    assert_eq!(
        exact_bytes(
            package
                .apply_table_cells(unsupported_leaf.patch())?
                .package()
        ),
        target
    );
    assert_eq!(
        exact_bytes(
            Package::from_bytes(&target)?
                .apply_table_cells(&unsupported_leaf.patch().inverse())?
                .package()
        ),
        source
    );
    Ok(())
}

#[test]
fn all_typed_caches_commit_through_native_tiles_and_reopen() -> TestResult {
    let package = Package::open(fixture_path())?;
    let source = exact_bytes(&package);
    assert!(
        data_list_entries(&source, tst::table_data_list::ListType::String)?
            .iter()
            .all(|entry| entry.string.as_deref() != Some("typed cache"))
    );
    let unsupported_if = |reference| {
        Expression::function(
            "IF",
            [
                Expression::boolean(true),
                Expression::cell(reference),
                Expression::cell(reference),
            ],
        )
    };
    let commit = package
        .edit_table_cells(0usize, 0usize)?
        .set_a1("B3", Input::number(1.5)?)?
        .set_a1("B5", Input::date(123_456.0)?)?
        .set_a1("B6", Input::duration(90.0)?)?
        .set_formula_cached_a1(
            "C3",
            Expression::cell(CellReference::relative(2, 1)),
            CachedValue::number(1.5)?,
        )?
        .set_formula_cached_a1(
            "D3",
            Expression::binary(
                BinaryOperator::GreaterThan,
                Expression::cell(CellReference::relative(2, 1)),
                Expression::number(0.0)?,
            )?,
            CachedValue::boolean(true),
        )?
        .set_formula_cached_a1(
            "E3",
            Expression::text("typed cache")?,
            CachedValue::text("typed cache")?,
        )?
        .set_formula_cached_a1(
            "F3",
            unsupported_if(CellReference::relative(4, 1))?,
            CachedValue::date(123_456.0)?,
        )?
        .set_formula_cached_a1(
            "G3",
            unsupported_if(CellReference::relative(5, 1))?,
            CachedValue::duration(90.0)?,
        )?
        .set_formula_cached_a1(
            "C4",
            Expression::negate(Expression::number(5.0)?)?,
            CachedValue::number(-5.0)?,
        )?
        .set_formula_cached_a1(
            "D4",
            Expression::percent(Expression::number(50.0)?)?,
            CachedValue::number(0.5)?,
        )?
        .set_formula_cached_a1(
            "E4",
            Expression::binary(
                BinaryOperator::Add,
                Expression::number(1.0)?,
                Expression::number(2.0)?,
            )?,
            CachedValue::number(3.0)?,
        )?
        .commit()?;

    assert_eq!(commit.diagnostics().requested_cells(), 11);
    assert_eq!(commit.diagnostics().changed_cells(), 11);
    assert_eq!(commit.diagnostics().refreshed_formula_caches(), 8);
    let target = exact_bytes(commit.package());
    assert_eq!(
        exact_bytes(package.apply_table_cells(commit.patch())?.package()),
        target
    );
    let reopened = Package::from_bytes(&target)?;
    for address in ["C3", "D3", "E3", "F3", "G3", "C4", "D4", "E4"] {
        assert_formula_source(&reopened, CellPosition::from_a1(address)?)?;
    }
    assert!(matches!(
        formula_cell(&target, CellPosition::from_a1("C3")?)?.cached_scalar()?,
        Some(CachedScalar::Number(value)) if value.get() == 1.5
    ));
    assert!(matches!(
        formula_cell(&target, CellPosition::from_a1("D3")?)?.cached_scalar()?,
        Some(CachedScalar::Boolean(true))
    ));
    assert!(matches!(
        formula_cell(&target, CellPosition::from_a1("E3")?)?.cached_scalar()?,
        Some(CachedScalar::Unsupported(_))
    ));
    let text_entries = data_list_entries(&target, tst::table_data_list::ListType::String)?;
    let typed_text = text_entries
        .iter()
        .filter(|entry| entry.string.as_deref() == Some("typed cache"))
        .collect::<Vec<_>>();
    assert_eq!(typed_text.len(), 1);
    assert_eq!(typed_text[0].refcount, 1);
    assert!(matches!(
        formula_cell(&target, CellPosition::from_a1("F3")?)?.cached_scalar()?,
        Some(CachedScalar::Date(value)) if value.get() == 123_456.0
    ));
    assert!(matches!(
        formula_cell(&target, CellPosition::from_a1("G3")?)?.cached_scalar()?,
        Some(CachedScalar::Duration(value)) if value.get() == 90.0
    ));
    for (address, expected) in [("C4", -5.0), ("D4", 0.5), ("E4", 3.0)] {
        assert!(matches!(
            formula_cell(&target, CellPosition::from_a1(address)?)?.cached_scalar()?,
            Some(CachedScalar::Number(value)) if value.get() == expected
        ));
    }
    let entries = formula_entries(&target)?;
    assert_eq!(entries.len(), 8);
    assert!(entries.iter().all(|entry| entry.refcount == 1));

    assert_eq!(
        exact_bytes(
            reopened
                .apply_table_cells(&commit.patch().inverse())?
                .package()
        ),
        source
    );

    let refreshed = reopened
        .edit_table_cells("Sheet 1", "Table 1")?
        .set_a1("B3", Input::number(44.0)?)?
        .commit()?;
    let refreshed_bytes = exact_bytes(refreshed.package());
    assert!(matches!(
        formula_cell(&refreshed_bytes, CellPosition::from_a1("C3")?)?.cached_scalar()?,
        Some(CachedScalar::Number(value)) if value.get() == 44.0
    ));
    assert!(matches!(
        formula_cell(&refreshed_bytes, CellPosition::from_a1("D3")?)?.cached_scalar()?,
        Some(CachedScalar::Boolean(true))
    ));
    assert!(matches!(
        formula_cell(&refreshed_bytes, CellPosition::from_a1("G3")?)?.cached_scalar()?,
        Some(CachedScalar::Duration(value)) if value.get() == 90.0
    ));
    assert_eq!(formula_entries(&refreshed_bytes)?.len(), 8);
    assert_eq!(
        exact_bytes(
            refreshed
                .package()
                .apply_table_cells(&refreshed.patch().inverse())?
                .package()
        ),
        target
    );

    let cleared = Package::from_bytes(&refreshed_bytes)?
        .edit_table_cells("Sheet 1", "Table 1")?
        .clear_a1("G3")?
        .commit()?;
    let cleared_bytes = exact_bytes(cleared.package());
    assert_eq!(formula_entries(&cleared_bytes)?.len(), 7);
    assert!(matches!(
        cleared
            .package()
            .table_cell(0usize, 0usize, CellPosition::from_a1("G3")?)?
            .storage(),
        Storage::Stored(Value::Empty)
    ));
    for address in ["C3", "D3", "E3", "F3", "C4", "D4", "E4"] {
        assert_formula_source(cleared.package(), CellPosition::from_a1(address)?)?;
    }
    assert_eq!(
        exact_bytes(
            cleared
                .package()
                .apply_table_cells(&cleared.patch().inverse())?
                .package()
        ),
        refreshed_bytes
    );
    Ok(())
}

#[test]
fn text_cache_exact_noop_resolves_native_string_key() -> TestResult {
    let expression = || Expression::text("text cache exact");
    let authored = Package::open(fixture_path())?
        .edit_table_cells(0usize, 0usize)?
        .set_formula_cached_a1("E3", expression()?, CachedValue::text("text cache exact")?)?
        .commit()?
        .into_package();
    let authored_bytes = exact_bytes(&authored);
    let reopened = Package::from_bytes(&authored_bytes)?;

    let noop = reopened
        .edit_table_cells(0usize, 0usize)?
        .set_formula_cached_a1("E3", expression()?, CachedValue::text("text cache exact")?)?
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(noop.diagnostics().changed_cells(), 0);
    assert_eq!(exact_bytes(noop.package()), authored_bytes);

    let changed = reopened
        .edit_table_cells(0usize, 0usize)?
        .set_formula_cached_a1(
            "E3",
            expression()?,
            CachedValue::text("text cache changed")?,
        )?
        .commit()?;
    assert!(!changed.patch().is_noop());
    assert_eq!(changed.diagnostics().changed_cells(), 1);
    Ok(())
}

#[test]
fn locked_table_allows_formula_noop_and_refuses_formula_change() -> TestResult {
    let expression = || {
        Expression::function(
            "SUM",
            [
                Expression::cell(CellReference::relative(2, 1)),
                Expression::number(8.0).expect("finite literal"),
            ],
        )
        .expect("supported native formula")
    };
    let authored = Package::open(fixture_path())?
        .edit_table_cells(0usize, 0usize)?
        .set_formula_cached_a1("C3", expression(), CachedValue::number(50.0)?)?
        .commit()?
        .into_package();
    let mut lock = authored.edit_table_lock(0usize, 0usize)?;
    lock.lock();
    let locked = lock.commit()?.into_package();
    let source = exact_bytes(&locked);

    let noop = locked
        .edit_table_cells(0usize, 0usize)?
        .set_formula_cached_a1("C3", expression(), CachedValue::number(50.0)?)?
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(exact_bytes(noop.package()), source);

    let error = locked
        .edit_table_cells(0usize, 0usize)?
        .set_formula_cached_a1(
            "C3",
            Expression::function(
                "SUM",
                [
                    Expression::cell(CellReference::relative(2, 1)),
                    Expression::number(9.0)?,
                ],
            )?,
            CachedValue::number(51.0)?,
        )?
        .commit()
        .expect_err("a changed authored formula must respect the table lock");
    assert!(matches!(error, CellError::TableLocked { .. }));
    assert_eq!(exact_bytes(&locked), source);
    Ok(())
}

#[test]
fn native_basic_formula_is_exact_reversible_noop_replaceable_and_clearable() -> TestResult {
    let package = Package::open(fixture_path())?;
    let source = exact_bytes(&package);
    let c3 = CellPosition::from_a1("C3")?;
    let authored_expression = || {
        Expression::function(
            "SUM",
            [
                Expression::cell(CellReference::relative(2, 1)),
                Expression::number(8.0).expect("finite literal"),
            ],
        )
        .expect("supported native formula")
    };

    let mismatch = package
        .edit_table_cells("Sheet 1", "Table 1")?
        .set_formula_cached_a1("C3", authored_expression(), CachedValue::number(999.0)?)?
        .commit()
        .expect_err("a supplied cache inconsistent with a supported formula must refuse");
    assert!(
        matches!(
            mismatch,
            CellError::UnsupportedDependency {
                kind: DependencyKind::FormulaCache,
                ..
            }
        ),
        "unexpected cache-mismatch error: {mismatch:?}"
    );
    assert_eq!(exact_bytes(&package), source);

    let authored = package
        .edit_table_cells("Sheet 1", "Table 1")?
        .set_formula_a1("C3", authored_expression())?
        .commit()?;
    let authored_bytes = exact_bytes(authored.package());
    assert_ne!(authored_bytes, source);
    assert_formula_source(authored.package(), c3)?;
    assert!(matches!(
        formula_cell(&authored_bytes, c3)?.cached_scalar()?,
        Some(CachedScalar::Number(value)) if value.get() == 50.0
    ));
    let created_entries = formula_entries(&authored_bytes)?;
    assert_eq!(created_entries.len(), 1);
    assert_eq!(created_entries[0].refcount, 1);
    let created_formula = created_entries[0].formula.clone();
    assert_eq!(authored.diagnostics().changed_cells(), 1);
    assert_eq!(authored.diagnostics().refreshed_formula_caches(), 1);

    assert_eq!(
        exact_bytes(package.apply_table_cells(authored.patch())?.package()),
        authored_bytes
    );
    let reopened = Package::from_bytes(&authored_bytes)?;
    assert!(matches!(
        reopened.apply_table_cells(authored.patch()),
        Err(CellError::PatchConflict)
    ));
    assert_eq!(
        exact_bytes(
            reopened
                .apply_table_cells(&authored.patch().inverse())?
                .package()
        ),
        source
    );

    let noop = reopened
        .edit_table_cells(0usize, 0usize)?
        .set_formula_cached(c3, authored_expression(), CachedValue::number(50.0)?)?
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(noop.diagnostics().changed_cells(), 0);
    assert_eq!(noop.diagnostics().refreshed_formula_caches(), 0);
    assert_eq!(exact_bytes(noop.package()), authored_bytes);

    let replacement = reopened
        .edit_table_cells(0usize, 0usize)?
        .set_formula_cached_a1(
            "C3",
            Expression::function(
                "SUM",
                [
                    Expression::cell(CellReference::relative(2, 1)),
                    Expression::number(9.0)?,
                ],
            )?,
            CachedValue::number(51.0)?,
        )?
        .commit()?;
    let replacement_bytes = exact_bytes(replacement.package());
    let replacement_entries = formula_entries(&replacement_bytes)?;
    assert_eq!(replacement_entries.len(), 1);
    assert_eq!(replacement_entries[0].refcount, 1);
    assert_ne!(replacement_entries[0].formula, created_formula);
    assert!(matches!(
        formula_cell(&replacement_bytes, c3)?.cached_scalar()?,
        Some(CachedScalar::Number(value)) if value.get() == 51.0
    ));

    let cleared = replacement
        .package()
        .edit_table_cells(0usize, 0usize)?
        .clear(c3)?
        .commit()?;
    let cleared_bytes = exact_bytes(cleared.package());
    assert!(formula_entries(&cleared_bytes)?.is_empty());
    assert!(matches!(
        cleared.package().table_cell(0usize, 0usize, c3)?.storage(),
        Storage::Stored(Value::Empty)
    ));
    assert_eq!(
        exact_bytes(
            cleared
                .package()
                .apply_table_cells(&cleared.patch().inverse())?
                .package()
        ),
        replacement_bytes
    );
    Ok(())
}

#[test]
fn sequential_formula_overlays_preserve_survivors_and_exact_inverse() -> TestResult {
    let formula = |literal| {
        Expression::function(
            "SUM",
            [
                Expression::cell(CellReference::relative(2, 1)),
                Expression::number(literal).expect("finite literal"),
            ],
        )
        .expect("supported native formula")
    };
    let c3 = CellPosition::from_a1("C3")?;
    let d3 = CellPosition::from_a1("D3")?;
    let e3 = CellPosition::from_a1("E3")?;
    let base = Package::open(fixture_path())?;

    let first = base
        .edit_table_cells(0usize, 0usize)?
        .set_formula(c3, formula(8.0))?
        .commit()?;
    let first_bytes = exact_bytes(first.package());
    let first_reopened = Package::from_bytes(&first_bytes)?;

    let second = first_reopened
        .edit_table_cells(0usize, 0usize)?
        .set_formula(d3, formula(9.0))?
        .commit()?;
    let second_bytes = exact_bytes(second.package());
    assert_eq!(formula_entries(&second_bytes)?.len(), 2);
    assert_formula_source(second.package(), c3)?;
    assert_formula_source(second.package(), d3)?;
    assert!(matches!(
        formula_cell(&second_bytes, c3)?.cached_scalar()?,
        Some(CachedScalar::Number(value)) if value.get() == 50.0
    ));
    assert!(matches!(
        formula_cell(&second_bytes, d3)?.cached_scalar()?,
        Some(CachedScalar::Number(value)) if value.get() == 51.0
    ));
    assert_eq!(
        exact_bytes(
            second
                .package()
                .apply_table_cells(&second.patch().inverse())?
                .package()
        ),
        first_bytes
    );

    let replaced = Package::from_bytes(&second_bytes)?
        .edit_table_cells(0usize, 0usize)?
        .set_formula(c3, formula(10.0))?
        .commit()?;
    let replaced_bytes = exact_bytes(replaced.package());
    assert_eq!(formula_entries(&replaced_bytes)?.len(), 2);
    assert!(matches!(
        formula_cell(&replaced_bytes, c3)?.cached_scalar()?,
        Some(CachedScalar::Number(value)) if value.get() == 52.0
    ));
    assert!(matches!(
        formula_cell(&replaced_bytes, d3)?.cached_scalar()?,
        Some(CachedScalar::Number(value)) if value.get() == 51.0
    ));

    let mixed = Package::from_bytes(&replaced_bytes)?
        .edit_table_cells(0usize, 0usize)?
        .clear(c3)?
        .set_formula(e3, formula(11.0))?
        .commit()?;
    let mixed_bytes = exact_bytes(mixed.package());
    assert_eq!(formula_entries(&mixed_bytes)?.len(), 2);
    assert!(matches!(
        mixed.package().table_cell(0usize, 0usize, c3)?.storage(),
        Storage::Stored(Value::Empty)
    ));
    assert_formula_source(mixed.package(), d3)?;
    assert_formula_source(mixed.package(), e3)?;
    assert!(matches!(
        formula_cell(&mixed_bytes, d3)?.cached_scalar()?,
        Some(CachedScalar::Number(value)) if value.get() == 51.0
    ));
    assert!(matches!(
        formula_cell(&mixed_bytes, e3)?.cached_scalar()?,
        Some(CachedScalar::Number(value)) if value.get() == 53.0
    ));
    assert_eq!(
        exact_bytes(
            mixed
                .package()
                .apply_table_cells(&mixed.patch().inverse())?
                .package()
        ),
        replaced_bytes
    );

    let cleared = Package::from_bytes(&mixed_bytes)?
        .edit_table_cells(0usize, 0usize)?
        .clear(e3)?
        .commit()?;
    let cleared_bytes = exact_bytes(cleared.package());
    assert_eq!(formula_entries(&cleared_bytes)?.len(), 1);
    assert_formula_source(cleared.package(), d3)?;
    assert!(matches!(
        cleared.package().table_cell(0usize, 0usize, e3)?.storage(),
        Storage::Stored(Value::Empty)
    ));
    assert_eq!(
        exact_bytes(
            cleared
                .package()
                .apply_table_cells(&cleared.patch().inverse())?
                .package()
        ),
        mixed_bytes
    );
    Ok(())
}

#[test]
fn native_formula_evaluates_against_the_same_batch_final_scalar_overlay() -> TestResult {
    let package = Package::open(fixture_path())?;
    let source = exact_bytes(&package);
    let b3 = CellPosition::from_a1("B3")?;
    let c3 = CellPosition::from_a1("C3")?;
    let expression = Expression::function(
        "SUM",
        [
            Expression::cell(CellReference::relative(2, 1)),
            Expression::number(8.0)?,
        ],
    )?;

    let commit = package
        .edit_table_cells(0usize, 0usize)?
        .set(b3, Input::number(43.0)?)?
        .set_formula(c3, expression)?
        .commit()?;
    let target = exact_bytes(commit.package());
    assert!(matches!(
        formula_cell(&target, c3)?.cached_scalar()?,
        Some(CachedScalar::Number(value)) if value.get() == 51.0
    ));
    assert_eq!(
        exact_bytes(package.apply_table_cells(commit.patch())?.package()),
        target
    );
    assert_eq!(
        exact_bytes(
            commit
                .package()
                .apply_table_cells(&commit.patch().inverse())?
                .package()
        ),
        source
    );
    Ok(())
}

#[test]
fn local_formula_and_clear_in_multi_owner_engine_preserve_unrelated_owner() -> TestResult {
    let package = two_table_formula_package()?;
    let source = exact_bytes(&package);
    let c3 = CellPosition::from_a1("C3")?;
    let authored = package
        .edit_table_cells("Sheet 1", "Table 1")?
        .set_formula(c3, Expression::cell(CellReference::relative(2, 1)))?
        .commit()?;
    let authored_bytes = exact_bytes(authored.package());
    assert_formula_source(authored.package(), c3)?;
    assert_eq!(formula_entries(&authored_bytes)?.len(), 1);
    assert!(matches!(
        formula_cell(&authored_bytes, c3)?.cached_scalar()?,
        Some(CachedScalar::Number(value)) if value.get() == 42.0
    ));
    assert_eq!(
        exact_bytes(package.apply_table_cells(authored.patch())?.package()),
        authored_bytes
    );
    assert_eq!(
        exact_bytes(
            authored
                .package()
                .apply_table_cells(&authored.patch().inverse())?
                .package()
        ),
        source
    );

    let cleared = Package::from_bytes(&authored_bytes)?
        .edit_table_cells("Sheet 1", "Table 1")?
        .clear(c3)?
        .commit()?;
    let cleared_bytes = exact_bytes(cleared.package());
    assert!(formula_entries(&cleared_bytes)?.is_empty());
    assert!(matches!(
        cleared.package().table_cell(0usize, 0usize, c3)?.storage(),
        Storage::Stored(Value::Empty)
    ));
    assert_eq!(
        exact_bytes(
            cleared
                .package()
                .apply_table_cells(&cleared.patch().inverse())?
                .package()
        ),
        authored_bytes
    );
    Ok(())
}

#[test]
fn every_local_and_distinct_owner_reference_form_commits_and_reopens() -> TestResult {
    let package = two_table_formula_package()?;
    let source = exact_bytes(&package);
    let edit = package.edit_table_cells("Sheet 1", "Table 1")?;
    let table = edit.formula_table("Sheet 1", "External")?;
    let sum = |expression| Expression::function("SUM", [expression]);
    let commit = edit
        .set_a1("B3", Input::number(43.0)?)?
        .set_formula_a1("C3", Expression::cell(CellReference::relative(2, 1)))?
        .set_formula_a1(
            "D3",
            sum(Expression::range(
                CellReference::relative(1, 1),
                CellReference::relative(3, 1),
            )?)?,
        )?
        .set_formula_a1(
            "E3",
            sum(Expression::rows(
                AxisReference::absolute(21),
                AxisReference::absolute(21),
            )?)?,
        )?
        .set_formula_a1(
            "F3",
            sum(Expression::columns(
                AxisReference::absolute(0),
                AxisReference::absolute(0),
            )?)?,
        )?
        .set_formula_a1(
            "G3",
            Expression::table_cell(&table, CellReference::relative(2, 1)),
        )?
        .set_formula_a1(
            "C4",
            sum(Expression::table_range(
                &table,
                CellReference::relative(1, 1),
                CellReference::relative(3, 1),
            )?)?,
        )?
        .set_formula_a1(
            "D4",
            sum(Expression::table_rows(
                &table,
                AxisReference::absolute(21),
                AxisReference::absolute(21),
            )?)?,
        )?
        .set_formula_a1(
            "E4",
            sum(Expression::table_columns(
                &table,
                AxisReference::absolute(0),
                AxisReference::absolute(0),
            )?)?,
        )?
        .commit()?;

    assert_eq!(commit.diagnostics().requested_cells(), 9);
    assert_eq!(commit.diagnostics().changed_cells(), 9);
    assert_eq!(commit.diagnostics().refreshed_formula_caches(), 8);
    let target = exact_bytes(commit.package());
    assert_eq!(
        exact_bytes(package.apply_table_cells(commit.patch())?.package()),
        target
    );
    let reopened = Package::from_bytes(&target)?;
    for address in ["C3", "D3", "E3", "F3", "G3", "C4", "D4", "E4"] {
        assert_formula_source(&reopened, CellPosition::from_a1(address)?)?;
    }
    for (address, expected) in [
        ("C3", 43.0),
        ("D3", 43.0),
        ("E3", 0.0),
        ("F3", 0.0),
        ("G3", 42.0),
        ("C4", 42.0),
        ("D4", 0.0),
        ("E4", 0.0),
    ] {
        assert!(matches!(
            formula_cell(&target, CellPosition::from_a1(address)?)?.cached_scalar()?,
            Some(CachedScalar::Number(value)) if value.get() == expected
        ));
    }
    let entries = formula_entries(&target)?;
    assert_eq!(entries.len(), 8);
    assert!(entries.iter().all(|entry| entry.refcount == 1));

    let refreshed = reopened
        .edit_table_cells("Sheet 1", "Table 1")?
        .set_a1("B3", Input::number(44.0)?)?
        .commit()?;
    let refreshed_bytes = exact_bytes(refreshed.package());
    for address in ["C3", "D3"] {
        assert!(matches!(
            formula_cell(&refreshed_bytes, CellPosition::from_a1(address)?)?.cached_scalar()?,
            Some(CachedScalar::Number(value)) if value.get() == 44.0
        ));
    }
    assert!(matches!(
        formula_cell(&refreshed_bytes, CellPosition::from_a1("G3")?)?.cached_scalar()?,
        Some(CachedScalar::Number(value)) if value.get() == 42.0
    ));
    assert_eq!(
        exact_bytes(reopened.apply_table_cells(refreshed.patch())?.package()),
        refreshed_bytes
    );
    assert_eq!(
        exact_bytes(
            refreshed
                .package()
                .apply_table_cells(&refreshed.patch().inverse())?
                .package()
        ),
        target
    );
    assert_eq!(
        exact_bytes(
            reopened
                .apply_table_cells(&commit.patch().inverse())?
                .package()
        ),
        source
    );
    Ok(())
}
