//! Exact-source integration coverage for Numbers table appearance ownership.
//!
//! The fixture deliberately keeps the table model in a Document component and
//! the shared parent style plus stylesheet in a separate component.  Metadata
//! therefore has to be part of the transaction: a copy-on-write style changes
//! the selected model, the stylesheet registry, the UUID map, the external
//! edge, and the selected current-component save tokens together.

use std::io;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::wire::{WireView, append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, FieldType, RawMessage, SnappyStream};
use litchi_iwa_protos::{tn, tsd, tsp, tss, tst};
use litchi_numbers::{
    Appearance, Banding, GridlineVisibility, Gridlines, Package, PackageLimits, PackageReadOptions,
    PackageSemanticLimits, RowSizing,
};
use prost::Message as _;

use litchi_numbers::table::appearance::transaction::{Error, Path};

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const STYLESHEET_MEMBER: &str = "Index/DocumentStylesheet.iwa";
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const VIEW_STATE_MEMBER: &str = "Index/ViewState.iwa";
const PACKAGE_METADATA_TYPE: u32 = 11_006;
const TABLE_INFO_TYPE: u32 = 6_000;
const TABLE_MODEL_TYPE: u32 = 6_001;
const TABLE_STYLE_TYPE: u32 = 6_003;
const STYLESHEET_TYPE: u32 = 401;
const DATA_LIST_TYPE: u32 = 6_005;

const DOCUMENT_ID: u64 = 1;
const SHEET_ID: u64 = 2;
const FIRST_INFO_ID: u64 = 10;
const SECOND_INFO_ID: u64 = 11;
const FIRST_MODEL_ID: u64 = 20;
const SECOND_MODEL_ID: u64 = 21;
const SIDECARS_ID: u64 = 90;
const PARENT_STYLE_ID: u64 = 30;
const UNUSED_STYLE_ID: u64 = 31;
const STYLESHEET_ID: u64 = 40;
const VIEW_STATE_ID: u64 = 300;
const METADATA_OBJECT_ID: u64 = 900;

const PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

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

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn object(identifier: u64, type_: u32, data: Vec<u8>) -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage { type_, data }],
    )?)
}

fn add_model_archive_metadata(object: &mut ArchiveObject, style_identifier: u64) {
    let info = &mut object.archive_info.message_infos[0];
    info.data_references = vec![700 + style_identifier];
    let mut style_field = FieldInfo::new(vec![3]);
    style_field.object_references = vec![style_identifier];
    style_field.data_references = vec![800 + style_identifier];
    let mut unknown_tail = FieldInfo::new(vec![99, 1]);
    unknown_tail.data_references = vec![900 + style_identifier];
    info.field_infos = vec![style_field, unknown_tail];
}

fn compressed(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

fn table_model(identifier: u64) -> tst::TableModelArchive {
    tst::TableModelArchive {
        table_id: format!("appearance-table-{identifier}"),
        table_name: format!("Table {identifier}"),
        number_of_rows: 2,
        number_of_columns: 2,
        table_style: reference(PARENT_STYLE_ID),
        // Match native Numbers producers that retain a non-zero preset hint
        // alongside the direct style reference.  The direct style remains
        // authoritative for appearance resolution and COW publication.
        table_style_preset: (identifier == FIRST_MODEL_ID).then(|| reference(UNUSED_STYLE_ID)),
        base_data_store: tst::DataStore {
            string_table: reference(SIDECARS_ID),
            formula_table: reference(SIDECARS_ID),
            ..Default::default()
        },
        ..Default::default()
    }
}

fn table_info(model_identifier: u64) -> TestResult<ArchiveObject> {
    let payload = tst::TableInfoArchive {
        super_: tsd::DrawableArchive::default(),
        table_model: reference(model_identifier),
        ..Default::default()
    }
    .encode_to_vec();
    let mut info = object(FIRST_INFO_ID, TABLE_INFO_TYPE, payload)?;
    info.archive_info.message_infos[0].object_references = vec![model_identifier];
    Ok(info)
}

fn second_table_info() -> TestResult<ArchiveObject> {
    let payload = tst::TableInfoArchive {
        super_: tsd::DrawableArchive::default(),
        table_model: reference(SECOND_MODEL_ID),
        ..Default::default()
    }
    .encode_to_vec();
    let mut info = object(SECOND_INFO_ID, TABLE_INFO_TYPE, payload)?;
    info.archive_info.message_infos[0].object_references = vec![SECOND_MODEL_ID];
    Ok(info)
}

fn table_style_properties() -> tst::TableStylePropertiesArchive {
    tst::TableStylePropertiesArchive {
        banded_rows: Some(false),
        auto_resize: Some(false),
        h_strokes_visible: Some(true),
        v_strokes_visible: Some(true),
        table_hc_divider_visible: Some(true),
        table_hr_divider_visible: Some(true),
        table_footer_divider_visible: Some(true),
        ..Default::default()
    }
}

fn table_style(identifier: u64, parent: Option<u64>) -> TestResult<ArchiveObject> {
    let mut properties = table_style_properties();
    if identifier == UNUSED_STYLE_ID {
        properties.banded_rows = Some(true);
        properties.auto_resize = Some(true);
    }
    let mut data = tst::TableStyleArchive {
        super_: tss::StyleArchive {
            name: Some(format!("Shared table style {identifier}")),
            style_identifier: Some(format!("appearance-style-{identifier}")),
            parent: parent.map(reference),
            is_variation: Some(parent.is_some()),
            stylesheet: Some(reference(STYLESHEET_ID)),
        },
        override_count: Some(0),
        table_properties: Some(properties),
    }
    .encode_to_vec();

    // This unknown field is intentionally retained by a COW style rewrite.
    append_varint_field(&mut data, 90, 901 + identifier)?;
    object(identifier, TABLE_STYLE_TYPE, data)
}

fn stylesheet() -> TestResult<ArchiveObject> {
    let payload = tss::StylesheetArchive {
        styles: vec![reference(PARENT_STYLE_ID), reference(UNUSED_STYLE_ID)],
        identifier_to_style_map: vec![
            tss::stylesheet_archive::IdentifiedStyleEntry {
                identifier: "shared-appearance".to_owned(),
                style: reference(PARENT_STYLE_ID),
            },
            tss::stylesheet_archive::IdentifiedStyleEntry {
                identifier: "unused-appearance".to_owned(),
                style: reference(UNUSED_STYLE_ID),
            },
        ],
        parent_to_children_style_map: vec![tss::stylesheet_archive::StyleChildrenEntry {
            parent: reference(PARENT_STYLE_ID),
            children: vec![reference(UNUSED_STYLE_ID)],
        }],
        is_locked: Some(false),
        can_cull_styles: Some(true),
        ..Default::default()
    }
    .encode_to_vec();
    let mut stylesheet = object(STYLESHEET_ID, STYLESHEET_TYPE, payload)?;
    stylesheet.archive_info.message_infos[0].object_references =
        vec![PARENT_STYLE_ID, UNUSED_STYLE_ID];
    Ok(stylesheet)
}

fn sidecars() -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        SIDECARS_ID,
        [
            tst::table_data_list::ListType::String,
            tst::table_data_list::ListType::Formula,
        ]
        .into_iter()
        .map(|list_type| RawMessage {
            type_: DATA_LIST_TYPE,
            data: tst::TableDataList {
                list_type: list_type as i32,
                next_list_id: 1,
                ..Default::default()
            }
            .encode_to_vec(),
        })
        .collect(),
    )?)
}

fn uuid_entry(identifier: u64) -> tsp::ObjectUuidMapEntry {
    tsp::ObjectUuidMapEntry {
        identifier,
        uuid: tsp::Uuid {
            lower: identifier,
            upper: identifier.saturating_add(1000),
        },
    }
}

fn metadata_component(
    identifier: u64,
    locator: &str,
    token: u64,
    object_ids: &[u64],
    external_references: &[tsp::ComponentExternalReference],
) -> TestResult<Vec<u8>> {
    let mut data = tsp::ComponentInfo {
        identifier,
        preferred_locator: locator.to_owned(),
        locator: Some(locator.to_owned()),
        save_token: Some(token),
        object_uuid_map_entries: object_ids.iter().copied().map(uuid_entry).collect(),
        external_references: external_references.to_vec(),
        ..Default::default()
    }
    .encode_to_vec();
    append_varint_field(&mut data, 90, 10_000 + identifier)?;
    Ok(data)
}

fn metadata_payload() -> TestResult<Vec<u8>> {
    let document_external = tsp::ComponentExternalReference {
        component_identifier: 200,
        object_identifier: Some(PARENT_STYLE_ID),
        is_weak: None,
    };
    let stylesheet = metadata_component(
        200,
        "DocumentStylesheet",
        8,
        &[PARENT_STYLE_ID, UNUSED_STYLE_ID, STYLESHEET_ID],
        &[],
    )?;
    let document = metadata_component(
        100,
        "Document",
        9,
        &[
            DOCUMENT_ID,
            SHEET_ID,
            FIRST_INFO_ID,
            SECOND_INFO_ID,
            FIRST_MODEL_ID,
            SECOND_MODEL_ID,
            SIDECARS_ID,
        ],
        &[document_external],
    )?;
    let view = metadata_component(300, "ViewState", 7, &[VIEW_STATE_ID], &[])?;
    let versioned = metadata_component(100, "Document", 3, &[999], &[])?;
    let mut data = tsp::PackageMetadata {
        last_object_identifier: 1_000,
        save_token: Some(10),
        ..Default::default()
    }
    .encode_to_vec();
    append_length_delimited_field(&mut data, 3, &document)?;
    append_length_delimited_field(&mut data, 3, &stylesheet)?;
    append_length_delimited_field(&mut data, 3, &view)?;
    append_length_delimited_field(&mut data, 11, &versioned)?;
    append_varint_field(&mut data, 90, 900)?;
    Ok(data)
}

fn metadata_entry() -> TestResult<Vec<u8>> {
    compressed(vec![object(
        METADATA_OBJECT_ID,
        PACKAGE_METADATA_TYPE,
        metadata_payload()?,
    )?])
}

fn metadata_payload_from_bytes(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == METADATA_MEMBER)
        .ok_or_else(|| io::Error::other("metadata member is missing"))?;
    let stream = SnappyStream::decompress(entry.data())?;
    let archive = Archive::parse(stream.as_bytes())?;
    archive
        .objects
        .iter()
        .find(|object| object.archive_info.identifier == Some(METADATA_OBJECT_ID))
        .and_then(|object| {
            object
                .messages
                .iter()
                .find(|message| message.type_ == PACKAGE_METADATA_TYPE)
        })
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("PackageMetadata object is missing").into())
}

fn metadata_components(
    source: &[u8],
    field_number: u32,
) -> TestResult<Vec<(tsp::ComponentInfo, Vec<u8>)>> {
    let payload = metadata_payload_from_bytes(source)?;
    WireView::parse(&payload)?
        .fields()
        .filter(|field| field.number() == field_number)
        .map(|field| {
            let raw = field.payload().to_owned();
            let component = tsp::ComponentInfo::decode(raw.as_slice())
                .map_err(|error| io::Error::other(error.to_string()))?;
            Ok((component, raw))
        })
        .collect::<TestResult<Vec<_>>>()
}

fn find_metadata_component(
    source: &[u8],
    field_number: u32,
    locator: &str,
    identifier: u64,
) -> TestResult<(tsp::ComponentInfo, Vec<u8>)> {
    metadata_components(source, field_number)?
        .into_iter()
        .find(|(component, _)| {
            component.identifier == identifier && component.preferred_locator == locator
        })
        .ok_or_else(|| io::Error::other(format!("metadata component {locator} is missing")).into())
}

fn metadata_root(source: &[u8]) -> TestResult<tsp::PackageMetadata> {
    Ok(tsp::PackageMetadata::decode(
        metadata_payload_from_bytes(source)?.as_slice(),
    )?)
}

fn object_messages(
    source: &[u8],
    member_name: &str,
    identifier: u64,
) -> TestResult<Vec<(u32, Vec<u8>)>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member_name)
        .ok_or_else(|| io::Error::other(format!("member {member_name} is missing")))?;
    let stream = SnappyStream::decompress(entry.data())?;
    let archive = Archive::parse(stream.as_bytes())?;
    let object = archive
        .objects
        .iter()
        .find(|object| object.archive_info.identifier == Some(identifier))
        .ok_or_else(|| io::Error::other(format!("object {identifier} is missing")))?;
    Ok(object
        .messages
        .iter()
        .map(|message| (message.type_, message.data.clone()))
        .collect())
}

fn object_references(
    source: &[u8],
    member_name: &str,
    identifier: u64,
) -> TestResult<Vec<Vec<u64>>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member_name)
        .ok_or_else(|| io::Error::other(format!("member {member_name} is missing")))?;
    let stream = SnappyStream::decompress(entry.data())?;
    let archive = Archive::parse(stream.as_bytes())?;
    let object = archive
        .objects
        .iter()
        .find(|object| object.archive_info.identifier == Some(identifier))
        .ok_or_else(|| io::Error::other(format!("object {identifier} is missing")))?;
    Ok(object
        .archive_info
        .message_infos
        .iter()
        .map(|info| info.object_references.clone())
        .collect())
}

fn object_message_metadata(
    source: &[u8],
    member_name: &str,
    identifier: u64,
) -> TestResult<litchi_iwa_core::MessageInfo> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member_name)
        .ok_or_else(|| io::Error::other(format!("member {member_name} is missing")))?;
    let stream = SnappyStream::decompress(entry.data())?;
    let archive = Archive::parse(stream.as_bytes())?;
    let object = archive
        .objects
        .iter()
        .find(|object| object.archive_info.identifier == Some(identifier))
        .ok_or_else(|| io::Error::other(format!("object {identifier} is missing")))?;
    object
        .archive_info
        .message_infos
        .first()
        .cloned()
        .ok_or_else(|| io::Error::other(format!("object {identifier} has no metadata")).into())
}

fn rewrite_object_message<F>(
    source: &[u8],
    member_name: &str,
    identifier: u64,
    rewrite: F,
) -> TestResult<Vec<u8>>
where
    F: FnOnce(&mut RawMessage) -> TestResult,
{
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member_name)
        .ok_or_else(|| io::Error::other(format!("member {member_name} is missing")))?;
    let stream = SnappyStream::decompress(entry.data())?;
    let mut archive = Archive::parse(stream.as_bytes())?;
    let object = archive
        .object_mut(identifier)
        .ok_or_else(|| io::Error::other(format!("object {identifier} is missing")))?;
    let message = object
        .messages
        .first_mut()
        .ok_or_else(|| io::Error::other(format!("object {identifier} has no message")))?;
    let mut replacement = RawMessage {
        type_: message.type_,
        data: message.data.clone(),
    };
    rewrite(&mut replacement)?;
    object.replace_message_preserving_header(0, replacement)?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            member_name,
            &compressed,
        )],
        Limits::default(),
    )?)
}

fn rewrite_object_message_type(
    source: &[u8],
    member_name: &str,
    identifier: u64,
    type_: u32,
) -> TestResult<Vec<u8>> {
    rewrite_object_message(source, member_name, identifier, |message| {
        message.type_ = type_;
        Ok(())
    })
}

fn rewrite_object_archive_info<F>(
    source: &[u8],
    member_name: &str,
    identifier: u64,
    rewrite: F,
) -> TestResult<Vec<u8>>
where
    F: FnOnce(&mut litchi_iwa_core::MessageInfo) -> TestResult,
{
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member_name)
        .ok_or_else(|| io::Error::other(format!("member {member_name} is missing")))?;
    let stream = SnappyStream::decompress(entry.data())?;
    let mut archive = Archive::parse(stream.as_bytes())?;
    let object = archive
        .object_mut(identifier)
        .ok_or_else(|| io::Error::other(format!("object {identifier} is missing")))?;
    let info = object
        .archive_info
        .message_infos
        .first_mut()
        .ok_or_else(|| io::Error::other(format!("object {identifier} has no metadata")))?;
    rewrite(info)?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            member_name,
            &compressed,
        )],
        Limits::default(),
    )?)
}

fn rewrite_metadata<F>(source: &[u8], rewrite: F) -> TestResult<Vec<u8>>
where
    F: FnOnce(&mut tsp::PackageMetadata),
{
    rewrite_object_message(source, METADATA_MEMBER, METADATA_OBJECT_ID, |message| {
        let mut metadata = tsp::PackageMetadata::decode(message.data.as_slice())?;
        rewrite(&mut metadata);
        message.data = metadata.encode_to_vec();
        Ok(())
    })
}

fn assert_appearance_read_rejects(source: &[u8]) -> TestResult {
    match Package::from_bytes(source) {
        Err(_) => Ok(()),
        Ok(package) => match package.table_appearance(0usize, 0usize) {
            Err(Error::InvalidSource { .. })
            | Err(Error::UnsupportedDependency { .. })
            | Err(Error::UnsupportedSource) => Ok(()),
            Ok(_) => Err(io::Error::other("malformed appearance graph was accepted").into()),
            Err(error) => Err(io::Error::other(format!("unexpected error: {error}")).into()),
        },
    }
}

fn assert_changed_edit_rejects_atomically(source: &[u8], label: &str) -> TestResult {
    match Package::from_bytes(source) {
        Err(_) => Ok(()),
        Ok(package) => {
            let before = package.exact_bytes();
            assert!(
                package
                    .edit_table_appearance(0usize, 0usize)?
                    .set(custom_appearance())
                    .commit()
                    .is_err(),
                "hostile archive metadata must not publish an appearance edit: {label}"
            );
            assert_eq!(package.exact_bytes(), before);
            Ok(())
        },
    }
}

fn assert_reserved_identifier_is_not_reused(source: &[u8]) -> TestResult {
    let package = Package::from_bytes(source)?;
    let commit = package
        .edit_table_appearance(0usize, 0usize)?
        .set(custom_appearance())
        .commit()
        .expect("reserved metadata identifiers must be skipped by COW allocation");
    let target = commit.package().exact_bytes();
    assert_eq!(metadata_root(&target)?.last_object_identifier, 1_002);
    assert_eq!(
        object_references(&target, DOCUMENT_MEMBER, FIRST_MODEL_ID)?,
        vec![vec![SIDECARS_ID, 1_002]]
    );
    assert_eq!(
        object_references(&target, STYLESHEET_MEMBER, 1_002)?,
        vec![vec![PARENT_STYLE_ID, STYLESHEET_ID]]
    );
    assert_eq!(
        find_metadata_component(&target, 3, "DocumentStylesheet", 200)?
            .0
            .object_uuid_map_entries
            .iter()
            .filter(|entry| entry.identifier == 1_002)
            .count(),
        1
    );
    assert_eq!(
        Package::from_bytes(&target)?
            .apply_table_appearance(&commit.patch().inverse())?
            .package()
            .exact_bytes(),
        source
    );
    Ok(())
}

fn raw_fields(source: &[u8], field_number: u32) -> TestResult<Vec<Vec<u8>>> {
    Ok(WireView::parse(&metadata_payload_from_bytes(source)?)?
        .fields()
        .filter(|field| field.number() == field_number)
        .map(|field| field.raw().to_owned())
        .collect())
}

fn without_wire_field(source: &[u8], field_number: u32) -> TestResult<Vec<u8>> {
    without_wire_fields(source, &[field_number])
}

fn without_wire_fields(source: &[u8], field_numbers: &[u32]) -> TestResult<Vec<u8>> {
    let view = WireView::parse(source)?;
    let mut output = Vec::new();
    for field in view
        .fields()
        .filter(|field| !field_numbers.contains(&field.number()))
    {
        output.extend_from_slice(field.raw());
    }
    Ok(output)
}

fn append_nested_wire_field(
    source: &[u8],
    outer_field: u32,
    nested_field: &[u8],
) -> TestResult<Vec<u8>> {
    let view = WireView::parse(source)?;
    let mut output = Vec::new();
    let mut appended = false;
    for field in view.fields() {
        if field.number() != outer_field {
            output.extend_from_slice(field.raw());
            continue;
        }
        let nested = WireView::parse(field.payload())?;
        let mut payload = Vec::new();
        for child in nested.fields() {
            payload.extend_from_slice(child.raw());
        }
        payload.extend_from_slice(nested_field);
        append_length_delimited_field(&mut output, outer_field, &payload)?;
        appended = true;
    }
    if !appended {
        return Err(io::Error::other("nested protobuf parent field is missing").into());
    }
    Ok(output)
}

fn replace_nested_wire_field(
    source: &[u8],
    outer_field: u32,
    nested_field: u32,
    replacement: &[u8],
) -> TestResult<Vec<u8>> {
    let view = WireView::parse(source)?;
    let mut output = Vec::new();
    let mut replaced = false;
    for field in view.fields() {
        if field.number() != outer_field {
            output.extend_from_slice(field.raw());
            continue;
        }
        let nested = WireView::parse(field.payload())?;
        let mut payload = Vec::new();
        for child in nested.fields() {
            if child.number() == nested_field && !replaced {
                payload.extend_from_slice(replacement);
                replaced = true;
            } else {
                payload.extend_from_slice(child.raw());
            }
        }
        append_length_delimited_field(&mut output, outer_field, &payload)?;
    }
    if !replaced {
        return Err(io::Error::other("nested protobuf field is missing").into());
    }
    Ok(output)
}

fn unknown_component_fields(raw: &[u8]) -> TestResult<Vec<Vec<u8>>> {
    Ok(WireView::parse(raw)?
        .fields()
        .filter(|field| field.number() == 90)
        .map(|field| field.raw().to_owned())
        .collect())
}

fn wire_fields(source: &[u8], field_number: u32) -> TestResult<Vec<Vec<u8>>> {
    Ok(WireView::parse(source)?
        .fields()
        .filter(|field| field.number() == field_number)
        .map(|field| field.raw().to_owned())
        .collect())
}

fn metadata_transition_assertions(source: &[u8], target: &[u8]) -> TestResult {
    let source_root = metadata_root(source)?;
    let target_root = metadata_root(target)?;
    assert_eq!(source_root.last_object_identifier, 1_000);
    assert_eq!(target_root.last_object_identifier, 1_001);
    assert_eq!(source_root.save_token, Some(10));
    assert_eq!(target_root.save_token, Some(11));

    let (document_source, _) = find_metadata_component(source, 3, "Document", 100)?;
    let (document_target, _) = find_metadata_component(target, 3, "Document", 100)?;
    let (stylesheet_source, _) = find_metadata_component(source, 3, "DocumentStylesheet", 200)?;
    let (stylesheet_target, _) = find_metadata_component(target, 3, "DocumentStylesheet", 200)?;
    let (view_source, view_source_raw) = find_metadata_component(source, 3, "ViewState", 300)?;
    let (view_target, view_target_raw) = find_metadata_component(target, 3, "ViewState", 300)?;
    let (_, versioned_source_raw) = find_metadata_component(source, 11, "Document", 100)?;
    let (_, versioned_target_raw) = find_metadata_component(target, 11, "Document", 100)?;

    assert_eq!(document_source.save_token, Some(9));
    assert_eq!(document_target.save_token, Some(11));
    assert_eq!(stylesheet_source.save_token, Some(8));
    assert_eq!(stylesheet_target.save_token, Some(11));
    assert_eq!(view_source.save_token, Some(7));
    assert_eq!(view_target.save_token, Some(7));
    assert_eq!(view_source_raw, view_target_raw);
    assert_eq!(versioned_source_raw, versioned_target_raw);
    assert_eq!(raw_fields(source, 90)?, raw_fields(target, 90)?);
    assert_eq!(
        unknown_component_fields(&find_metadata_component(source, 3, "Document", 100)?.1)?,
        unknown_component_fields(&find_metadata_component(target, 3, "Document", 100)?.1)?
    );
    assert_eq!(
        unknown_component_fields(
            &find_metadata_component(source, 3, "DocumentStylesheet", 200,)?.1
        )?,
        unknown_component_fields(
            &find_metadata_component(target, 3, "DocumentStylesheet", 200,)?.1
        )?
    );

    let source_document_ids: Vec<_> = document_source
        .object_uuid_map_entries
        .iter()
        .map(|entry| entry.identifier)
        .collect();
    let target_document_ids: Vec<_> = document_target
        .object_uuid_map_entries
        .iter()
        .map(|entry| entry.identifier)
        .collect();
    assert!(!source_document_ids.contains(&1_001));
    assert!(!target_document_ids.contains(&1_001));
    assert_eq!(
        stylesheet_target
            .object_uuid_map_entries
            .iter()
            .filter(|entry| entry.identifier == 1_001)
            .count(),
        1
    );
    assert_eq!(
        target_root
            .components
            .iter()
            .flat_map(|component| component.object_uuid_map_entries.iter())
            .filter(|entry| entry.identifier == 1_001)
            .count(),
        1
    );
    assert_eq!(
        target_root
            .versioned_components
            .iter()
            .flat_map(|component| component.object_uuid_map_entries.iter())
            .filter(|entry| entry.identifier == 1_001)
            .count(),
        0
    );

    let new_edges: Vec<_> = document_target
        .external_references
        .iter()
        .filter(|reference| {
            reference.component_identifier == 200 && reference.object_identifier == Some(1_001)
        })
        .collect();
    assert_eq!(new_edges.len(), 1);
    assert_eq!(
        document_source
            .external_references
            .iter()
            .filter(|reference| {
                reference.component_identifier == 200
                    && reference.object_identifier == Some(PARENT_STYLE_ID)
            })
            .count(),
        1
    );
    assert_eq!(
        document_target
            .external_references
            .iter()
            .filter(|reference| {
                reference.component_identifier == 200
                    && reference.object_identifier == Some(PARENT_STYLE_ID)
            })
            .count(),
        1
    );
    let source_model = object_messages(source, DOCUMENT_MEMBER, FIRST_MODEL_ID)?;
    let target_model = object_messages(target, DOCUMENT_MEMBER, FIRST_MODEL_ID)?;
    assert_eq!(source_model.len(), target_model.len());
    let source_style = object_messages(source, STYLESHEET_MEMBER, PARENT_STYLE_ID)?;
    let target_style = object_messages(target, STYLESHEET_MEMBER, PARENT_STYLE_ID)?;
    assert_eq!(source_style, target_style);
    let sibling_source = object_messages(source, DOCUMENT_MEMBER, SECOND_MODEL_ID)?;
    let sibling_target = object_messages(target, DOCUMENT_MEMBER, SECOND_MODEL_ID)?;
    assert_eq!(sibling_source, sibling_target);
    assert_eq!(
        object_references(source, DOCUMENT_MEMBER, FIRST_MODEL_ID)?,
        vec![vec![SIDECARS_ID, PARENT_STYLE_ID]]
    );
    assert_eq!(
        object_references(target, DOCUMENT_MEMBER, FIRST_MODEL_ID)?,
        vec![vec![SIDECARS_ID, 1_001]]
    );
    assert_eq!(
        object_references(target, STYLESHEET_MEMBER, 1_001)?,
        vec![vec![PARENT_STYLE_ID, STYLESHEET_ID]]
    );
    let source_model_info = object_message_metadata(source, DOCUMENT_MEMBER, FIRST_MODEL_ID)?;
    let target_model_info = object_message_metadata(target, DOCUMENT_MEMBER, FIRST_MODEL_ID)?;
    assert_eq!(
        source_model_info.data_references,
        target_model_info.data_references
    );
    assert_eq!(source_model_info.field_infos.len(), 2);
    assert_eq!(target_model_info.field_infos.len(), 2);
    assert_eq!(source_model_info.field_infos[0].path.as_slice(), [3]);
    assert_eq!(target_model_info.field_infos[0].path.as_slice(), [3]);
    assert_eq!(
        source_model_info.field_infos[0].object_references,
        [PARENT_STYLE_ID]
    );
    assert_eq!(target_model_info.field_infos[0].object_references, [1_001]);
    assert_eq!(
        source_model_info.field_infos[0].data_references,
        target_model_info.field_infos[0].data_references
    );
    assert_eq!(
        source_model_info.field_infos[1],
        target_model_info.field_infos[1]
    );
    let source_stylesheet_info = object_message_metadata(source, STYLESHEET_MEMBER, STYLESHEET_ID)?;
    let target_stylesheet_info = object_message_metadata(target, STYLESHEET_MEMBER, STYLESHEET_ID)?;
    assert_eq!(
        source_stylesheet_info.object_references,
        vec![PARENT_STYLE_ID, UNUSED_STYLE_ID]
    );
    assert_eq!(
        target_stylesheet_info.object_references,
        vec![PARENT_STYLE_ID, UNUSED_STYLE_ID, 1_001]
    );
    assert_eq!(
        source_stylesheet_info.data_references,
        target_stylesheet_info.data_references
    );
    assert_eq!(source_stylesheet_info.field_infos.len(), 2);
    assert_eq!(target_stylesheet_info.field_infos.len(), 2);
    assert_eq!(source_stylesheet_info.field_infos[0].path.as_slice(), [1]);
    assert_eq!(target_stylesheet_info.field_infos[0].path.as_slice(), [1]);
    assert_eq!(
        source_stylesheet_info.field_infos[0].object_references,
        [PARENT_STYLE_ID, UNUSED_STYLE_ID]
    );
    assert_eq!(
        target_stylesheet_info.field_infos[0].object_references,
        [PARENT_STYLE_ID, UNUSED_STYLE_ID, 1_001]
    );
    assert_eq!(
        source_stylesheet_info.field_infos[0].data_references,
        target_stylesheet_info.field_infos[0].data_references
    );
    assert_eq!(
        source_stylesheet_info.field_infos[1],
        target_stylesheet_info.field_infos[1]
    );
    let source_stylesheet_payload = object_messages(source, STYLESHEET_MEMBER, STYLESHEET_ID)?
        .first()
        .ok_or_else(|| io::Error::other("stylesheet payload is missing"))?
        .1
        .clone();
    let target_stylesheet_payload = object_messages(target, STYLESHEET_MEMBER, STYLESHEET_ID)?
        .first()
        .ok_or_else(|| io::Error::other("target stylesheet payload is missing"))?
        .1
        .clone();
    let source_registry = tss::StylesheetArchive::decode(source_stylesheet_payload.as_slice())?;
    let target_registry = tss::StylesheetArchive::decode(target_stylesheet_payload.as_slice())?;
    assert_eq!(
        source_registry.identifier_to_style_map,
        target_registry.identifier_to_style_map
    );
    assert_eq!(
        source_registry.styles,
        vec![reference(PARENT_STYLE_ID), reference(UNUSED_STYLE_ID)]
    );
    assert_eq!(
        target_registry
            .styles
            .iter()
            .filter(|style| style.identifier == 1_001)
            .count(),
        1
    );
    let parent_entries: Vec<_> = target_registry
        .parent_to_children_style_map
        .iter()
        .filter(|entry| entry.parent == reference(PARENT_STYLE_ID))
        .collect();
    assert_eq!(parent_entries.len(), 1);
    assert_eq!(
        parent_entries[0]
            .children
            .iter()
            .filter(|child| child.identifier == 1_001)
            .count(),
        1
    );
    let target_new_style = object_messages(target, STYLESHEET_MEMBER, 1_001)?;
    assert_eq!(target_new_style.len(), 1);

    Ok(())
}

fn synthetic_package() -> TestResult<Vec<u8>> {
    let mut document = object(
        DOCUMENT_ID,
        1,
        tn::DocumentArchive {
            sheets: vec![reference(SHEET_ID)],
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    document.archive_info.message_infos[0].object_references = vec![SHEET_ID];

    let mut sheet = object(
        SHEET_ID,
        2,
        tn::SheetArchive {
            name: "Appearance Sheet".to_owned(),
            drawable_infos: vec![reference(FIRST_INFO_ID), reference(SECOND_INFO_ID)],
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    sheet.archive_info.message_infos[0].object_references = vec![FIRST_INFO_ID, SECOND_INFO_ID];

    let mut first_model = object(
        FIRST_MODEL_ID,
        TABLE_MODEL_TYPE,
        table_model(FIRST_MODEL_ID).encode_to_vec(),
    )?;
    first_model.archive_info.message_infos[0].object_references =
        vec![SIDECARS_ID, PARENT_STYLE_ID];
    add_model_archive_metadata(&mut first_model, PARENT_STYLE_ID);
    let mut second_model = object(
        SECOND_MODEL_ID,
        TABLE_MODEL_TYPE,
        table_model(SECOND_MODEL_ID).encode_to_vec(),
    )?;
    second_model.archive_info.message_infos[0].object_references =
        vec![SIDECARS_ID, PARENT_STYLE_ID];
    add_model_archive_metadata(&mut second_model, PARENT_STYLE_ID);

    let document_archive = vec![
        document,
        sheet,
        table_info(FIRST_MODEL_ID)?,
        second_table_info()?,
        first_model,
        second_model,
        sidecars()?,
    ];
    let mut stylesheet_object = stylesheet()?;
    stylesheet_object.archive_info.message_infos[0].data_references = vec![6_000];
    let mut stylesheet_reference_field = FieldInfo::new(vec![1]);
    stylesheet_reference_field.object_references = vec![PARENT_STYLE_ID, UNUSED_STYLE_ID];
    stylesheet_reference_field.data_references = vec![6_002];
    let mut stylesheet_tail = FieldInfo::new(vec![77, 1]);
    stylesheet_tail.data_references = vec![6_001];
    stylesheet_object.archive_info.message_infos[0]
        .field_infos
        .extend([stylesheet_reference_field, stylesheet_tail]);
    let stylesheet_archive = compressed(vec![
        table_style(PARENT_STYLE_ID, None)?,
        table_style(UNUSED_STYLE_ID, Some(PARENT_STYLE_ID))?,
        stylesheet_object,
    ])?;
    let document_archive = compressed(document_archive)?;
    let view_archive = compressed(vec![object(
        VIEW_STATE_ID,
        210,
        b"unselected view state is byte exact".to_vec(),
    )?])?;
    let metadata = metadata_entry()?;

    Ok(litchi_iwa_archive::package::to_bytes(
        [
            (DOCUMENT_MEMBER, document_archive.as_slice()),
            (STYLESHEET_MEMBER, stylesheet_archive.as_slice()),
            (VIEW_STATE_MEMBER, view_archive.as_slice()),
            (METADATA_MEMBER, metadata.as_slice()),
            ("preview.jpg", b"appearance preview".as_slice()),
            ("preview-micro.jpg", b"appearance micro".as_slice()),
            ("preview-web.jpg", b"appearance web".as_slice()),
            ("Data/sentinel.bin", b"unrelated data".as_slice()),
        ],
        Limits::default(),
    )?)
}

fn custom_appearance() -> Appearance {
    Appearance {
        row_banding: Banding::Enabled,
        row_sizing: RowSizing::FitCellContents,
        gridlines: Gridlines {
            body_horizontal: GridlineVisibility::Hidden,
            header_columns_horizontal: GridlineVisibility::Visible,
            body_vertical: GridlineVisibility::Hidden,
            header_rows_vertical: GridlineVisibility::Visible,
            footer_rows_vertical: GridlineVisibility::Hidden,
        },
    }
}

fn changed_members(source: &[u8], target: &[u8]) -> TestResult<Vec<String>> {
    let source = Catalog::from_bytes(source)?;
    let target = Catalog::from_bytes(target)?;
    let mut changed = Vec::new();
    for before in source.iter() {
        let after = target
            .iter()
            .find(|candidate| candidate.name() == before.name());
        if PREVIEWS.contains(&before.name()) {
            assert!(after.is_none(), "preview {} was retained", before.name());
            continue;
        }
        let after = after.ok_or_else(|| {
            io::Error::other(format!("member {} was unexpectedly deleted", before.name()))
        })?;
        if before.data() == after.data() {
            assert_eq!(
                before.raw_record().local_record(),
                after.raw_record().local_record(),
                "unchanged member {} lost its exact local record",
                before.name()
            );
        } else {
            changed.push(before.name().to_owned());
        }
    }
    Ok(changed)
}

fn rewrite_document_table_info(source: &[u8], locked: bool) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("Document member is missing"))?;
    let stream = SnappyStream::decompress(entry.data())?;
    let mut archive = Archive::parse(stream.as_bytes())?;
    let info = archive
        .object_mut(FIRST_INFO_ID)
        .ok_or_else(|| io::Error::other("first table info is missing"))?;
    let mut decoded = tst::TableInfoArchive::decode(info.messages[0].data.as_slice())?;
    decoded.super_.locked = Some(locked);
    info.replace_message_preserving_header(
        0,
        RawMessage {
            type_: TABLE_INFO_TYPE,
            data: decoded.encode_to_vec(),
        },
    )?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            DOCUMENT_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?)
}

fn remove_metadata(source: &[u8]) -> TestResult<Vec<u8>> {
    Ok(
        Catalog::from_bytes(source)?.reassemble_with_deletions_to_bytes(
            &[],
            &[METADATA_MEMBER],
            Limits::default(),
        )?,
    )
}

#[test]
fn shared_parent_read_noop_cow_apply_inverse_and_locality() -> TestResult {
    let source = synthetic_package()?;
    let package = Package::from_bytes(&source)?;
    let before = package.table_appearance(0usize, 0usize)?;
    assert_eq!(
        before.row_banding,
        Banding::Disabled,
        "the direct table_style must win over the nonzero table_style_preset hint"
    );
    assert_eq!(before, package.table_appearance(0usize, 1usize)?);

    let noop = package
        .edit_table_appearance(0usize, 0usize)?
        .set(before)
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package().exact_bytes(), source);

    let after = custom_appearance();
    let commit = package
        .edit_table_appearance(0usize, 0usize)?
        .set(after)
        .commit()?;
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.patch().path(), Path::Table { sheet: 0, table: 0 });
    assert_eq!(commit.patch().before(), before);
    assert_eq!(commit.patch().after(), after);
    assert_eq!(commit.package().table_appearance(0usize, 0usize)?, after);
    assert_eq!(commit.package().table_appearance(0usize, 1usize)?, before);
    assert_eq!(commit.diagnostics().deleted_previews(), 3);

    let target = commit.package().exact_bytes();
    metadata_transition_assertions(&source, &target)?;
    for (member, identifier) in [
        (DOCUMENT_MEMBER, DOCUMENT_ID),
        (DOCUMENT_MEMBER, SHEET_ID),
        (DOCUMENT_MEMBER, FIRST_INFO_ID),
        (DOCUMENT_MEMBER, SECOND_INFO_ID),
        (DOCUMENT_MEMBER, SECOND_MODEL_ID),
        (DOCUMENT_MEMBER, SIDECARS_ID),
        (STYLESHEET_MEMBER, PARENT_STYLE_ID),
        (STYLESHEET_MEMBER, UNUSED_STYLE_ID),
    ] {
        assert_eq!(
            object_messages(&source, member, identifier)?,
            object_messages(&target, member, identifier)?,
            "unselected object {identifier} in {member} must remain byte exact"
        );
    }
    let source_model_payload = object_messages(&source, DOCUMENT_MEMBER, FIRST_MODEL_ID)?
        .first()
        .ok_or_else(|| io::Error::other("source model payload is missing"))?
        .1
        .clone();
    let target_model_payload = object_messages(&target, DOCUMENT_MEMBER, FIRST_MODEL_ID)?
        .first()
        .ok_or_else(|| io::Error::other("target model payload is missing"))?
        .1
        .clone();
    assert_eq!(
        wire_fields(&source_model_payload, 48)?,
        wire_fields(&target_model_payload, 48)?,
        "direct field 3 COW must preserve the conflicting field 48 preset raw"
    );
    let changed = changed_members(&source, &target)?;
    assert_eq!(
        changed,
        vec![
            DOCUMENT_MEMBER.to_owned(),
            STYLESHEET_MEMBER.to_owned(),
            METADATA_MEMBER.to_owned(),
        ]
    );

    let applied = package.apply_table_appearance(commit.patch())?;
    assert_eq!(applied.package().exact_bytes(), target);
    assert!(matches!(
        Package::from_bytes(&target)?.apply_table_appearance(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(
        Package::from_bytes(&target)?
            .apply_table_appearance(&commit.patch().inverse())?
            .package()
            .exact_bytes(),
        source
    );
    Ok(())
}

#[test]
fn missing_metadata_and_locked_table_fail_atomically() -> TestResult {
    let source = synthetic_package()?;
    let without_metadata = remove_metadata(&source)?;
    let package = Package::from_bytes(&without_metadata)?;
    let before = package.exact_bytes();
    let error = package
        .edit_table_appearance(0usize, 0usize)?
        .set(custom_appearance())
        .commit()
        .expect_err("changed appearance without metadata must fail closed");
    assert!(matches!(error, Error::InvalidSource { .. }));
    assert_eq!(package.exact_bytes(), before);

    let locked = Package::from_bytes(&rewrite_document_table_info(&source, true)?)?;
    let before = locked.exact_bytes();
    let error = locked
        .edit_table_appearance(0usize, 0usize)?
        .set(custom_appearance())
        .commit()
        .expect_err("locked table must refuse changed appearance");
    assert!(matches!(error, Error::TableLocked { .. }));
    assert_eq!(locked.exact_bytes(), before);
    Ok(())
}

#[test]
fn duplicate_edges_missing_stylesheet_and_inheritance_cycle_fail_closed() -> TestResult {
    let source = synthetic_package()?;

    let duplicate_model_style =
        rewrite_object_message(&source, DOCUMENT_MEMBER, FIRST_MODEL_ID, |message| {
            append_length_delimited_field(
                &mut message.data,
                3,
                &reference(PARENT_STYLE_ID).encode_to_vec(),
            )?;
            Ok(())
        })?;
    assert_appearance_read_rejects(&duplicate_model_style)?;

    let wrong_model_style_wire =
        rewrite_object_message(&source, DOCUMENT_MEMBER, FIRST_MODEL_ID, |message| {
            append_varint_field(&mut message.data, 3, PARENT_STYLE_ID)?;
            Ok(())
        })?;
    assert_appearance_read_rejects(&wrong_model_style_wire)?;

    let preset_only_model =
        rewrite_object_message(&source, DOCUMENT_MEMBER, FIRST_MODEL_ID, |message| {
            message.data = without_wire_field(&message.data, 3)?;
            Ok(())
        })?;
    assert_appearance_read_rejects(&preset_only_model)?;

    let duplicate_stylesheet_edge =
        rewrite_object_message(&source, STYLESHEET_MEMBER, PARENT_STYLE_ID, |message| {
            let mut duplicate = Vec::new();
            append_length_delimited_field(
                &mut duplicate,
                5,
                &reference(STYLESHEET_ID).encode_to_vec(),
            )?;
            message.data = append_nested_wire_field(&message.data, 1, &duplicate)?;
            Ok(())
        })?;
    assert_appearance_read_rejects(&duplicate_stylesheet_edge)?;

    let wrong_parent_wire =
        rewrite_object_message(&source, STYLESHEET_MEMBER, PARENT_STYLE_ID, |message| {
            let mut replacement = Vec::new();
            append_varint_field(&mut replacement, 3, PARENT_STYLE_ID)?;
            message.data = append_nested_wire_field(&message.data, 1, &replacement)?;
            Ok(())
        })?;
    assert_appearance_read_rejects(&wrong_parent_wire)?;

    let wrong_stylesheet_wire =
        rewrite_object_message(&source, STYLESHEET_MEMBER, PARENT_STYLE_ID, |message| {
            let mut replacement = Vec::new();
            append_varint_field(&mut replacement, 5, STYLESHEET_ID)?;
            message.data = replace_nested_wire_field(&message.data, 1, 5, &replacement)?;
            Ok(())
        })?;
    assert_appearance_read_rejects(&wrong_stylesheet_wire)?;

    let wrong_stylesheet_type = rewrite_object_message_type(
        &source,
        STYLESHEET_MEMBER,
        STYLESHEET_ID,
        STYLESHEET_TYPE + 1,
    )?;
    assert_appearance_read_rejects(&wrong_stylesheet_type)?;

    let missing_stylesheet =
        rewrite_object_message(&source, STYLESHEET_MEMBER, PARENT_STYLE_ID, |message| {
            let sanitized = without_wire_fields(&message.data, &[90, 91])?;
            let mut style = tst::TableStyleArchive::decode(sanitized.as_slice())?;
            style.super_.stylesheet = None;
            message.data = style.encode_to_vec();
            Ok(())
        })?;
    assert_appearance_read_rejects(&missing_stylesheet)?;

    let inheritance_cycle =
        rewrite_object_message(&source, STYLESHEET_MEMBER, PARENT_STYLE_ID, |message| {
            let sanitized = without_wire_fields(&message.data, &[90, 91])?;
            let mut style = tst::TableStyleArchive::decode(sanitized.as_slice())?;
            style.super_.parent = Some(reference(PARENT_STYLE_ID));
            message.data = style.encode_to_vec();
            Ok(())
        })?;
    assert_appearance_read_rejects(&inheritance_cycle)?;

    let foreign_stylesheet_payload = stylesheet()?
        .messages
        .first()
        .ok_or_else(|| io::Error::other("stylesheet payload is missing"))?
        .data
        .clone();
    let style_with_foreign_stylesheet =
        rewrite_object_message(&source, STYLESHEET_MEMBER, PARENT_STYLE_ID, |message| {
            let sanitized = without_wire_fields(&message.data, &[90])?;
            let mut style = tst::TableStyleArchive::decode(sanitized.as_slice())?;
            style.super_.stylesheet = Some(reference(SIDECARS_ID));
            message.data = style.encode_to_vec();
            Ok(())
        })?;
    let style_with_foreign_stylesheet = rewrite_object_message_type(
        &style_with_foreign_stylesheet,
        DOCUMENT_MEMBER,
        SIDECARS_ID,
        STYLESHEET_TYPE,
    )?;
    let style_with_foreign_stylesheet = rewrite_object_message(
        &style_with_foreign_stylesheet,
        DOCUMENT_MEMBER,
        SIDECARS_ID,
        |message| {
            message.data = foreign_stylesheet_payload.clone();
            Ok(())
        },
    )?;
    assert_changed_edit_rejects_atomically(
        &style_with_foreign_stylesheet,
        "style and stylesheet in distinct components",
    )?;

    let duplicate_stylesheet_style =
        rewrite_object_message(&source, STYLESHEET_MEMBER, STYLESHEET_ID, |message| {
            let mut stylesheet = tss::StylesheetArchive::decode(message.data.as_slice())?;
            stylesheet.styles.push(reference(PARENT_STYLE_ID));
            message.data = stylesheet.encode_to_vec();
            Ok(())
        })?;
    assert_changed_edit_rejects_atomically(
        &duplicate_stylesheet_style,
        "duplicate stylesheet style",
    )?;

    let duplicate_parent_registry =
        rewrite_object_message(&source, STYLESHEET_MEMBER, STYLESHEET_ID, |message| {
            let mut stylesheet = tss::StylesheetArchive::decode(message.data.as_slice())?;
            let entry = stylesheet
                .parent_to_children_style_map
                .first()
                .cloned()
                .ok_or_else(|| io::Error::other("stylesheet parent registry is missing"))?;
            stylesheet.parent_to_children_style_map.push(entry);
            message.data = stylesheet.encode_to_vec();
            Ok(())
        })?;
    assert_changed_edit_rejects_atomically(
        &duplicate_parent_registry,
        "duplicate parent registry",
    )?;

    let aggregate_only_reference =
        rewrite_object_archive_info(&source, DOCUMENT_MEMBER, FIRST_MODEL_ID, |info| {
            info.object_references.push(1_001);
            Ok(())
        })?;
    assert_changed_edit_rejects_atomically(&aggregate_only_reference, "aggregate-only model ref")?;

    let field_only_reference =
        rewrite_object_archive_info(&source, DOCUMENT_MEMBER, FIRST_MODEL_ID, |info| {
            info.field_infos[0].object_references = vec![1_001];
            Ok(())
        })?;
    assert_changed_edit_rejects_atomically(&field_only_reference, "field-only model ref")?;

    let duplicate_canonical_path =
        rewrite_object_archive_info(&source, DOCUMENT_MEMBER, FIRST_MODEL_ID, |info| {
            info.field_infos.push(info.field_infos[0].clone());
            Ok(())
        })?;
    assert_changed_edit_rejects_atomically(&duplicate_canonical_path, "duplicate model path")?;

    let conflicting_reference_counts =
        rewrite_object_archive_info(&source, DOCUMENT_MEMBER, FIRST_MODEL_ID, |info| {
            info.object_references.push(PARENT_STYLE_ID);
            Ok(())
        })?;
    assert_changed_edit_rejects_atomically(
        &conflicting_reference_counts,
        "conflicting model refs",
    )?;

    let wrong_field_type =
        rewrite_object_archive_info(&source, DOCUMENT_MEMBER, FIRST_MODEL_ID, |info| {
            info.field_infos[0].r#type = Some(FieldType::Value);
            Ok(())
        })?;
    assert_changed_edit_rejects_atomically(&wrong_field_type, "wrong model field type")?;

    let missing_current_field_info =
        rewrite_object_archive_info(&source, DOCUMENT_MEMBER, FIRST_MODEL_ID, |info| {
            info.field_infos
                .retain(|field| field.path.as_slice() != [3]);
            Ok(())
        })?;
    let missing_current_field_info = rewrite_object_archive_info(
        &missing_current_field_info,
        STYLESHEET_MEMBER,
        STYLESHEET_ID,
        |info| {
            info.field_infos
                .retain(|field| field.path.as_slice() != [1]);
            info.object_references
                .retain(|identifier| *identifier != PARENT_STYLE_ID);
            Ok(())
        },
    )?;
    let missing_current_package = Package::from_bytes(&missing_current_field_info)?;
    let missing_current_commit = missing_current_package
        .edit_table_appearance(0usize, 0usize)?
        .set(custom_appearance())
        .commit()?;
    let restored = missing_current_commit
        .package()
        .apply_table_appearance(&missing_current_commit.patch().inverse())?;
    assert_eq!(restored.package().exact_bytes(), missing_current_field_info);

    let duplicate_stylesheet_field_info =
        rewrite_object_archive_info(&source, STYLESHEET_MEMBER, STYLESHEET_ID, |info| {
            let field = info
                .field_infos
                .iter()
                .find(|field| field.path.as_slice() == [1])
                .cloned()
                .ok_or_else(|| io::Error::other("stylesheet FieldInfo path [1] is missing"))?;
            info.field_infos.push(field);
            Ok(())
        })?;
    assert_changed_edit_rejects_atomically(
        &duplicate_stylesheet_field_info,
        "duplicate stylesheet FieldInfo path [1]",
    )?;

    let conflicting_stylesheet_field_info =
        rewrite_object_archive_info(&source, STYLESHEET_MEMBER, STYLESHEET_ID, |info| {
            let field = info
                .field_infos
                .iter_mut()
                .find(|field| field.path.as_slice() == [1])
                .ok_or_else(|| io::Error::other("stylesheet FieldInfo path [1] is missing"))?;
            field.object_references = vec![PARENT_STYLE_ID];
            Ok(())
        })?;
    assert_changed_edit_rejects_atomically(
        &conflicting_stylesheet_field_info,
        "conflicting stylesheet FieldInfo path [1]",
    )?;

    let duplicate_parent_edge = rewrite_metadata(&source, |metadata| {
        let document = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 100)
            .expect("document component");
        document
            .external_references
            .push(tsp::ComponentExternalReference {
                component_identifier: 200,
                object_identifier: Some(PARENT_STYLE_ID),
                is_weak: None,
            });
    })?;
    assert_changed_edit_rejects_atomically(&duplicate_parent_edge, "duplicate old-parent edge")?;

    let conflicting_parent_weakness = rewrite_metadata(&source, |metadata| {
        let document = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 100)
            .expect("document component");
        document
            .external_references
            .push(tsp::ComponentExternalReference {
                component_identifier: 200,
                object_identifier: Some(PARENT_STYLE_ID),
                is_weak: Some(true),
            });
    })?;
    assert_changed_edit_rejects_atomically(
        &conflicting_parent_weakness,
        "conflicting parent weakness",
    )?;

    let versioned_parent_edge = rewrite_metadata(&source, |metadata| {
        let document = metadata
            .versioned_components
            .iter_mut()
            .find(|component| component.identifier == 100)
            .expect("versioned document component");
        document
            .external_references
            .push(tsp::ComponentExternalReference {
                component_identifier: 200,
                object_identifier: Some(PARENT_STYLE_ID),
                is_weak: None,
            });
    })?;
    assert_changed_edit_rejects_atomically(&versioned_parent_edge, "versioned old-parent edge")?;

    Ok(())
}

#[test]
fn old_style_metadata_ownership_and_aliases_fail_closed() -> TestResult {
    let source = synthetic_package()?;

    let component_root_stylesheet = rewrite_metadata(&source, |metadata| {
        let stylesheet = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 200)
            .expect("stylesheet component");
        stylesheet.identifier = STYLESHEET_ID;
        stylesheet
            .object_uuid_map_entries
            .retain(|entry| entry.identifier != STYLESHEET_ID);
        let document = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 100)
            .expect("document component");
        for reference in &mut document.external_references {
            if reference.component_identifier == 200 {
                reference.component_identifier = STYLESHEET_ID;
            }
        }
    })?;
    let component_root_package = Package::from_bytes(&component_root_stylesheet)?;
    let component_root_commit = component_root_package
        .edit_table_appearance(0, 0)?
        .set(custom_appearance())
        .commit()?;
    assert_eq!(
        component_root_commit.package().table_appearance(0, 0)?,
        custom_appearance()
    );

    let missing_old_style_uuid = rewrite_metadata(&source, |metadata| {
        metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 200)
            .expect("stylesheet component")
            .object_uuid_map_entries
            .retain(|entry| entry.identifier != PARENT_STYLE_ID);
    })?;
    assert_changed_edit_rejects_atomically(&missing_old_style_uuid, "missing old-style UUID")?;

    let versioned_old_style_uuid = rewrite_metadata(&source, |metadata| {
        metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 200)
            .expect("stylesheet component")
            .object_uuid_map_entries
            .retain(|entry| entry.identifier != PARENT_STYLE_ID);
        metadata
            .versioned_components
            .iter_mut()
            .find(|component| component.identifier == 100)
            .expect("versioned document component")
            .object_uuid_map_entries
            .push(uuid_entry(PARENT_STYLE_ID));
    })?;
    assert_changed_edit_rejects_atomically(
        &versioned_old_style_uuid,
        "versioned-only old-style UUID",
    )?;

    let cross_component_old_style_uuid = rewrite_metadata(&source, |metadata| {
        metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 200)
            .expect("stylesheet component")
            .object_uuid_map_entries
            .retain(|entry| entry.identifier != PARENT_STYLE_ID);
        metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 100)
            .expect("document component")
            .object_uuid_map_entries
            .push(uuid_entry(PARENT_STYLE_ID));
    })?;
    assert_changed_edit_rejects_atomically(
        &cross_component_old_style_uuid,
        "cross-component old-style UUID",
    )?;

    let duplicate_old_style_uuid = rewrite_metadata(&source, |metadata| {
        metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 100)
            .expect("document component")
            .object_uuid_map_entries
            .push(uuid_entry(PARENT_STYLE_ID));
    })?;
    assert_changed_edit_rejects_atomically(&duplicate_old_style_uuid, "duplicate old-style UUID")?;

    let ambiguous_old_style = rewrite_metadata(&source, |metadata| {
        metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 100)
            .expect("document component")
            .ambiguous_object_identifiers
            .push(PARENT_STYLE_ID);
    })?;
    assert_changed_edit_rejects_atomically(&ambiguous_old_style, "ambiguous old-style ID")?;

    let data_owned_old_style = rewrite_metadata(&source, |metadata| {
        metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 100)
            .expect("document component")
            .data_references
            .push(tsp::ComponentDataReference {
                data_identifier: PARENT_STYLE_ID,
                object_reference_list: vec![tsp::component_data_reference::ObjectReference {
                    object_identifier: PARENT_STYLE_ID,
                    count: 1,
                }],
            });
    })?;
    assert_changed_edit_rejects_atomically(&data_owned_old_style, "data-owned old-style ID")?;

    let root_map_old_style = rewrite_metadata(&source, |metadata| {
        metadata.data_metadata_map = Some(reference(PARENT_STYLE_ID));
    })?;
    assert_changed_edit_rejects_atomically(&root_map_old_style, "root-map old-style ID")?;
    Ok(())
}

#[test]
fn tight_ingress_limit_rejects_before_package_publication() -> TestResult {
    let source = synthetic_package()?;
    let archive_limits = PackageLimits::new(
        u64::try_from(source.len().saturating_sub(1))?,
        PackageLimits::MAX_ENTRIES,
        PackageLimits::MAX_ENTRY_BYTES,
        PackageLimits::MAX_TOTAL_BYTES,
        PackageLimits::MAX_IWA_STREAM_BYTES,
    )?;
    assert!(
        Package::from_bytes_with_options(
            &source,
            PackageReadOptions::new(archive_limits, PackageSemanticLimits::default()),
        )
        .is_err()
    );
    Ok(())
}

#[test]
fn tight_candidate_output_limit_rejects_before_publication() -> TestResult {
    let source = synthetic_package()?;
    let source_limit = u64::try_from(source.len())?;
    let archive_limits = PackageLimits::new(
        source_limit,
        PackageLimits::MAX_ENTRIES,
        PackageLimits::MAX_ENTRY_BYTES,
        source_limit,
        PackageLimits::MAX_IWA_STREAM_BYTES,
    )?;
    let package = Package::from_bytes_with_options(
        &source,
        PackageReadOptions::new(archive_limits, PackageSemanticLimits::default()),
    )?;
    let before = package.exact_bytes();
    assert!(
        package
            .edit_table_appearance(0usize, 0usize)?
            .set(custom_appearance())
            .commit()
            .is_err(),
        "candidate growth must respect the package output ceiling"
    );
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

#[test]
fn metadata_owned_identifiers_are_reserved_across_registry_axes() -> TestResult {
    let source = synthetic_package()?;

    let uuid_reserved = rewrite_metadata(&source, |metadata| {
        metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 100)
            .expect("document component")
            .object_uuid_map_entries
            .push(uuid_entry(1_001));
    })?;
    assert_reserved_identifier_is_not_reused(&uuid_reserved)?;

    let external_reserved = rewrite_metadata(&source, |metadata| {
        metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 100)
            .expect("document component")
            .external_references
            .push(tsp::ComponentExternalReference {
                component_identifier: 200,
                object_identifier: Some(1_001),
                is_weak: None,
            });
    })?;
    assert_reserved_identifier_is_not_reused(&external_reserved)?;

    let data_reserved = rewrite_metadata(&source, |metadata| {
        metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 100)
            .expect("document component")
            .data_references
            .push(tsp::ComponentDataReference {
                data_identifier: 1_001,
                object_reference_list: vec![tsp::component_data_reference::ObjectReference {
                    object_identifier: 1_001,
                    count: 1,
                }],
            });
    })?;
    assert_reserved_identifier_is_not_reused(&data_reserved)?;

    let root_map_reserved = rewrite_metadata(&source, |metadata| {
        metadata.data_metadata_map = Some(reference(1_001));
    })?;
    assert_reserved_identifier_is_not_reused(&root_map_reserved)?;

    let ambiguous_reserved = rewrite_metadata(&source, |metadata| {
        metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 100)
            .expect("document component")
            .ambiguous_object_identifiers
            .push(1_001);
    })?;
    assert_reserved_identifier_is_not_reused(&ambiguous_reserved)?;
    Ok(())
}

#[test]
fn metadata_preferred_and_effective_locators_are_resolved_strictly() -> TestResult {
    let source = synthetic_package()?;
    for (identifier, physical_locator) in [(100, "Document"), (200, "DocumentStylesheet")] {
        let (component, _) = find_metadata_component(&source, 3, physical_locator, identifier)?;
        assert_eq!(
            component.preferred_locator, physical_locator,
            "component {identifier} must use the exact normalized physical locator"
        );
    }
    let explicit = rewrite_metadata(&source, |metadata| {
        metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 100)
            .expect("document component")
            .locator = Some("Document-Explicit".to_owned());
        metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 200)
            .expect("stylesheet component")
            .locator = Some("Stylesheet-Explicit".to_owned());
    })?;
    let package = Package::from_bytes(&explicit)?;
    let commit = package
        .edit_table_appearance(0usize, 0usize)?
        .set(custom_appearance())
        .commit()?;
    metadata_transition_assertions(&explicit, &commit.package().exact_bytes())?;
    assert_eq!(
        Package::from_bytes(&commit.package().exact_bytes())?
            .apply_table_appearance(&commit.patch().inverse())?
            .package()
            .exact_bytes(),
        explicit
    );

    let preferred_only = rewrite_metadata(&source, |metadata| {
        for component in &mut metadata.components {
            if component.identifier == 100 || component.identifier == 200 {
                component.locator = None;
            }
        }
    })?;
    let package = Package::from_bytes(&preferred_only)?;
    let commit = package
        .edit_table_appearance(0usize, 0usize)?
        .set(custom_appearance())
        .commit()?;
    assert_eq!(
        Package::from_bytes(&commit.package().exact_bytes())?
            .apply_table_appearance(&commit.patch().inverse())?
            .package()
            .exact_bytes(),
        preferred_only
    );

    let duplicate_effective = rewrite_metadata(&source, |metadata| {
        for component in &mut metadata.components {
            if component.identifier == 100 || component.identifier == 200 {
                component.locator = Some("Shared-Effective".to_owned());
            }
        }
    })?;
    let package = Package::from_bytes(&duplicate_effective)?;
    let before = package.exact_bytes();
    assert!(
        package
            .edit_table_appearance(0usize, 0usize)?
            .set(custom_appearance())
            .commit()
            .is_err(),
        "duplicate effective metadata locators must fail closed"
    );
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

#[test]
fn public_transaction_values_are_redacted_and_thread_safe() -> TestResult {
    fn assert_send_sync_debug<T: Send + Sync + std::fmt::Debug>() {}

    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<Appearance>();
    assert_send_sync_debug::<litchi_numbers::table::appearance::transaction::Edit<'static>>();
    assert_send_sync_debug::<litchi_numbers::table::appearance::transaction::Commit>();
    assert_send_sync_debug::<litchi_numbers::table::appearance::transaction::Patch>();
    assert_send_sync_debug::<litchi_numbers::table::appearance::transaction::Diagnostics>();
    assert_send_sync_debug::<Error>();
    assert_send_sync_debug::<Path>();
    let package = Package::from_bytes(&synthetic_package()?)?;
    let edit = package.edit_table_appearance(0usize, 0usize)?;
    let rendered = format!("{edit:?}");
    assert!(!rendered.contains("Appearance Sheet"));
    assert!(!rendered.contains(DOCUMENT_MEMBER));
    Ok(())
}
