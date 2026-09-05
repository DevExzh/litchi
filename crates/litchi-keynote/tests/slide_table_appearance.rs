//! Exact-source integration coverage for Keynote slide-table appearance.
//!
//! The fixture is intentionally synthetic but follows the native ownership
//! shape: Document -> Show -> SlideNode -> Slide -> TableInfo in one member,
//! TableModel objects in CalculationEngine, and TableStyle/preset/network plus
//! the stylesheet registry in a co-located stylesheet member.  Metadata,
//! previews, and an unrelated sentinel make the package-level locality and
//! source-atomicity assertions observable.

use std::error::Error as StdError;
use std::io;

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::wire::{append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{
    Archive, ArchiveObject, FieldInfo, FieldPath, FieldType, RawMessage, SnappyStream,
};
use litchi_iwa_protos::{kn, tsa, tsd, tsk, tsp, tss, tst, tswp};
use litchi_keynote::slide::table::appearance::transaction::{
    Commit, Diagnostics, Edit, Error, LimitKind, Patch, Path,
};
use litchi_keynote::slide::table::appearance::{
    Appearance, Banding, GridlineVisibility, Gridlines, RowSizing,
};
use litchi_keynote::{
    MAX_OBJECTS, MAX_SLIDES, MAX_TEXT_BYTES, MAX_TEXT_FRAGMENTS, MAX_TEXT_STORAGES, Package,
    ReadOptions, SemanticLimits, SlideSelector, TableSelector,
};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const MODEL_MEMBER: &str = "Index/CalculationEngine.iwa";
const STYLESHEET_MEMBER: &str = "Index/DocumentStylesheet.iwa";
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const SENTINEL_MEMBER: &str = "Data/appearance-sentinel.bin";
const PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

const SHOW_ID: u64 = 2;
const SLIDE_NODE_ID: u64 = 3;
const SLIDE_ID: u64 = 4;
const TABLE_INFO_IDS: [u64; 4] = [100, 101, 102, 103];
const TABLE_MODEL_IDS: [u64; 4] = [110, 111, 112, 113];
const NON_TABLE_DRAWABLE_ID: u64 = 130;

const SHARED_STYLE_ID: u64 = 200;
const PARENT_STYLE_ID: u64 = 201;
const PRESET_STYLE_ID: u64 = 202;
const PRESET_ID: u64 = 300;
const NETWORK_ID: u64 = 301;
const STYLESHEET_ID: u64 = 302;

const METADATA_OBJECT_ID: u64 = 900;
const FOREIGN_INBOUND_ID: u64 = 902;

const DOCUMENT_MESSAGE_TYPE: u32 = 1;
const SHOW_MESSAGE_TYPE: u32 = 2;
const SLIDE_NODE_MESSAGE_TYPE: u32 = 4;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;
const TABLE_STYLE_MESSAGE_TYPE: u32 = 6_003;
const TABLE_STYLE_PRESET_MESSAGE_TYPE: u32 = 6_008;
const TABLE_STYLE_NETWORK_MESSAGE_TYPE: u32 = 6_247;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const METADATA_MESSAGE_TYPE: u32 = 11_006;

const TABLE_MODEL_STYLE_FIELD: u32 = 3;
const TABLE_MODEL_PRESET_FIELD: u32 = 48;
const TABLE_STYLE_SUPER_FIELD: u32 = 1;
const TABLE_STYLE_OVERRIDE_COUNT_FIELD: u32 = 10;
const TABLE_STYLE_PARENT_FIELD: u32 = 3;
const TABLE_STYLE_STYLESHEET_FIELD: u32 = 5;
const UNKNOWN_MODEL_FIELD: u32 = 90;
const UNKNOWN_MODEL_VALUE: u64 = 0xfeed_beef;

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn field_reference(path: impl Into<FieldPath>, identifiers: &[u64]) -> FieldInfo {
    let mut field = FieldInfo::new(path);
    field.r#type = Some(FieldType::ObjectReference);
    field.object_references.extend_from_slice(identifiers);
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

fn model_payload(
    name: &str,
    style_identifier: u64,
    preset_identifier: Option<u64>,
    with_unknown: bool,
) -> TestResult<Vec<u8>> {
    let mut payload = tst::TableModelArchive {
        table_id: format!("appearance-table-{name}"),
        table_name: name.to_owned(),
        number_of_rows: 5,
        number_of_columns: 4,
        table_style: reference(style_identifier),
        table_style_preset: preset_identifier.map(reference),
        default_row_height: 20.0,
        default_column_width: 64.0,
        ..Default::default()
    }
    .encode_to_vec();
    if with_unknown {
        append_varint_field(&mut payload, UNKNOWN_MODEL_FIELD, UNKNOWN_MODEL_VALUE)?;
    }
    Ok(payload)
}

fn style_properties(overrides: [Option<bool>; 7]) -> tst::TableStylePropertiesArchive {
    tst::TableStylePropertiesArchive {
        banded_rows: overrides[0],
        auto_resize: overrides[1],
        v_strokes_visible: overrides[2],
        h_strokes_visible: overrides[3],
        table_hc_divider_visible: overrides[4],
        table_hr_divider_visible: overrides[5],
        table_footer_divider_visible: overrides[6],
        ..Default::default()
    }
}

fn style_payload(
    identifier: u64,
    parent_identifier: Option<u64>,
    overrides: [Option<bool>; 7],
    with_unknown_group: bool,
) -> TestResult<Vec<u8>> {
    let mut payload = tst::TableStyleArchive {
        super_: tss::StyleArchive {
            name: Some(format!("Appearance style {identifier}")),
            style_identifier: Some(format!("appearance-style-{identifier}")),
            parent: parent_identifier.map(reference),
            is_variation: Some(parent_identifier.is_some()),
            stylesheet: Some(reference(STYLESHEET_ID)),
        },
        override_count: Some(u32::try_from(
            overrides.iter().filter(|value| value.is_some()).count(),
        )?),
        table_properties: Some(style_properties(overrides)),
    }
    .encode_to_vec();
    if with_unknown_group {
        // Canonical balanced unknown group: strict reads accept it and the
        // raw-preserving path has to retain it in the source style.
        payload.extend_from_slice(&[0x93, 0x05, 0x98, 0x05, 0x01, 0x94, 0x05]);
    }
    Ok(payload)
}

fn preset_payload() -> Vec<u8> {
    tst::TableStylePresetArchive {
        style_network: Some(reference(NETWORK_ID)),
        ..Default::default()
    }
    .encode_to_vec()
}

fn network_payload() -> Vec<u8> {
    tst::TableStyleNetworkArchive {
        body_text_style: reference(PRESET_STYLE_ID),
        header_row_text_style: reference(PRESET_STYLE_ID),
        header_column_text_style: reference(PRESET_STYLE_ID),
        footer_row_text_style: reference(PRESET_STYLE_ID),
        body_cell_style: reference(PRESET_STYLE_ID),
        header_row_style: reference(PRESET_STYLE_ID),
        header_column_style: reference(PRESET_STYLE_ID),
        footer_row_style: reference(PRESET_STYLE_ID),
        table_style: reference(PRESET_STYLE_ID),
        ..Default::default()
    }
    .encode_to_vec()
}

fn stylesheet_payload() -> Vec<u8> {
    tss::StylesheetArchive {
        styles: vec![
            reference(SHARED_STYLE_ID),
            reference(PARENT_STYLE_ID),
            reference(PRESET_STYLE_ID),
        ],
        identifier_to_style_map: vec![
            tss::stylesheet_archive::IdentifiedStyleEntry {
                identifier: "shared".to_owned(),
                style: reference(SHARED_STYLE_ID),
            },
            tss::stylesheet_archive::IdentifiedStyleEntry {
                identifier: "parent".to_owned(),
                style: reference(PARENT_STYLE_ID),
            },
            tss::stylesheet_archive::IdentifiedStyleEntry {
                identifier: "preset".to_owned(),
                style: reference(PRESET_STYLE_ID),
            },
        ],
        parent_to_children_style_map: vec![tss::stylesheet_archive::StyleChildrenEntry {
            parent: reference(PARENT_STYLE_ID),
            children: vec![reference(SHARED_STYLE_ID)],
        }],
        is_locked: Some(false),
        can_cull_styles: Some(true),
        ..Default::default()
    }
    .encode_to_vec()
}

fn uuid_entry(identifier: u64) -> tsp::ObjectUuidMapEntry {
    tsp::ObjectUuidMapEntry {
        identifier,
        uuid: tsp::Uuid {
            lower: identifier.saturating_add(10_000),
            upper: identifier.saturating_add(20_000),
        },
    }
}

fn metadata_component(
    identifier: u64,
    locator: &str,
    token: u64,
    object_ids: &[u64],
    external_references: &[tsp::ComponentExternalReference],
) -> Vec<u8> {
    tsp::ComponentInfo {
        identifier,
        preferred_locator: locator.to_owned(),
        locator: Some(locator.to_owned()),
        save_token: Some(token),
        object_uuid_map_entries: object_ids.iter().copied().map(uuid_entry).collect(),
        external_references: external_references.to_vec(),
        ..Default::default()
    }
    .encode_to_vec()
}

fn metadata_payload() -> TestResult<Vec<u8>> {
    let document_external_references = TABLE_MODEL_IDS
        .iter()
        .copied()
        .map(|object_identifier| tsp::ComponentExternalReference {
            component_identifier: 200,
            object_identifier: Some(object_identifier),
            is_weak: Some(false),
        })
        .collect::<Vec<_>>();
    let document = metadata_component(
        100,
        "Document",
        1,
        &[
            1,
            SHOW_ID,
            SLIDE_NODE_ID,
            SLIDE_ID,
            TABLE_INFO_IDS[0],
            TABLE_INFO_IDS[1],
            TABLE_INFO_IDS[2],
            TABLE_INFO_IDS[3],
            NON_TABLE_DRAWABLE_ID,
        ],
        &document_external_references,
    );
    let calculation_external_references = vec![
        tsp::ComponentExternalReference {
            component_identifier: 300,
            object_identifier: Some(SHARED_STYLE_ID),
            is_weak: Some(false),
        },
        tsp::ComponentExternalReference {
            component_identifier: 300,
            object_identifier: Some(PARENT_STYLE_ID),
            is_weak: Some(false),
        },
        tsp::ComponentExternalReference {
            component_identifier: 300,
            object_identifier: Some(PRESET_STYLE_ID),
            is_weak: Some(false),
        },
        tsp::ComponentExternalReference {
            component_identifier: 300,
            object_identifier: Some(PRESET_ID),
            is_weak: Some(false),
        },
        tsp::ComponentExternalReference {
            component_identifier: 300,
            object_identifier: Some(NETWORK_ID),
            is_weak: Some(false),
        },
        tsp::ComponentExternalReference {
            component_identifier: 300,
            object_identifier: Some(STYLESHEET_ID),
            is_weak: Some(false),
        },
    ];
    let calculation = metadata_component(
        200,
        "CalculationEngine",
        2,
        &TABLE_MODEL_IDS,
        &calculation_external_references,
    );
    let stylesheet = metadata_component(
        300,
        "DocumentStylesheet",
        3,
        &[
            SHARED_STYLE_ID,
            PARENT_STYLE_ID,
            PRESET_STYLE_ID,
            PRESET_ID,
            NETWORK_ID,
            STYLESHEET_ID,
        ],
        &[],
    );
    let mut payload = tsp::PackageMetadata {
        last_object_identifier: 1_000,
        save_token: Some(10),
        ..Default::default()
    }
    .encode_to_vec();
    append_length_delimited_field(&mut payload, 3, &document)?;
    append_length_delimited_field(&mut payload, 3, &calculation)?;
    append_length_delimited_field(&mut payload, 3, &stylesheet)?;
    Ok(payload)
}

fn shape_info_payload() -> Vec<u8> {
    tswp::ShapeInfoArchive {
        super_: tsd::ShapeArchive {
            super_: tsd::DrawableArchive::default(),
            ..Default::default()
        },
        ..Default::default()
    }
    .encode_to_vec()
}

fn synthetic_package(locked: bool, with_unknown_group: bool) -> TestResult<Vec<u8>> {
    let document = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..Default::default()
        },
        show: reference(SHOW_ID),
        ..Default::default()
    };
    let show = kn::ShowArchive {
        slide_tree: kn::SlideTreeArchive {
            slides: vec![reference(SLIDE_NODE_ID)],
            ..Default::default()
        },
        ..Default::default()
    };
    #[allow(deprecated, reason = "native schema retains legacy slide-node flags")]
    let node = kn::SlideNodeArchive {
        slide: Some(reference(SLIDE_ID)),
        ..Default::default()
    };
    let drawables = [
        TABLE_INFO_IDS[0],
        NON_TABLE_DRAWABLE_ID,
        TABLE_INFO_IDS[1],
        TABLE_INFO_IDS[2],
        TABLE_INFO_IDS[3],
    ];
    let slide = kn::SlideArchive {
        owned_drawables: drawables.iter().copied().map(reference).collect(),
        drawables_z_order: drawables.iter().copied().map(reference).collect(),
        name: Some("Appearance Tables".to_owned()),
        in_document: true,
        ..Default::default()
    };

    let mut document_objects = vec![
        object(
            1,
            DOCUMENT_MESSAGE_TYPE,
            document.encode_to_vec(),
            &[SHOW_ID],
        )?,
        object(
            SHOW_ID,
            SHOW_MESSAGE_TYPE,
            show.encode_to_vec(),
            &[SLIDE_NODE_ID],
        )?,
        object(
            SLIDE_NODE_ID,
            SLIDE_NODE_MESSAGE_TYPE,
            node.encode_to_vec(),
            &[SLIDE_ID],
        )?,
    ];
    let mut slide_object = object(
        SLIDE_ID,
        SLIDE_MESSAGE_TYPE,
        slide.encode_to_vec(),
        &drawables,
    )?;
    slide_object.archive_info.message_infos[0]
        .field_infos
        .extend([
            field_reference(vec![7], &drawables),
            field_reference(vec![42], &drawables),
        ]);
    document_objects.push(slide_object);
    document_objects.push(object(
        NON_TABLE_DRAWABLE_ID,
        SHAPE_INFO_MESSAGE_TYPE,
        shape_info_payload(),
        &[SLIDE_ID],
    )?);

    let model_style = [SHARED_STYLE_ID, SHARED_STYLE_ID, 0, 0];
    let model_preset = [None, None, Some(PRESET_ID), None];
    for index in 0..TABLE_INFO_IDS.len() {
        let info = tst::TableInfoArchive {
            super_: tsd::DrawableArchive {
                parent: Some(reference(SLIDE_ID)),
                locked: Some(locked),
                ..Default::default()
            },
            table_model: reference(TABLE_MODEL_IDS[index]),
            ..Default::default()
        };
        let mut info_object = object(
            TABLE_INFO_IDS[index],
            TABLE_INFO_MESSAGE_TYPE,
            info.encode_to_vec(),
            &[TABLE_MODEL_IDS[index]],
        )?;
        info_object.archive_info.message_infos[0]
            .field_infos
            .push(field_reference(vec![2], &[TABLE_MODEL_IDS[index]]));
        document_objects.push(info_object);
    }

    let mut model_objects = Vec::new();
    for index in 0..TABLE_MODEL_IDS.len() {
        let style = model_style[index];
        let preset = model_preset[index];
        let model_references = if style == 0 {
            preset.into_iter().collect::<Vec<_>>()
        } else {
            vec![style]
        };
        let mut model = object(
            TABLE_MODEL_IDS[index],
            TABLE_MODEL_MESSAGE_TYPE,
            model_payload(
                match index {
                    0 => "Direct",
                    1 => "Shared",
                    2 => "Preset",
                    _ => "Default",
                },
                style,
                preset,
                with_unknown_group && index == 0,
            )?,
            &model_references,
        )?;
        if style != 0 {
            model.archive_info.message_infos[0]
                .field_infos
                .push(field_reference(vec![3], &[style]));
        } else if let Some(preset) = preset {
            model.archive_info.message_infos[0]
                .field_infos
                .push(field_reference(vec![48], &[preset]));
        }
        model_objects.push(model);
    }

    let direct_overrides = [
        Some(true),
        None,
        Some(false),
        None,
        Some(true),
        None,
        Some(false),
    ];
    let parent_overrides = [
        Some(false),
        Some(true),
        None,
        Some(true),
        None,
        Some(false),
        Some(true),
    ];
    let preset_overrides = [
        Some(false),
        Some(false),
        Some(true),
        Some(false),
        Some(false),
        Some(true),
        Some(true),
    ];
    let mut stylesheet_objects = vec![
        {
            let mut style = object(
                SHARED_STYLE_ID,
                TABLE_STYLE_MESSAGE_TYPE,
                style_payload(
                    SHARED_STYLE_ID,
                    Some(PARENT_STYLE_ID),
                    direct_overrides,
                    with_unknown_group,
                )?,
                &[PARENT_STYLE_ID, STYLESHEET_ID],
            )?;
            style.archive_info.message_infos[0].field_infos.extend([
                field_reference(
                    vec![TABLE_STYLE_SUPER_FIELD, TABLE_STYLE_PARENT_FIELD],
                    &[PARENT_STYLE_ID],
                ),
                field_reference(
                    vec![TABLE_STYLE_SUPER_FIELD, TABLE_STYLE_STYLESHEET_FIELD],
                    &[STYLESHEET_ID],
                ),
            ]);
            style
        },
        {
            let mut style = object(
                PARENT_STYLE_ID,
                TABLE_STYLE_MESSAGE_TYPE,
                style_payload(PARENT_STYLE_ID, None, parent_overrides, false)?,
                &[STYLESHEET_ID],
            )?;
            style.archive_info.message_infos[0]
                .field_infos
                .push(field_reference(
                    vec![TABLE_STYLE_SUPER_FIELD, TABLE_STYLE_STYLESHEET_FIELD],
                    &[STYLESHEET_ID],
                ));
            style
        },
        {
            let mut style = object(
                PRESET_STYLE_ID,
                TABLE_STYLE_MESSAGE_TYPE,
                style_payload(PRESET_STYLE_ID, None, preset_overrides, false)?,
                &[STYLESHEET_ID],
            )?;
            style.archive_info.message_infos[0]
                .field_infos
                .push(field_reference(
                    vec![TABLE_STYLE_SUPER_FIELD, TABLE_STYLE_STYLESHEET_FIELD],
                    &[STYLESHEET_ID],
                ));
            style
        },
        object(
            PRESET_ID,
            TABLE_STYLE_PRESET_MESSAGE_TYPE,
            preset_payload(),
            &[NETWORK_ID],
        )?,
        object(
            NETWORK_ID,
            TABLE_STYLE_NETWORK_MESSAGE_TYPE,
            network_payload(),
            &[PRESET_STYLE_ID; 9],
        )?,
    ];
    stylesheet_objects.push({
        let mut registry = object(
            STYLESHEET_ID,
            STYLESHEET_MESSAGE_TYPE,
            stylesheet_payload(),
            &[SHARED_STYLE_ID, PARENT_STYLE_ID, PRESET_STYLE_ID],
        )?;
        registry.archive_info.message_infos[0]
            .field_infos
            .push(field_reference(
                vec![1],
                &[SHARED_STYLE_ID, PARENT_STYLE_ID, PRESET_STYLE_ID],
            ));
        registry
    });

    let document_component = SnappyStream::compress(
        &Archive {
            objects: document_objects,
        }
        .to_bytes()?,
    )?;
    let model_component = SnappyStream::compress(
        &Archive {
            objects: model_objects,
        }
        .to_bytes()?,
    )?;
    let stylesheet_component = SnappyStream::compress(
        &Archive {
            objects: stylesheet_objects,
        }
        .to_bytes()?,
    )?;
    let metadata_component = SnappyStream::compress(
        &Archive {
            objects: vec![object(
                METADATA_OBJECT_ID,
                METADATA_MESSAGE_TYPE,
                metadata_payload()?,
                &[],
            )?],
        }
        .to_bytes()?,
    )?;

    let mut members: Vec<(&str, &[u8])> = vec![
        (SENTINEL_MEMBER, b"unrelated appearance sentinel"),
        (DOCUMENT_MEMBER, document_component.as_slice()),
        (MODEL_MEMBER, model_component.as_slice()),
        (STYLESHEET_MEMBER, stylesheet_component.as_slice()),
        (METADATA_MEMBER, metadata_component.as_slice()),
    ];
    members.extend(
        PREVIEWS
            .into_iter()
            .map(|name| (name, b"appearance preview".as_slice())),
    );
    Ok(litchi_iwa_archive::package::to_bytes(
        members,
        Limits::default(),
    )?)
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn member_bytes(source: &[u8], name: &str) -> TestResult<Vec<u8>> {
    Ok(Catalog::from_bytes(source)?
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or_else(|| io::Error::other(format!("missing member {name}")))?
        .data()
        .to_vec())
}

fn member_archive(source: &[u8], name: &str) -> TestResult<Archive> {
    Ok(Archive::parse(
        SnappyStream::decompress(&member_bytes(source, name)?)?.as_bytes(),
    )?)
}

fn replace_member_archive(
    source: &[u8],
    name: &str,
    mutate: impl FnOnce(&mut Archive) -> TestResult,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let mut archive = member_archive(source, name)?;
    mutate(&mut archive)?;
    let replacement = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(&[EntryEdit::new(name, &replacement)], Limits::default())?)
}

fn rewrite_object_message<F>(
    source: &[u8],
    member: &str,
    identifier: u64,
    rewrite: F,
) -> TestResult<Vec<u8>>
where
    F: FnOnce(&mut RawMessage) -> TestResult,
{
    replace_member_archive(source, member, |archive| {
        let object = archive
            .object_mut(identifier)
            .ok_or_else(|| io::Error::other(format!("missing object {identifier}")))?;
        let message = object
            .messages
            .first_mut()
            .ok_or_else(|| io::Error::other("object has no message"))?;
        let mut replacement = RawMessage {
            type_: message.type_,
            data: message.data.clone(),
        };
        rewrite(&mut replacement)?;
        object.replace_message_preserving_header(0, replacement)?;
        Ok(())
    })
}

fn rewrite_model_payload<F>(source: &[u8], identifier: u64, rewrite: F) -> TestResult<Vec<u8>>
where
    F: FnOnce(&mut Vec<u8>) -> TestResult,
{
    rewrite_object_message(source, MODEL_MEMBER, identifier, |message| {
        rewrite(&mut message.data)
    })
}

fn rewrite_style_payload<F>(source: &[u8], identifier: u64, rewrite: F) -> TestResult<Vec<u8>>
where
    F: FnOnce(&mut Vec<u8>) -> TestResult,
{
    rewrite_object_message(source, STYLESHEET_MEMBER, identifier, |message| {
        rewrite(&mut message.data)
    })
}

fn rewrite_archive_info<F>(
    source: &[u8],
    member: &str,
    identifier: u64,
    rewrite: F,
) -> TestResult<Vec<u8>>
where
    F: FnOnce(&mut litchi_iwa_core::MessageInfo),
{
    replace_member_archive(source, member, |archive| {
        let object = archive
            .object_mut(identifier)
            .ok_or_else(|| io::Error::other(format!("missing object {identifier}")))?;
        let info = object
            .archive_info
            .message_infos
            .first_mut()
            .ok_or_else(|| io::Error::other("object has no archive info"))?;
        rewrite(info);
        Ok(())
    })
}

fn replace_model_type(source: &[u8], identifier: u64, type_: u32) -> TestResult<Vec<u8>> {
    rewrite_object_message(source, MODEL_MEMBER, identifier, |message| {
        message.type_ = type_;
        Ok(())
    })
}

fn metadata_payload_from_source(source: &[u8]) -> TestResult<Vec<u8>> {
    let archive = member_archive(source, METADATA_MEMBER)?;
    archive
        .objects
        .iter()
        .find(|object| object.archive_info.identifier == Some(METADATA_OBJECT_ID))
        .and_then(|object| {
            object
                .messages
                .iter()
                .find(|message| message.type_ == METADATA_MESSAGE_TYPE)
        })
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("metadata payload is missing").into())
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

fn direct_appearance() -> Appearance {
    Appearance {
        row_banding: Banding::Enabled,
        row_sizing: RowSizing::FitCellContents,
        gridlines: Gridlines {
            body_horizontal: GridlineVisibility::Hidden,
            header_columns_horizontal: GridlineVisibility::Visible,
            body_vertical: GridlineVisibility::Visible,
            header_rows_vertical: GridlineVisibility::Hidden,
            footer_rows_vertical: GridlineVisibility::Hidden,
        },
    }
}

fn preset_appearance() -> Appearance {
    Appearance {
        row_banding: Banding::Disabled,
        row_sizing: RowSizing::Fixed,
        gridlines: Gridlines {
            body_horizontal: GridlineVisibility::Visible,
            header_columns_horizontal: GridlineVisibility::Hidden,
            body_vertical: GridlineVisibility::Hidden,
            header_rows_vertical: GridlineVisibility::Visible,
            footer_rows_vertical: GridlineVisibility::Visible,
        },
    }
}

fn assert_rejected_atomically(source: &[u8], label: &str) -> TestResult {
    assert_rejected_atomically_at(source, 0, label)
}

fn assert_rejected_atomically_at(source: &[u8], table: usize, label: &str) -> TestResult {
    let Ok(package) = Package::from_bytes(source) else {
        return Ok(());
    };
    let before = exact_bytes(&package)?;
    let result = package
        .edit_slide_table_appearance(0usize, table)
        .and_then(|edit| edit.set(custom_appearance()).commit());
    assert!(
        result.is_err(),
        "hostile appearance graph was accepted: {label}"
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

fn member_names(source: &[u8]) -> TestResult<Vec<String>> {
    let mut names = Catalog::from_bytes(source)?
        .iter()
        .map(|entry| entry.name().to_owned())
        .collect::<Vec<_>>();
    names.sort_unstable();
    Ok(names)
}

fn common_changed_member_names(source: &[u8], target: &[u8]) -> TestResult<Vec<String>> {
    let source_names = member_names(source)?;
    let target_names = member_names(target)?;
    let mut changed = Vec::new();
    for name in source_names {
        if target_names.binary_search(&name).is_ok()
            && member_bytes(source, &name)? != member_bytes(target, &name)?
        {
            changed.push(name);
        }
    }
    changed.sort_unstable();
    Ok(changed)
}

fn assert_unaffected_members_exact(source: &[u8], target: &[u8], names: &[&str]) -> TestResult {
    for name in names {
        assert_eq!(
            member_bytes(source, name)?,
            member_bytes(target, name)?,
            "member {name} changed"
        );
    }
    Ok(())
}

fn assert_send_sync_debug<T: Send + Sync + std::fmt::Debug>() {}

#[test]
fn typed_selectors_effective_values_and_z_order_are_publicly_stable() -> TestResult {
    let source = synthetic_package(false, false)?;
    let package = Package::from_bytes(&source)?;

    assert_eq!(
        package.slide_table_appearance(
            SlideSelector::name("Appearance Tables"),
            TableSelector::index(0)
        )?,
        direct_appearance()
    );
    assert_eq!(
        package.slide_table_appearance(SlideSelector::index(0), TableSelector::index(1))?,
        direct_appearance()
    );
    assert_eq!(
        package.slide_table_appearance(SlideSelector::index(0), TableSelector::index(2))?,
        preset_appearance()
    );
    assert_eq!(
        package.slide_table_appearance(SlideSelector::index(0), TableSelector::index(3))?,
        Appearance::default()
    );
    assert!(
        package
            .slide_table_appearance(0usize, TableSelector::index(4))
            .is_err()
    );
    assert!(SlideSelector::try_name("").is_err());
    Ok(())
}

#[cfg(feature = "internal-iwork-source")]
#[test]
fn focused_batch_appearance_read_preserves_selector_order() -> TestResult {
    let source = synthetic_package(false, false)?;
    let package = Package::from_bytes(&source)?;
    let batch = package.__slide_table_appearances(0)?;

    assert_eq!(
        batch,
        vec![
            direct_appearance(),
            direct_appearance(),
            preset_appearance(),
            Appearance::default(),
        ]
    );
    for (table, expected) in batch.iter().copied().enumerate() {
        assert_eq!(
            package.slide_table_appearance(0usize, table)?,
            expected,
            "batch result diverged at table position {table}"
        );
    }
    Ok(())
}

#[cfg(feature = "internal-iwork-source")]
#[test]
fn focused_batch_appearance_read_validates_unselected_tables() -> TestResult {
    let source = rewrite_style_payload(
        &synthetic_package(false, false)?,
        PRESET_STYLE_ID,
        |payload| {
            *payload = style_payload(PRESET_STYLE_ID, Some(9_999), [Some(true); 7], false)?;
            Ok(())
        },
    )?;
    let package = Package::from_bytes(&source)?;
    assert!(package.slide_table_appearance(0usize, 0usize).is_err());
    assert!(package.__slide_table_appearances(0).is_err());
    Ok(())
}

#[test]
fn direct_preset_and_default_resolution_preserve_precedence() -> TestResult {
    let source = synthetic_package(false, false)?;
    let malformed_preset = rewrite_model_payload(&source, TABLE_MODEL_IDS[0], |payload| {
        append_length_delimited_field(
            payload,
            TABLE_MODEL_PRESET_FIELD,
            &reference_payload(9_999),
        )?;
        Ok(())
    })?;
    let package = Package::from_bytes(&malformed_preset)?;
    assert_eq!(
        package.slide_table_appearance(0usize, 0usize)?,
        direct_appearance(),
        "direct style must win over an unavailable preset"
    );

    let missing_parent = rewrite_style_payload(&source, SHARED_STYLE_ID, |payload| {
        let replacement = style_payload(SHARED_STYLE_ID, Some(9_999), [Some(true); 7], false)?;
        *payload = replacement;
        Ok(())
    })?;
    let package = Package::from_bytes(&missing_parent)?;
    assert!(package.slide_table_appearance(0usize, 0usize).is_err());

    let cycle = rewrite_style_payload(&source, PARENT_STYLE_ID, |payload| {
        *payload = style_payload(
            PARENT_STYLE_ID,
            Some(SHARED_STYLE_ID),
            [Some(false); 7],
            false,
        )?;
        Ok(())
    })?;
    let package = Package::from_bytes(&cycle)?;
    assert!(package.slide_table_appearance(0usize, 0usize).is_err());
    Ok(())
}

#[test]
fn unknown_groups_are_admitted_for_read_and_source_remains_exact() -> TestResult {
    let source = synthetic_package(false, true)?;
    let package = Package::from_bytes(&source)
        .map_err(|error| io::Error::other(format!("producer-shape ingress: {error:?}")))?;
    assert_eq!(
        package
            .slide_table_appearance(0usize, 0usize)
            .map_err(|error| io::Error::other(format!("producer-shape read: {error:?}")))?,
        direct_appearance()
    );
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn changed_appearance_cows_shared_style_and_preserves_locality() -> TestResult {
    let source = synthetic_package(false, true)?;
    let package = Package::from_bytes(&source)?;
    let before = package.slide_table_appearance(0usize, 0usize)?;
    let commit = package
        .edit_slide_table_appearance(0usize, 0usize)?
        .set(custom_appearance())
        .commit()
        .map_err(|error| io::Error::other(format!("producer-shape commit: {error:?}")))?;
    let target = exact_bytes(commit.package())?;
    assert_eq!(
        commit.package().slide_table_appearance(0usize, 0usize)?,
        custom_appearance()
    );
    assert_eq!(
        commit.package().slide_table_appearance(0usize, 1usize)?,
        before
    );
    assert_eq!(
        commit.package().slide_table_appearance(0usize, 2usize)?,
        preset_appearance()
    );
    let target_styles = member_archive(&target, STYLESHEET_MEMBER)?;
    let fresh_style = target_styles
        .objects
        .iter()
        .find(|object| {
            object.messages.iter().any(|message| {
                message.type_ == TABLE_STYLE_MESSAGE_TYPE
                    && ![SHARED_STYLE_ID, PARENT_STYLE_ID, PRESET_STYLE_ID]
                        .contains(&object.archive_info.identifier.unwrap_or(0))
            })
        })
        .ok_or_else(|| io::Error::other("missing fresh COW style"))?;
    let fresh_info = fresh_style
        .archive_info
        .message_infos
        .iter()
        .find(|info| info.type_ == TABLE_STYLE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("missing fresh style ArchiveInfo"))?;
    assert_eq!(
        fresh_info.object_references,
        [SHARED_STYLE_ID, STYLESHEET_ID]
    );
    assert_eq!(fresh_info.field_infos.len(), 2);
    assert_eq!(
        fresh_info.field_infos[0].path.as_slice(),
        [TABLE_STYLE_SUPER_FIELD, TABLE_STYLE_PARENT_FIELD]
    );
    assert_eq!(
        fresh_info.field_infos[0].object_references,
        [SHARED_STYLE_ID]
    );
    assert_eq!(
        fresh_info.field_infos[1].path.as_slice(),
        [TABLE_STYLE_SUPER_FIELD, TABLE_STYLE_STYLESHEET_FIELD]
    );
    assert_eq!(fresh_info.field_infos[1].object_references, [STYLESHEET_ID]);
    let source_names = member_names(&source)?;
    let target_names = member_names(&target)?;
    let mut expected_target_names = vec![
        SENTINEL_MEMBER,
        DOCUMENT_MEMBER,
        MODEL_MEMBER,
        STYLESHEET_MEMBER,
        METADATA_MEMBER,
    ]
    .into_iter()
    .map(str::to_owned)
    .collect::<Vec<_>>();
    expected_target_names.sort_unstable();
    assert_eq!(target_names, expected_target_names);
    let removed_names = source_names
        .iter()
        .filter(|name| target_names.binary_search(name).is_err())
        .cloned()
        .collect::<Vec<_>>();
    let mut expected_removed_names = PREVIEWS.into_iter().map(str::to_owned).collect::<Vec<_>>();
    expected_removed_names.sort_unstable();
    assert_eq!(removed_names, expected_removed_names);
    assert_eq!(
        common_changed_member_names(&source, &target)?,
        vec![
            MODEL_MEMBER.to_owned(),
            STYLESHEET_MEMBER.to_owned(),
            METADATA_MEMBER.to_owned(),
        ]
    );
    assert_unaffected_members_exact(&source, &target, &[DOCUMENT_MEMBER, SENTINEL_MEMBER])?;
    for preview in PREVIEWS {
        assert!(
            Catalog::from_bytes(&target)?
                .iter()
                .all(|entry| entry.name() != preview)
        );
    }
    assert_ne!(
        member_bytes(&source, MODEL_MEMBER)?,
        member_bytes(&target, MODEL_MEMBER)?
    );
    assert_ne!(
        member_bytes(&source, STYLESHEET_MEMBER)?,
        member_bytes(&target, STYLESHEET_MEMBER)?
    );
    assert_ne!(
        member_bytes(&source, METADATA_MEMBER)?,
        member_bytes(&target, METADATA_MEMBER)?
    );
    assert!(commit.diagnostics().changed());
    assert!(commit.diagnostics().full_reparse_performed());
    let inverse = commit.patch().inverse();
    let restored = commit.package().apply_slide_table_appearance(&inverse)?;
    let restored_bytes = exact_bytes(restored.package())?;
    assert_eq!(restored_bytes, source);
    assert_eq!(member_names(&restored_bytes)?, source_names);
    Ok(())
}

#[test]
fn no_op_is_exact_and_apply_conflict_is_source_bound() -> TestResult {
    let source = synthetic_package(false, false)?;
    let package = Package::from_bytes(&source)?;
    let appearance = package.slide_table_appearance(0usize, 0usize)?;
    let edit = package.edit_slide_table_appearance(0usize, 0usize)?;
    assert_eq!(edit.before(), appearance);
    let commit = edit.set(appearance).commit()?;
    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    assert_eq!(exact_bytes(commit.package())?, source);

    let changed = Package::from_bytes(&source)?
        .edit_slide_table_appearance(0usize, 0usize)?
        .set(custom_appearance())
        .commit()?;
    let unrelated = Catalog::from_bytes(&source)?.reassemble_to_bytes(
        &[EntryEdit::new(SENTINEL_MEMBER, b"a different sentinel")],
        Limits::default(),
    )?;
    let unrelated_package = Package::from_bytes(&unrelated)?;
    assert!(matches!(
        unrelated_package.apply_slide_table_appearance(changed.patch()),
        Err(litchi_keynote::SlideTableAppearanceError::PatchConflict)
    ));
    Ok(())
}

#[test]
fn locked_table_allows_noop_but_refuses_changed_source_atomically() -> TestResult {
    let source = synthetic_package(true, false)?;
    let package = Package::from_bytes(&source)?;
    let before = package.slide_table_appearance(0usize, 0usize)?;
    let noop = package
        .edit_slide_table_appearance(0usize, 0usize)?
        .set(before)
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(exact_bytes(noop.package())?, source);
    assert_rejected_atomically(&source, "locked table")?;
    Ok(())
}

#[test]
fn malformed_role_and_archive_info_routes_fail_closed_atomically() -> TestResult {
    let source = synthetic_package(false, false)?;
    let duplicate_style = rewrite_style_payload(&source, SHARED_STYLE_ID, |payload| {
        append_varint_field(payload, TABLE_STYLE_OVERRIDE_COUNT_FIELD, 7)?;
        Ok(())
    })?;
    assert_rejected_atomically(&duplicate_style, "duplicate style field")?;

    let duplicate_model_style = rewrite_model_payload(&source, TABLE_MODEL_IDS[0], |payload| {
        append_length_delimited_field(
            payload,
            TABLE_MODEL_STYLE_FIELD,
            &reference_payload(SHARED_STYLE_ID),
        )?;
        Ok(())
    })?;
    assert_rejected_atomically(&duplicate_model_style, "duplicate model style edge")?;

    let wrong_style_role =
        replace_model_type(&source, TABLE_MODEL_IDS[0], TABLE_STYLE_MESSAGE_TYPE)?;
    assert_rejected_atomically(&wrong_style_role, "model/style role alias")?;

    let missing_archive_edge =
        rewrite_archive_info(&source, MODEL_MEMBER, TABLE_MODEL_IDS[0], |info| {
            info.object_references.clear();
            info.field_infos.clear();
        })?;
    assert_rejected_atomically(&missing_archive_edge, "missing style ArchiveInfo edge")?;
    Ok(())
}

#[test]
fn duplicate_style_parent_field_without_stylesheet_field_is_rejected_atomically() -> TestResult {
    let source = synthetic_package(false, false)?;
    let duplicate_parent =
        rewrite_archive_info(&source, STYLESHEET_MEMBER, SHARED_STYLE_ID, |info| {
            info.field_infos = vec![
                field_reference(
                    vec![TABLE_STYLE_SUPER_FIELD, TABLE_STYLE_PARENT_FIELD],
                    &[PARENT_STYLE_ID],
                ),
                field_reference(
                    vec![TABLE_STYLE_SUPER_FIELD, TABLE_STYLE_PARENT_FIELD],
                    &[PARENT_STYLE_ID],
                ),
            ];
        })?;
    assert_rejected_atomically(
        &duplicate_parent,
        "duplicate style parent field without stylesheet field",
    )?;
    Ok(())
}

#[test]
fn broad_native_archive_metadata_and_omitted_false_flags_roundtrip_exactly() -> TestResult {
    let source = synthetic_package(false, false)?;
    let source = rewrite_archive_info(&source, DOCUMENT_MEMBER, SLIDE_ID, |info| {
        info.object_references.push(SHOW_ID);
        let mut unrelated = FieldInfo::new(vec![45]);
        unrelated.r#type = Some(FieldType::Message);
        info.field_infos = vec![unrelated];
    })?;
    let source = rewrite_archive_info(&source, DOCUMENT_MEMBER, TABLE_INFO_IDS[0], |info| {
        info.object_references.push(SHOW_ID);
        let mut unrelated = FieldInfo::new(vec![10]);
        unrelated.r#type = Some(FieldType::Message);
        info.field_infos = vec![unrelated];
    })?;
    let source = rewrite_archive_info(&source, MODEL_MEMBER, TABLE_MODEL_IDS[0], |info| {
        info.object_references.push(SHOW_ID);
        info.object_references.push(NON_TABLE_DRAWABLE_ID);
        let mut unrelated = FieldInfo::new(vec![70]);
        unrelated.r#type = Some(FieldType::Message);
        unrelated.object_references.push(NON_TABLE_DRAWABLE_ID);
        info.field_infos = vec![unrelated];
    })?;
    let source = rewrite_archive_info(&source, STYLESHEET_MEMBER, SHARED_STYLE_ID, |info| {
        info.object_references.clear();
        info.field_infos.clear();
    })?;
    let source = rewrite_archive_info(&source, STYLESHEET_MEMBER, STYLESHEET_ID, |info| {
        info.object_references
            .retain(|identifier| *identifier != PRESET_STYLE_ID);
        info.object_references.push(NON_TABLE_DRAWABLE_ID);
        let mut unrelated = FieldInfo::new(vec![14]);
        unrelated.r#type = Some(FieldType::Message);
        unrelated.object_references.push(NON_TABLE_DRAWABLE_ID);
        info.field_infos = vec![unrelated];
    })?;
    let source = rewrite_metadata(&source, |metadata| {
        if let Some(component) = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 300)
        {
            component
                .object_uuid_map_entries
                .retain(|entry| entry.identifier != STYLESHEET_ID);
        }
        if let Some(component) = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 200)
        {
            if let Some(edge) = component.external_references.iter_mut().find(|edge| {
                edge.component_identifier == 300 && edge.object_identifier == Some(SHARED_STYLE_ID)
            }) {
                edge.is_weak = None;
            }
        }
    })?;

    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.slide_table_appearance(0usize, 0usize)?,
        direct_appearance()
    );
    let changed = package
        .edit_slide_table_appearance(0usize, 0usize)?
        .set(custom_appearance())
        .commit()?;
    assert_eq!(
        changed.package().slide_table_appearance(0usize, 0usize)?,
        custom_appearance()
    );
    let changed_bytes = exact_bytes(changed.package())?;
    let changed_styles = member_archive(&changed_bytes, STYLESHEET_MEMBER)?;
    let fresh_style = changed_styles
        .objects
        .iter()
        .find(|object| {
            object.messages.iter().any(|message| {
                message.type_ == TABLE_STYLE_MESSAGE_TYPE
                    && ![SHARED_STYLE_ID, PARENT_STYLE_ID, PRESET_STYLE_ID]
                        .contains(&object.archive_info.identifier.unwrap_or(0))
            })
        })
        .ok_or_else(|| io::Error::other("missing omitted-metadata COW style"))?;
    let fresh_info = fresh_style
        .archive_info
        .message_infos
        .iter()
        .find(|info| info.type_ == TABLE_STYLE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("missing omitted-metadata COW ArchiveInfo"))?;
    assert!(fresh_info.object_references.is_empty());
    assert!(fresh_info.field_infos.is_empty());
    let restored = changed
        .package()
        .apply_slide_table_appearance(&changed.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn preset_and_network_archive_reference_shapes_are_exact() -> TestResult {
    let source = synthetic_package(false, false)?;

    let missing_preset_network =
        rewrite_archive_info(&source, STYLESHEET_MEMBER, PRESET_ID, |info| {
            info.object_references.clear();
        })?;
    assert_rejected_atomically_at(
        &missing_preset_network,
        2,
        "missing preset network aggregate reference",
    )?;

    let extra_network_style =
        rewrite_archive_info(&source, STYLESHEET_MEMBER, NETWORK_ID, |info| {
            info.object_references.push(SHARED_STYLE_ID);
        })?;
    assert_rejected_atomically_at(
        &extra_network_style,
        2,
        "extra network style aggregate reference",
    )?;

    let wrong_network_field =
        rewrite_archive_info(&source, STYLESHEET_MEMBER, NETWORK_ID, |info| {
            info.field_infos
                .push(field_reference(vec![8], &[PRESET_STYLE_ID]));
        })?;
    assert_rejected_atomically_at(
        &wrong_network_field,
        2,
        "unexpected network field reference",
    )?;
    Ok(())
}

#[test]
fn stale_same_component_style_stylesheet_edge_is_rejected_atomically() -> TestResult {
    let source = synthetic_package(false, false)?;
    let stale_edge = rewrite_metadata(&source, |metadata| {
        if let Some(component) = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 300)
        {
            component
                .external_references
                .push(tsp::ComponentExternalReference {
                    component_identifier: 300,
                    object_identifier: Some(STYLESHEET_ID),
                    is_weak: Some(false),
                });
        }
    })?;
    assert_rejected_atomically(&stale_edge, "stale same-component style stylesheet edge")?;
    Ok(())
}

#[test]
fn unrelated_archive_field_and_data_references_are_authoritative() -> TestResult {
    let source = synthetic_package(false, false)?;

    let undeclared_field_reference =
        rewrite_archive_info(&source, MODEL_MEMBER, TABLE_MODEL_IDS[0], |info| {
            let mut field = FieldInfo::new(vec![70]);
            field.r#type = Some(FieldType::Message);
            field.object_references.push(9_999_999);
            info.field_infos.push(field);
        })?;
    assert_rejected_atomically(
        &undeclared_field_reference,
        "unresolved unrelated FieldInfo reference",
    )?;

    let unresolved_aggregate_reference =
        rewrite_archive_info(&source, MODEL_MEMBER, TABLE_MODEL_IDS[0], |info| {
            info.object_references.push(9_999_999);
        })?;
    assert_rejected_atomically(
        &unresolved_aggregate_reference,
        "unresolved unrelated aggregate reference",
    )?;

    let unregistered_data_reference =
        rewrite_archive_info(&source, MODEL_MEMBER, TABLE_MODEL_IDS[0], |info| {
            info.data_references.push(9_999_998);
        })?;
    assert_rejected_atomically(
        &unregistered_data_reference,
        "ArchiveInfo data reference absent from current Metadata",
    )?;
    Ok(())
}

#[test]
fn component_external_edges_are_complete_current_and_object_specific() -> TestResult {
    let source = synthetic_package(false, false)?;

    let wildcard = rewrite_metadata(&source, |metadata| {
        if let Some(component) = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 200)
        {
            component
                .external_references
                .push(tsp::ComponentExternalReference {
                    component_identifier: 300,
                    object_identifier: None,
                    is_weak: Some(false),
                });
        }
    })?;
    assert_rejected_atomically(&wildcard, "wildcard model-to-style component edge")?;

    let duplicate = rewrite_metadata(&source, |metadata| {
        if let Some(component) = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 200)
        {
            component
                .external_references
                .push(tsp::ComponentExternalReference {
                    component_identifier: 300,
                    object_identifier: Some(PARENT_STYLE_ID),
                    is_weak: Some(false),
                });
        }
    })?;
    assert_rejected_atomically(&duplicate, "duplicate model component object edge")?;

    let weak_sibling = rewrite_metadata(&source, |metadata| {
        if let Some(component) = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 200)
            && let Some(edge) = component.external_references.iter_mut().find(|edge| {
                edge.component_identifier == 300 && edge.object_identifier == Some(PARENT_STYLE_ID)
            })
        {
            edge.is_weak = Some(true);
        }
    })?;
    assert_rejected_atomically(&weak_sibling, "weak sibling style component edge")?;
    Ok(())
}

#[test]
fn inbound_roles_are_classified_by_the_exact_referencing_message() -> TestResult {
    let source = synthetic_package(false, false)?;
    let mixed_role = replace_member_archive(&source, DOCUMENT_MEMBER, |archive| {
        let mut object = ArchiveObject::new(
            FOREIGN_INBOUND_ID,
            vec![
                RawMessage {
                    type_: 9_999,
                    data: vec![0x08, 0x01],
                },
                RawMessage {
                    type_: TABLE_STYLE_MESSAGE_TYPE,
                    data: style_payload(FOREIGN_INBOUND_ID, None, [None; 7], false)?,
                },
            ],
        )?;
        object.archive_info.message_infos[0]
            .object_references
            .push(SHARED_STYLE_ID);
        archive.objects.push(object);
        Ok(())
    })?;
    assert_rejected_atomically(
        &mixed_role,
        "wrong-type message on a mixed-role object references the selected style",
    )?;
    Ok(())
}

#[test]
fn surviving_style_uuid_ownership_accepts_unregistered_registry_and_rejects_conflicts() -> TestResult
{
    let source = synthetic_package(false, false)?;

    let duplicate_parent_uuid = rewrite_metadata(&source, |metadata| {
        if let Some(component) = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 300)
        {
            component
                .object_uuid_map_entries
                .push(uuid_entry(PARENT_STYLE_ID));
        }
    })?;
    assert_rejected_atomically(
        &duplicate_parent_uuid,
        "duplicate current inherited style UUID",
    )?;

    let missing_model_uuid = rewrite_metadata(&source, |metadata| {
        if let Some(component) = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 200)
        {
            component
                .object_uuid_map_entries
                .retain(|entry| entry.identifier != TABLE_MODEL_IDS[0]);
        }
    })?;
    assert_rejected_atomically(&missing_model_uuid, "missing current model UUID")?;

    let missing_stylesheet_uuid = rewrite_metadata(&source, |metadata| {
        if let Some(component) = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 300)
        {
            component
                .object_uuid_map_entries
                .retain(|entry| entry.identifier != STYLESHEET_ID);
        }
    })?;
    let package = Package::from_bytes(&missing_stylesheet_uuid)?;
    let changed = package
        .edit_slide_table_appearance(0usize, 0usize)?
        .set(custom_appearance())
        .commit()?;
    assert_eq!(
        changed.package().slide_table_appearance(0usize, 0usize)?,
        custom_appearance()
    );

    let versioned_model_uuid = rewrite_metadata(&source, |metadata| {
        metadata.versioned_components.push(tsp::ComponentInfo {
            identifier: 400,
            preferred_locator: "HistoricalModel".to_owned(),
            locator: Some("HistoricalModel".to_owned()),
            object_uuid_map_entries: vec![uuid_entry(TABLE_MODEL_IDS[0])],
            ..Default::default()
        });
    })?;
    assert_rejected_atomically(&versioned_model_uuid, "versioned model UUID")?;

    let versioned_stylesheet_uuid = rewrite_metadata(&source, |metadata| {
        metadata.versioned_components.push(tsp::ComponentInfo {
            identifier: 401,
            preferred_locator: "HistoricalStylesheet".to_owned(),
            locator: Some("HistoricalStylesheet".to_owned()),
            object_uuid_map_entries: vec![uuid_entry(STYLESHEET_ID)],
            ..Default::default()
        });
    })?;
    assert_rejected_atomically(&versioned_stylesheet_uuid, "versioned stylesheet UUID")?;

    let versioned_parent_uuid = rewrite_metadata(&source, |metadata| {
        metadata.versioned_components.push(tsp::ComponentInfo {
            identifier: 402,
            preferred_locator: "HistoricalParentStyle".to_owned(),
            locator: Some("HistoricalParentStyle".to_owned()),
            object_uuid_map_entries: vec![uuid_entry(PARENT_STYLE_ID)],
            ..Default::default()
        });
    })?;
    assert_rejected_atomically(&versioned_parent_uuid, "versioned inherited style UUID")?;
    Ok(())
}

#[test]
fn missing_stylesheet_metadata_and_foreign_inbound_are_atomic() -> TestResult {
    let source = synthetic_package(false, false)?;
    let missing_metadata = Catalog::from_bytes(&source)?.reassemble_with_deletions_to_bytes(
        &[],
        &[METADATA_MEMBER],
        Limits::default(),
    )?;
    assert_rejected_atomically(&missing_metadata, "missing PackageMetadata")?;

    let missing_style = replace_member_archive(&source, STYLESHEET_MEMBER, |archive| {
        archive
            .objects
            .retain(|object| object.archive_info.identifier != Some(SHARED_STYLE_ID));
        Ok(())
    })?;
    assert_rejected_atomically(&missing_style, "missing selected style")?;

    let foreign = replace_member_archive(&source, DOCUMENT_MEMBER, |archive| {
        archive.objects.push(object(
            FOREIGN_INBOUND_ID,
            9_999,
            vec![0x08, 0x01],
            &[SHARED_STYLE_ID],
        )?);
        Ok(())
    })?;
    assert_rejected_atomically(&foreign, "foreign inbound style owner")?;
    Ok(())
}

#[test]
fn metadata_uuid_and_locator_collisions_fail_before_publication() -> TestResult {
    let source = synthetic_package(false, false)?;
    let missing_uuid = rewrite_metadata(&source, |metadata| {
        if let Some(component) = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 300)
        {
            component
                .object_uuid_map_entries
                .retain(|entry| entry.identifier != SHARED_STYLE_ID);
        }
    })?;
    assert_rejected_atomically(&missing_uuid, "missing current style UUID")?;

    let locator_disagreement = rewrite_metadata(&source, |metadata| {
        if let Some(component) = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 300)
        {
            component.locator = Some("CalculationEngine".to_owned());
        }
    })?;
    assert_rejected_atomically(&locator_disagreement, "metadata locator disagreement")?;

    let stale_payload = metadata_payload_from_source(&source)?;
    assert!(!stale_payload.is_empty());
    Ok(())
}

#[test]
fn finite_semantic_limits_refuse_before_candidate_publication() -> TestResult {
    // Variation work is private to the prepared rewrite and has no public
    // ReadOptions/SemanticLimits setter, so a deterministic max-minus-one
    // variation case is not caller-accessible at this integration boundary.
    let source = synthetic_package(false, false)?;
    let mut saw_limit = false;
    for references in 1..=64 {
        let semantic = SemanticLimits::new(
            MAX_OBJECTS,
            MAX_SLIDES,
            references,
            MAX_TEXT_STORAGES,
            MAX_TEXT_FRAGMENTS,
            MAX_TEXT_BYTES,
        )?;
        let options = ReadOptions::new(Limits::default(), semantic);
        let Ok(package) = Package::from_bytes_with_options(&source, options) else {
            continue;
        };
        let before = exact_bytes(&package)?;
        let result = package
            .edit_slide_table_appearance(0usize, 0usize)
            .and_then(|edit| edit.set(custom_appearance()).commit());
        if matches!(
            result,
            Err(litchi_keynote::SlideTableAppearanceError::LimitExceeded { .. })
        ) {
            assert_eq!(exact_bytes(&package)?, before);
            saw_limit = true;
            break;
        }
    }
    assert!(saw_limit, "semantic reference limit never reached");
    Ok(())
}

#[test]
fn transaction_values_are_typed_archive_free_and_redacted() {
    assert_send_sync_debug::<Edit<'static>>();
    assert_send_sync_debug::<Patch>();
    assert_send_sync_debug::<Commit>();
    assert_send_sync_debug::<Diagnostics>();
    assert_send_sync_debug::<Error>();
    assert_send_sync_debug::<LimitKind>();
    assert_send_sync_debug::<Path>();
    let appearance = custom_appearance();
    let debug = format!("{appearance:?}");
    assert!(!debug.contains("Archive"));
    assert!(!debug.contains("RawMessage"));
    assert!(!debug.contains("CalculationEngine"));
}

fn reference_payload(identifier: u64) -> Vec<u8> {
    let mut payload = Vec::new();
    append_varint_field(&mut payload, 1, identifier).expect("Vec cannot fail");
    payload
}
