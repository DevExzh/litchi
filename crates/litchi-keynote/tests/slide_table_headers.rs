//! Native integration coverage for selector-first Keynote slide-table
//! header/footer settings.
//!
//! The fixture is intentionally small but rooted like a real presentation:
//! Document -> Show -> SlideNode -> Slide -> z-order drawables.  The model
//! objects live in a separate CalculationEngine component so every changed
//! transaction has to prove its cross-member route while preserving the
//! document, Metadata, previews, and the unselected table.

use std::error::Error as StdError;
use std::io;

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::{
    decode_varint_from_bytes,
    wire::{
        WireView, append_length_delimited_field, append_varint_field,
        repeated_length_delimited_payloads,
    },
};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, FieldPath, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsce, tsd, tsk, tsp, tss, tst, tswp};
use litchi_keynote::slide::table::headers::{Count, Settings};
use litchi_keynote::{
    MAX_OBJECTS, MAX_SLIDES, MAX_TEXT_BYTES, MAX_TEXT_FRAGMENTS, MAX_TEXT_STORAGES, Package,
    ReadOptions, SemanticLimits, SlideSelector, SlideTableHeaderCommit,
    SlideTableHeaderDiagnostics, SlideTableHeaderEdit, SlideTableHeaderError,
    SlideTableHeaderLimitKind, SlideTableHeaderPatch, SlideTableHeaderPath, TableSelector,
};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const MODEL_MEMBER: &str = "Index/CalculationEngine.iwa";
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const SENTINEL_MEMBER: &str = "Data/header-sentinel.bin";

const SLIDE_NODE: u64 = 3;
const SLIDE: u64 = 4;
const TABLE_INFOS: [u64; 2] = [100, 101];
const MODELS: [u64; 2] = [110, 111];
const NON_TABLE_DRAWABLE: u64 = 130;
const TITLE_STYLE: u64 = 120;
const SHAPE_STYLE: u64 = 121;
const METADATA_OBJECT: u64 = 900;
const FOREIGN_INBOUND_OBJECT: u64 = 902;
const CALCULATION_ENGINE: u64 = 903;
const HEADER_NAME_MANAGER: u64 = 904;
const PIVOT_OWNER: u64 = 905;
const TABLE_INFO_CACHE: u64 = 906;

const DOCUMENT_MESSAGE_TYPE: u32 = 1;
const SHOW_MESSAGE_TYPE: u32 = 2;
const SLIDE_NODE_MESSAGE_TYPE: u32 = 4;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;
const METADATA_MESSAGE_TYPE: u32 = 11_006;
const CALCULATION_ENGINE_MESSAGE_TYPE: u32 = 4_000;
const HEADER_NAME_MANAGER_MESSAGE_TYPE: u32 = 6_366;
const PIVOT_OWNER_MESSAGE_TYPE: u32 = 6_370;

const TABLE_MODEL_FIELD_HEADER_ROWS: u32 = 9;
const TABLE_MODEL_FIELD_HEADER_COLUMNS: u32 = 10;
const TABLE_MODEL_FIELD_FOOTER_ROWS: u32 = 11;
const TABLE_MODEL_FIELD_HEADER_ROWS_FROZEN: u32 = 12;
const TABLE_MODEL_FIELD_HEADER_COLUMNS_FROZEN: u32 = 13;
const TABLE_MODEL_FIELD_REPEAT_ROWS: u32 = 29;
const TABLE_MODEL_FIELD_REPEAT_COLUMNS: u32 = 32;
const TABLE_MODEL_FIELD_TRACKER: u32 = 45;
const UNKNOWN_MODEL_FIELD: u32 = 90;
const UNKNOWN_MODEL_VALUE: u64 = 0xfeed_beef;
const TRACKER_BYTES: &[u8] = &[0x0a, 0x02, 0x08, 0x01];
const PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn field_reference(path: impl Into<FieldPath>, references: &[u64]) -> FieldInfo {
    let mut field = FieldInfo::new(path);
    field.object_references.extend_from_slice(references);
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

fn fixture_settings(index: usize) -> Settings {
    if index == 0 {
        Settings {
            header_rows: Some(Count::TWO),
            header_columns: Some(Count::ONE),
            footer_rows: Some(Count::THREE),
            header_rows_frozen: Some(false),
            header_columns_frozen: Some(true),
            repeating_header_rows_enabled: Some(true),
            repeating_header_columns_enabled: Some(false),
        }
    } else {
        Settings {
            header_rows: Some(Count::ONE),
            header_columns: Some(Count::TWO),
            footer_rows: Some(Count::ONE),
            header_rows_frozen: Some(true),
            header_columns_frozen: Some(false),
            repeating_header_rows_enabled: Some(false),
            repeating_header_columns_enabled: Some(true),
        }
    }
}

fn changed_settings() -> Settings {
    Settings {
        header_rows: Some(Count::THREE),
        header_columns: Some(Count::TWO),
        footer_rows: Some(Count::ONE),
        header_rows_frozen: Some(true),
        header_columns_frozen: Some(false),
        repeating_header_rows_enabled: Some(false),
        repeating_header_columns_enabled: Some(true),
    }
}

fn toggled_flags(settings: Settings) -> Settings {
    Settings {
        header_rows_frozen: Some(!settings.header_rows_frozen.unwrap_or(false)),
        header_columns_frozen: Some(!settings.header_columns_frozen.unwrap_or(false)),
        repeating_header_rows_enabled: Some(
            !settings.repeating_header_rows_enabled.unwrap_or(false),
        ),
        repeating_header_columns_enabled: Some(
            !settings.repeating_header_columns_enabled.unwrap_or(false),
        ),
        ..settings
    }
}

fn table_model(
    name: &str,
    settings: Settings,
    with_unknowns: bool,
    with_tracker: bool,
) -> TestResult<Vec<u8>> {
    let mut payload = tst::TableModelArchive {
        table_id: format!("table-{name}"),
        table_name: name.to_owned(),
        number_of_rows: 8,
        number_of_columns: 4,
        number_of_header_rows: settings.header_rows.map(|count| count.get() as u32),
        number_of_header_columns: settings.header_columns.map(|count| count.get() as u32),
        number_of_footer_rows: settings.footer_rows.map(|count| count.get() as u32),
        header_rows_frozen: settings.header_rows_frozen,
        header_columns_frozen: settings.header_columns_frozen,
        repeating_header_rows_enabled: settings.repeating_header_rows_enabled,
        repeating_header_columns_enabled: settings.repeating_header_columns_enabled,
        table_name_style: Some(reference(TITLE_STYLE)),
        table_name_shape_style: Some(reference(SHAPE_STYLE)),
        default_row_height: 20.0,
        default_column_width: 64.0,
        ..tst::TableModelArchive::default()
    }
    .encode_to_vec();

    if with_unknowns {
        append_varint_field(&mut payload, UNKNOWN_MODEL_FIELD, UNKNOWN_MODEL_VALUE)?;
    }
    if with_tracker {
        append_length_delimited_field(&mut payload, TABLE_MODEL_FIELD_TRACKER, TRACKER_BYTES)?;
    }
    Ok(payload)
}

fn paragraph_style_payload() -> Vec<u8> {
    tswp::ParagraphStyleArchive {
        super_: tss_style("slide-table-headers"),
        ..tswp::ParagraphStyleArchive::default()
    }
    .encode_to_vec()
}

fn shape_style_payload() -> Vec<u8> {
    tswp::ShapeStyleArchive {
        super_: tsd::ShapeStyleArchive {
            super_: tss_style("slide-table-headers-shape"),
            ..tsd::ShapeStyleArchive::default()
        },
        ..tswp::ShapeStyleArchive::default()
    }
    .encode_to_vec()
}

fn tss_style(identifier: &str) -> tss::StyleArchive {
    tss::StyleArchive {
        style_identifier: Some(identifier.to_owned()),
        ..tss::StyleArchive::default()
    }
}

fn shape_info_payload() -> Vec<u8> {
    tswp::ShapeInfoArchive {
        super_: tsd::ShapeArchive {
            super_: tsd::DrawableArchive::default(),
            ..tsd::ShapeArchive::default()
        },
        ..tswp::ShapeInfoArchive::default()
    }
    .encode_to_vec()
}

fn component_payload(
    identifier: u64,
    locator: &str,
    identifiers: impl IntoIterator<Item = u64>,
) -> Vec<u8> {
    tsp::ComponentInfo {
        identifier,
        preferred_locator: locator.to_owned(),
        locator: Some(locator.to_owned()),
        save_token: Some(1),
        object_uuid_map_entries: identifiers
            .into_iter()
            .map(|identifier| tsp::ObjectUuidMapEntry {
                identifier,
                uuid: tsp::Uuid {
                    lower: identifier + 10_000,
                    upper: identifier + 20_000,
                },
            })
            .collect(),
        ..tsp::ComponentInfo::default()
    }
    .encode_to_vec()
}

fn metadata_payload() -> TestResult<Vec<u8>> {
    let document_component = component_payload(
        1,
        "Document",
        [
            1,
            2,
            SLIDE_NODE,
            SLIDE,
            TABLE_INFOS[0],
            TABLE_INFOS[1],
            NON_TABLE_DRAWABLE,
        ],
    );
    let model_component = component_payload(
        2,
        "CalculationEngine",
        [MODELS[0], MODELS[1], TITLE_STYLE, SHAPE_STYLE],
    );
    let mut payload = Vec::new();
    append_varint_field(&mut payload, 1, METADATA_OBJECT)?;
    append_length_delimited_field(&mut payload, 3, &document_component)?;
    append_length_delimited_field(&mut payload, 3, &model_component)?;
    Ok(payload)
}

/// Build a strict rooted Keynote package with the table models in a separate
/// physical IWA member.  The shape drawable is deliberately interleaved in
/// z-order so a table position is not a raw drawable position.
fn synthetic_package(locked: bool, with_unknowns: bool, with_tracker: bool) -> TestResult<Vec<u8>> {
    let document = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..tsa::DocumentArchive::default()
        },
        show: reference(2),
        ..kn::DocumentArchive::default()
    };
    let show = kn::ShowArchive {
        theme: reference(80),
        slide_tree: kn::SlideTreeArchive {
            slides: vec![reference(SLIDE_NODE)],
            ..kn::SlideTreeArchive::default()
        },
        size: tsp::Size {
            width: 1_024.0,
            height: 768.0,
        },
        stylesheet: reference(81),
        ..kn::ShowArchive::default()
    };
    #[allow(deprecated, reason = "native schema retains cache fields")]
    let node = kn::SlideNodeArchive {
        slide: Some(reference(SLIDE)),
        is_skipped: false,
        has_builds: false,
        has_transition: false,
        ..kn::SlideNodeArchive::default()
    };

    let drawables = [TABLE_INFOS[0], NON_TABLE_DRAWABLE, TABLE_INFOS[1]];
    let slide = kn::SlideArchive {
        style: reference(90),
        transition: kn::TransitionArchive::default(),
        owned_drawables: drawables.iter().copied().map(reference).collect(),
        drawables_z_order: drawables.iter().copied().map(reference).collect(),
        name: Some("Tables".to_owned()),
        in_document: true,
        ..kn::SlideArchive::default()
    };
    let mut slide_object = object(SLIDE, SLIDE_MESSAGE_TYPE, slide.encode_to_vec(), &drawables)?;
    slide_object.archive_info.message_infos[0]
        .field_infos
        .extend([
            field_reference(vec![7], &drawables),
            field_reference(vec![42], &drawables),
        ]);

    let mut document_objects = vec![
        object(1, DOCUMENT_MESSAGE_TYPE, document.encode_to_vec(), &[2])?,
        object(
            2,
            SHOW_MESSAGE_TYPE,
            show.encode_to_vec(),
            &[SLIDE_NODE, 80, 81],
        )?,
        object(
            SLIDE_NODE,
            SLIDE_NODE_MESSAGE_TYPE,
            node.encode_to_vec(),
            &[SLIDE],
        )?,
        slide_object,
        object(
            NON_TABLE_DRAWABLE,
            SHAPE_INFO_MESSAGE_TYPE,
            shape_info_payload(),
            &[SLIDE],
        )?,
    ];
    let mut model_objects = Vec::new();

    for index in 0..TABLE_INFOS.len() {
        let info = tst::TableInfoArchive {
            super_: tsd::DrawableArchive {
                parent: Some(reference(SLIDE)),
                locked: Some(locked),
                ..tsd::DrawableArchive::default()
            },
            table_model: reference(MODELS[index]),
            ..tst::TableInfoArchive::default()
        };
        let mut info_object = object(
            TABLE_INFOS[index],
            TABLE_INFO_MESSAGE_TYPE,
            info.encode_to_vec(),
            // Native Keynote keeps the drawable parent only in the TableInfo
            // payload.  The strong ArchiveInfo edge is the model reference;
            // requiring the weak parent here would reject app-authored files.
            &[MODELS[index]],
        )?;
        info_object.archive_info.message_infos[0]
            .field_infos
            .push(field_reference(vec![2], &[MODELS[index]]));
        document_objects.push(info_object);

        let mut model_object = object(
            MODELS[index],
            TABLE_MODEL_MESSAGE_TYPE,
            table_model(
                if index == 0 { "Revenue" } else { "Costs" },
                fixture_settings(index),
                with_unknowns,
                with_tracker,
            )?,
            &[TITLE_STYLE, SHAPE_STYLE],
        )?;
        model_object.archive_info.message_infos[0]
            .field_infos
            .extend([
                field_reference(vec![30], &[TITLE_STYLE]),
                field_reference(vec![36], &[SHAPE_STYLE]),
            ]);
        model_objects.push(model_object);
    }
    model_objects.push(object(TITLE_STYLE, 2_022, paragraph_style_payload(), &[])?);
    model_objects.push(object(SHAPE_STYLE, 2_025, shape_style_payload(), &[])?);

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
    let metadata_component = SnappyStream::compress(
        &Archive {
            objects: vec![object(
                METADATA_OBJECT,
                METADATA_MESSAGE_TYPE,
                metadata_payload()?,
                &[],
            )?],
        }
        .to_bytes()?,
    )?;

    let mut members: Vec<(&str, &[u8])> = vec![
        (SENTINEL_MEMBER, b"untouched-header-sentinel"),
        (DOCUMENT_MEMBER, document_component.as_slice()),
        (MODEL_MEMBER, model_component.as_slice()),
        (METADATA_MEMBER, metadata_component.as_slice()),
    ];
    members.extend(
        PREVIEWS
            .into_iter()
            .map(|name| (name, b"header preview".as_slice())),
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

fn member_bytes(package: &[u8], name: &str) -> TestResult<Vec<u8>> {
    Ok(Catalog::from_bytes(package)?
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or_else(|| io::Error::other(format!("missing package member {name}")))?
        .data()
        .to_vec())
}

fn member_archive(package: &[u8], name: &str) -> TestResult<Archive> {
    let bytes = member_bytes(package, name)?;
    Ok(Archive::parse(
        SnappyStream::decompress(&bytes)?.as_bytes(),
    )?)
}

fn model_payload(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let archive = member_archive(package, MODEL_MEMBER)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing table model object"))?;
    Ok(object
        .messages
        .iter()
        .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("missing table-model message"))?
        .data
        .clone())
}

fn object_messages(package: &[u8], member: &str, identifier: u64) -> TestResult<Vec<RawMessage>> {
    Ok(member_archive(package, member)?
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing archive object"))?
        .messages
        .clone())
}

fn replace_member_archive(
    package: &[u8],
    name: &str,
    mutate: impl FnOnce(&mut Archive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let mut archive = member_archive(package, name)?;
    mutate(&mut archive)?;
    let replacement = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(&[EntryEdit::new(name, &replacement)], Limits::default())?)
}

fn replace_model_payload(package: &[u8], identifier: u64, payload: Vec<u8>) -> TestResult<Vec<u8>> {
    replace_member_archive(package, MODEL_MEMBER, |archive| {
        let object = archive
            .object_mut(identifier)
            .ok_or_else(|| io::Error::other("missing table model object"))?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing table-model message"))?;
        object.replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: TABLE_MODEL_MESSAGE_TYPE,
                data: payload,
            },
        )?;
        Ok(())
    })
}

fn append_model_raw(package: &[u8], raw: &[u8]) -> TestResult<Vec<u8>> {
    let payload = {
        let mut payload = model_payload(package, MODELS[0])?;
        payload.extend_from_slice(raw);
        payload
    };
    replace_model_payload(package, MODELS[0], payload)
}

fn append_model_message_field(
    package: &[u8],
    field_number: u32,
    payload: &[u8],
) -> TestResult<Vec<u8>> {
    let mut raw = Vec::new();
    append_length_delimited_field(&mut raw, field_number, payload)?;
    append_model_raw(package, &raw)
}

fn with_haunted_owner(package: &[u8]) -> TestResult<Vec<u8>> {
    let owner = tsce::HauntedOwnerArchive {
        owner_uid: tsp::Uuid {
            lower: 41,
            upper: 42,
        },
    }
    .encode_to_vec();
    append_model_message_field(package, 84, &owner)
}

fn with_header_name_manager(package: &[u8]) -> TestResult<Vec<u8>> {
    let package = replace_member_archive(package, DOCUMENT_MEMBER, |archive| {
        let root = archive
            .object_mut(1)
            .ok_or_else(|| io::Error::other("missing root document"))?;
        let message_index = root
            .messages
            .iter()
            .position(|message| message.type_ == DOCUMENT_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing root document message"))?;
        let mut document =
            kn::DocumentArchive::decode(root.messages[message_index].data.as_slice())?;
        document.super_.calculation_engine = Some(reference(CALCULATION_ENGINE));
        root.replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: DOCUMENT_MESSAGE_TYPE,
                data: document.encode_to_vec(),
            },
        )?;
        let info = &mut root.archive_info.message_infos[message_index];
        info.object_references.push(CALCULATION_ENGINE);
        info.field_infos
            .push(field_reference(vec![3, 4], &[CALCULATION_ENGINE]));
        Ok(())
    })?;

    replace_member_archive(&package, MODEL_MEMBER, |archive| {
        let engine_payload = tsce::CalculationEngineArchive {
            dependency_tracker: tsce::DependencyTrackerArchive::default(),
            header_name_manager: Some(reference(HEADER_NAME_MANAGER)),
            ..tsce::CalculationEngineArchive::default()
        }
        .encode_to_vec();
        let mut engine = object(
            CALCULATION_ENGINE,
            CALCULATION_ENGINE_MESSAGE_TYPE,
            engine_payload,
            &[HEADER_NAME_MANAGER],
        )?;
        engine.archive_info.message_infos[0]
            .field_infos
            .push(field_reference(vec![14], &[HEADER_NAME_MANAGER]));
        let manager_payload = tst::HeaderNameMgrArchive {
            owner_uid: tsp::Uuid {
                lower: 51,
                upper: 52,
            },
            ..tst::HeaderNameMgrArchive::default()
        }
        .encode_to_vec();
        archive.objects.extend([
            engine,
            object(
                HEADER_NAME_MANAGER,
                HEADER_NAME_MANAGER_MESSAGE_TYPE,
                manager_payload,
                &[],
            )?,
        ]);
        Ok(())
    })
}

fn with_malformed_category_owner(package: &[u8], missing_group_uid: bool) -> TestResult<Vec<u8>> {
    let mut group = Vec::new();
    if !missing_group_uid {
        append_length_delimited_field(
            &mut group,
            1,
            &tsp::Uuid {
                lower: 61,
                upper: 62,
            }
            .encode_to_vec(),
        )?;
    }
    append_varint_field(&mut group, 6, 0)?;
    let mut owner = Vec::new();
    if missing_group_uid {
        append_length_delimited_field(
            &mut owner,
            1,
            &tsp::Uuid {
                lower: 63,
                upper: 64,
            }
            .encode_to_vec(),
        )?;
    }
    append_length_delimited_field(&mut owner, 2, &group)?;
    append_model_message_field(package, 81, &owner)
}

fn with_invalid_group_owner_index(package: &[u8], duplicate: bool) -> TestResult<Vec<u8>> {
    let mut group = Vec::new();
    append_length_delimited_field(
        &mut group,
        1,
        &tsp::Uuid {
            lower: 65,
            upper: 66,
        }
        .encode_to_vec(),
    )?;
    append_varint_field(&mut group, 6, 0)?;
    append_varint_field(&mut group, 14, 8)?;
    append_varint_field(
        &mut group,
        14,
        if duplicate {
            9
        } else {
            u64::from(u32::MAX) + 1
        },
    )?;
    let mut owner = Vec::new();
    append_length_delimited_field(
        &mut owner,
        1,
        &tsp::Uuid {
            lower: 67,
            upper: 68,
        }
        .encode_to_vec(),
    )?;
    append_length_delimited_field(&mut owner, 2, &group)?;
    append_model_message_field(package, 81, &owner)
}

fn with_pivot_owner(package: &[u8]) -> TestResult<Vec<u8>> {
    replace_member_archive(package, MODEL_MEMBER, |archive| {
        let model = archive
            .object_mut(MODELS[0])
            .ok_or_else(|| io::Error::other("missing table model object"))?;
        let message_index = model
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing table-model message"))?;
        let mut payload = model.messages[message_index].data.clone();
        append_length_delimited_field(&mut payload, 85, &reference(PIVOT_OWNER).encode_to_vec())?;
        model.replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: TABLE_MODEL_MESSAGE_TYPE,
                data: payload,
            },
        )?;
        let info = &mut model.archive_info.message_infos[message_index];
        info.object_references.push(PIVOT_OWNER);
        info.field_infos
            .push(field_reference(vec![85], &[PIVOT_OWNER]));

        archive.objects.push(object(
            PIVOT_OWNER,
            PIVOT_OWNER_MESSAGE_TYPE,
            tst::PivotOwnerArchive {
                pivot_owner_uid: Some(tsp::Uuid {
                    lower: 69,
                    upper: 70,
                }),
                ..tst::PivotOwnerArchive::default()
            }
            .encode_to_vec(),
            &[],
        )?);
        Ok(())
    })
}

fn with_table_info_cache(package: &[u8]) -> TestResult<Vec<u8>> {
    replace_member_archive(package, DOCUMENT_MEMBER, |archive| {
        let info_object = archive
            .object_mut(TABLE_INFOS[0])
            .ok_or_else(|| io::Error::other("missing table-info object"))?;
        let message_index = info_object
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing table-info message"))?;
        let mut info =
            tst::TableInfoArchive::decode(info_object.messages[message_index].data.as_slice())?;
        info.summary_model = Some(reference(TABLE_INFO_CACHE));
        info_object.replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: TABLE_INFO_MESSAGE_TYPE,
                data: info.encode_to_vec(),
            },
        )?;
        let message_info = &mut info_object.archive_info.message_infos[message_index];
        message_info.object_references.push(TABLE_INFO_CACHE);
        archive
            .objects
            .push(object(TABLE_INFO_CACHE, 7_001, Vec::new(), &[])?);
        Ok(())
    })
}

fn replace_table_info(
    package: &[u8],
    identifier: u64,
    mutate: impl FnOnce(&mut tst::TableInfoArchive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    replace_member_archive(package, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(identifier)
            .ok_or_else(|| io::Error::other("missing table-info object"))?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing table-info message"))?;
        let mut info =
            tst::TableInfoArchive::decode(object.messages[message_index].data.as_slice())?;
        mutate(&mut info)?;
        object.replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: TABLE_INFO_MESSAGE_TYPE,
                data: info.encode_to_vec(),
            },
        )?;
        Ok(())
    })
}

fn replace_slide(
    package: &[u8],
    mutate: impl FnOnce(&mut kn::SlideArchive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    replace_member_archive(package, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(SLIDE)
            .ok_or_else(|| io::Error::other("missing slide object"))?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == SLIDE_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing slide message"))?;
        let mut slide = kn::SlideArchive::decode(object.messages[message_index].data.as_slice())?;
        mutate(&mut slide)?;
        object.replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: SLIDE_MESSAGE_TYPE,
                data: slide.encode_to_vec(),
            },
        )?;
        Ok(())
    })
}

fn set_model_message_type(package: &[u8], type_: u32) -> TestResult<Vec<u8>> {
    replace_member_archive(package, MODEL_MEMBER, |archive| {
        let object = archive
            .object_mut(MODELS[0])
            .ok_or_else(|| io::Error::other("missing table model object"))?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing table-model message"))?;
        let data = object.messages[message_index].data.clone();
        object.replace_message_preserving_header(message_index, RawMessage { type_, data })?;
        Ok(())
    })
}

fn duplicate_canonical_model_message(package: &[u8]) -> TestResult<Vec<u8>> {
    replace_member_archive(package, MODEL_MEMBER, |archive| {
        let object = archive
            .object_mut(MODELS[0])
            .ok_or_else(|| io::Error::other("missing table model object"))?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing table-model message"))?;
        let duplicate_message = object.messages[message_index].clone();
        let duplicate_info = object.archive_info.message_infos[message_index].clone();
        object.messages.push(duplicate_message);
        object.archive_info.message_infos.push(duplicate_info);
        Ok(())
    })
}

fn add_foreign_inbound(package: &[u8]) -> TestResult<Vec<u8>> {
    replace_member_archive(package, MODEL_MEMBER, |archive| {
        archive.objects.push(object(
            FOREIGN_INBOUND_OBJECT,
            7_000,
            Vec::new(),
            &[MODELS[0]],
        )?);
        Ok(())
    })
}

fn duplicate_model_identity(package: &[u8]) -> TestResult<Vec<u8>> {
    let duplicate = member_archive(package, MODEL_MEMBER)?
        .object(MODELS[0])
        .ok_or_else(|| io::Error::other("missing first model object"))?
        .clone();
    replace_member_archive(package, DOCUMENT_MEMBER, |archive| {
        archive.objects.push(duplicate);
        Ok(())
    })
}

fn archive_info_missing_model(package: &[u8]) -> TestResult<Vec<u8>> {
    replace_member_archive(package, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(TABLE_INFOS[0])
            .ok_or_else(|| io::Error::other("missing table-info object"))?;
        object.archive_info.message_infos[0]
            .object_references
            .clear();
        object.archive_info.message_infos[0].field_infos.clear();
        Ok(())
    })
}

fn field_info_wrong_path(package: &[u8]) -> TestResult<Vec<u8>> {
    replace_member_archive(package, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(TABLE_INFOS[0])
            .ok_or_else(|| io::Error::other("missing table-info object"))?;
        object.archive_info.message_infos[0].field_infos[0].path = FieldPath::new(vec![99]);
        Ok(())
    })
}

fn unknown_model_value(payload: &[u8]) -> TestResult<u64> {
    let field = WireView::parse(payload)?
        .fields()
        .find(|field| field.number() == UNKNOWN_MODEL_FIELD)
        .ok_or_else(|| io::Error::other("missing unknown model field"))?;
    Ok(decode_varint_from_bytes(field.payload())?.0)
}

fn model_field_payloads(package: &[u8], field: u32) -> TestResult<Vec<Vec<u8>>> {
    Ok(
        repeated_length_delimited_payloads(&model_payload(package, MODELS[0])?, field)?
            .into_iter()
            .map(ToOwned::to_owned)
            .collect(),
    )
}

fn field_numbers(package: &[u8], field: u32) -> TestResult<usize> {
    Ok(WireView::parse(&model_payload(package, MODELS[0])?)?
        .fields()
        .filter(|candidate| candidate.number() == field)
        .count())
}

fn assert_rejected_atomically(source: &[u8]) -> TestResult {
    let Ok(package) = Package::from_bytes(source) else {
        return Ok(());
    };
    let before = exact_bytes(&package)?;
    assert!(
        package
            .slide_table_header_settings("Tables", TableSelector::index(0))
            .is_err()
    );
    assert!(
        package
            .edit_slide_table_headers("Tables", TableSelector::index(0))
            .is_err()
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn transaction_values_are_typed_and_archive_free() {
    fn assert_send_sync_debug<T: Send + Sync + std::fmt::Debug>() {}

    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<Count>();
    assert_send_sync_debug::<Settings>();
    assert_send_sync_debug::<SlideTableHeaderEdit<'static>>();
    assert_send_sync_debug::<SlideTableHeaderPatch>();
    assert_send_sync_debug::<SlideTableHeaderCommit>();
    assert_send_sync_debug::<SlideTableHeaderDiagnostics>();
    assert_send_sync_debug::<SlideTableHeaderError>();
    assert_send_sync_debug::<SlideTableHeaderLimitKind>();
    assert_send_sync_debug::<SlideTableHeaderPath>();
}

#[test]
fn selectors_read_all_seven_fields_and_positions_count_only_tables() -> TestResult {
    let source = synthetic_package(false, false, false)?;
    let package = Package::from_bytes(&source)?;

    let first = package.slide_table_header_settings("Tables", TableSelector::index(0))?;
    assert_eq!(first, fixture_settings(0));
    assert_eq!(first.header_row_count(), 2);
    assert_eq!(first.header_column_count(), 1);
    assert_eq!(first.footer_row_count(), 3);
    assert!(!first.header_rows_are_frozen());
    assert!(first.header_columns_are_frozen());
    assert!(first.repeats_header_rows());
    assert!(!first.repeats_header_columns());

    assert_eq!(
        package.slide_table_header_settings(SlideSelector::index(0), 1usize)?,
        fixture_settings(1)
    );
    assert_eq!(
        package.slide_table_header_settings("Tables", 0usize)?,
        package.slide_table_header_settings(SlideSelector::index(0), 0usize)?,
    );
    assert!(matches!(
        package.slide_table_header_settings("Tables", TableSelector::index(2)),
        Err(SlideTableHeaderError::TablePositionNotFound { .. })
    ));

    let edit = package.edit_slide_table_headers("Tables", 0usize)?;
    assert_eq!(edit.before(), first);
    assert_eq!(edit.settings(), first);
    assert_eq!(
        edit.path(),
        SlideTableHeaderPath::Table {
            slide: litchi_core::Position::new(0),
            table: litchi_core::Position::new(0),
        }
    );
    Ok(())
}

#[test]
fn exact_noop_and_default_presence_are_source_bound() -> TestResult {
    let source = synthetic_package(false, false, true)?;
    let package = Package::from_bytes(&source)?;
    let before = package.slide_table_header_settings("Tables", 0usize)?;

    let noop = package
        .edit_slide_table_headers("Tables", 0usize)?
        .set(before)
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(noop.patch().before(), before);
    assert_eq!(noop.patch().after(), before);
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert_eq!(noop.diagnostics().deleted_previews(), 0);
    assert!(!noop.diagnostics().full_reparse_performed());
    assert_eq!(exact_bytes(noop.package())?, source);
    assert_eq!(
        exact_bytes(package.apply_slide_table_headers(noop.patch())?.package())?,
        source
    );

    // Settings::default() means all seven native fields are absent.  It is
    // the lossless clear/default operation exposed by the frozen API.
    let cleared = package
        .edit_slide_table_headers("Tables", 0usize)?
        .set(Settings::default())
        .commit()?;
    assert_eq!(
        cleared
            .package()
            .slide_table_header_settings("Tables", 0usize)?,
        Settings::default()
    );
    for field in [
        TABLE_MODEL_FIELD_HEADER_ROWS,
        TABLE_MODEL_FIELD_HEADER_COLUMNS,
        TABLE_MODEL_FIELD_FOOTER_ROWS,
        TABLE_MODEL_FIELD_HEADER_ROWS_FROZEN,
        TABLE_MODEL_FIELD_HEADER_COLUMNS_FROZEN,
        TABLE_MODEL_FIELD_REPEAT_ROWS,
        TABLE_MODEL_FIELD_REPEAT_COLUMNS,
    ] {
        assert_eq!(field_numbers(&exact_bytes(cleared.package())?, field)?, 0);
    }
    assert_eq!(
        model_field_payloads(&source, TABLE_MODEL_FIELD_TRACKER)?,
        model_field_payloads(&exact_bytes(cleared.package())?, TABLE_MODEL_FIELD_TRACKER)?,
    );
    assert_eq!(
        exact_bytes(
            cleared
                .package()
                .apply_slide_table_headers(&cleared.patch().inverse())?
                .package(),
        )?,
        source
    );
    Ok(())
}

#[test]
fn changed_settings_rewrite_one_cross_member_model_and_preserve_locality() -> TestResult {
    let source = synthetic_package(false, true, true)?;
    let package = Package::from_bytes(&source)?;
    let before_first = model_payload(&source, MODELS[0])?;
    let before_second = model_payload(&source, MODELS[1])?;
    let before_document = member_bytes(&source, DOCUMENT_MEMBER)?;
    let before_metadata = member_bytes(&source, METADATA_MEMBER)?;
    let before_sentinel = member_bytes(&source, SENTINEL_MEMBER)?;
    let before_tracker = model_field_payloads(&source, TABLE_MODEL_FIELD_TRACKER)?;
    let before_title_style = object_messages(&source, MODEL_MEMBER, TITLE_STYLE)?;
    let before_shape_style = object_messages(&source, MODEL_MEMBER, SHAPE_STYLE)?;

    let desired = changed_settings();
    let commit = package
        .edit_slide_table_headers("Tables", 0usize)?
        .set(desired)
        .commit()?;
    assert_eq!(
        commit
            .package()
            .slide_table_header_settings("Tables", 0usize)?,
        desired
    );
    assert_eq!(commit.patch().before(), fixture_settings(0));
    assert_eq!(commit.patch().after(), desired);
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    assert!(commit.diagnostics().full_reparse_performed());

    let target = exact_bytes(commit.package())?;
    assert_ne!(target, source);
    assert_eq!(member_bytes(&target, DOCUMENT_MEMBER)?, before_document);
    assert_eq!(member_bytes(&target, METADATA_MEMBER)?, before_metadata);
    assert_eq!(member_bytes(&target, SENTINEL_MEMBER)?, before_sentinel);
    assert_eq!(
        PREVIEWS
            .iter()
            .map(|name| member_bytes(&source, name))
            .collect::<TestResult<Vec<_>>>()?,
        PREVIEWS
            .iter()
            .map(|name| member_bytes(&target, name))
            .collect::<TestResult<Vec<_>>>()?,
    );
    assert_ne!(
        member_bytes(&target, MODEL_MEMBER)?,
        member_bytes(&source, MODEL_MEMBER)?
    );

    let after_first = model_payload(&target, MODELS[0])?;
    assert_ne!(after_first, before_first);
    assert_eq!(model_payload(&target, MODELS[1])?, before_second);
    assert_eq!(
        object_messages(&target, MODEL_MEMBER, TITLE_STYLE)?,
        before_title_style
    );
    assert_eq!(
        object_messages(&target, MODEL_MEMBER, SHAPE_STYLE)?,
        before_shape_style
    );
    assert_eq!(
        model_field_payloads(&target, TABLE_MODEL_FIELD_TRACKER)?,
        before_tracker
    );
    assert_eq!(unknown_model_value(&after_first)?, UNKNOWN_MODEL_VALUE);
    Ok(())
}

#[test]
fn apply_inverse_and_conflict_reopen_exactly() -> TestResult {
    let source = synthetic_package(false, false, true)?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_slide_table_headers(0usize, 0usize)?
        .set(changed_settings())
        .commit()?;
    let target = exact_bytes(commit.package())?;

    let applied = package.apply_slide_table_headers(commit.patch())?;
    assert_eq!(exact_bytes(applied.package())?, target);
    assert!(matches!(
        commit.package().apply_slide_table_headers(commit.patch()),
        Err(SlideTableHeaderError::PatchConflict)
    ));

    let inverse = commit.patch().inverse();
    assert_eq!(inverse.inverse(), *commit.patch());
    let restored = commit.package().apply_slide_table_headers(&inverse)?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_eq!(
        restored
            .package()
            .slide_table_header_settings("Tables", 0usize)?,
        fixture_settings(0)
    );

    let reopened = Package::from_bytes(&target)?;
    assert_eq!(
        reopened.slide_table_header_settings("Tables", 0usize)?,
        changed_settings()
    );
    Ok(())
}

#[test]
fn count_and_dimension_validation_refuse_invalid_edits_atomically() -> TestResult {
    assert!(Count::new(0).is_err());
    assert_eq!(Count::new(1)?, Count::ONE);
    assert_eq!(Count::new(5)?, Count::FIVE);
    assert!(Count::new(6).is_err());
    assert!(Count::new(usize::MAX).is_err());

    let source = synthetic_package(false, false, false)?;
    let package = Package::from_bytes(&source)?;
    let row_invalid = Settings {
        header_rows: Some(Count::FIVE),
        footer_rows: Some(Count::FIVE),
        ..Settings::default()
    };
    let column_invalid = Settings {
        header_columns: Some(Count::FIVE),
        ..Settings::default()
    };
    for desired in [row_invalid, column_invalid] {
        let before = exact_bytes(&package)?;
        assert!(
            package
                .edit_slide_table_headers("Tables", 0usize)?
                .set(desired)
                .commit()
                .is_err()
        );
        assert_eq!(exact_bytes(&package)?, before);
    }
    Ok(())
}

#[test]
fn locked_table_refuses_changed_settings_but_allows_exact_noop() -> TestResult {
    let source = synthetic_package(true, false, false)?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.slide_table_header_settings("Tables", 0usize)?,
        fixture_settings(0)
    );

    let noop = package
        .edit_slide_table_headers("Tables", 0usize)?
        .set(fixture_settings(0))
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(exact_bytes(noop.package())?, source);
    assert!(matches!(
        package
            .edit_slide_table_headers("Tables", 0usize)?
            .set(changed_settings())
            .commit(),
        Err(SlideTableHeaderError::Locked)
    ));
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn coordinate_dependency_graphs_block_counts_but_allow_safe_flags() -> TestResult {
    let base = synthetic_package(false, false, false)?;
    for source in [
        with_haunted_owner(&base)?,
        with_header_name_manager(&base)?,
        with_table_info_cache(&base)?,
    ] {
        let package = Package::from_bytes(&source)?;
        let before = exact_bytes(&package)?;
        assert!(matches!(
            package
                .edit_slide_table_headers("Tables", 0usize)?
                .set(changed_settings())
                .commit(),
            Err(SlideTableHeaderError::UnsupportedDependency)
        ));
        assert_eq!(exact_bytes(&package)?, before);

        let flags = toggled_flags(fixture_settings(0));
        let commit = package
            .edit_slide_table_headers("Tables", 0usize)?
            .set(flags)
            .commit()?;
        assert_eq!(
            commit
                .package()
                .slide_table_header_settings("Tables", 0usize)?,
            flags
        );
        assert_eq!(commit.patch().inverse().after(), fixture_settings(0));
    }
    Ok(())
}

#[test]
fn pivot_owner_blocks_even_flag_only_edits() -> TestResult {
    let source = with_pivot_owner(&synthetic_package(false, false, false)?)?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    assert_eq!(
        package.slide_table_header_settings("Tables", 0usize)?,
        fixture_settings(0)
    );
    assert!(matches!(
        package
            .edit_slide_table_headers("Tables", 0usize)?
            .set(toggled_flags(fixture_settings(0)))
            .commit(),
        Err(SlideTableHeaderError::UnsupportedDependency)
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn malformed_inactive_category_routes_fail_closed() -> TestResult {
    let source = synthetic_package(false, false, false)?;
    for source in [
        with_malformed_category_owner(&source, false)?,
        with_malformed_category_owner(&source, true)?,
        with_invalid_group_owner_index(&source, true)?,
        with_invalid_group_owner_index(&source, false)?,
    ] {
        let package = Package::from_bytes(&source)?;
        let before = exact_bytes(&package)?;
        assert_eq!(
            package.slide_table_header_settings("Tables", 0usize)?,
            fixture_settings(0)
        );
        assert!(
            package
                .edit_slide_table_headers("Tables", 0usize)?
                .set(changed_settings())
                .commit()
                .is_err()
        );
        assert_eq!(exact_bytes(&package)?, before);
    }
    Ok(())
}

#[test]
fn malformed_model_wire_duplicate_alias_and_legacy_routes_fail_closed() -> TestResult {
    let source = synthetic_package(false, false, false)?;
    for raw in [
        vec![0x48, 0],                // header rows = zero
        vec![0x50, 7],                // columns exceed the model
        vec![0x58, 0],                // footer rows = zero
        vec![0x60, 2],                // invalid Boolean scalar
        vec![0x48, 1, 0x48, 2],       // duplicate known field
        vec![0x4d, 0, 0, 0, 0],       // known field, wrong wire type
        vec![0x48, 0x80, 0],          // known scalar, non-canonical varint
        vec![0x80, 0x02, 0x01],       // repeat-columns field (32) is valid
        vec![0x80, 0x02, 0x80, 0x00], // repeated known field, non-canonical
    ] {
        assert_rejected_atomically(&append_model_raw(&source, &raw)?)?;
    }
    assert_rejected_atomically(&duplicate_canonical_model_message(&source)?)?;
    assert_rejected_atomically(&set_model_message_type(&source, TABLE_INFO_MESSAGE_TYPE)?)?;

    // A canonical model carrying the legacy alias is rejected rather than
    // allowing a malformed/ambiguous source to fall through.
    let aliased = replace_member_archive(&source, MODEL_MEMBER, |archive| {
        let object = archive
            .object_mut(MODELS[0])
            .ok_or_else(|| io::Error::other("missing table model object"))?;
        let info = object.archive_info.message_infos[0].clone();
        object.messages.push(RawMessage {
            type_: TABLE_INFO_MESSAGE_TYPE,
            data: object.messages[0].data.clone(),
        });
        let mut alias_info = info;
        alias_info.type_ = TABLE_INFO_MESSAGE_TYPE;
        object.archive_info.message_infos.push(alias_info);
        Ok(())
    })?;
    assert_rejected_atomically(&aliased)?;
    Ok(())
}

#[test]
fn duplicate_missing_and_noncanonical_graph_routes_fail_closed() -> TestResult {
    let source = synthetic_package(false, false, false)?;
    assert_rejected_atomically(&replace_slide(&source, |slide| {
        slide.drawables_z_order.push(reference(TABLE_INFOS[0]));
        Ok(())
    })?)?;
    assert_rejected_atomically(&replace_table_info(&source, TABLE_INFOS[0], |info| {
        info.table_model = reference(999_999);
        Ok(())
    })?)?;
    assert_rejected_atomically(&duplicate_model_identity(&source)?)?;
    assert_rejected_atomically(&archive_info_missing_model(&source)?)?;
    assert_rejected_atomically(&field_info_wrong_path(&source)?)?;
    assert_rejected_atomically(&add_foreign_inbound(&source)?)?;
    Ok(())
}

#[test]
fn finite_resource_limits_reject_before_publication() -> TestResult {
    let source = synthetic_package(false, false, false)?;
    let mut rejected = false;
    // Semantic reference admission is lazy, so the package remains a valid
    // read snapshot even when the transaction's six slide drawable edges are
    // one over the configured budget.  Every rejected attempt must leave the
    // immutable source exact.
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
            .edit_slide_table_headers(0usize, 0usize)
            .and_then(|edit| edit.set(changed_settings()).commit());
        if matches!(result, Err(SlideTableHeaderError::LimitExceeded { .. })) {
            assert_eq!(exact_bytes(&package)?, before);
            rejected = true;
            break;
        }
    }
    assert!(
        rejected,
        "a finite ingress-compatible budget must reject the edit"
    );
    Ok(())
}
