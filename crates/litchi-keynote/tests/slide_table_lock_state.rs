//! Integration coverage for selector-first Keynote slide-table lock state.
//!
//! The fixture keeps the presentation graph in `Index/Document.iwa` and the
//! table models in a separate calculation component.  Table positions are
//! counted from the slide's z-order table drawables, with a shape interleaved
//! to ensure that native object identifiers never become public selectors.

use std::error::Error as StdError;
use std::io;

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::{
    decode_varint_from_bytes,
    wire::{WireView, append_length_delimited_field, append_varint_field},
};
use litchi_iwa_core::{
    Archive, ArchiveObject, FieldInfo, FieldPath, FieldType, RawMessage, SnappyStream,
};
use litchi_iwa_protos::{kn, tsa, tsd, tsk, tsp, tst, tswp};
use litchi_keynote::slide::table::lock::State;
use litchi_keynote::{
    Package, ReadOptions, SemanticLimits, SlideSelector, SlideTableLockStateCommit,
    SlideTableLockStateDiagnostics, SlideTableLockStateEdit, SlideTableLockStateError,
    SlideTableLockStateLimitKind, SlideTableLockStatePatch, SlideTableLockStatePath, TableSelector,
};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const MODEL_MEMBER: &str = "Index/CalculationEngine.iwa";
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const SENTINEL_MEMBER: &str = "Data/lock-sentinel.bin";
const PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

const SLIDE_NODE: u64 = 3;
const SLIDE: u64 = 4;
const TABLE_INFOS: [u64; 3] = [100, 101, 102];
const MODELS: [u64; 3] = [110, 111, 112];
const NON_TABLE_DRAWABLE: u64 = 130;
const METADATA_OBJECT: u64 = 900;

const DOCUMENT_MESSAGE_TYPE: u32 = 1;
const SHOW_MESSAGE_TYPE: u32 = 2;
const SLIDE_NODE_MESSAGE_TYPE: u32 = 4;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;
const METADATA_MESSAGE_TYPE: u32 = 11_006;

const TABLE_SUPER_FIELD: u32 = 1;
const TABLE_MODEL_FIELD: u32 = 2;
const DRAWABLE_PARENT_FIELD: u32 = 2;
const DRAWABLE_LOCKED_FIELD: u32 = 5;
const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const SLIDE_Z_ORDER_FIELD: u32 = 42;
const UNKNOWN_TABLE_INFO_FIELD: u32 = 99;
const UNKNOWN_TABLE_INFO_VALUE: u64 = 0xfeed_beef;

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

fn table_model(name: &str) -> Vec<u8> {
    tst::TableModelArchive {
        table_id: format!("table-{name}"),
        table_name: name.to_owned(),
        number_of_rows: 8,
        number_of_columns: 4,
        ..tst::TableModelArchive::default()
    }
    .encode_to_vec()
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
            TABLE_INFOS[2],
            NON_TABLE_DRAWABLE,
        ],
    );
    let model_component = component_payload(2, "CalculationEngine", MODELS);
    let mut payload = Vec::new();
    append_varint_field(&mut payload, 1, METADATA_OBJECT)?;
    append_length_delimited_field(&mut payload, 3, &document_component)?;
    append_length_delimited_field(&mut payload, 3, &model_component)?;
    Ok(payload)
}

fn synthetic_package(states: [Option<bool>; 3]) -> TestResult<Vec<u8>> {
    let table_drawables = [
        TABLE_INFOS[0],
        NON_TABLE_DRAWABLE,
        TABLE_INFOS[1],
        TABLE_INFOS[2],
    ];
    let document = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..tsa::DocumentArchive::default()
        },
        show: reference(2),
        ..kn::DocumentArchive::default()
    };
    let show = kn::ShowArchive {
        slide_tree: kn::SlideTreeArchive {
            slides: vec![reference(SLIDE_NODE)],
            ..kn::SlideTreeArchive::default()
        },
        size: tsp::Size {
            width: 1_024.0,
            height: 768.0,
        },
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
    let slide = kn::SlideArchive {
        style: reference(90),
        owned_drawables: table_drawables.iter().copied().map(reference).collect(),
        drawables_z_order: table_drawables.iter().copied().map(reference).collect(),
        name: Some("Tables".to_owned()),
        in_document: true,
        ..kn::SlideArchive::default()
    };

    let mut slide_object = object(
        SLIDE,
        SLIDE_MESSAGE_TYPE,
        slide.encode_to_vec(),
        &table_drawables,
    )?;
    slide_object.archive_info.message_infos[0]
        .field_infos
        .extend([
            field_reference(vec![SLIDE_OWNED_DRAWABLES_FIELD], &table_drawables),
            field_reference(vec![SLIDE_Z_ORDER_FIELD], &table_drawables),
        ]);

    let mut document_objects = vec![
        object(1, DOCUMENT_MESSAGE_TYPE, document.encode_to_vec(), &[2])?,
        object(2, SHOW_MESSAGE_TYPE, show.encode_to_vec(), &[SLIDE_NODE])?,
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

    for (index, table_info_identifier) in TABLE_INFOS.iter().copied().enumerate() {
        let info = tst::TableInfoArchive {
            super_: tsd::DrawableArchive {
                parent: Some(reference(SLIDE)),
                locked: states[index],
                ..tsd::DrawableArchive::default()
            },
            table_model: reference(MODELS[index]),
            ..tst::TableInfoArchive::default()
        };
        let mut info_payload = info.encode_to_vec();
        append_varint_field(
            &mut info_payload,
            UNKNOWN_TABLE_INFO_FIELD,
            UNKNOWN_TABLE_INFO_VALUE,
        )?;
        let mut info_object = object(
            table_info_identifier,
            TABLE_INFO_MESSAGE_TYPE,
            info_payload,
            &[SLIDE, MODELS[index]],
        )?;
        info_object.archive_info.message_infos[0]
            .field_infos
            .extend([
                field_reference(vec![TABLE_SUPER_FIELD, DRAWABLE_PARENT_FIELD], &[SLIDE]),
                field_reference(vec![TABLE_MODEL_FIELD], &[MODELS[index]]),
            ]);
        document_objects.push(info_object);
    }

    let model_objects = MODELS
        .iter()
        .enumerate()
        .map(|(index, identifier)| {
            object(
                *identifier,
                TABLE_MODEL_MESSAGE_TYPE,
                table_model(["Revenue", "Costs", "Profit"][index]),
                &[],
            )
        })
        .collect::<TestResult<Vec<_>>>()?;

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
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            (SENTINEL_MEMBER, b"unrelated ZIP sentinel".as_slice()),
            (DOCUMENT_MEMBER, document_component.as_slice()),
            (MODEL_MEMBER, model_component.as_slice()),
            (METADATA_MEMBER, metadata_component.as_slice()),
            (PREVIEWS[0], b"large preview".as_slice()),
            (PREVIEWS[1], b"micro preview".as_slice()),
            (PREVIEWS[2], b"web preview".as_slice()),
        ],
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
        .ok_or_else(|| io::Error::other("missing package member"))?
        .data()
        .to_vec())
}

fn archive(package: &[u8], member: &str) -> TestResult<Archive> {
    let bytes = member_bytes(package, member)?;
    Ok(Archive::parse(
        SnappyStream::decompress(&bytes)?.as_bytes(),
    )?)
}

fn rewrite_member_archive(
    package: &[u8],
    member: &str,
    mutate: impl FnOnce(&mut Archive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let mut archive = archive(package, member)?;
    mutate(&mut archive)?;
    let replacement = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(&[EntryEdit::new(member, &replacement)], Limits::default())?)
}

fn rewrite_document_archive(
    package: &[u8],
    mutate: impl FnOnce(&mut Archive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    rewrite_member_archive(package, DOCUMENT_MEMBER, mutate)
}

fn table_info_payload(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let archive = archive(package, DOCUMENT_MEMBER)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing table-info object"))?;
    object
        .messages
        .iter()
        .find(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("missing table-info message").into())
}

fn metadata_payload_bytes(package: &[u8]) -> TestResult<Vec<u8>> {
    let archive = archive(package, METADATA_MEMBER)?;
    let object = archive
        .object(METADATA_OBJECT)
        .ok_or_else(|| io::Error::other("missing package metadata object"))?;
    object
        .messages
        .iter()
        .find(|message| message.type_ == METADATA_MESSAGE_TYPE)
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("missing package metadata message").into())
}

fn replace_metadata_payload(package: &[u8], payload: Vec<u8>) -> TestResult<Vec<u8>> {
    rewrite_member_archive(package, METADATA_MEMBER, |archive| {
        let object = archive
            .object_mut(METADATA_OBJECT)
            .ok_or_else(|| io::Error::other("missing package metadata object"))?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == METADATA_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing package metadata message"))?;
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: METADATA_MESSAGE_TYPE,
                data: payload,
            },
        )?;
        Ok(())
    })
}

fn unknown_package_metadata(package: &[u8]) -> TestResult<Vec<u8>> {
    let mut payload = metadata_payload_bytes(package)?;
    append_varint_field(&mut payload, 99, UNKNOWN_TABLE_INFO_VALUE)?;
    replace_metadata_payload(package, payload)
}

fn replace_table_info_payload(
    package: &[u8],
    identifier: u64,
    payload: Vec<u8>,
) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let object = archive
            .object_mut(identifier)
            .ok_or_else(|| io::Error::other("missing table-info object"))?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing table-info message"))?;
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: TABLE_INFO_MESSAGE_TYPE,
                data: payload,
            },
        )?;
        Ok(())
    })
}

fn rewrite_table_info_super(
    package: &[u8],
    identifier: u64,
    mutate: impl FnOnce(&mut Vec<u8>) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let payload = table_info_payload(package, identifier)?;
    let fields = WireView::parse(&payload)?;
    let super_field = fields
        .fields()
        .find(|field| field.number() == TABLE_SUPER_FIELD)
        .ok_or_else(|| io::Error::other("missing table-info super field"))?;
    let mut super_payload = super_field.payload().to_vec();
    mutate(&mut super_payload)?;
    let mut rewritten = Vec::with_capacity(payload.len() + super_payload.len());
    for field in fields.fields() {
        if field.number() == TABLE_SUPER_FIELD {
            append_length_delimited_field(&mut rewritten, TABLE_SUPER_FIELD, &super_payload)?;
        } else {
            rewritten.extend_from_slice(field.raw());
        }
    }
    replace_table_info_payload(package, identifier, rewritten)
}

fn lock_field(package: &[u8], identifier: u64) -> TestResult<Option<bool>> {
    let payload = table_info_payload(package, identifier)?;
    let outer = WireView::parse(&payload)?;
    let super_field = outer
        .fields()
        .find(|field| field.number() == TABLE_SUPER_FIELD)
        .ok_or_else(|| io::Error::other("missing table-info super field"))?;
    let drawable = WireView::parse(super_field.payload())?;
    let Some(field) = drawable
        .fields()
        .find(|field| field.number() == DRAWABLE_LOCKED_FIELD)
    else {
        return Ok(None);
    };
    let (value, width) = decode_varint_from_bytes(field.payload())?;
    if width != field.payload().len() {
        return Err(io::Error::other("non-canonical lock field").into());
    }
    Ok(Some(match value {
        0 => false,
        1 => true,
        _ => return Err(io::Error::other("non-boolean lock field").into()),
    }))
}

fn has_unknown_table_info_field(package: &[u8], identifier: u64) -> TestResult<bool> {
    Ok(WireView::parse(&table_info_payload(package, identifier)?)?
        .fields()
        .any(|field| {
            field.number() == UNKNOWN_TABLE_INFO_FIELD
                && decode_varint_from_bytes(field.payload()).is_ok_and(|(value, width)| {
                    value == UNKNOWN_TABLE_INFO_VALUE && width == field.payload().len()
                })
        }))
}

fn assert_locality(source: &[u8], target: &[u8]) -> TestResult {
    let before = Catalog::from_bytes(source)?;
    let after = Catalog::from_bytes(target)?;
    let changed = before
        .iter()
        .filter(|entry| {
            after
                .iter()
                .find(|candidate| candidate.name() == entry.name())
                .is_some_and(|candidate| candidate.data() != entry.data())
        })
        .map(|entry| entry.name().to_owned())
        .collect::<Vec<_>>();
    assert_eq!(changed, vec![DOCUMENT_MEMBER.to_owned()]);
    for entry in before.iter() {
        let candidate = after
            .iter()
            .find(|value| value.name() == entry.name())
            .ok_or_else(|| io::Error::other("candidate removed a package member"))?;
        if entry.name() != DOCUMENT_MEMBER {
            assert_eq!(entry.data(), candidate.data(), "unselected member changed");
            assert_eq!(
                entry.raw_record().local_record(),
                candidate.raw_record().local_record(),
                "unselected local ZIP record changed"
            );
        }
    }
    assert_eq!(before.len(), after.len());
    Ok(())
}

fn assert_rejected_atomically(source: &[u8]) -> TestResult {
    let Ok(package) = Package::from_bytes(source) else {
        return Ok(());
    };
    let before = exact_bytes(&package)?;
    assert!(
        package
            .slide_table_lock_state(SlideSelector::index(0), TableSelector::index(0))
            .is_err()
    );
    assert!(
        package
            .edit_slide_table_lock_state(SlideSelector::index(0), TableSelector::index(0))
            .is_err()
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

fn duplicate_info_message(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let object = archive
            .object_mut(TABLE_INFOS[0])
            .ok_or_else(|| io::Error::other("missing table-info object"))?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing table-info message"))?;
        let duplicate_message = object.messages[index].clone();
        let duplicate_info = object.archive_info.message_infos[index].clone();
        object.messages.push(duplicate_message);
        object.archive_info.message_infos.push(duplicate_info);
        Ok(())
    })
}

fn info_role_alias(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let object = archive
            .object_mut(TABLE_INFOS[0])
            .ok_or_else(|| io::Error::other("missing table-info object"))?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing table-info message"))?;
        object.messages[index].type_ = TABLE_MODEL_MESSAGE_TYPE;
        object.archive_info.message_infos[index].type_ = TABLE_MODEL_MESSAGE_TYPE;
        Ok(())
    })
}

fn info_model_role_alias(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let object = archive
            .object_mut(TABLE_INFOS[0])
            .ok_or_else(|| io::Error::other("missing table-info object"))?;
        let model = table_model("alias");
        let mut info = object.archive_info.message_infos[0].clone();
        info.type_ = TABLE_MODEL_MESSAGE_TYPE;
        info.length = u32::try_from(model.len())?;
        object.messages.push(RawMessage {
            type_: TABLE_MODEL_MESSAGE_TYPE,
            data: model,
        });
        object.archive_info.message_infos.push(info);
        Ok(())
    })
}

fn duplicate_model_identifier(package: &[u8]) -> TestResult<Vec<u8>> {
    let duplicate = archive(package, MODEL_MEMBER)?
        .object(MODELS[0])
        .ok_or_else(|| io::Error::other("missing model object"))?
        .clone();
    rewrite_document_archive(package, |archive| {
        archive.objects.push(duplicate);
        Ok(())
    })
}

fn foreign_inbound(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        archive
            .objects
            .push(object(902, 7_000, vec![0x08, 0x01], &[TABLE_INFOS[0]])?);
        Ok(())
    })
}

fn foreign_data_inbound(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let mut foreign = object(903, 7_001, vec![0x08, 0x01], &[])?;
        foreign.archive_info.message_infos[0]
            .data_references
            .push(TABLE_INFOS[0]);
        archive.objects.push(foreign);
        Ok(())
    })
}

fn foreign_field_data_inbound(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let mut foreign = object(904, 7_002, vec![0x08, 0x01], &[])?;
        let mut field = FieldInfo::new(vec![77]);
        field.r#type = Some(FieldType::DataReference);
        field.data_references.push(TABLE_INFOS[0]);
        foreign.archive_info.message_infos[0]
            .field_infos
            .push(field);
        archive.objects.push(foreign);
        Ok(())
    })
}

fn foreign_field_object_inbound(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let mut foreign = object(905, 7_003, vec![0x08, 0x01], &[])?;
        let mut field = FieldInfo::new(vec![78]);
        field.r#type = Some(FieldType::ObjectReference);
        field.object_references.push(TABLE_INFOS[0]);
        foreign.archive_info.message_infos[0]
            .field_infos
            .push(field);
        archive.objects.push(foreign);
        Ok(())
    })
}

fn missing_route_identifier(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let object = archive
            .object_mut(TABLE_INFOS[0])
            .ok_or_else(|| io::Error::other("missing table-info object"))?;
        object.archive_info.identifier = None;
        Ok(())
    })
}

fn archive_info_mismatch(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
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

fn missing_model_route(package: &[u8]) -> TestResult<Vec<u8>> {
    let payload = table_info_payload(package, TABLE_INFOS[0])?;
    let mut info = tst::TableInfoArchive::decode(payload.as_slice())?;
    info.table_model = reference(999_999);
    replace_table_info_payload(package, TABLE_INFOS[0], info.encode_to_vec())
}

fn duplicate_slide_z_order(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let object = archive
            .object_mut(SLIDE)
            .ok_or_else(|| io::Error::other("missing slide object"))?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == SLIDE_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing slide message"))?;
        let mut slide = kn::SlideArchive::decode(object.messages[index].data.as_slice())?;
        slide.drawables_z_order.push(reference(TABLE_INFOS[0]));
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: SLIDE_MESSAGE_TYPE,
                data: slide.encode_to_vec(),
            },
        )?;
        Ok(())
    })
}

#[test]
fn transaction_values_are_typed_archive_free_and_redacted() {
    fn assert_send_sync_debug<T: Send + Sync + std::fmt::Debug>() {}

    assert_send_sync_debug::<State>();
    assert_send_sync_debug::<SlideTableLockStateEdit<'static>>();
    assert_send_sync_debug::<SlideTableLockStateCommit>();
    assert_send_sync_debug::<SlideTableLockStatePatch>();
    assert_send_sync_debug::<SlideTableLockStateDiagnostics>();
    assert_send_sync_debug::<SlideTableLockStateError>();
    assert_send_sync_debug::<SlideTableLockStateLimitKind>();
    assert_send_sync_debug::<SlideTableLockStatePath>();
}

#[test]
fn selectors_count_only_tables_in_z_order_and_read_all_presence_states() -> TestResult {
    let source = synthetic_package([None, Some(false), Some(true)])?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.slide_table_lock_state("Tables", TableSelector::index(0))?,
        State::Unlocked
    );
    assert_eq!(
        package.slide_table_lock_state(SlideSelector::index(0), 1usize)?,
        State::Unlocked
    );
    assert_eq!(
        package.slide_table_lock_state(SlideSelector::index(0), TableSelector::index(2))?,
        State::Locked
    );
    assert!(
        package
            .slide_table_lock_state(SlideSelector::index(0), TableSelector::index(3))
            .is_err()
    );
    assert!(
        package
            .slide_table_lock_state(SlideSelector::name("Missing"), TableSelector::index(0))
            .is_err()
    );
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn exact_noop_preserves_absent_false_true_semantics_and_metadata() -> TestResult {
    let source = synthetic_package([None, Some(false), Some(true)])?;
    for (index, state) in [State::Unlocked, State::Unlocked, State::Locked]
        .into_iter()
        .enumerate()
    {
        let package = Package::from_bytes(&source)?;
        let mut edit = package.edit_slide_table_lock_state(0usize, index)?;
        if state == State::Locked {
            edit.lock();
        } else {
            edit.unlock();
        }
        let commit = edit.commit()?;
        assert!(commit.patch().is_noop());
        assert!(!commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().touched_components(), 0);
        assert_eq!(commit.diagnostics().deleted_previews(), 0);
        assert!(!commit.diagnostics().full_reparse_performed());
        assert_eq!(exact_bytes(commit.package())?, source);
    }
    Ok(())
}

#[test]
fn lock_and_unlock_are_exact_source_bound_local_and_reversible() -> TestResult {
    let source = synthetic_package([None, Some(false), Some(true)])?;
    let package = Package::from_bytes(&source)?;
    let mut edit = package.edit_slide_table_lock_state("Tables", TableSelector::index(0))?;
    assert_eq!(edit.before(), State::Unlocked);
    edit.lock();
    assert_eq!(edit.state(), State::Locked);
    let commit = edit.commit()?;
    assert_eq!(
        commit.package().slide_table_lock_state(0usize, 0usize)?,
        State::Locked
    );
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    assert!(commit.diagnostics().full_reparse_performed());
    assert!(!commit.patch().is_noop());
    assert_eq!(commit.patch().before(), State::Unlocked);
    assert_eq!(commit.patch().after(), State::Locked);
    assert!(has_unknown_table_info_field(
        &exact_bytes(commit.package())?,
        TABLE_INFOS[0]
    )?);
    assert_eq!(
        lock_field(&exact_bytes(commit.package())?, TABLE_INFOS[0])?,
        Some(true)
    );

    let candidate = exact_bytes(commit.package())?;
    assert_locality(&source, &candidate)?;
    assert!(
        PREVIEWS
            .iter()
            .all(|name| member_bytes(&source, name).ok() == member_bytes(&candidate, name).ok())
    );

    let applied = package.apply_slide_table_lock_state(commit.patch())?;
    assert_eq!(exact_bytes(applied.package())?, candidate);
    assert!(matches!(
        commit
            .package()
            .apply_slide_table_lock_state(commit.patch()),
        Err(SlideTableLockStateError::PatchConflict)
    ));
    assert!(matches!(
        package.apply_slide_table_lock_state(&commit.patch().inverse()),
        Err(SlideTableLockStateError::PatchConflict)
    ));
    let restored = commit
        .package()
        .apply_slide_table_lock_state(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_eq!(
        restored.package().slide_table_lock_state(0usize, 0usize)?,
        State::Unlocked
    );
    Ok(())
}

#[test]
fn unlock_round_trip_preserves_explicit_false_and_inverse() -> TestResult {
    let source = synthetic_package([Some(true), Some(false), None])?;
    let package = Package::from_bytes(&source)?;
    let mut edit = package.edit_slide_table_lock_state(0usize, 0usize)?;
    assert_eq!(edit.before(), State::Locked);
    edit.unlock();
    let commit = edit.commit()?;
    assert_eq!(
        commit.package().slide_table_lock_state(0usize, 0usize)?,
        State::Unlocked
    );
    assert_eq!(
        lock_field(&exact_bytes(commit.package())?, TABLE_INFOS[0])?,
        Some(false)
    );
    let restored = commit
        .package()
        .apply_slide_table_lock_state(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_eq!(
        lock_field(&exact_bytes(restored.package())?, TABLE_INFOS[0])?,
        Some(true)
    );
    Ok(())
}

#[test]
fn malformed_lock_wire_and_role_routes_fail_closed_without_source_mutation() -> TestResult {
    let source = synthetic_package([None, Some(false), Some(true)])?;
    let duplicate = rewrite_table_info_super(&source, TABLE_INFOS[0], |super_payload| {
        append_varint_field(super_payload, DRAWABLE_LOCKED_FIELD, 1)?;
        append_varint_field(super_payload, DRAWABLE_LOCKED_FIELD, 0)?;
        Ok(())
    })?;
    let wrong_wire = rewrite_table_info_super(&source, TABLE_INFOS[0], |super_payload| {
        append_length_delimited_field(super_payload, DRAWABLE_LOCKED_FIELD, &[1])?;
        Ok(())
    })?;
    let noncanonical = rewrite_table_info_super(&source, TABLE_INFOS[0], |super_payload| {
        super_payload.extend_from_slice(&[0x28, 0x80, 0x00]);
        Ok(())
    })?;
    let duplicate_info = duplicate_info_message(&source)?;
    let info_alias = info_role_alias(&source)?;
    let model_alias = info_model_role_alias(&source)?;
    for malformed in [
        duplicate,
        wrong_wire,
        noncanonical,
        duplicate_info,
        info_alias,
        model_alias,
    ] {
        assert_rejected_atomically(&malformed)?;
    }
    Ok(())
}

#[test]
fn archive_info_routes_and_global_inbound_identity_are_atomic() -> TestResult {
    let source = synthetic_package([None, Some(false), Some(true)])?;
    let malformed = [
        archive_info_mismatch(&source)?,
        missing_model_route(&source)?,
        duplicate_slide_z_order(&source)?,
        duplicate_model_identifier(&source)?,
        foreign_inbound(&source)?,
    ];
    for source in malformed {
        assert_rejected_atomically(&source)?;
    }
    Ok(())
}

#[test]
fn opaque_metadata_and_data_or_missing_identifier_routes_fail_closed() -> TestResult {
    let source = synthetic_package([None, Some(false), Some(true)])?;
    let hostile = [
        unknown_package_metadata(&source)?,
        foreign_data_inbound(&source)?,
        foreign_field_data_inbound(&source)?,
        foreign_field_object_inbound(&source)?,
    ];
    for source in hostile {
        assert_rejected_atomically(&source)?;
    }
    // The core Archive writer itself refuses to serialize an object whose
    // ArchiveInfo identifier is missing, which is an earlier fail-closed gate.
    assert!(missing_route_identifier(&source).is_err());
    Ok(())
}

#[test]
fn lock_state_limits_fail_before_publication_and_keep_source_exact() -> TestResult {
    let source = synthetic_package([None, Some(false), Some(true)])?;
    let semantic =
        SemanticLimits::new(1_000_000, 65_536, 1, 1_000_000, 1_000_000, 64 * 1024 * 1024)?;
    let package =
        Package::from_bytes_with_options(&source, ReadOptions::new(Limits::default(), semantic))?;
    let before = exact_bytes(&package)?;
    assert!(matches!(
        package.slide_table_lock_state(0usize, 0usize),
        Err(SlideTableLockStateError::LimitExceeded { .. })
    ));
    assert!(matches!(
        package.edit_slide_table_lock_state(0usize, 0usize),
        Err(SlideTableLockStateError::LimitExceeded { .. })
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}
