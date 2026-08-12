use std::io;

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::decode_varint_from_bytes;
use litchi_iwa_core::{
    Archive, ArchiveInfo, ArchiveObject, FieldInfo, FieldPath, FieldType, RawMessage, SnappyStream,
};
use litchi_iwa_protos::{kn, tsa, tsk, tsp};
use litchi_keynote::{Limits, Package, Position, SlideSelector, slide::delete};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const NODE_IDS: [u64; 3] = [3, 5, 7];
const SLIDE_IDS: [u64; 3] = [4, 6, 8];
const NAMES: [&str; 3] = ["Intro", "Plan", "Evidence"];
const OWNED_CHILD: u64 = 100;
const COLOCATED_SENTINEL: u64 = 101;
const HOSTILE_SURVIVOR: u64 = 200;

type TestResult<T> = Result<T, Box<dyn std::error::Error>>;

#[derive(Debug, Clone, Copy)]
enum MetadataMode {
    Exact,
    DuplicateAggregate,
    DuplicateFieldInfo,
    MissingFieldInfoType,
    WrongFieldInfoType,
    WrongFieldInfoPath,
    ShowAggregateOnlyObject,
    NodeAggregateOnlyObject,
    SurvivingInbound,
    SurvivingInboundNode,
    MissingNodeUuid,
    DuplicateNodeUuid,
    WrongComponentNodeUuid,
    MissingSlideUuid,
    DuplicateSlideUuid,
    WrongComponentSlideUuid,
    VersionedSlideUuid,
    CurrentAndVersionedSlideUuid,
    MissingExternalReference,
    ComponentOnlyExternalReference,
    DuplicateComponentOnlyExternalReference,
    DuplicateExternalReference,
    VersionedExternalReference,
    CurrentAndVersionedExternalReference,
    DuplicateNodeDataReferenceOwner,
    DuplicateSlideDataReferenceOwner,
    MismatchedNodeDataReferenceCount,
    MismatchedSlideDataReferenceCount,
    MissingNodeDataIdentifier,
    MissingSlideDataIdentifier,
    VersionedNodeDataReferenceOwner,
    VersionedSlideDataReferenceOwner,
    AmbiguousNodeIdentifier,
    AmbiguousSlideIdentifier,
    VersionedAmbiguousNodeIdentifier,
    VersionedAmbiguousSlideIdentifier,
    DuplicateComponentIdentifier,
    DuplicateEffectiveLocator,
    MismatchedComponentLocator,
    UnrelatedComponentSelectedUuid,
    NearNameRootPreview,
    SurvivorAggregateDataNode,
    SurvivorFieldDataNode,
    SurvivorAggregateDataSlide,
    SurvivorFieldDataSlide,
    NodeFieldOnlyData,
    NodeAggregateOnlyData,
    NodeWrongDataFieldType,
    SlideFieldOnlyData,
    SlideAggregateOnlyData,
    SlideWrongDataFieldType,
    NodeSecondMessageData,
    SlideSecondMessageData,
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        deprecated_type: Some(-1),
        deprecated_is_external: Some(false),
    }
}

fn object(identifier: u64, type_: u32, data: Vec<u8>) -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage { type_, data }],
    )?)
}

fn referenced_object(
    identifier: u64,
    type_: u32,
    data: Vec<u8>,
    path: &[u32],
    references: &[u64],
) -> TestResult<ArchiveObject> {
    let mut object = object(identifier, type_, data)?;
    let info = &mut object.archive_info.message_infos[0];
    info.object_references.extend_from_slice(references);
    info.field_infos.push(FieldInfo {
        path: FieldPath::new(path.to_vec()),
        r#type: Some(FieldType::ObjectReference),
        object_references: references.to_vec(),
        ..FieldInfo::default()
    });
    Ok(object)
}

fn component(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

fn encode_varint_into(output: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        output.push(((value as u8) & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn component_with_unknown_header(
    objects: Vec<ArchiveObject>,
    target_identifier: u64,
) -> TestResult<Vec<u8>> {
    let bytes = Archive { objects }.to_bytes()?;
    let parsed = Archive::parse(&bytes)?;
    let object = parsed
        .object(target_identifier)
        .ok_or_else(|| io::Error::other("synthetic header target is missing"))?;
    let header_offset = usize::try_from(object.header_offset)?;
    let data_offset = usize::try_from(object.data_offset)?;
    let (header_length, prefix_length) = decode_varint_from_bytes(&bytes[header_offset..])?;
    let header_start = header_offset + prefix_length;
    let header_end = header_start + usize::try_from(header_length)?;
    if header_end != data_offset {
        return Err(io::Error::other("synthetic header offsets disagree").into());
    }
    let mut header = bytes[header_start..header_end].to_vec();
    litchi_iwa_common::wire::append_varint_field(&mut header, 99, 9_999)?;
    let mut modified = Vec::with_capacity(bytes.len().saturating_add(8));
    modified.extend_from_slice(&bytes[..header_offset]);
    encode_varint_into(&mut modified, u64::try_from(header.len())?);
    modified.extend_from_slice(&header);
    modified.extend_from_slice(&bytes[data_offset..]);
    assert_eq!(Archive::parse(&modified)?.to_bytes()?, modified);
    Ok(SnappyStream::compress(&modified)?)
}

fn node(slide_identifier: u64, thumbnail: bool) -> Vec<u8> {
    #[allow(
        deprecated,
        reason = "native schema retains required legacy cache fields"
    )]
    kn::SlideNodeArchive {
        slide: Some(reference(slide_identifier)),
        thumbnails: thumbnail
            .then_some(tsp::DataReference { identifier: 500 })
            .into_iter()
            .collect(),
        thumbnail_sizes: thumbnail
            .then_some(tsp::Size {
                width: 320.0,
                height: 240.0,
            })
            .into_iter()
            .collect(),
        thumbnails_are_dirty: thumbnail.then_some(false),
        is_skipped: false,
        has_builds: false,
        has_transition: false,
        ..Default::default()
    }
    .encode_to_vec()
}

fn slide(name: &str, child: Option<u64>) -> Vec<u8> {
    kn::SlideArchive {
        style: reference(90),
        transition: kn::TransitionArchive::default(),
        owned_drawables: child.into_iter().map(reference).collect(),
        name: Some(name.to_owned()),
        in_document: true,
        ..Default::default()
    }
    .encode_to_vec()
}

fn uuid_entry(identifier: u64) -> tsp::ObjectUuidMapEntry {
    tsp::ObjectUuidMapEntry {
        identifier,
        uuid: tsp::Uuid {
            lower: identifier.saturating_add(1_000),
            upper: identifier.saturating_add(2_000),
        },
    }
}

fn external_reference(
    component_identifier: u64,
    object_identifier: u64,
) -> tsp::ComponentExternalReference {
    tsp::ComponentExternalReference {
        component_identifier,
        object_identifier: Some(object_identifier),
        is_weak: None,
    }
}

fn component_external_reference(component_identifier: u64) -> tsp::ComponentExternalReference {
    tsp::ComponentExternalReference {
        component_identifier,
        object_identifier: None,
        is_weak: None,
    }
}

fn data_reference(owner_identifier: u64) -> tsp::ComponentDataReference {
    tsp::ComponentDataReference {
        data_identifier: 500,
        object_reference_list: vec![tsp::component_data_reference::ObjectReference {
            object_identifier: owner_identifier,
            count: 1,
        }],
    }
}

fn data_reference_with_owners(
    data_identifier: u64,
    owners: impl IntoIterator<Item = (u64, u32)>,
) -> tsp::ComponentDataReference {
    tsp::ComponentDataReference {
        data_identifier,
        object_reference_list: owners
            .into_iter()
            .map(
                |(object_identifier, count)| tsp::component_data_reference::ObjectReference {
                    object_identifier,
                    count,
                },
            )
            .collect(),
    }
}

fn package_metadata(slide_count: usize, mode: MetadataMode) -> tsp::PackageMetadata {
    let mut document_uuids = NODE_IDS[..slide_count]
        .iter()
        .copied()
        .map(uuid_entry)
        .collect::<Vec<_>>();
    match mode {
        MetadataMode::MissingNodeUuid => {
            document_uuids.retain(|entry| entry.identifier != NODE_IDS[0])
        },
        MetadataMode::DuplicateNodeUuid => document_uuids.push(uuid_entry(NODE_IDS[0])),
        MetadataMode::WrongComponentNodeUuid => {
            document_uuids.retain(|entry| entry.identifier != NODE_IDS[0]);
        },
        MetadataMode::WrongComponentSlideUuid => document_uuids.push(uuid_entry(SLIDE_IDS[0])),
        _ => {},
    }
    let mut document_external = SLIDE_IDS[..slide_count]
        .iter()
        .copied()
        .map(|identifier| external_reference(identifier, identifier))
        .collect::<Vec<_>>();
    match mode {
        MetadataMode::MissingExternalReference | MetadataMode::VersionedExternalReference => {
            document_external.retain(|reference| reference.object_identifier != Some(SLIDE_IDS[0]));
        },
        MetadataMode::ComponentOnlyExternalReference => {
            document_external.retain(|reference| reference.object_identifier != Some(SLIDE_IDS[0]));
            document_external.push(component_external_reference(SLIDE_IDS[0]));
        },
        MetadataMode::DuplicateComponentOnlyExternalReference => {
            document_external.retain(|reference| reference.object_identifier != Some(SLIDE_IDS[0]));
            document_external.extend([
                component_external_reference(SLIDE_IDS[0]),
                component_external_reference(SLIDE_IDS[0]),
            ]);
        },
        MetadataMode::DuplicateExternalReference => {
            document_external.push(external_reference(SLIDE_IDS[0], SLIDE_IDS[0]));
        },
        _ => {},
    }
    let mut document_data_owners =
        vec![data_reference_with_owners(500, [(NODE_IDS[0], 1), (2, 1)])];
    match mode {
        MetadataMode::DuplicateNodeDataReferenceOwner => document_data_owners[0]
            .object_reference_list
            .push(tsp::component_data_reference::ObjectReference {
                object_identifier: NODE_IDS[0],
                count: 1,
            }),
        MetadataMode::MismatchedNodeDataReferenceCount => {
            document_data_owners[0].object_reference_list[0].count = 2;
        },
        MetadataMode::MissingNodeDataIdentifier => {
            document_data_owners[0].data_identifier = 999;
        },
        _ => {},
    }
    let mut components = vec![tsp::ComponentInfo {
        identifier: 1,
        preferred_locator: "Document".to_owned(),
        object_uuid_map_entries: document_uuids,
        external_references: document_external,
        data_references: document_data_owners,
        versioned_external_references: matches!(
            mode,
            MetadataMode::VersionedExternalReference
                | MetadataMode::CurrentAndVersionedExternalReference
        )
        .then(|| external_reference(SLIDE_IDS[0], SLIDE_IDS[0]))
        .into_iter()
        .collect(),
        ambiguous_object_identifiers: matches!(mode, MetadataMode::AmbiguousNodeIdentifier)
            .then_some(NODE_IDS[0])
            .into_iter()
            .collect(),
        ..Default::default()
    }];
    for &identifier in &SLIDE_IDS[..slide_count] {
        let mut object_uuid_map_entries = vec![uuid_entry(identifier)];
        if identifier == SLIDE_IDS[0]
            && matches!(
                mode,
                MetadataMode::MissingSlideUuid | MetadataMode::VersionedSlideUuid
            )
        {
            object_uuid_map_entries.clear();
        }
        if identifier == SLIDE_IDS[0] && matches!(mode, MetadataMode::DuplicateSlideUuid) {
            object_uuid_map_entries.push(uuid_entry(SLIDE_IDS[0]));
        }
        if identifier == SLIDE_IDS[0] && matches!(mode, MetadataMode::WrongComponentNodeUuid) {
            object_uuid_map_entries.push(uuid_entry(NODE_IDS[0]));
        }
        let mut component_data_owners = (identifier == SLIDE_IDS[0])
            .then(|| data_reference_with_owners(501, [(SLIDE_IDS[0], 1), (OWNED_CHILD, 1)]))
            .into_iter()
            .collect::<Vec<_>>();
        if identifier == SLIDE_IDS[0] {
            match mode {
                MetadataMode::DuplicateSlideDataReferenceOwner => component_data_owners[0]
                    .object_reference_list
                    .push(tsp::component_data_reference::ObjectReference {
                        object_identifier: SLIDE_IDS[0],
                        count: 1,
                    }),
                MetadataMode::MismatchedSlideDataReferenceCount => {
                    component_data_owners[0].object_reference_list[0].count = 2;
                },
                MetadataMode::MissingSlideDataIdentifier => {
                    component_data_owners[0].data_identifier = 999;
                },
                _ => {},
            }
        }
        components.push(tsp::ComponentInfo {
            identifier,
            preferred_locator: "Slide".to_owned(),
            locator: Some(format!("Slide-{identifier}")),
            object_uuid_map_entries,
            data_references: component_data_owners,
            ambiguous_object_identifiers: (identifier == SLIDE_IDS[0]
                && matches!(mode, MetadataMode::AmbiguousSlideIdentifier))
            .then_some(SLIDE_IDS[0])
            .into_iter()
            .collect(),
            ..Default::default()
        });
    }
    if matches!(mode, MetadataMode::SurvivingInbound) {
        components.push(tsp::ComponentInfo {
            identifier: HOSTILE_SURVIVOR,
            preferred_locator: "Hostile-Survivor".to_owned(),
            object_uuid_map_entries: vec![uuid_entry(HOSTILE_SURVIVOR)],
            ..Default::default()
        });
    }
    if matches!(mode, MetadataMode::SurvivingInboundNode) {
        components.push(tsp::ComponentInfo {
            identifier: HOSTILE_SURVIVOR,
            preferred_locator: "Hostile-Survivor".to_owned(),
            object_uuid_map_entries: vec![uuid_entry(HOSTILE_SURVIVOR)],
            external_references: vec![external_reference(1, NODE_IDS[0])],
            ..Default::default()
        });
    }
    if matches!(
        mode,
        MetadataMode::SurvivorAggregateDataNode
            | MetadataMode::SurvivorFieldDataNode
            | MetadataMode::SurvivorAggregateDataSlide
            | MetadataMode::SurvivorFieldDataSlide
    ) {
        components.push(tsp::ComponentInfo {
            identifier: HOSTILE_SURVIVOR,
            preferred_locator: "Hostile-Survivor".to_owned(),
            object_uuid_map_entries: vec![uuid_entry(HOSTILE_SURVIVOR)],
            ..Default::default()
        });
    }
    if matches!(mode, MetadataMode::DuplicateComponentIdentifier) {
        components.push(tsp::ComponentInfo {
            identifier: SLIDE_IDS[0],
            preferred_locator: "Duplicate-Slide".to_owned(),
            ..Default::default()
        });
    }
    if matches!(mode, MetadataMode::DuplicateEffectiveLocator) {
        components.push(tsp::ComponentInfo {
            identifier: 250,
            preferred_locator: "Slide".to_owned(),
            locator: Some(format!("Slide-{}", SLIDE_IDS[0])),
            ..Default::default()
        });
    }
    if matches!(mode, MetadataMode::MismatchedComponentLocator) {
        let selected = components
            .iter_mut()
            .find(|component| component.identifier == SLIDE_IDS[0])
            .expect("synthetic selected component exists");
        selected.locator = Some("Slide-999".to_owned());
    }
    if matches!(mode, MetadataMode::UnrelatedComponentSelectedUuid) {
        components.push(tsp::ComponentInfo {
            identifier: 250,
            preferred_locator: "Unrelated".to_owned(),
            object_uuid_map_entries: vec![uuid_entry(SLIDE_IDS[0])],
            ..Default::default()
        });
    }
    let mut versioned_components = Vec::new();
    if matches!(
        mode,
        MetadataMode::VersionedSlideUuid | MetadataMode::CurrentAndVersionedSlideUuid
    ) {
        versioned_components.push(tsp::ComponentInfo {
            identifier: SLIDE_IDS[0],
            preferred_locator: "Slide".to_owned(),
            locator: Some(format!("Slide-{}", SLIDE_IDS[0])),
            object_uuid_map_entries: vec![uuid_entry(SLIDE_IDS[0])],
            ..Default::default()
        });
    }
    if matches!(mode, MetadataMode::VersionedNodeDataReferenceOwner) {
        versioned_components.push(tsp::ComponentInfo {
            identifier: 1,
            preferred_locator: "Document".to_owned(),
            data_references: vec![data_reference(NODE_IDS[0])],
            ..Default::default()
        });
    }
    if matches!(mode, MetadataMode::VersionedSlideDataReferenceOwner) {
        versioned_components.push(tsp::ComponentInfo {
            identifier: SLIDE_IDS[0],
            preferred_locator: "Slide".to_owned(),
            locator: Some(format!("Slide-{}", SLIDE_IDS[0])),
            data_references: vec![data_reference(SLIDE_IDS[0])],
            ..Default::default()
        });
    }
    if matches!(mode, MetadataMode::VersionedAmbiguousNodeIdentifier) {
        versioned_components.push(tsp::ComponentInfo {
            identifier: 1,
            preferred_locator: "Document".to_owned(),
            ambiguous_object_identifiers: vec![NODE_IDS[0]],
            ..Default::default()
        });
    }
    if matches!(mode, MetadataMode::VersionedAmbiguousSlideIdentifier) {
        versioned_components.push(tsp::ComponentInfo {
            identifier: SLIDE_IDS[0],
            preferred_locator: "Slide".to_owned(),
            locator: Some(format!("Slide-{}", SLIDE_IDS[0])),
            ambiguous_object_identifiers: vec![SLIDE_IDS[0]],
            ..Default::default()
        });
    }
    tsp::PackageMetadata {
        last_object_identifier: 300,
        components,
        datas: vec![
            tsp::DataInfo {
                identifier: 500,
                digest: b"thumbnail-digest".to_vec(),
                preferred_file_name: "thumbnail.bin".to_owned(),
                ..Default::default()
            },
            tsp::DataInfo {
                identifier: 501,
                digest: b"slide-cache-digest".to_vec(),
                preferred_file_name: "slide-cache.bin".to_owned(),
                ..Default::default()
            },
        ],
        versioned_components,
        ..Default::default()
    }
}

fn package_bytes(names: &[&str], mode: MetadataMode) -> TestResult<Vec<u8>> {
    if names.is_empty() || names.len() > NODE_IDS.len() {
        return Err(io::Error::other("synthetic fixture needs one to three slides").into());
    }

    let document = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..Default::default()
        },
        show: reference(2),
        ..Default::default()
    };
    let show = kn::ShowArchive {
        theme: reference(80),
        slide_tree: kn::SlideTreeArchive {
            slides: NODE_IDS[..names.len()]
                .iter()
                .copied()
                .map(reference)
                .collect(),
            ..Default::default()
        },
        size: tsp::Size {
            width: 1_024.0,
            height: 768.0,
        },
        stylesheet: reference(81),
        mode: Some(-1),
        ..Default::default()
    };
    let mut show_object = ArchiveObject::new(
        2,
        vec![
            RawMessage {
                type_: 777,
                data: b"before-show-sentinel".to_vec(),
            },
            RawMessage {
                type_: 2,
                data: show.encode_to_vec(),
            },
            RawMessage {
                type_: 778,
                data: b"after-show-sentinel".to_vec(),
            },
        ],
    )?;
    let show_info = &mut show_object.archive_info.message_infos[1];
    show_info
        .object_references
        .extend_from_slice(&NODE_IDS[..names.len()]);
    show_info.field_infos.push(FieldInfo {
        path: FieldPath::new(vec![3, 2]),
        r#type: Some(FieldType::ObjectReference),
        object_references: NODE_IDS[..names.len()].to_vec(),
        ..FieldInfo::default()
    });
    if matches!(mode, MetadataMode::ShowAggregateOnlyObject) {
        show_info.field_infos.clear();
    }

    let mut document_objects = vec![
        referenced_object(1, 1, document.encode_to_vec(), &[2], &[2])?,
        show_object,
    ];
    for (index, (&node_identifier, &slide_identifier)) in NODE_IDS
        .iter()
        .zip(&SLIDE_IDS)
        .take(names.len())
        .enumerate()
    {
        let mut node_object = referenced_object(
            node_identifier,
            4,
            node(slide_identifier, index == 0),
            &[2],
            &[slide_identifier],
        )?;
        if index == 0 {
            let info = &mut node_object.archive_info.message_infos[0];
            info.data_references.push(500);
            info.field_infos.push(FieldInfo {
                path: FieldPath::new(vec![16]),
                r#type: Some(FieldType::DataReference),
                data_references: vec![500],
                ..FieldInfo::default()
            });
            match mode {
                MetadataMode::DuplicateAggregate => {
                    info.object_references.push(SLIDE_IDS[0]);
                },
                MetadataMode::DuplicateFieldInfo => info.field_infos.push(FieldInfo {
                    path: FieldPath::new(vec![2]),
                    r#type: Some(FieldType::ObjectReference),
                    object_references: vec![SLIDE_IDS[0]],
                    ..FieldInfo::default()
                }),
                MetadataMode::MissingFieldInfoType => {
                    info.field_infos[0].r#type = None;
                },
                MetadataMode::WrongFieldInfoType => {
                    info.field_infos[0].r#type = Some(FieldType::Value);
                },
                MetadataMode::WrongFieldInfoPath => {
                    info.field_infos[0].path = FieldPath::new(vec![99]);
                },
                MetadataMode::NodeAggregateOnlyObject => {
                    info.field_infos
                        .retain(|field| field.path.as_slice() != [2]);
                },
                _ => {},
            }
            match mode {
                MetadataMode::NodeFieldOnlyData => info.data_references.clear(),
                MetadataMode::NodeAggregateOnlyData => {
                    info.field_infos
                        .retain(|field| field.path.as_slice() != [16]);
                },
                MetadataMode::NodeWrongDataFieldType => {
                    info.field_infos
                        .iter_mut()
                        .find(|field| field.path.as_slice() == [16])
                        .expect("synthetic thumbnail FieldInfo exists")
                        .r#type = Some(FieldType::Value);
                },
                _ => {},
            }
            if matches!(mode, MetadataMode::NodeSecondMessageData) {
                let payload = b"second node data metadata".to_vec();
                let mut metadata = litchi_iwa_core::MessageInfo::new(991, payload.len() as u32);
                metadata.data_references.push(500);
                metadata.field_infos.push(FieldInfo {
                    path: FieldPath::new(vec![1]),
                    r#type: Some(FieldType::DataReference),
                    data_references: vec![500],
                    ..FieldInfo::default()
                });
                node_object.archive_info.message_infos.push(metadata);
                node_object.messages.push(RawMessage {
                    type_: 991,
                    data: payload,
                });
            }
        }
        document_objects.push(node_object);
    }

    let mut entries: Vec<(String, Vec<u8>)> = vec![
        (
            "Data/sentinel.bin".to_owned(),
            b"unrelated opaque package sentinel".to_vec(),
        ),
        (
            "Data/thumbnail.bin".to_owned(),
            b"selected node thumbnail data".to_vec(),
        ),
        (
            "Data/slide-cache.bin".to_owned(),
            b"selected slide retained cache data".to_vec(),
        ),
        (DOCUMENT_MEMBER.to_owned(), component(document_objects)?),
    ];
    for (index, (&identifier, name)) in SLIDE_IDS.iter().zip(names.iter().copied()).enumerate() {
        let child = (index == 0).then_some(OWNED_CHILD);
        let child_references = if index == 0 { &[OWNED_CHILD][..] } else { &[] };
        let mut slide_object =
            referenced_object(identifier, 5, slide(name, child), &[7], child_references)?;
        if index == 0 {
            let info = &mut slide_object.archive_info.message_infos[0];
            info.data_references.push(501);
            info.field_infos.push(FieldInfo {
                path: FieldPath::new(vec![99]),
                r#type: Some(FieldType::DataReference),
                data_references: vec![501],
                ..FieldInfo::default()
            });
            match mode {
                MetadataMode::SlideFieldOnlyData => info.data_references.clear(),
                MetadataMode::SlideAggregateOnlyData => {
                    info.field_infos
                        .retain(|field| field.path.as_slice() != [99]);
                },
                MetadataMode::SlideWrongDataFieldType => {
                    info.field_infos
                        .iter_mut()
                        .find(|field| field.path.as_slice() == [99])
                        .expect("synthetic slide-cache FieldInfo exists")
                        .r#type = Some(FieldType::Value);
                },
                _ => {},
            }
            if matches!(mode, MetadataMode::SlideSecondMessageData) {
                let payload = b"second slide data metadata".to_vec();
                let mut metadata = litchi_iwa_core::MessageInfo::new(992, payload.len() as u32);
                metadata.data_references.push(501);
                metadata.field_infos.push(FieldInfo {
                    path: FieldPath::new(vec![1]),
                    r#type: Some(FieldType::DataReference),
                    data_references: vec![501],
                    ..FieldInfo::default()
                });
                slide_object.archive_info.message_infos.push(metadata);
                slide_object.messages.push(RawMessage {
                    type_: 992,
                    data: payload,
                });
            }
        }
        let mut objects = vec![slide_object];
        if index == 0 {
            objects.push(object(
                OWNED_CHILD,
                880,
                b"selected-slide-owned-child".to_vec(),
            )?);
            objects.push(object(
                COLOCATED_SENTINEL,
                881,
                b"unknown-colocated-sentinel".to_vec(),
            )?);
        }
        let component = if index == 0 {
            component_with_unknown_header(objects, COLOCATED_SENTINEL)?
        } else {
            component(objects)?
        };
        entries.push((format!("Index/Slide-{identifier}.iwa"), component));
    }
    if matches!(mode, MetadataMode::SurvivingInbound) {
        entries.push((
            "Index/Hostile-Survivor.iwa".to_owned(),
            component(vec![referenced_object(
                HOSTILE_SURVIVOR,
                990,
                reference(SLIDE_IDS[0]).encode_to_vec(),
                &[1],
                &[SLIDE_IDS[0]],
            )?])?,
        ));
    }
    if matches!(
        mode,
        MetadataMode::SurvivorAggregateDataNode
            | MetadataMode::SurvivorFieldDataNode
            | MetadataMode::SurvivorAggregateDataSlide
            | MetadataMode::SurvivorFieldDataSlide
    ) {
        let identifier = if matches!(
            mode,
            MetadataMode::SurvivorAggregateDataNode | MetadataMode::SurvivorFieldDataNode
        ) {
            NODE_IDS[0]
        } else {
            SLIDE_IDS[0]
        };
        let mut survivor = object(HOSTILE_SURVIVOR, 990, b"hostile data metadata".to_vec())?;
        let info = &mut survivor.archive_info.message_infos[0];
        if matches!(
            mode,
            MetadataMode::SurvivorAggregateDataNode | MetadataMode::SurvivorAggregateDataSlide
        ) {
            info.data_references.push(identifier);
        } else {
            info.field_infos.push(FieldInfo {
                path: FieldPath::new(vec![1]),
                r#type: Some(FieldType::DataReference),
                data_references: vec![identifier],
                ..FieldInfo::default()
            });
        }
        entries.push((
            "Index/Hostile-Survivor.iwa".to_owned(),
            component(vec![survivor])?,
        ));
    }
    entries.push((
        "Index/Metadata.iwa".to_owned(),
        component(vec![object(
            300,
            11_006,
            package_metadata(names.len(), mode).encode_to_vec(),
        )?])?,
    ));
    entries.extend([
        ("preview.jpg".to_owned(), b"root preview".to_vec()),
        (
            "preview-micro.jpg".to_owned(),
            b"root micro preview".to_vec(),
        ),
        ("preview-web.jpg".to_owned(), b"root web preview".to_vec()),
        (
            "Index/preview.jpg".to_owned(),
            b"nested preview sentinel".to_vec(),
        ),
    ]);
    if matches!(mode, MetadataMode::NearNameRootPreview) {
        entries.push(("Preview.jpg".to_owned(), b"ambiguous root preview".to_vec()));
    }

    Ok(litchi_iwa_archive::package::to_bytes(
        entries
            .iter()
            .map(|(name, data)| (name.as_str(), data.as_slice())),
        Limits::default(),
    )?)
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn names(package: &Package) -> TestResult<Vec<String>> {
    Ok(package
        .show()?
        .slides()
        .iter()
        .map(|slide| slide.name().unwrap_or_default().to_owned())
        .collect())
}

fn component_archive(package: &[u8], member: &str) -> TestResult<Archive> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or_else(|| io::Error::other(format!("missing synthetic member {member}")))?;
    Ok(Archive::parse(
        &SnappyStream::decompress(entry.data())?.into_bytes(),
    )?)
}

fn member_bytes(package: &[u8], member: &str) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    Ok(catalog
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or_else(|| io::Error::other(format!("missing synthetic member {member}")))?
        .data()
        .to_vec())
}

fn object_contents(
    package: &[u8],
    member: &str,
    identifier: u64,
) -> TestResult<(ArchiveInfo, Vec<RawMessage>)> {
    let archive = component_archive(package, member)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other(format!("missing synthetic object {identifier}")))?;
    Ok((object.archive_info.clone(), object.messages.clone()))
}

fn raw_object_record(package: &[u8], member: &str, identifier: u64) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or_else(|| io::Error::other(format!("missing synthetic member {member}")))?;
    let stream = SnappyStream::decompress(entry.data())?.into_bytes();
    let archive = Archive::parse(&stream)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other(format!("missing synthetic object {identifier}")))?;
    let start = usize::try_from(object.header_offset)?;
    let end = usize::try_from(
        object
            .data_offset
            .checked_add(object.data_length)
            .ok_or_else(|| io::Error::other("synthetic object record range overflow"))?,
    )?;
    Ok(stream[start..end].to_vec())
}

fn metadata(package: &[u8]) -> TestResult<tsp::PackageMetadata> {
    let archive = component_archive(package, "Index/Metadata.iwa")?;
    let message = archive
        .object(300)
        .and_then(|object| {
            object
                .messages
                .iter()
                .find(|message| message.type_ == 11_006)
        })
        .ok_or_else(|| io::Error::other("missing synthetic PackageMetadata"))?;
    Ok(tsp::PackageMetadata::decode(message.data.as_slice())?)
}

#[test]
fn selector_first_deletion_is_exact_applicable_and_reversible() -> TestResult<()> {
    let source = package_bytes(&NAMES, MetadataMode::Exact)?;
    let package = Package::from_bytes(&source)?;
    let source_snapshot = exact_bytes(&package)?;
    let before_child = object_contents(&source, "Index/Slide-4.iwa", OWNED_CHILD)?;
    let before_sentinel = object_contents(&source, "Index/Slide-4.iwa", COLOCATED_SENTINEL)?;
    let before_sentinel_raw = raw_object_record(&source, "Index/Slide-4.iwa", COLOCATED_SENTINEL)?;
    let before_package_sentinel = member_bytes(&source, "Data/sentinel.bin")?;
    let before_thumbnail = member_bytes(&source, "Data/thumbnail.bin")?;
    let before_slide_cache = member_bytes(&source, "Data/slide-cache.bin")?;
    let before_nested_preview = member_bytes(&source, "Index/preview.jpg")?;
    let mut expected_metadata = metadata(&source)?;
    let document_component = expected_metadata
        .components
        .iter_mut()
        .find(|component| component.identifier == 1)
        .ok_or_else(|| io::Error::other("missing synthetic Document component"))?;
    document_component
        .object_uuid_map_entries
        .retain(|entry| entry.identifier != NODE_IDS[0]);
    document_component.external_references.retain(|reference| {
        reference.component_identifier != SLIDE_IDS[0]
            || reference.object_identifier != Some(SLIDE_IDS[0])
    });
    for data_reference in &mut document_component.data_references {
        data_reference
            .object_reference_list
            .retain(|owner| owner.object_identifier != NODE_IDS[0]);
    }
    let slide_component = expected_metadata
        .components
        .iter_mut()
        .find(|component| component.identifier == SLIDE_IDS[0])
        .ok_or_else(|| io::Error::other("missing synthetic slide component"))?;
    slide_component
        .object_uuid_map_entries
        .retain(|entry| entry.identifier != SLIDE_IDS[0]);
    for data_reference in &mut slide_component.data_references {
        data_reference
            .object_reference_list
            .retain(|owner| owner.object_identifier != SLIDE_IDS[0]);
    }

    let mut edit = package.edit_slide_deletion();
    edit.remove_slide(SlideSelector::name("Intro"))?;
    let commit = edit.commit()?;
    assert_eq!(names(commit.package())?, ["Plan", "Evidence"]);
    assert_eq!(commit.patch().position(), Position::new(0));
    assert_ne!(
        commit.patch().source_fingerprint(),
        commit.patch().target_fingerprint()
    );
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().slides_removed(), 1);
    assert_eq!(commit.diagnostics().slides_restored(), 0);
    assert_eq!(commit.diagnostics().touched_components(), 3);
    assert!(commit.diagnostics().full_reparse_performed());
    assert_eq!(exact_bytes(&package)?, source_snapshot);
    assert_eq!(source_snapshot, source);

    let changed = exact_bytes(commit.package())?;
    assert_eq!(
        object_contents(&changed, "Index/Slide-4.iwa", OWNED_CHILD)?,
        before_child
    );
    assert_eq!(
        object_contents(&changed, "Index/Slide-4.iwa", COLOCATED_SENTINEL)?,
        before_sentinel
    );
    assert_eq!(
        raw_object_record(&changed, "Index/Slide-4.iwa", COLOCATED_SENTINEL)?,
        before_sentinel_raw
    );
    assert!(
        component_archive(&changed, "Index/Slide-4.iwa")?
            .object(SLIDE_IDS[0])
            .is_none()
    );
    assert_eq!(metadata(&changed)?, expected_metadata);
    assert_eq!(
        member_bytes(&changed, "Data/thumbnail.bin")?,
        before_thumbnail
    );
    assert_eq!(
        member_bytes(&changed, "Data/slide-cache.bin")?,
        before_slide_cache
    );
    assert_eq!(
        member_bytes(&changed, "Data/sentinel.bin")?,
        before_package_sentinel
    );
    for root_preview in ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"] {
        assert!(
            Catalog::from_bytes(&changed)?
                .iter()
                .all(|entry| entry.name() != root_preview)
        );
    }
    assert_eq!(
        member_bytes(&changed, "Index/preview.jpg")?,
        before_nested_preview
    );

    let applied = package.apply_slide_deletion(commit.patch())?;
    assert_eq!(exact_bytes(applied.package())?, changed);
    let inverse = commit.patch().inverse();
    assert_eq!(inverse.inverse(), commit.patch().clone());
    let restored = commit.package().apply_slide_deletion(&inverse)?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn selectors_final_slide_and_transaction_state_fail_closed() -> TestResult<()> {
    let source = package_bytes(&NAMES, MetadataMode::Exact)?;
    let package = Package::from_bytes(&source)?;

    let mut missing_position = package.edit_slide_deletion();
    assert!(matches!(
        missing_position.remove_slide(Position::new(99)),
        Err(delete::Error::SlidePositionNotFound { position })
            if position == Position::new(99)
    ));
    let mut missing_name = package.edit_slide_deletion();
    assert!(matches!(
        missing_name.remove_slide("Missing"),
        Err(delete::Error::SlideNameNotFound)
    ));

    let ambiguous_bytes = package_bytes(&["Same", "Same", "Other"], MetadataMode::Exact)?;
    let ambiguous = Package::from_bytes(&ambiguous_bytes)?;
    let mut ambiguous_edit = ambiguous.edit_slide_deletion();
    assert!(matches!(
        ambiguous_edit.remove_slide("Same"),
        Err(delete::Error::AmbiguousSelector)
    ));

    let one_bytes = package_bytes(&["Only"], MetadataMode::Exact)?;
    let one = Package::from_bytes(&one_bytes)?;
    let mut final_edit = one.edit_slide_deletion();
    assert!(matches!(
        final_edit.remove_slide(0usize),
        Err(delete::Error::CannotDeleteFinalSlide)
    ));

    let mut staged = package.edit_slide_deletion();
    staged.remove_slide(0usize)?;
    assert!(matches!(
        staged.remove_slide(1usize),
        Err(delete::Error::OperationAlreadyStaged)
    ));
    assert!(matches!(
        package.edit_slide_deletion().commit(),
        Err(delete::Error::NoStagedOperation)
    ));
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn surviving_inbound_reference_refuses_before_publication() -> TestResult<()> {
    for mode in [
        MetadataMode::SurvivingInbound,
        MetadataMode::SurvivingInboundNode,
    ] {
        let source = package_bytes(&NAMES, mode)?;
        let package = Package::from_bytes(&source)?;
        let snapshot = exact_bytes(&package)?;
        let mut edit = package.edit_slide_deletion();
        edit.remove_slide("Intro")?;
        match edit.commit() {
            Err(delete::Error::AmbiguousOwnership) => {},
            Err(error) => panic!("{mode:?} returned {error:?}"),
            Ok(_) => panic!("{mode:?} unexpectedly committed"),
        }
        assert_eq!(exact_bytes(&package)?, snapshot, "{mode:?}");
        assert_eq!(snapshot, source, "{mode:?}");
    }
    Ok(())
}

#[test]
fn duplicate_aggregate_or_field_metadata_is_invalid_source() -> TestResult<()> {
    for mode in [
        MetadataMode::DuplicateAggregate,
        MetadataMode::DuplicateFieldInfo,
        MetadataMode::MissingFieldInfoType,
        MetadataMode::WrongFieldInfoType,
        MetadataMode::WrongFieldInfoPath,
        MetadataMode::SurvivorAggregateDataNode,
        MetadataMode::SurvivorFieldDataNode,
        MetadataMode::SurvivorAggregateDataSlide,
        MetadataMode::SurvivorFieldDataSlide,
        MetadataMode::NodeFieldOnlyData,
        MetadataMode::NodeWrongDataFieldType,
        MetadataMode::SlideFieldOnlyData,
        MetadataMode::SlideWrongDataFieldType,
    ] {
        let source = package_bytes(&NAMES, mode)?;
        let package = Package::from_bytes(&source)?;
        let snapshot = exact_bytes(&package)?;
        let mut edit = package.edit_slide_deletion();
        edit.remove_slide(0usize)?;
        match edit.commit() {
            Err(delete::Error::InvalidSource) => {},
            Err(error) => panic!("{mode:?} returned {error:?}"),
            Ok(_) => panic!("{mode:?} unexpectedly committed"),
        }
        assert_eq!(exact_bytes(&package)?, snapshot, "{mode:?}");
        assert_eq!(snapshot, source, "{mode:?}");
    }
    Ok(())
}

#[test]
fn additional_selected_message_is_invalid_source_before_publication() -> TestResult<()> {
    for mode in [
        MetadataMode::NodeSecondMessageData,
        MetadataMode::SlideSecondMessageData,
    ] {
        let source = package_bytes(&NAMES, mode)?;
        let package = Package::from_bytes(&source)?;
        let snapshot = exact_bytes(&package)?;
        let mut edit = package.edit_slide_deletion();
        edit.remove_slide(0usize)?;
        match edit.commit() {
            Err(delete::Error::InvalidSource) => {},
            Err(error) => panic!("{mode:?} returned {error:?}"),
            Ok(_) => panic!("{mode:?} unexpectedly committed"),
        }
        assert_eq!(exact_bytes(&package)?, snapshot, "{mode:?}");
        assert_eq!(snapshot, source, "{mode:?}");
    }
    Ok(())
}

#[test]
fn optional_field_attribution_and_component_external_edges_are_exactly_supported() -> TestResult<()>
{
    for mode in [
        MetadataMode::ShowAggregateOnlyObject,
        MetadataMode::NodeAggregateOnlyObject,
        MetadataMode::NodeAggregateOnlyData,
        MetadataMode::SlideAggregateOnlyData,
        MetadataMode::ComponentOnlyExternalReference,
    ] {
        let source = package_bytes(&NAMES, mode)?;
        let package = Package::from_bytes(&source)?;
        let source_raw_sentinel =
            raw_object_record(&source, "Index/Slide-4.iwa", COLOCATED_SENTINEL)?;
        let mut edit = package.edit_slide_deletion();
        edit.remove_slide(0usize)?;
        let commit = edit.commit()?;
        let changed = exact_bytes(commit.package())?;
        assert_eq!(names(commit.package())?, ["Plan", "Evidence"], "{mode:?}");
        assert_eq!(exact_bytes(&package)?, source, "{mode:?}");
        assert_eq!(
            raw_object_record(&changed, "Index/Slide-4.iwa", COLOCATED_SENTINEL)?,
            source_raw_sentinel,
            "{mode:?}"
        );

        let applied = package.apply_slide_deletion(commit.patch())?;
        assert_eq!(exact_bytes(applied.package())?, changed, "{mode:?}");
        let restored = commit
            .package()
            .apply_slide_deletion(&commit.patch().inverse())?;
        assert_eq!(exact_bytes(restored.package())?, source, "{mode:?}");

        if matches!(mode, MetadataMode::ComponentOnlyExternalReference) {
            let changed_metadata = metadata(&changed)?;
            let document = changed_metadata
                .components
                .iter()
                .find(|component| component.identifier == 1)
                .ok_or_else(|| io::Error::other("missing changed Document component"))?;
            assert!(document.external_references.iter().any(|reference| {
                reference.component_identifier == SLIDE_IDS[0]
                    && reference.object_identifier.is_none()
            }));
        }
    }
    Ok(())
}

#[test]
fn malformed_package_metadata_ownership_is_invalid_source() -> TestResult<()> {
    for mode in [
        MetadataMode::MissingNodeUuid,
        MetadataMode::DuplicateNodeUuid,
        MetadataMode::WrongComponentNodeUuid,
        MetadataMode::MissingSlideUuid,
        MetadataMode::DuplicateSlideUuid,
        MetadataMode::WrongComponentSlideUuid,
        MetadataMode::VersionedSlideUuid,
        MetadataMode::CurrentAndVersionedSlideUuid,
        MetadataMode::MissingExternalReference,
        MetadataMode::DuplicateComponentOnlyExternalReference,
        MetadataMode::DuplicateExternalReference,
        MetadataMode::VersionedExternalReference,
        MetadataMode::CurrentAndVersionedExternalReference,
        MetadataMode::DuplicateNodeDataReferenceOwner,
        MetadataMode::DuplicateSlideDataReferenceOwner,
        MetadataMode::MismatchedNodeDataReferenceCount,
        MetadataMode::MismatchedSlideDataReferenceCount,
        MetadataMode::MissingNodeDataIdentifier,
        MetadataMode::MissingSlideDataIdentifier,
        MetadataMode::VersionedNodeDataReferenceOwner,
        MetadataMode::VersionedSlideDataReferenceOwner,
        MetadataMode::AmbiguousNodeIdentifier,
        MetadataMode::AmbiguousSlideIdentifier,
        MetadataMode::VersionedAmbiguousNodeIdentifier,
        MetadataMode::VersionedAmbiguousSlideIdentifier,
        MetadataMode::DuplicateComponentIdentifier,
        MetadataMode::DuplicateEffectiveLocator,
        MetadataMode::MismatchedComponentLocator,
        MetadataMode::UnrelatedComponentSelectedUuid,
    ] {
        let source = package_bytes(&NAMES, mode)?;
        let package = Package::from_bytes(&source)?;
        let snapshot = exact_bytes(&package)?;
        let mut edit = package.edit_slide_deletion();
        edit.remove_slide(0usize)?;
        match edit.commit() {
            Err(delete::Error::InvalidSource) => {},
            Err(error) => panic!("{mode:?} returned {error:?}"),
            Ok(_) => panic!("{mode:?} unexpectedly committed"),
        }
        assert_eq!(exact_bytes(&package)?, snapshot, "{mode:?}");
        assert_eq!(snapshot, source, "{mode:?}");
    }
    Ok(())
}

#[test]
fn case_distinct_root_preview_name_is_preserved() -> TestResult<()> {
    let source = package_bytes(&NAMES, MetadataMode::NearNameRootPreview)?;
    let package = Package::from_bytes(&source)?;
    let before = member_bytes(&source, "Preview.jpg")?;
    let mut edit = package.edit_slide_deletion();
    edit.remove_slide(0usize)?;
    let commit = edit.commit()?;
    assert_eq!(
        member_bytes(&exact_bytes(commit.package())?, "Preview.jpg")?,
        before
    );
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn stale_patch_conflicts_and_debug_is_content_free() -> TestResult<()> {
    let source = package_bytes(&NAMES, MetadataMode::Exact)?;
    let package = Package::from_bytes(&source)?;
    let mut edit = package.edit_slide_deletion();
    edit.remove_slide("Intro")?;
    let commit = edit.commit()?;

    let unrelated_bytes = package_bytes(&["Changed", "Plan", "Evidence"], MetadataMode::Exact)?;
    let unrelated = Package::from_bytes(&unrelated_bytes)?;
    assert!(matches!(
        unrelated.apply_slide_deletion(commit.patch()),
        Err(delete::Error::PatchConflict)
    ));
    assert_eq!(exact_bytes(&unrelated)?, unrelated_bytes);

    let debug = format!("{:?}", commit.patch());
    assert!(!debug.contains("Index/"));
    assert!(!debug.contains("Intro"));
    assert!(!debug.contains("sentinel"));
    assert!(!debug.contains("fingerprint"));
    assert!(!debug.contains("bytes"));
    Ok(())
}

#[test]
fn public_deletion_values_are_send_sync() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<delete::Patch>();
    assert_send_sync::<delete::Commit>();
    assert_send_sync::<delete::Diagnostics>();
    assert_send_sync::<delete::Error>();
    assert_send_sync::<delete::LimitKind>();
    assert_send_sync::<delete::Path>();
}
