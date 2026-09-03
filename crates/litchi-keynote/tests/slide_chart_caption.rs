use std::io;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::{
    decode_varint_from_bytes,
    wire::{WireView, append_length_delimited_field, append_varint_field},
};
use litchi_iwa_core::{
    Archive, ArchiveInfo, ArchiveObject, FieldInfo, FieldPath, FieldType, RawMessage, SnappyStream,
};
use litchi_iwa_protos::{kn, tsa, tsch, tsd, tsk, tsp, tswp};
use litchi_keynote::{
    ChartCaptionError, ChartCaptionLimitKind, ChartSelector, Package, Position, ReadOptions,
    SemanticLimits, SlideSelector,
};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const METADATA_OBJECT: u64 = 300;
const METADATA_LAST_IDENTIFIER: u64 = 1_000;
const DOCUMENT_COMPONENT: u64 = 1;
const UNRELATED_COMPONENT: u64 = 2;
const FOREIGN_COMPONENT: u64 = 3;
const METADATA_ROOT_UNKNOWN_FIELD: u32 = 4_001;
const METADATA_COMPONENT_UNKNOWN_FIELD: u32 = 4_002;
const METADATA_UNKNOWN_MARKER: &[u8] = b"chart-caption metadata extension";
const SLIDE_NODE: u64 = 3;
const SLIDE: u64 = 4;
const CHARTS: [u64; 2] = [100, 101];
const TITLES: [u64; 2] = [110, 111];
const NON_STYLES: [u64; 2] = [120, 121];
const CAPTION_INFOS: [u64; 2] = [130, 131];
const STORAGES: [u64; 2] = [140, 141];
const PLACEMENTS: [u64; 2] = [150, 151];
const STYLES: [u64; 2] = [160, 161];
const CHART_MESSAGE_TYPE: u32 = 5_021;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
const CHART_NON_STYLE_MESSAGE_TYPE: u32 = 5_023;
const CAPTION_INFO_MESSAGE_TYPE: u32 = 633;
const CAPTION_PLACEMENT_MESSAGE_TYPE: u32 = 634;
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const SHAPE_STYLE_MESSAGE_TYPE: u32 = 2_025;
const ARCHIVE_HEADER_UNKNOWN_FIELD: u32 = 4_003;
const ARCHIVE_HEADER_UNKNOWN_MARKER: &[u8] = b"chart-caption archive header extension";

type TestResult<T> = Result<T, Box<dyn std::error::Error>>;

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn object(identifier: u64, type_: u32, data: Vec<u8>) -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage { type_, data }],
    )?)
}

fn object_with_references(
    identifier: u64,
    type_: u32,
    data: Vec<u8>,
    references: Vec<u64>,
) -> TestResult<ArchiveObject> {
    let mut object = object(identifier, type_, data)?;
    object.archive_info.message_infos[0].object_references = references;
    Ok(object)
}

fn chart_payload(chart: usize, caption_identifier: Option<u64>) -> TestResult<Vec<u8>> {
    let drawable = tsd::DrawableArchive {
        geometry: Some(tsd::GeometryArchive {
            size: Some(tsp::Size {
                width: 640.0,
                height: 360.0,
            }),
            ..tsd::GeometryArchive::default()
        }),
        parent: Some(reference(SLIDE)),
        title: Some(reference(TITLES[chart])),
        caption: caption_identifier.map(reference),
        ..tsd::DrawableArchive::default()
    };
    let chart_data = tsch::ChartArchive {
        chart_non_style: Some(reference(NON_STYLES[chart])),
        ..tsch::ChartArchive::default()
    };
    let mut payload = Vec::new();
    append_length_delimited_field(&mut payload, 1, &drawable.encode_to_vec())?;
    append_length_delimited_field(&mut payload, 10_000, &chart_data.encode_to_vec())?;
    Ok(payload)
}

fn non_style_payload(title: &str) -> TestResult<Vec<u8>> {
    let mut extension = Vec::new();
    append_varint_field(&mut extension, 21, 1)?;
    append_length_delimited_field(&mut extension, 23, title.as_bytes())?;
    let mut payload = Vec::new();
    append_length_delimited_field(&mut payload, 10_000, &extension)?;
    Ok(payload)
}

fn caption_info_payload(chart: usize, storage_identifier: u64) -> Vec<u8> {
    #[allow(
        deprecated,
        reason = "native caption graph retains the legacy storage edge"
    )]
    let info = tsa::CaptionInfoArchive {
        super_: tswp::ShapeInfoArchive {
            super_: tsd::ShapeArchive {
                super_: tsd::DrawableArchive {
                    parent: Some(reference(CHARTS[chart])),
                    caption_hidden: Some(false),
                    ..tsd::DrawableArchive::default()
                },
                style: Some(reference(STYLES[chart])),
                ..tsd::ShapeArchive::default()
            },
            deprecated_storage: Some(reference(storage_identifier)),
            owned_storage: Some(reference(storage_identifier)),
            is_text_box: Some(true),
            ..tswp::ShapeInfoArchive::default()
        },
        placement: Some(reference(PLACEMENTS[chart])),
        child_info_kind: Some(1),
    };
    info.encode_to_vec()
}

fn storage_payload(text: &str, unknown: bool) -> TestResult<Vec<u8>> {
    let mut payload = tswp::StorageArchive {
        kind: Some(3),
        text: vec![text.to_owned()],
        in_document: Some(true),
        ..tswp::StorageArchive::default()
    }
    .encode_to_vec();
    if unknown {
        append_length_delimited_field(&mut payload, 4_001, b"opaque caption extension")?;
    }
    Ok(payload)
}

fn component(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

fn synthetic_package() -> TestResult<Vec<u8>> {
    synthetic_package_with_captions([Some("North"), Some("South")])
}

fn synthetic_package_with_captions(captions: [Option<&str>; 2]) -> TestResult<Vec<u8>> {
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
    let slide = kn::SlideArchive {
        style: reference(90),
        transition: kn::TransitionArchive::default(),
        owned_drawables: CHARTS.iter().copied().map(reference).collect(),
        drawables_z_order: CHARTS.iter().copied().map(reference).collect(),
        name: Some("Charts".to_owned()),
        in_document: true,
        ..kn::SlideArchive::default()
    };
    let mut objects = vec![
        object(1, 1, document.encode_to_vec())?,
        object(2, 2, show.encode_to_vec())?,
        object(SLIDE_NODE, 4, node.encode_to_vec())?,
        object(SLIDE, 5, slide.encode_to_vec())?,
    ];
    for chart in 0..CHARTS.len() {
        let caption_reference = captions[chart]
            .map(|_text| CAPTION_INFOS[chart])
            .unwrap_or(CAPTION_INFOS[chart]);
        objects.push(object_with_references(
            CHARTS[chart],
            CHART_MESSAGE_TYPE,
            chart_payload(chart, Some(caption_reference))?,
            vec![TITLES[chart], NON_STYLES[chart], caption_reference],
        )?);
        objects.push(object(TITLES[chart], STANDIN_MESSAGE_TYPE, Vec::new())?);
        objects.push(object(
            NON_STYLES[chart],
            CHART_NON_STYLE_MESSAGE_TYPE,
            non_style_payload(if chart == 0 { "Revenue" } else { "Costs" })?,
        )?);
        if let Some(text) = captions[chart] {
            objects.push(object_with_references(
                CAPTION_INFOS[chart],
                CAPTION_INFO_MESSAGE_TYPE,
                caption_info_payload(chart, STORAGES[chart]),
                vec![STYLES[chart], STORAGES[chart], PLACEMENTS[chart]],
            )?);
            objects.push(object(
                STORAGES[chart],
                STORAGE_MESSAGE_TYPE,
                storage_payload(text, chart == 0)?,
            )?);
            objects.push(object(
                PLACEMENTS[chart],
                CAPTION_PLACEMENT_MESSAGE_TYPE,
                tsa::CaptionPlacementArchive::default().encode_to_vec(),
            )?);
            objects.push(object(STYLES[chart], SHAPE_STYLE_MESSAGE_TYPE, Vec::new())?);
        } else {
            objects.push(object(
                CAPTION_INFOS[chart],
                STANDIN_MESSAGE_TYPE,
                Vec::new(),
            )?);
        }
    }
    let document_component = component(objects)?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Data/sentinel.bin", b"unrelated ZIP sentinel".as_slice()),
            ("preview.jpg", b"large preview".as_slice()),
            ("preview-micro.jpg", b"micro preview".as_slice()),
            ("preview-web.jpg", b"web preview".as_slice()),
            (DOCUMENT_MEMBER, document_component.as_slice()),
        ],
        Limits::default(),
    )?)
}

fn metadata_uuid_entry(identifier: u64) -> tsp::ObjectUuidMapEntry {
    tsp::ObjectUuidMapEntry {
        identifier,
        uuid: tsp::Uuid {
            lower: identifier.saturating_add(10_000),
            upper: identifier.saturating_add(20_000),
        },
    }
}

fn external_reference(
    component_identifier: u64,
    object_identifier: Option<u64>,
    is_weak: Option<bool>,
) -> tsp::ComponentExternalReference {
    tsp::ComponentExternalReference {
        component_identifier,
        object_identifier,
        is_weak,
    }
}

fn metadata_component_payload(
    identifier: u64,
    preferred_locator: &str,
    locator: Option<&str>,
    save_token: u64,
    object_identifiers: &[u64],
) -> TestResult<Vec<u8>> {
    metadata_component_payload_with_external_references(
        identifier,
        preferred_locator,
        locator,
        save_token,
        object_identifiers,
        &[],
    )
}

fn metadata_component_payload_with_external_references(
    identifier: u64,
    preferred_locator: &str,
    locator: Option<&str>,
    save_token: u64,
    object_identifiers: &[u64],
    external_references: &[tsp::ComponentExternalReference],
) -> TestResult<Vec<u8>> {
    let mut payload = tsp::ComponentInfo {
        identifier,
        preferred_locator: preferred_locator.to_owned(),
        locator: locator.map(str::to_owned),
        save_token: Some(save_token),
        external_references: external_references.to_vec(),
        object_uuid_map_entries: object_identifiers
            .iter()
            .copied()
            .map(metadata_uuid_entry)
            .collect(),
        ..tsp::ComponentInfo::default()
    }
    .encode_to_vec();
    append_length_delimited_field(
        &mut payload,
        METADATA_COMPONENT_UNKNOWN_FIELD,
        METADATA_UNKNOWN_MARKER,
    )?;
    Ok(payload)
}

fn metadata_payload(
    last_identifier: u64,
    collision_identifier: Option<u64>,
) -> TestResult<Vec<u8>> {
    let mut document_objects = vec![
        1, 2, 3, 4, 80, 81, 82, 90, 100, 101, 110, 111, 120, 121, 130, 131, 140, 141, 150, 151,
        160, 161,
    ];
    if let Some(identifier) = collision_identifier {
        document_objects.push(identifier);
    }
    let document = metadata_component_payload(
        DOCUMENT_COMPONENT,
        "Document",
        Some("Document"),
        10,
        &document_objects,
    )?;
    let unrelated = metadata_component_payload(
        UNRELATED_COMPONENT,
        "Unrelated",
        Some("Unrelated"),
        7,
        &[900],
    )?;
    let versioned = tsp::ComponentInfo {
        identifier: DOCUMENT_COMPONENT,
        preferred_locator: "Document".to_owned(),
        locator: Some("Document".to_owned()),
        save_token: Some(3),
        object_uuid_map_entries: vec![metadata_uuid_entry(901)],
        ..tsp::ComponentInfo::default()
    }
    .encode_to_vec();

    let mut payload = Vec::new();
    append_varint_field(&mut payload, 1, last_identifier)?;
    append_length_delimited_field(&mut payload, 3, &document)?;
    append_length_delimited_field(&mut payload, 3, &unrelated)?;
    append_varint_field(&mut payload, 8, 10)?;
    append_length_delimited_field(&mut payload, 11, &versioned)?;
    append_length_delimited_field(
        &mut payload,
        METADATA_ROOT_UNKNOWN_FIELD,
        METADATA_UNKNOWN_MARKER,
    )?;
    Ok(payload)
}

fn metadata_payload_with_foreign_dependencies(
    last_identifier: u64,
    external_references: &[tsp::ComponentExternalReference],
) -> TestResult<Vec<u8>> {
    let document = metadata_component_payload_with_external_references(
        DOCUMENT_COMPONENT,
        "Document",
        Some("Document"),
        10,
        &[
            1, 2, 3, 4, 80, 90, 100, 101, 110, 111, 120, 121, 130, 131, 140, 141, 150, 151, 160,
            161,
        ],
        external_references,
    )?;
    let unrelated = metadata_component_payload(
        UNRELATED_COMPONENT,
        "Unrelated",
        Some("Unrelated"),
        7,
        &[900],
    )?;
    let stylesheet = metadata_component_payload(
        FOREIGN_COMPONENT,
        "Stylesheet",
        Some("Stylesheet"),
        8,
        &[81, 82],
    )?;
    let versioned = tsp::ComponentInfo {
        identifier: DOCUMENT_COMPONENT,
        preferred_locator: "Document".to_owned(),
        locator: Some("Document".to_owned()),
        save_token: Some(3),
        object_uuid_map_entries: vec![metadata_uuid_entry(901)],
        ..tsp::ComponentInfo::default()
    }
    .encode_to_vec();
    let mut payload = Vec::new();
    append_varint_field(&mut payload, 1, last_identifier)?;
    append_length_delimited_field(&mut payload, 3, &document)?;
    append_length_delimited_field(&mut payload, 3, &unrelated)?;
    append_length_delimited_field(&mut payload, 3, &stylesheet)?;
    append_varint_field(&mut payload, 8, 10)?;
    append_length_delimited_field(&mut payload, 11, &versioned)?;
    append_length_delimited_field(
        &mut payload,
        METADATA_ROOT_UNKNOWN_FIELD,
        METADATA_UNKNOWN_MARKER,
    )?;
    Ok(payload)
}

fn caption_theme_payload() -> TestResult<Vec<u8>> {
    let presets = {
        let mut bytes = Vec::new();
        append_length_delimited_field(&mut bytes, 1, &reference(82).encode_to_vec())?;
        bytes
    };
    let mut theme_super = Vec::new();
    append_length_delimited_field(&mut theme_super, 210, &presets)?;
    let mut theme = Vec::new();
    append_length_delimited_field(&mut theme, 1, &theme_super)?;
    Ok(theme)
}

fn synthetic_metadata_package_with_captions(
    captions: [Option<&str>; 2],
    last_identifier: u64,
    collision_identifier: Option<u64>,
) -> TestResult<Vec<u8>> {
    let source = synthetic_package_with_captions(captions)?;
    let mut document = Archive::parse(&document_stream(&source)?)?;
    document
        .objects
        .push(object(80, 10, caption_theme_payload()?)?);
    document.objects.push(object(81, 9_002, Vec::new())?);
    document.objects.push(object(82, 9_003, Vec::new())?);
    let source = replace_document_stream(&source, document)?;
    let metadata = component(vec![object(
        METADATA_OBJECT,
        11_006,
        metadata_payload(last_identifier, collision_identifier)?,
    )?])?;
    let catalog = Catalog::from_bytes(&source)?;
    let mut entries = catalog
        .iter()
        .map(|entry| (entry.name(), entry.data()))
        .collect::<Vec<_>>();
    entries.push((METADATA_MEMBER, metadata.as_slice()));
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        Limits::default(),
    )?)
}

fn synthetic_metadata_package_with_one_byte_standin() -> TestResult<Vec<u8>> {
    let source = synthetic_metadata_package_with_captions(
        [None, Some("South")],
        METADATA_LAST_IDENTIFIER,
        None,
    )?;
    let mut document = Archive::parse(&document_stream(&source)?)?;
    let chart = document
        .object_mut(CHARTS[0])
        .ok_or_else(|| io::Error::other("missing synthetic chart"))?;
    let mut chart_data = chart_payload(0, Some(7))?;
    append_length_delimited_field(&mut chart_data, 4_004, &vec![0; 4_096])?;
    chart.messages[0].data = chart_data;
    chart.archive_info.message_infos[0].object_references = vec![TITLES[0], NON_STYLES[0], 7];
    let standin = document
        .object_mut(CAPTION_INFOS[0])
        .ok_or_else(|| io::Error::other("missing synthetic stand-in"))?;
    standin.archive_info.identifier = Some(7);
    replace_document_stream(&source, document)
}

fn synthetic_cross_component_metadata_package(
    captions: [Option<&str>; 2],
    external_references: &[tsp::ComponentExternalReference],
) -> TestResult<Vec<u8>> {
    let source = synthetic_package_with_captions(captions)?;
    let mut document = Archive::parse(&document_stream(&source)?)?;
    document
        .objects
        .push(object(80, 10, caption_theme_payload()?)?);
    let source = replace_document_stream(&source, document)?;
    let stylesheet = component(vec![
        object(81, 9_002, Vec::new())?,
        object(82, 9_003, Vec::new())?,
    ])?;
    let metadata = component(vec![object(
        METADATA_OBJECT,
        11_006,
        metadata_payload_with_foreign_dependencies(METADATA_LAST_IDENTIFIER, external_references)?,
    )?])?;
    let catalog = Catalog::from_bytes(&source)?;
    let mut entries = catalog
        .iter()
        .map(|entry| (entry.name(), entry.data()))
        .collect::<Vec<_>>();
    entries.push(("Index/Stylesheet.iwa", stylesheet.as_slice()));
    entries.push((METADATA_MEMBER, metadata.as_slice()));
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        Limits::default(),
    )?)
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

#[derive(Debug, Clone, Copy)]
struct PackageMetrics {
    total_uncompressed_bytes: u64,
    total_iwa_stream_bytes: u64,
    max_entry_bytes: u64,
    max_iwa_stream_bytes: usize,
    total_objects: usize,
    max_archive_objects: usize,
    max_archive_messages: usize,
    max_message_bytes: usize,
}

impl PackageMetrics {
    fn total_bytes_limit(self) -> u64 {
        self.total_uncompressed_bytes
            .max(self.total_iwa_stream_bytes)
    }
}

fn package_metrics(package: &[u8]) -> TestResult<PackageMetrics> {
    let catalog = Catalog::from_bytes(package)?;
    let mut metrics = PackageMetrics {
        total_uncompressed_bytes: 0,
        total_iwa_stream_bytes: 0,
        max_entry_bytes: 0,
        max_iwa_stream_bytes: 0,
        total_objects: 0,
        max_archive_objects: 0,
        max_archive_messages: 0,
        max_message_bytes: 0,
    };
    for entry in catalog.iter() {
        metrics.total_uncompressed_bytes = metrics
            .total_uncompressed_bytes
            .checked_add(entry.metadata().uncompressed_size())
            .ok_or_else(|| io::Error::other("synthetic package total overflow"))?;
        metrics.max_entry_bytes = metrics
            .max_entry_bytes
            .max(entry.metadata().uncompressed_size());
        if !entry.name().ends_with(".iwa") {
            continue;
        }
        let stream = SnappyStream::decompress(entry.data())?;
        metrics.total_iwa_stream_bytes = metrics
            .total_iwa_stream_bytes
            .checked_add(u64::try_from(stream.as_bytes().len())?)
            .ok_or_else(|| io::Error::other("synthetic IWA stream total overflow"))?;
        metrics.max_iwa_stream_bytes = metrics.max_iwa_stream_bytes.max(stream.as_bytes().len());
        let archive = Archive::parse(stream.as_bytes())?;
        metrics.total_objects = metrics
            .total_objects
            .checked_add(archive.objects.len())
            .ok_or_else(|| io::Error::other("synthetic package object count overflow"))?;
        metrics.max_archive_objects = metrics.max_archive_objects.max(archive.objects.len());
        let message_count = archive
            .objects
            .iter()
            .map(|object| object.messages.len())
            .sum::<usize>();
        metrics.max_archive_messages = metrics.max_archive_messages.max(message_count);
        metrics.max_message_bytes = metrics.max_message_bytes.max(
            archive
                .objects
                .iter()
                .flat_map(|object| object.messages.iter())
                .map(|message| message.data.len())
                .max()
                .unwrap_or(0),
        );
    }
    Ok(metrics)
}

fn target_read_options(target: &[u8]) -> TestResult<ReadOptions> {
    let metrics = package_metrics(target)?;
    let defaults = Limits::default();
    let max_input_bytes = u64::try_from(target.len())?;
    let archive_limits = Limits::new(
        max_input_bytes,
        defaults.max_entries(),
        metrics.max_entry_bytes,
        metrics.total_bytes_limit(),
        metrics.max_iwa_stream_bytes,
    )?;
    let iwa_limits = litchi_iwa_core::Limits::default()
        .with_archive_bytes(metrics.max_iwa_stream_bytes)?
        .with_objects(metrics.max_archive_objects)?
        .with_messages(metrics.max_archive_messages)?
        .with_message_bytes(metrics.max_message_bytes)?;
    let archive_limits = archive_limits.with_archive_limits(iwa_limits)?;
    let defaults = SemanticLimits::default();
    let semantic_limits = SemanticLimits::new(
        metrics.total_objects,
        defaults.max_slides(),
        defaults.max_references(),
        defaults.max_text_storages(),
        defaults.max_text_fragments(),
        defaults.max_text_bytes(),
    )?;
    Ok(ReadOptions::new(archive_limits, semantic_limits))
}

fn replace_archive_limit(
    options: ReadOptions,
    max_input_bytes: u64,
    max_entry_bytes: u64,
    max_total_bytes: u64,
    max_iwa_stream_bytes: usize,
    max_archive_objects: usize,
    max_archive_messages: usize,
    max_message_bytes: usize,
) -> TestResult<ReadOptions> {
    let defaults = options.archive();
    let archive_limits = Limits::new(
        max_input_bytes,
        defaults.max_entries(),
        max_entry_bytes,
        max_total_bytes,
        max_iwa_stream_bytes,
    )?;
    let iwa_limits = litchi_iwa_core::Limits::default()
        .with_archive_bytes(max_iwa_stream_bytes)?
        .with_objects(max_archive_objects)?
        .with_messages(max_archive_messages)?
        .with_message_bytes(max_message_bytes)?;
    Ok(ReadOptions::new(
        archive_limits.with_archive_limits(iwa_limits)?,
        options.semantic(),
    ))
}

fn run_caption_create_with_options(
    source: &[u8],
    options: ReadOptions,
) -> TestResult<Result<Vec<u8>, ChartCaptionError>> {
    let package = Package::from_bytes_with_options(source, options)?;
    let result = package
        .edit_slide_chart_caption(0usize, 0usize)
        .and_then(|edit| edit.set("aggregate budget caption"))
        .and_then(|edit| edit.commit());
    if result.is_err() {
        assert_eq!(exact_bytes(&package)?, source);
    }
    match result {
        Ok(commit) => Ok(Ok(exact_bytes(commit.package())?)),
        Err(error) => Ok(Err(error)),
    }
}

fn run_caption_apply_with_options(
    source: &[u8],
    patch: &litchi_keynote::ChartCaptionPatch,
    options: ReadOptions,
) -> TestResult<Result<Vec<u8>, ChartCaptionError>> {
    let package = Package::from_bytes_with_options(source, options)?;
    let result = package.apply_slide_chart_caption(patch);
    if result.is_err() {
        assert_eq!(exact_bytes(&package)?, source);
    }
    match result {
        Ok(commit) => Ok(Ok(exact_bytes(commit.package())?)),
        Err(error) => Ok(Err(error)),
    }
}

fn assert_caption_limit(result: Result<Vec<u8>, ChartCaptionError>, label: &str) -> TestResult<()> {
    assert!(
        matches!(result, Err(ChartCaptionError::LimitExceeded { .. })),
        "{label} did not return a bounded limit error: {result:?}"
    );
    Ok(())
}

fn document_stream(package: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("missing synthetic document component"))?;
    Ok(SnappyStream::decompress(entry.data())?.into_bytes())
}

fn metadata_stream(package: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == METADATA_MEMBER)
        .ok_or_else(|| io::Error::other("missing synthetic metadata component"))?;
    Ok(SnappyStream::decompress(entry.data())?.into_bytes())
}

fn metadata_message_payload(package: &[u8]) -> TestResult<Vec<u8>> {
    let archive = Archive::parse(&metadata_stream(package)?)?;
    let object = archive
        .objects
        .iter()
        .find(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == 11_006)
        })
        .ok_or_else(|| io::Error::other("missing synthetic metadata object"))?;
    object
        .messages
        .iter()
        .find(|message| message.type_ == 11_006)
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("missing synthetic metadata message").into())
}

fn decoded_metadata(package: &[u8]) -> TestResult<tsp::PackageMetadata> {
    let payload = metadata_message_payload(package)?;
    Ok(tsp::PackageMetadata::decode(payload.as_slice())?)
}

fn raw_fields(payload: &[u8], number: u32) -> TestResult<Vec<Vec<u8>>> {
    Ok(WireView::parse(payload)?
        .fields()
        .filter(|field| field.number() == number)
        .map(|field| field.raw().to_vec())
        .collect())
}

fn push_varint_width(mut value: u64, width: usize, output: &mut Vec<u8>) {
    assert!((1..=10).contains(&width));
    for index in 0..width {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if index + 1 != width {
            byte |= 0x80;
        }
        output.push(byte);
    }
    assert_eq!(value, 0, "requested varint width is too narrow");
}

fn length_delimited_field_with_key_width(number: u32, payload: &[u8], key_width: usize) -> Vec<u8> {
    let mut output = Vec::new();
    push_varint_width((u64::from(number) << 3) | 2, key_width, &mut output);
    push_varint(payload.len() as u64, &mut output);
    output.extend_from_slice(payload);
    output
}

fn varint_field_with_width(
    number: u32,
    value: u64,
    key_width: usize,
    value_width: usize,
) -> Vec<u8> {
    let mut output = Vec::new();
    push_varint_width(u64::from(number) << 3, key_width, &mut output);
    push_varint_width(value, value_width, &mut output);
    output
}

fn fixed32_field_with_key_width(number: u32, value: f32, key_width: usize) -> Vec<u8> {
    let mut output = Vec::new();
    push_varint_width((u64::from(number) << 3) | 5, key_width, &mut output);
    output.extend_from_slice(&value.to_le_bytes());
    output
}

fn append_duplicate_field(payload: &[u8], number: u32) -> TestResult<Vec<u8>> {
    let view = WireView::parse(payload)?;
    let duplicate = view
        .fields()
        .find(|field| field.number() == number)
        .map(|field| field.raw().to_vec())
        .ok_or_else(|| io::Error::other("missing field to duplicate"))?;
    let mut output = payload.to_vec();
    output.extend_from_slice(&duplicate);
    Ok(output)
}

fn replace_nested_field_raw(
    payload: &[u8],
    path: &[u32],
    replacement: &[u8],
) -> TestResult<Vec<u8>> {
    let number = *path
        .first()
        .ok_or_else(|| io::Error::other("empty nested field path"))?;
    let view = WireView::parse(payload)?;
    let mut output = Vec::with_capacity(payload.len() + replacement.len());
    let mut replaced = false;
    for field in view.fields() {
        if !replaced && field.number() == number {
            if path.len() == 1 {
                output.extend_from_slice(replacement);
            } else {
                if field.wire_type() != 2 {
                    return Err(io::Error::other("nested field is not length-delimited").into());
                }
                let nested = replace_nested_field_raw(field.payload(), &path[1..], replacement)?;
                output.extend_from_slice(&length_delimited_field_with_key_width(
                    number,
                    &nested,
                    field.key().len(),
                ));
            }
            replaced = true;
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    if !replaced {
        return Err(io::Error::other("missing nested field").into());
    }
    Ok(output)
}

fn append_nested_duplicate(payload: &[u8], path: &[u32]) -> TestResult<Vec<u8>> {
    if path.len() == 1 {
        return append_duplicate_field(payload, path[0]);
    }
    let view = WireView::parse(payload)?;
    let parent_number = path[0];
    let parent = view
        .fields()
        .find(|field| field.number() == parent_number)
        .ok_or_else(|| io::Error::other("missing nested parent"))?;
    if parent.wire_type() != 2 {
        return Err(io::Error::other("nested parent is not length-delimited").into());
    }
    let nested = append_nested_duplicate(parent.payload(), &path[1..])?;
    let replacement =
        length_delimited_field_with_key_width(parent_number, &nested, parent.key().len());
    replace_nested_field_raw(payload, &[parent_number], &replacement)
}

fn chart_reference_identifier(package: &[u8]) -> TestResult<Option<u64>> {
    let payload = message_payload(package, CHARTS[0], CHART_MESSAGE_TYPE)?;
    let outer = WireView::parse(&payload)?;
    let drawable = outer
        .fields()
        .find(|field| field.number() == 1)
        .ok_or_else(|| io::Error::other("missing chart drawable"))?;
    let drawable = WireView::parse(drawable.payload())?;
    let Some(reference) = drawable.fields().find(|field| field.number() == 11) else {
        return Ok(None);
    };
    Ok(Some(
        tsp::Reference::decode(reference.payload())?.identifier,
    ))
}

fn metadata_component(
    metadata: &tsp::PackageMetadata,
    identifier: u64,
    versioned: bool,
) -> TestResult<&tsp::ComponentInfo> {
    let components = if versioned {
        &metadata.versioned_components
    } else {
        &metadata.components
    };
    components
        .iter()
        .find(|component| component.identifier == identifier)
        .ok_or_else(|| io::Error::other("missing synthetic metadata component").into())
}

fn metadata_component_raw_fields(
    payload: &[u8],
    identifier: u64,
    number: u32,
) -> TestResult<Vec<Vec<u8>>> {
    for field in WireView::parse(payload)?
        .fields()
        .filter(|field| field.number() == 3)
    {
        let component = tsp::ComponentInfo::decode(field.payload())?;
        if component.identifier == identifier {
            return Ok(WireView::parse(field.payload())?
                .fields()
                .filter(|field| field.number() == number)
                .map(|field| field.raw().to_vec())
                .collect());
        }
    }
    Err(io::Error::other("missing synthetic current metadata component").into())
}

fn metadata_component_raw_payload(
    payload: &[u8],
    identifier: u64,
    versioned: bool,
) -> TestResult<Vec<u8>> {
    let field_number = if versioned { 11 } else { 3 };
    for field in WireView::parse(payload)?
        .fields()
        .filter(|field| field.number() == field_number)
    {
        let component = tsp::ComponentInfo::decode(field.payload())?;
        if component.identifier == identifier {
            return Ok(field.payload().to_vec());
        }
    }
    Err(io::Error::other("missing synthetic metadata component payload").into())
}

fn replace_component_field(
    component_payload: &[u8],
    number: u32,
    replacement: &[u8],
) -> TestResult<Vec<u8>> {
    replace_nested_field_raw(component_payload, &[number], replacement)
}

fn with_document_preferred_locator(source: &[u8], preferred_locator: &str) -> TestResult<Vec<u8>> {
    let payload = metadata_message_payload(source)?;
    let root = WireView::parse(&payload)?;
    let mut rewritten = Vec::with_capacity(payload.len() + preferred_locator.len());
    let mut replaced = false;
    for field in root.fields() {
        if !replaced && field.number() == 3 {
            let component = tsp::ComponentInfo::decode(field.payload())?;
            if component.identifier == DOCUMENT_COMPONENT {
                let mut component_payload = field.payload().to_vec();
                let mut preferred = Vec::new();
                append_length_delimited_field(&mut preferred, 2, preferred_locator.as_bytes())?;
                component_payload = replace_component_field(&component_payload, 2, &preferred)?;
                append_length_delimited_field(&mut rewritten, 3, &component_payload)?;
                replaced = true;
                continue;
            }
        }
        rewritten.extend_from_slice(field.raw());
    }
    if !replaced {
        return Err(io::Error::other("missing document metadata component").into());
    }
    replace_metadata_payload(source, rewritten)
}

#[derive(Clone, Copy)]
enum MetadataReservedIdentifierKind {
    UuidObject,
    ExternalObject,
    DataObject,
    RootDataMetadataMap,
    AmbiguousObject,
}

fn with_metadata_reserved_identifier(
    source: &[u8],
    identifier: u64,
    kind: MetadataReservedIdentifierKind,
) -> TestResult<Vec<u8>> {
    let payload = metadata_message_payload(source)?;
    let mut rewritten = Vec::with_capacity(payload.len().saturating_add(32));
    let mut document_seen = false;
    for field in WireView::parse(&payload)?.fields() {
        if field.number() != 3 {
            rewritten.extend_from_slice(field.raw());
            continue;
        }
        let component = tsp::ComponentInfo::decode(field.payload())?;
        if component.identifier != DOCUMENT_COMPONENT || document_seen {
            rewritten.extend_from_slice(field.raw());
            continue;
        }
        document_seen = true;
        let mut component_payload = field.payload().to_vec();
        match kind {
            MetadataReservedIdentifierKind::UuidObject => {
                let entry = metadata_uuid_entry(identifier).encode_to_vec();
                append_length_delimited_field(&mut component_payload, 11, &entry)?;
            },
            MetadataReservedIdentifierKind::ExternalObject => {
                let reference = external_reference(UNRELATED_COMPONENT, Some(identifier), None);
                append_length_delimited_field(
                    &mut component_payload,
                    6,
                    &reference.encode_to_vec(),
                )?;
            },
            MetadataReservedIdentifierKind::DataObject => {
                let reference = tsp::ComponentDataReference {
                    data_identifier: 9_999,
                    object_reference_list: vec![tsp::component_data_reference::ObjectReference {
                        object_identifier: identifier,
                        count: 1,
                    }],
                };
                append_length_delimited_field(
                    &mut component_payload,
                    7,
                    &reference.encode_to_vec(),
                )?;
            },
            MetadataReservedIdentifierKind::RootDataMetadataMap
            | MetadataReservedIdentifierKind::AmbiguousObject => {},
        }
        if matches!(kind, MetadataReservedIdentifierKind::AmbiguousObject) {
            append_varint_field(&mut component_payload, 20, identifier)?;
        }
        append_length_delimited_field(&mut rewritten, 3, &component_payload)?;
    }
    if !document_seen {
        return Err(io::Error::other("missing document metadata component").into());
    }
    if matches!(kind, MetadataReservedIdentifierKind::RootDataMetadataMap) {
        append_length_delimited_field(&mut rewritten, 10, &reference(identifier).encode_to_vec())?;
    }
    replace_metadata_payload(source, rewritten)
}

fn message_payload(package: &[u8], identifier: u64, type_: u32) -> TestResult<Vec<u8>> {
    let archive = Archive::parse(&document_stream(package)?)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing synthetic object"))?;
    object
        .messages
        .iter()
        .find(|message| message.type_ == type_)
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("missing synthetic message").into())
}

fn replace_document_stream(source: &[u8], archive: Archive) -> TestResult<Vec<u8>> {
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(Catalog::from_bytes(source)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            DOCUMENT_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?)
}

fn push_varint(value: u64, output: &mut Vec<u8>) {
    let mut value = value;
    while value >= 0x80 {
        output.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn object_header_location(stream: &[u8], identifier: u64) -> TestResult<(usize, usize, usize)> {
    let mut object_start = 0usize;
    while object_start < stream.len() {
        let (encoded_header_length, prefix_length) =
            decode_varint_from_bytes(&stream[object_start..])?;
        let header_length = usize::try_from(encoded_header_length)?;
        let header_start = object_start
            .checked_add(prefix_length)
            .ok_or_else(|| io::Error::other("synthetic header offset overflow"))?;
        let header_end = header_start
            .checked_add(header_length)
            .ok_or_else(|| io::Error::other("synthetic header range overflow"))?;
        let header = stream
            .get(header_start..header_end)
            .ok_or_else(|| io::Error::other("synthetic header is truncated"))?;
        let info = ArchiveInfo::decode(header)?;
        if info.identifier == Some(identifier) {
            return Ok((object_start, prefix_length, header_length));
        }
        let payload_length = info
            .message_infos
            .iter()
            .try_fold(0usize, |total, message| {
                total.checked_add(usize::try_from(message.length).ok()?)
            })
            .ok_or_else(|| io::Error::other("synthetic payload length overflow"))?;
        object_start = header_end
            .checked_add(payload_length)
            .ok_or_else(|| io::Error::other("synthetic object range overflow"))?;
    }
    Err(io::Error::other("synthetic object is missing").into())
}

fn object_header(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let stream = document_stream(package)?;
    let (object_start, prefix_length, header_length) = object_header_location(&stream, identifier)?;
    let header_start = object_start
        .checked_add(prefix_length)
        .ok_or_else(|| io::Error::other("synthetic header offset overflow"))?;
    let header_end = header_start
        .checked_add(header_length)
        .ok_or_else(|| io::Error::other("synthetic header range overflow"))?;
    Ok(stream[header_start..header_end].to_vec())
}

fn with_unknown_archive_header(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let mut stream = document_stream(package)?;
    let (object_start, prefix_length, header_length) = object_header_location(&stream, identifier)?;
    let header_start = object_start
        .checked_add(prefix_length)
        .ok_or_else(|| io::Error::other("synthetic header offset overflow"))?;
    let header_end = header_start
        .checked_add(header_length)
        .ok_or_else(|| io::Error::other("synthetic header range overflow"))?;
    let mut header = stream[header_start..header_end].to_vec();
    append_length_delimited_field(
        &mut header,
        ARCHIVE_HEADER_UNKNOWN_FIELD,
        ARCHIVE_HEADER_UNKNOWN_MARKER,
    )?;
    let mut rewritten = Vec::with_capacity(
        stream
            .len()
            .saturating_add(header.len().saturating_sub(header_length)),
    );
    rewritten.extend_from_slice(&stream[..object_start]);
    push_varint(header.len() as u64, &mut rewritten);
    rewritten.extend_from_slice(&header);
    rewritten.extend_from_slice(&stream[header_end..]);
    stream = rewritten;
    Archive::parse(&stream)?;
    let compressed = SnappyStream::compress(&stream)?;
    Ok(Catalog::from_bytes(package)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            DOCUMENT_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?)
}

fn with_overlong_object_length_prefix(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let mut stream = document_stream(package)?;
    let archive = Archive::parse(&stream)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing synthetic object"))?;
    let offset = usize::try_from(object.header_offset)?;
    let (_length, prefix_length) = decode_varint_from_bytes(&stream[offset..])?;
    if prefix_length != 1 {
        return Err(io::Error::other("synthetic prefix is not one byte").into());
    }
    stream[offset] |= 0x80;
    stream.insert(offset + 1, 0);
    Archive::parse(&stream)?;
    let compressed = SnappyStream::compress(&stream)?;
    Ok(Catalog::from_bytes(package)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            DOCUMENT_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?)
}

fn replace_metadata_stream(source: &[u8], archive: Archive) -> TestResult<Vec<u8>> {
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(Catalog::from_bytes(source)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            METADATA_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?)
}

fn replace_metadata_payload(source: &[u8], payload: Vec<u8>) -> TestResult<Vec<u8>> {
    let mut archive = Archive::parse(&metadata_stream(source)?)?;
    let object = archive
        .objects
        .iter_mut()
        .find(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == 11_006)
        })
        .ok_or_else(|| io::Error::other("missing synthetic metadata object"))?;
    let message = object
        .messages
        .iter_mut()
        .find(|message| message.type_ == 11_006)
        .ok_or_else(|| io::Error::other("missing synthetic metadata message"))?;
    message.data = payload;
    replace_metadata_stream(source, archive)
}

fn duplicate_metadata_object(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut archive = Archive::parse(&metadata_stream(source)?)?;
    let mut duplicate = archive
        .objects
        .iter()
        .find(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == 11_006)
        })
        .cloned()
        .ok_or_else(|| io::Error::other("missing synthetic metadata object"))?;
    duplicate.archive_info.identifier = Some(METADATA_OBJECT + 1);
    archive.objects.push(duplicate);
    replace_metadata_stream(source, archive)
}

fn shared_metadata_graph_source() -> TestResult<Vec<u8>> {
    let source = synthetic_metadata_package_with_captions(
        [Some("North"), Some("South")],
        METADATA_LAST_IDENTIFIER,
        None,
    )?;
    let mut archive = Archive::parse(&document_stream(&source)?)?;
    let second = archive
        .object_mut(CAPTION_INFOS[1])
        .ok_or_else(|| io::Error::other("missing second caption info"))?;
    second.messages[0].data = caption_info_payload(1, STORAGES[0]);
    second.archive_info.message_infos[0].object_references =
        vec![STYLES[1], STORAGES[0], PLACEMENTS[1]];
    replace_document_stream(&source, archive)
}

fn assert_graph_change_rejected(source: &[u8]) -> TestResult<()> {
    let package = Package::from_bytes(source)?;
    let result = package
        .edit_slide_chart_caption(0usize, 0usize)
        .and_then(|edit| edit.set("fresh caption"))
        .and_then(|edit| edit.commit());
    assert!(matches!(
        result,
        Err(ChartCaptionError::InvalidSource | ChartCaptionError::UnsupportedDependency)
    ));
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

fn assert_graph_or_ingress_rejected(source: &[u8]) -> TestResult<()> {
    let package = match Package::from_bytes(source) {
        Ok(package) => package,
        Err(_) => return Ok(()),
    };
    let result = package
        .edit_slide_chart_caption(0usize, 0usize)
        .and_then(|edit| edit.set("fresh caption"))
        .and_then(|edit| edit.commit());
    assert!(matches!(
        result,
        Err(ChartCaptionError::InvalidSource | ChartCaptionError::UnsupportedDependency)
    ));
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

fn assert_replacement_rejected(source: &[u8]) -> TestResult<()> {
    let package = Package::from_bytes(source)?;
    let result = package
        .edit_slide_chart_caption(0usize, 0usize)
        .and_then(|edit| edit.set("replacement"))
        .and_then(|edit| edit.commit());
    assert!(matches!(
        result,
        Err(ChartCaptionError::InvalidSource | ChartCaptionError::UnsupportedDependency)
    ));
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

fn with_document_message_payload(
    source: &[u8],
    identifier: u64,
    type_: u32,
    data: Vec<u8>,
) -> TestResult<Vec<u8>> {
    let mut archive = Archive::parse(&document_stream(source)?)?;
    let object = archive
        .object_mut(identifier)
        .ok_or_else(|| io::Error::other("missing synthetic document object"))?;
    let message = object
        .messages
        .iter_mut()
        .find(|message| message.type_ == type_)
        .ok_or_else(|| io::Error::other("missing synthetic document message"))?;
    message.data = data;
    replace_document_stream(source, archive)
}

fn with_document_message_type(
    source: &[u8],
    identifier: u64,
    type_: u32,
    replacement_type: u32,
) -> TestResult<Vec<u8>> {
    let mut archive = Archive::parse(&document_stream(source)?)?;
    let object = archive
        .object_mut(identifier)
        .ok_or_else(|| io::Error::other("missing synthetic document object"))?;
    let message = object
        .messages
        .iter_mut()
        .find(|message| message.type_ == type_)
        .ok_or_else(|| io::Error::other("missing synthetic document message"))?;
    message.type_ = replacement_type;
    replace_document_stream(source, archive)
}

fn with_unrelated_caption_owner(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut archive = Archive::parse(&document_stream(source)?)?;
    let object = archive
        .object_mut(NON_STYLES[1])
        .ok_or_else(|| io::Error::other("missing unrelated chart object"))?;
    let info = object
        .archive_info
        .message_infos
        .first_mut()
        .ok_or_else(|| io::Error::other("missing unrelated chart message info"))?;
    info.object_references.push(STORAGES[0]);
    info.field_infos.push(FieldInfo {
        path: FieldPath::new(vec![77, 1]),
        r#type: Some(FieldType::ObjectReference),
        object_references: vec![STORAGES[0]],
        ..FieldInfo::default()
    });
    replace_document_stream(source, archive)
}

#[test]
fn existing_caption_replacement_is_exact_reversible_and_local() -> TestResult<()> {
    let source = synthetic_metadata_package_with_captions(
        [Some("North"), Some("South")],
        METADATA_LAST_IDENTIFIER,
        None,
    )?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.slide_chart_caption(
            SlideSelector::name("Charts"),
            ChartSelector::name("Revenue")
        )?,
        Some("North".to_owned())
    );

    let commit = package
        .edit_slide_chart_caption(SlideSelector::index(0), ChartSelector::index(0))?
        .set("北区 caption")?
        .commit()
        .map_err(|error| io::Error::other(format!("caption commit failed: {error:?}")))?;
    assert_eq!(
        commit.package().slide_chart_caption(0usize, 0usize)?,
        Some("北区 caption".to_owned())
    );
    assert_eq!(
        commit.package().slide_chart_caption(0usize, 1usize)?,
        Some("South".to_owned())
    );
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 2);
    assert_eq!(commit.diagnostics().deleted_previews(), 3);
    assert!(commit.diagnostics().full_reparse_performed());

    let target = exact_bytes(commit.package())?;
    let source_catalog = Catalog::from_bytes(&source)?;
    let target_catalog = Catalog::from_bytes(&target)?;
    assert_eq!(
        source_catalog
            .iter()
            .find(|entry| entry.name() == "Data/sentinel.bin")
            .map(|entry| entry.data()),
        target_catalog
            .iter()
            .find(|entry| entry.name() == "Data/sentinel.bin")
            .map(|entry| entry.data())
    );
    assert!(
        target_catalog
            .iter()
            .all(|entry| !entry.name().starts_with("preview"))
    );

    let restored = commit
        .package()
        .apply_slide_chart_caption(&commit.patch().inverse())
        .map_err(|error| io::Error::other(format!("caption inverse failed: {error:?}")))?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn no_op_is_exact_and_changed_graph_without_metadata_fails_closed() -> TestResult<()> {
    let source = synthetic_package_with_captions([Some("North"), None])?;
    let package = Package::from_bytes(&source)?;
    let no_op = package
        .edit_slide_chart_caption(0usize, 0usize)?
        .set("North")?
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert_eq!(exact_bytes(no_op.package())?, source);

    assert_eq!(package.slide_chart_caption(0usize, 1usize)?, None);
    assert!(matches!(
        package
            .edit_slide_chart_caption(0usize, 1usize)?
            .set("new graph")?
            .commit(),
        Err(ChartCaptionError::InvalidSource)
    ));
    assert!(matches!(
        package
            .edit_slide_chart_caption(0usize, 0usize)?
            .clear()?
            .commit(),
        Err(ChartCaptionError::InvalidSource)
    ));
    let absent_clear = package
        .edit_slide_chart_caption(0usize, 1usize)?
        .clear()?
        .commit()?;
    assert!(absent_clear.patch().is_noop());
    assert_eq!(exact_bytes(absent_clear.package())?, source);
    Ok(())
}

#[test]
fn semantic_noop_skips_changed_only_caption_ownership_guards() -> TestResult<()> {
    let source = shared_metadata_graph_source()?;
    let package = Package::from_bytes(&source)?;
    let no_op = package
        .edit_slide_chart_caption(0usize, 0usize)?
        .set("North")?
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert!(!no_op.diagnostics().changed());
    assert_eq!(no_op.diagnostics().touched_components(), 0);
    assert_eq!(no_op.diagnostics().deleted_previews(), 0);
    assert!(!no_op.diagnostics().full_reparse_performed());
    assert_eq!(exact_bytes(no_op.package())?, source);

    let applied = package.apply_slide_chart_caption(no_op.patch())?;
    assert!(applied.patch().is_noop());
    assert!(!applied.diagnostics().changed());
    assert_eq!(exact_bytes(applied.package())?, source);

    let changed = package
        .edit_slide_chart_caption(0usize, 0usize)?
        .set("fresh caption")?
        .commit();
    assert!(matches!(
        changed,
        Err(ChartCaptionError::InvalidSource | ChartCaptionError::UnsupportedDependency)
    ));
    assert_eq!(exact_bytes(&package)?, source);
    Ok(())
}

#[test]
fn semantic_selectors_report_missing_and_ambiguous_charts() -> TestResult<()> {
    let package = Package::from_bytes(&synthetic_package()?)?;
    assert!(matches!(
        package.slide_chart_caption(0usize, ChartSelector::name("Missing")),
        Err(ChartCaptionError::ChartNameNotFound)
    ));
    assert!(matches!(
        package.slide_chart_caption(Position::new(9), 0usize),
        Err(ChartCaptionError::SlidePositionNotFound { .. })
    ));

    let source = synthetic_package()?;
    let mut archive = Archive::parse(&document_stream(&source)?)?;
    let message = archive
        .object_mut(NON_STYLES[1])
        .and_then(|object| object.messages.first_mut())
        .ok_or_else(|| io::Error::other("missing second chart title"))?;
    message.data = non_style_payload("Revenue")?;
    let duplicate = Package::from_bytes(&replace_document_stream(&source, archive)?)?;
    assert!(matches!(
        duplicate.slide_chart_caption(0usize, ChartSelector::name("Revenue")),
        Err(ChartCaptionError::AmbiguousSelector)
    ));
    Ok(())
}

#[test]
fn shared_storage_and_wrong_parent_are_rejected_atomically() -> TestResult<()> {
    let source = synthetic_package()?;
    let mut archive = Archive::parse(&document_stream(&source)?)?;
    let second = archive
        .object_mut(CAPTION_INFOS[1])
        .ok_or_else(|| io::Error::other("missing second caption info"))?;
    second.messages[0].data = caption_info_payload(1, STORAGES[0]);
    second.archive_info.message_infos[0].object_references =
        vec![STYLES[1], STORAGES[0], PLACEMENTS[1]];
    let shared_bytes = replace_document_stream(&source, archive)?;
    let shared = Package::from_bytes(&shared_bytes)?;
    assert!(matches!(
        shared.slide_chart_caption(0usize, 0usize),
        Err(ChartCaptionError::UnsupportedDependency)
    ));
    assert_eq!(exact_bytes(&shared)?, shared_bytes);

    let mut archive = Archive::parse(&document_stream(&source)?)?;
    archive
        .object_mut(CAPTION_INFOS[0])
        .ok_or_else(|| io::Error::other("missing first caption info"))?
        .messages[0]
        .data = caption_info_payload(1, STORAGES[0]);
    let malformed_bytes = replace_document_stream(&source, archive)?;
    let malformed = Package::from_bytes(&malformed_bytes)?;
    assert!(matches!(
        malformed.slide_chart_caption(0usize, 0usize),
        Err(ChartCaptionError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&malformed)?, malformed_bytes);
    Ok(())
}

#[test]
fn patch_conflict_and_debug_output_are_content_redacted() -> TestResult<()> {
    let source = synthetic_metadata_package_with_captions(
        [Some("North"), Some("South")],
        METADATA_LAST_IDENTIFIER,
        None,
    )?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_slide_chart_caption(0usize, 0usize)?
        .set("private chart caption")?
        .commit()?;
    let other = Package::from_bytes(&synthetic_metadata_package_with_captions(
        [Some("Different"), Some("South")],
        METADATA_LAST_IDENTIFIER,
        None,
    )?)?;
    assert!(matches!(
        other.apply_slide_chart_caption(commit.patch()),
        Err(ChartCaptionError::PatchConflict)
    ));
    let debug = format!("{:?}", commit.patch());
    assert!(!debug.contains("private chart caption"));
    assert!(!debug.contains("North"));
    Ok(())
}

#[test]
fn malformed_caption_edge_and_storage_wire_are_refused() -> TestResult<()> {
    let source = synthetic_package()?;
    let mut archive = Archive::parse(&document_stream(&source)?)?;
    archive
        .object_mut(CHARTS[0])
        .ok_or_else(|| io::Error::other("missing chart"))?
        .messages[0]
        .data = vec![0x0a, 0x01, 0x5a];
    let malformed = Package::from_bytes(&replace_document_stream(&source, archive)?)?;
    assert!(matches!(
        malformed.slide_chart_caption(0usize, 0usize),
        Err(ChartCaptionError::InvalidSource)
    ));

    let mut archive = Archive::parse(&document_stream(&source)?)?;
    archive
        .object_mut(STORAGES[0])
        .ok_or_else(|| io::Error::other("missing storage"))?
        .messages[0]
        .data = vec![0x18, 0x01];
    let malformed = Package::from_bytes(&replace_document_stream(&source, archive)?)?;
    assert!(matches!(
        malformed.slide_chart_caption(0usize, 0usize),
        Err(ChartCaptionError::InvalidSource)
    ));
    Ok(())
}

#[test]
fn selected_storage_unknown_fields_survive_exactly() -> TestResult<()> {
    let source = synthetic_metadata_package_with_captions(
        [Some("North"), Some("South")],
        METADATA_LAST_IDENTIFIER,
        None,
    )?;
    let before = message_payload(&source, STORAGES[0], STORAGE_MESSAGE_TYPE)?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_slide_chart_caption(0usize, 0usize)?
        .set("changed")?
        .commit()?;
    let after = message_payload(
        &exact_bytes(commit.package())?,
        STORAGES[0],
        STORAGE_MESSAGE_TYPE,
    )?;
    let marker = b"opaque caption extension";
    assert_eq!(
        before
            .windows(marker.len())
            .filter(|window| *window == marker)
            .count(),
        1
    );
    assert_eq!(
        after
            .windows(marker.len())
            .filter(|window| *window == marker)
            .count(),
        1
    );
    Ok(())
}

#[test]
fn metadata_creation_from_canonical_standin_updates_graph_registry_and_inverse() -> TestResult<()> {
    let source = synthetic_metadata_package_with_captions(
        [None, Some("South")],
        METADATA_LAST_IDENTIFIER,
        None,
    )
    .map_err(|error| io::Error::other(format!("fixture: {error}")))?;
    let before_metadata = metadata_message_payload(&source)
        .map_err(|error| io::Error::other(format!("metadata before: {error}")))?;
    let before_field_one = raw_fields(&before_metadata, 1)?;
    let before_root_unknown = raw_fields(&before_metadata, METADATA_ROOT_UNKNOWN_FIELD)?;
    let before_component_unknown = metadata_component_raw_fields(
        &before_metadata,
        DOCUMENT_COMPONENT,
        METADATA_COMPONENT_UNKNOWN_FIELD,
    )?;
    let package = Package::from_bytes(&source)
        .map_err(|error| io::Error::other(format!("package: {error}")))?;
    assert_eq!(
        package
            .slide_chart_caption(0usize, 0usize)
            .map_err(|error| io::Error::other(format!("select before: {error}")))?,
        None
    );

    let commit = package
        .edit_slide_chart_caption(0usize, 0usize)
        .map_err(|error| io::Error::other(format!("edit: {error}")))?
        .set("")
        .map_err(|error| io::Error::other(format!("set: {error}")))?
        .commit()
        .map_err(|error| io::Error::other(format!("commit: {error}")))?;
    assert_eq!(
        commit
            .package()
            .slide_chart_caption(0usize, 0usize)
            .map_err(|error| io::Error::other(format!("select target: {error}")))?,
        Some(String::new())
    );
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 2);
    assert_eq!(commit.diagnostics().deleted_previews(), 3);
    assert!(commit.diagnostics().full_reparse_performed());

    let target = exact_bytes(commit.package())?;
    let metadata = decoded_metadata(&target)?;
    assert_eq!(metadata.last_object_identifier, 1_004);
    assert_eq!(metadata.save_token, Some(11));
    assert_eq!(
        metadata_component(&metadata, DOCUMENT_COMPONENT, false)?.save_token,
        Some(11)
    );
    assert_eq!(
        metadata_component(&metadata, UNRELATED_COMPONENT, false)?.save_token,
        Some(7)
    );
    assert_eq!(
        metadata_component(&metadata, DOCUMENT_COMPONENT, true)?.save_token,
        Some(3)
    );
    let mut expected_field_one = Vec::new();
    append_varint_field(&mut expected_field_one, 1, 1_004)?;
    assert_eq!(
        raw_fields(&metadata_message_payload(&target)?, 1)?,
        vec![expected_field_one]
    );
    assert_ne!(
        raw_fields(&metadata_message_payload(&target)?, 1)?,
        before_field_one
    );
    assert_eq!(
        metadata_component_raw_fields(
            &metadata_message_payload(&target)?,
            DOCUMENT_COMPONENT,
            METADATA_COMPONENT_UNKNOWN_FIELD,
        )?,
        before_component_unknown
    );
    assert_eq!(
        raw_fields(
            &metadata_message_payload(&target)?,
            METADATA_ROOT_UNKNOWN_FIELD
        )?,
        before_root_unknown
    );
    assert!(
        metadata_message_payload(&target)?
            .windows(METADATA_UNKNOWN_MARKER.len())
            .any(|window| window == METADATA_UNKNOWN_MARKER)
    );
    let selected = metadata_component(&metadata, DOCUMENT_COMPONENT, false)?;
    assert!(
        selected
            .object_uuid_map_entries
            .iter()
            .any(|entry| entry.identifier == 1_001)
    );
    assert!(
        selected
            .object_uuid_map_entries
            .iter()
            .any(|entry| entry.identifier == 1_002)
    );
    assert!(
        selected
            .object_uuid_map_entries
            .iter()
            .any(|entry| entry.identifier == 1_003)
    );
    assert!(
        selected
            .object_uuid_map_entries
            .iter()
            .any(|entry| entry.identifier == 1_004)
    );

    let document = Archive::parse(&document_stream(&target)?)?;
    assert_eq!(
        document
            .object(CAPTION_INFOS[0])
            .map(|object| object.messages[0].type_),
        Some(STANDIN_MESSAGE_TYPE)
    );
    assert_eq!(
        document
            .object(1_002)
            .map(|object| object.messages[0].type_),
        Some(CAPTION_INFO_MESSAGE_TYPE)
    );
    assert_eq!(
        document
            .object(1_003)
            .map(|object| object.messages[0].type_),
        Some(STORAGE_MESSAGE_TYPE)
    );
    assert_eq!(chart_reference_identifier(&target)?, Some(1_002));

    let source_catalog = Catalog::from_bytes(&source)?;
    let target_catalog = Catalog::from_bytes(&target)?;
    assert_eq!(
        source_catalog
            .iter()
            .find(|entry| entry.name() == "Data/sentinel.bin")
            .map(|entry| entry.data()),
        target_catalog
            .iter()
            .find(|entry| entry.name() == "Data/sentinel.bin")
            .map(|entry| entry.data())
    );
    assert!(
        target_catalog
            .iter()
            .all(|entry| !entry.name().starts_with("preview"))
    );

    let restored = commit
        .package()
        .apply_slide_chart_caption(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn metadata_removal_allocates_fresh_standin_and_retains_old_graph() -> TestResult<()> {
    let source = synthetic_metadata_package_with_captions(
        [Some("North"), Some("South")],
        METADATA_LAST_IDENTIFIER,
        None,
    )?;
    let before_metadata = metadata_message_payload(&source)?;
    let before_field_one = raw_fields(&before_metadata, 1)?;
    let before_root_unknown = raw_fields(&before_metadata, METADATA_ROOT_UNKNOWN_FIELD)?;
    let before_component_unknown = metadata_component_raw_fields(
        &before_metadata,
        DOCUMENT_COMPONENT,
        METADATA_COMPONENT_UNKNOWN_FIELD,
    )?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_slide_chart_caption(0usize, 0usize)?
        .clear()?
        .commit()?;
    assert_eq!(commit.package().slide_chart_caption(0usize, 0usize)?, None);
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 2);
    assert_eq!(commit.diagnostics().deleted_previews(), 3);

    let target = exact_bytes(commit.package())?;
    let metadata = decoded_metadata(&target)?;
    assert_eq!(metadata.last_object_identifier, 1_001);
    assert_eq!(metadata.save_token, Some(11));
    assert_eq!(
        metadata_component(&metadata, DOCUMENT_COMPONENT, false)?.save_token,
        Some(11)
    );
    assert_eq!(
        metadata_component(&metadata, UNRELATED_COMPONENT, false)?.save_token,
        Some(7)
    );
    assert_eq!(
        metadata_component(&metadata, DOCUMENT_COMPONENT, true)?.save_token,
        Some(3)
    );
    let mut expected_field_one = Vec::new();
    append_varint_field(&mut expected_field_one, 1, 1_001)?;
    assert_eq!(
        raw_fields(&metadata_message_payload(&target)?, 1)?,
        vec![expected_field_one]
    );
    assert_ne!(
        raw_fields(&metadata_message_payload(&target)?, 1)?,
        before_field_one
    );
    assert_eq!(
        metadata_component_raw_fields(
            &metadata_message_payload(&target)?,
            DOCUMENT_COMPONENT,
            METADATA_COMPONENT_UNKNOWN_FIELD,
        )?,
        before_component_unknown
    );
    assert_eq!(
        raw_fields(
            &metadata_message_payload(&target)?,
            METADATA_ROOT_UNKNOWN_FIELD
        )?,
        before_root_unknown
    );

    let selected = metadata_component(&metadata, DOCUMENT_COMPONENT, false)?;
    for identifier in [130, 140, 150, 160] {
        assert!(
            selected
                .object_uuid_map_entries
                .iter()
                .any(|entry| entry.identifier == identifier)
        );
    }
    assert!(
        selected
            .object_uuid_map_entries
            .iter()
            .any(|entry| entry.identifier == 1_001)
    );
    let document = Archive::parse(&document_stream(&target)?)?;
    for (identifier, type_) in [
        (CAPTION_INFOS[0], CAPTION_INFO_MESSAGE_TYPE),
        (STORAGES[0], STORAGE_MESSAGE_TYPE),
        (PLACEMENTS[0], CAPTION_PLACEMENT_MESSAGE_TYPE),
        (STYLES[0], SHAPE_STYLE_MESSAGE_TYPE),
    ] {
        assert_eq!(
            document
                .object(identifier)
                .map(|object| object.messages[0].type_),
            Some(type_)
        );
    }
    assert_eq!(
        document
            .object(1_001)
            .map(|object| object.messages[0].type_),
        Some(STANDIN_MESSAGE_TYPE)
    );
    assert_eq!(chart_reference_identifier(&target)?, Some(1_001));

    let target_catalog = Catalog::from_bytes(&target)?;
    assert!(
        target_catalog
            .iter()
            .all(|entry| !entry.name().starts_with("preview"))
    );
    assert_eq!(
        Catalog::from_bytes(&source)?
            .iter()
            .find(|entry| entry.name() == "Data/sentinel.bin")
            .map(|entry| entry.data()),
        target_catalog
            .iter()
            .find(|entry| entry.name() == "Data/sentinel.bin")
            .map(|entry| entry.data())
    );
    let restored = commit
        .package()
        .apply_slide_chart_caption(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn metadata_missing_ambiguous_malformed_and_overflow_fail_atomically() -> TestResult<()> {
    let missing = synthetic_package_with_captions([None, Some("South")])?;
    assert_graph_change_rejected(&missing)?;

    let valid = synthetic_metadata_package_with_captions(
        [None, Some("South")],
        METADATA_LAST_IDENTIFIER,
        None,
    )?;
    let ambiguous = duplicate_metadata_object(&valid)?;
    assert_graph_change_rejected(&ambiguous)?;

    let malformed = replace_metadata_payload(&valid, vec![0x80])?;
    assert_graph_change_rejected(&malformed)?;

    let overflow = synthetic_metadata_package_with_captions([None, Some("South")], u64::MAX, None)?;
    assert_graph_change_rejected(&overflow)?;

    Ok(())
}

#[test]
fn metadata_owned_identifiers_beyond_watermark_are_not_reused() -> TestResult<()> {
    let source = synthetic_metadata_package_with_captions(
        [None, Some("South")],
        METADATA_LAST_IDENTIFIER,
        None,
    )?;
    let reserved = METADATA_LAST_IDENTIFIER + 1;
    for kind in [
        MetadataReservedIdentifierKind::UuidObject,
        MetadataReservedIdentifierKind::ExternalObject,
        MetadataReservedIdentifierKind::DataObject,
        MetadataReservedIdentifierKind::RootDataMetadataMap,
        MetadataReservedIdentifierKind::AmbiguousObject,
    ] {
        let hostile = with_metadata_reserved_identifier(&source, reserved, kind)?;
        let package = Package::from_bytes(&hostile)?;
        let commit = package
            .edit_slide_chart_caption(0usize, 0usize)?
            .set("reserved identifier safe")?
            .commit()?;
        let target = exact_bytes(commit.package())?;
        assert_eq!(chart_reference_identifier(&target)?, Some(reserved + 2));
        let document = Archive::parse(&document_stream(&target)?)?;
        assert!(document.object(reserved).is_none());
        for identifier in reserved + 1..=reserved + 4 {
            assert!(document.object(identifier).is_some());
        }
        assert_eq!(
            decoded_metadata(&target)?.last_object_identifier,
            reserved + 4
        );
        let restored = commit
            .package()
            .apply_slide_chart_caption(&commit.patch().inverse())?;
        assert_eq!(exact_bytes(restored.package())?, hostile);
    }
    Ok(())
}

#[test]
fn extra_or_wrong_type_root_and_theme_candidates_fail_atomically() -> TestResult<()> {
    let source = synthetic_metadata_package_with_captions(
        [None, Some("South")],
        METADATA_LAST_IDENTIFIER,
        None,
    )?;
    let mut archive = Archive::parse(&document_stream(&source)?)?;
    let document_payload = archive
        .object(1)
        .and_then(|object| object.messages.first())
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("missing synthetic root document"))?;
    archive.objects.push(object(999, 1, document_payload)?);
    let extra_root = replace_document_stream(&source, archive)?;
    assert_graph_or_ingress_rejected(&extra_root)?;

    let wrong_theme = with_document_message_type(&source, 80, 10, 9_004)?;
    assert_graph_or_ingress_rejected(&wrong_theme)?;
    Ok(())
}

#[test]
fn canonical_standin_and_shared_or_aliased_graphs_are_rejected() -> TestResult<()> {
    let source = synthetic_metadata_package_with_captions(
        [None, Some("South")],
        METADATA_LAST_IDENTIFIER,
        None,
    )?;
    let mut archive = Archive::parse(&document_stream(&source)?)?;
    archive
        .object_mut(CAPTION_INFOS[0])
        .ok_or_else(|| io::Error::other("missing stand-in"))?
        .messages[0]
        .data = vec![0x08, 0x01];
    let malformed_standin = replace_document_stream(&source, archive)?;
    let package = Package::from_bytes(&malformed_standin)?;
    assert!(matches!(
        package.slide_chart_caption(0usize, 0usize),
        Err(ChartCaptionError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&package)?, malformed_standin);

    let shared = shared_metadata_graph_source()?;
    let package = Package::from_bytes(&shared)?;
    assert!(matches!(
        package.slide_chart_caption(0usize, 0usize),
        Err(ChartCaptionError::UnsupportedDependency)
    ));
    assert_graph_change_rejected(&shared)?;

    let mut alias_archive = Archive::parse(&document_stream(&source)?)?;
    let second_chart = alias_archive
        .object_mut(CHARTS[1])
        .ok_or_else(|| io::Error::other("missing second chart"))?;
    second_chart.messages[0].data = chart_payload(1, Some(CAPTION_INFOS[0]))?;
    second_chart.archive_info.message_infos[0].object_references =
        vec![TITLES[1], NON_STYLES[1], CAPTION_INFOS[0]];
    let aliased = replace_document_stream(&source, alias_archive)?;
    assert_graph_change_rejected(&aliased)?;
    Ok(())
}

#[test]
fn metadata_replacement_advances_selected_token_once_and_preserves_source_fields() -> TestResult<()>
{
    let source = with_document_preferred_locator(
        &synthetic_metadata_package_with_captions(
            [Some("North"), Some("South")],
            METADATA_LAST_IDENTIFIER,
            None,
        )?,
        "Document-preferred",
    )?;
    let before_metadata = metadata_message_payload(&source)?;
    let before_field_one = raw_fields(&before_metadata, 1)?;
    let before_root_unknown = raw_fields(&before_metadata, METADATA_ROOT_UNKNOWN_FIELD)?;
    let before_selected_unknown = metadata_component_raw_fields(
        &before_metadata,
        DOCUMENT_COMPONENT,
        METADATA_COMPONENT_UNKNOWN_FIELD,
    )?;
    let before_unrelated =
        metadata_component_raw_payload(&before_metadata, UNRELATED_COMPONENT, false)?;
    let before_versioned =
        metadata_component_raw_payload(&before_metadata, DOCUMENT_COMPONENT, true)?;
    let before_token = raw_fields(&before_metadata, 8)?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_slide_chart_caption(0usize, 0usize)?
        .set("East")?
        .commit()?;
    let target = exact_bytes(commit.package())?;
    let target_metadata_payload = metadata_message_payload(&target)?;
    let metadata = decoded_metadata(&target)?;
    assert_eq!(metadata.last_object_identifier, METADATA_LAST_IDENTIFIER);
    assert_eq!(metadata.save_token, Some(11));
    assert_eq!(
        metadata_component(&metadata, DOCUMENT_COMPONENT, false)?.save_token,
        Some(11)
    );
    assert_eq!(
        metadata_component(&metadata, UNRELATED_COMPONENT, false)?.save_token,
        Some(7)
    );
    assert_eq!(
        metadata_component(&metadata, DOCUMENT_COMPONENT, true)?.save_token,
        Some(3)
    );
    assert_eq!(raw_fields(&target_metadata_payload, 1)?, before_field_one);
    assert_eq!(
        raw_fields(&target_metadata_payload, METADATA_ROOT_UNKNOWN_FIELD)?,
        before_root_unknown
    );
    assert_eq!(
        metadata_component_raw_fields(
            &target_metadata_payload,
            DOCUMENT_COMPONENT,
            METADATA_COMPONENT_UNKNOWN_FIELD,
        )?,
        before_selected_unknown
    );
    assert_eq!(
        metadata_component_raw_payload(&target_metadata_payload, UNRELATED_COMPONENT, false)?,
        before_unrelated
    );
    assert_eq!(
        metadata_component_raw_payload(&target_metadata_payload, DOCUMENT_COMPONENT, true)?,
        before_versioned
    );
    assert_ne!(raw_fields(&target_metadata_payload, 8)?, before_token);
    assert_eq!(chart_reference_identifier(&target)?, Some(CAPTION_INFOS[0]));
    assert_eq!(commit.diagnostics().touched_components(), 2);
    let restored = commit
        .package()
        .apply_slide_chart_caption(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn metadata_missing_or_ambiguous_changed_replacement_fails_atomically() -> TestResult<()> {
    let missing = synthetic_package()?;
    assert_replacement_rejected(&missing)?;
    let valid = synthetic_metadata_package_with_captions(
        [Some("North"), Some("South")],
        METADATA_LAST_IDENTIFIER,
        None,
    )?;
    let ambiguous = duplicate_metadata_object(&valid)?;
    assert_replacement_rejected(&ambiguous)?;
    Ok(())
}

#[test]
fn cross_component_caption_dependencies_require_exact_metadata_edges() -> TestResult<()> {
    let missing = synthetic_cross_component_metadata_package(
        [None, Some("South")],
        &[external_reference(FOREIGN_COMPONENT, Some(81), None)],
    )?;
    assert_graph_change_rejected(&missing)?;

    let references = [
        external_reference(FOREIGN_COMPONENT, Some(81), None),
        external_reference(FOREIGN_COMPONENT, Some(82), None),
    ];
    let supported = synthetic_cross_component_metadata_package([None, Some("South")], &references)?;
    let before_metadata = metadata_message_payload(&supported)?;
    let before_external = metadata_component_raw_fields(&before_metadata, DOCUMENT_COMPONENT, 6)?;
    let package = Package::from_bytes(&supported)?;
    let commit = package
        .edit_slide_chart_caption(0usize, 0usize)?
        .set("cross-component caption")?
        .commit()?;
    let target = exact_bytes(commit.package())?;
    let target_metadata = metadata_message_payload(&target)?;
    assert_eq!(
        metadata_component_raw_fields(&target_metadata, DOCUMENT_COMPONENT, 6)?,
        before_external
    );
    assert_eq!(
        decoded_metadata(&target)?
            .components
            .iter()
            .find(|component| component.identifier == FOREIGN_COMPONENT)
            .map(|component| component.save_token),
        Some(Some(8))
    );
    let restored = commit
        .package()
        .apply_slide_chart_caption(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, supported);
    Ok(())
}

#[test]
fn unrelated_aggregate_and_field_caption_owner_is_rejected_atomically() -> TestResult<()> {
    let source = synthetic_metadata_package_with_captions(
        [Some("North"), Some("South")],
        METADATA_LAST_IDENTIFIER,
        None,
    )?;
    let hostile = with_unrelated_caption_owner(&source)?;
    assert_replacement_rejected(&hostile)?;
    Ok(())
}

#[test]
fn known_caption_reference_theme_and_width_wire_variants_fail_closed() -> TestResult<()> {
    let source = synthetic_metadata_package_with_captions(
        [None, Some("South")],
        METADATA_LAST_IDENTIFIER,
        None,
    )?;
    let chart = message_payload(&source, CHARTS[0], CHART_MESSAGE_TYPE)?;
    let drawable = WireView::parse(&chart)?
        .fields()
        .find(|field| field.number() == 1)
        .ok_or_else(|| io::Error::other("missing chart drawable"))?;
    let drawable_payload = drawable.payload();
    let reference_payload = reference(CAPTION_INFOS[0]).encode_to_vec();
    let chart_variants = [
        replace_nested_field_raw(
            &chart,
            &[1],
            &length_delimited_field_with_key_width(
                1,
                &append_nested_duplicate(drawable_payload, &[11])?,
                drawable.key().len(),
            ),
        )?,
        replace_nested_field_raw(
            &chart,
            &[1, 11],
            &varint_field_with_width(11, CAPTION_INFOS[0], 1, 2),
        )?,
        replace_nested_field_raw(
            &chart,
            &[1, 11],
            &length_delimited_field_with_key_width(11, &reference_payload, 2),
        )?,
        replace_nested_field_raw(
            &chart,
            &[1, 11, 1],
            &varint_field_with_width(1, CAPTION_INFOS[0], 1, 3),
        )?,
    ];
    for payload in chart_variants {
        assert_graph_change_rejected(&with_document_message_payload(
            &source,
            CHARTS[0],
            CHART_MESSAGE_TYPE,
            payload,
        )?)?;
    }

    let theme = message_payload(&source, 80, 10)?;
    let theme_root = WireView::parse(&theme)?
        .fields()
        .find(|field| field.number() == 1)
        .ok_or_else(|| io::Error::other("missing theme root"))?;
    let theme_variants = [
        append_nested_duplicate(&theme, &[1])?,
        replace_nested_field_raw(&theme, &[1], &varint_field_with_width(1, 1, 1, 1))?,
        replace_nested_field_raw(
            &theme,
            &[1],
            &length_delimited_field_with_key_width(1, theme_root.payload(), 2),
        )?,
    ];
    for payload in theme_variants {
        let hostile = with_document_message_payload(&source, 80, 10, payload)?;
        assert_graph_change_rejected(&hostile)?;
    }

    let width_variants = [
        append_nested_duplicate(&chart, &[1, 1, 2, 1])?,
        replace_nested_field_raw(
            &chart,
            &[1, 1, 2, 1],
            &varint_field_with_width(1, 640, 1, 2),
        )?,
        replace_nested_field_raw(
            &chart,
            &[1, 1, 2, 1],
            &fixed32_field_with_key_width(1, 640.0, 2),
        )?,
    ];
    for payload in width_variants {
        let hostile =
            with_document_message_payload(&source, CHARTS[0], CHART_MESSAGE_TYPE, payload)?;
        assert_graph_change_rejected(&hostile)?;
    }
    Ok(())
}

#[test]
fn archive_headers_are_retained_and_noncanonical_object_prefixes_fail_closed() -> TestResult<()> {
    let standin_source = synthetic_metadata_package_with_captions(
        [None, Some("South")],
        METADATA_LAST_IDENTIFIER,
        None,
    )?;
    let standin_header_source = with_unknown_archive_header(&standin_source, CAPTION_INFOS[0])?;
    assert!(
        object_header(&standin_header_source, CAPTION_INFOS[0])?
            .windows(ARCHIVE_HEADER_UNKNOWN_MARKER.len())
            .any(|window| window == ARCHIVE_HEADER_UNKNOWN_MARKER)
    );
    let package = Package::from_bytes(&standin_header_source)?;
    let commit = package
        .edit_slide_chart_caption(0usize, 0usize)?
        .set("header-preserved")?
        .commit()?;
    let target = exact_bytes(commit.package())?;
    assert!(
        object_header(&target, CAPTION_INFOS[0])?
            .windows(ARCHIVE_HEADER_UNKNOWN_MARKER.len())
            .any(|window| window == ARCHIVE_HEADER_UNKNOWN_MARKER)
    );
    let restored = commit
        .package()
        .apply_slide_chart_caption(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, standin_header_source);

    let active_source = synthetic_metadata_package_with_captions(
        [Some("North"), Some("South")],
        METADATA_LAST_IDENTIFIER,
        None,
    )?;
    let active_header_source = with_unknown_archive_header(&active_source, CAPTION_INFOS[0])?;
    let package = Package::from_bytes(&active_header_source)?;
    let commit = package
        .edit_slide_chart_caption(0usize, 0usize)?
        .clear()?
        .commit()?;
    let target = exact_bytes(commit.package())?;
    assert!(
        object_header(&target, CAPTION_INFOS[0])?
            .windows(ARCHIVE_HEADER_UNKNOWN_MARKER.len())
            .any(|window| window == ARCHIVE_HEADER_UNKNOWN_MARKER)
    );
    let restored = commit
        .package()
        .apply_slide_chart_caption(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, active_header_source);

    let noncanonical = with_overlong_object_length_prefix(&active_source, CHARTS[0])?;
    assert_replacement_rejected(&noncanonical)?;
    Ok(())
}

#[test]
fn caption_reference_varint_growth_is_source_preserving_and_output_bounded() -> TestResult<()> {
    let source = synthetic_metadata_package_with_one_byte_standin()?;
    let source = with_unknown_archive_header(&source, 7)?;
    assert_eq!(chart_reference_identifier(&source)?, Some(7));
    let source_header = object_header(&source, 7)?;
    assert!(
        source_header
            .windows(ARCHIVE_HEADER_UNKNOWN_MARKER.len())
            .any(|window| window == ARCHIVE_HEADER_UNKNOWN_MARKER)
    );

    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_slide_chart_caption(0usize, 0usize)?
        .set("one-byte stand-in growth")?
        .commit()?;
    let target = exact_bytes(commit.package())?;
    assert_eq!(chart_reference_identifier(&target)?, Some(1_002));
    assert!(
        message_payload(&target, CHARTS[0], CHART_MESSAGE_TYPE)?.len()
            > message_payload(&source, CHARTS[0], CHART_MESSAGE_TYPE)?.len()
    );
    let target_header = object_header(&target, 7)?;
    assert!(
        target_header
            .windows(ARCHIVE_HEADER_UNKNOWN_MARKER.len())
            .any(|window| window == ARCHIVE_HEADER_UNKNOWN_MARKER)
    );
    assert!(target.len() > source.len());

    let restored = commit
        .package()
        .apply_slide_chart_caption(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);

    let target_metrics = package_metrics(&target)?;
    let exact = target_read_options(&target)?;
    let source_chart_bytes = message_payload(&source, CHARTS[0], CHART_MESSAGE_TYPE)?.len();
    let under_output = replace_archive_limit(
        exact,
        u64::try_from(target.len())?,
        target_metrics.max_entry_bytes,
        target_metrics.total_bytes_limit(),
        target_metrics.max_iwa_stream_bytes,
        target_metrics.max_archive_objects,
        target_metrics.max_archive_messages,
        source_chart_bytes,
    )?;
    let result = run_caption_create_with_options(&source, under_output)?;
    assert!(
        matches!(
            &result,
            Err(ChartCaptionError::LimitExceeded {
                kind: ChartCaptionLimitKind::OutputBytes,
                ..
            })
        ),
        "unexpected low-output result: {result:?}"
    );
    assert_eq!(exact_bytes(&Package::from_bytes(&source)?)?, source);
    Ok(())
}

#[test]
fn public_package_limits_bound_caption_creation_and_patch_apply_atomically() -> TestResult<()> {
    let source = synthetic_metadata_package_with_captions(
        [None, Some("South")],
        METADATA_LAST_IDENTIFIER,
        None,
    )?;
    let baseline = Package::from_bytes(&source)?
        .edit_slide_chart_caption(0usize, 0usize)?
        .set("aggregate budget caption")?
        .commit()?;
    let target = exact_bytes(baseline.package())?;
    let source_metrics = package_metrics(&source)?;
    let target_metrics = package_metrics(&target)?;
    assert!(target.len() > source.len());
    assert!(target_metrics.total_uncompressed_bytes > source_metrics.total_uncompressed_bytes);
    assert!(target_metrics.total_iwa_stream_bytes > source_metrics.total_iwa_stream_bytes);
    assert!(target_metrics.max_iwa_stream_bytes > source_metrics.max_iwa_stream_bytes);
    assert!(target_metrics.total_objects > source_metrics.total_objects);
    assert!(target_metrics.max_archive_objects > source_metrics.max_archive_objects);
    assert!(target_metrics.max_archive_messages > source_metrics.max_archive_messages);

    let exact = target_read_options(&target)?;
    let exact_result = run_caption_create_with_options(&source, exact)?;
    assert_eq!(
        exact_result
            .map_err(|error| io::Error::other(format!("exact create failed: {error:?}")))?,
        target
    );
    let exact_apply = run_caption_apply_with_options(&source, baseline.patch(), exact)?;
    assert_eq!(
        exact_apply.map_err(|error| io::Error::other(format!("exact apply failed: {error:?}")))?,
        target
    );

    let under_output = replace_archive_limit(
        exact,
        u64::try_from(target.len() - 1)?,
        target_metrics.max_entry_bytes,
        target_metrics.total_bytes_limit(),
        target_metrics.max_iwa_stream_bytes,
        target_metrics.max_archive_objects,
        target_metrics.max_archive_messages,
        target_metrics.max_message_bytes,
    )?;
    assert_caption_limit(
        run_caption_create_with_options(&source, under_output)?,
        "output max-minus-one",
    )?;
    assert_caption_limit(
        run_caption_apply_with_options(&source, baseline.patch(), under_output)?,
        "patch output max-minus-one",
    )?;

    let under_entry = replace_archive_limit(
        exact,
        u64::try_from(target.len())?,
        target_metrics.max_entry_bytes.saturating_sub(1),
        target_metrics.total_bytes_limit(),
        target_metrics.max_iwa_stream_bytes,
        target_metrics.max_archive_objects,
        target_metrics.max_archive_messages,
        target_metrics.max_message_bytes,
    )?;
    assert_caption_limit(
        run_caption_create_with_options(&source, under_entry)?,
        "entry-byte max-minus-one",
    )?;

    let under_total = replace_archive_limit(
        exact,
        u64::try_from(target.len())?,
        target_metrics.max_entry_bytes,
        target_metrics.total_bytes_limit().saturating_sub(1),
        target_metrics.max_iwa_stream_bytes,
        target_metrics.max_archive_objects,
        target_metrics.max_archive_messages,
        target_metrics.max_message_bytes,
    )?;
    assert_caption_limit(
        run_caption_create_with_options(&source, under_total)?,
        "total-byte max-minus-one",
    )?;

    let under_iwa = replace_archive_limit(
        exact,
        u64::try_from(target.len())?,
        target_metrics.max_entry_bytes,
        target_metrics.total_bytes_limit(),
        target_metrics.max_iwa_stream_bytes.saturating_sub(1),
        target_metrics.max_archive_objects,
        target_metrics.max_archive_messages,
        target_metrics.max_message_bytes,
    )?;
    assert_caption_limit(
        run_caption_create_with_options(&source, under_iwa)?,
        "decompressed-IWA max-minus-one",
    )?;

    let under_objects = replace_archive_limit(
        exact,
        u64::try_from(target.len())?,
        target_metrics.max_entry_bytes,
        target_metrics.total_bytes_limit(),
        target_metrics.max_iwa_stream_bytes,
        target_metrics.max_archive_objects.saturating_sub(1),
        target_metrics.max_archive_messages,
        target_metrics.max_message_bytes,
    )?;
    assert_caption_limit(
        run_caption_create_with_options(&source, under_objects)?,
        "IWA-object max-minus-one",
    )?;

    let under_messages = replace_archive_limit(
        exact,
        u64::try_from(target.len())?,
        target_metrics.max_entry_bytes,
        target_metrics.total_bytes_limit(),
        target_metrics.max_iwa_stream_bytes,
        target_metrics.max_archive_objects,
        target_metrics.max_archive_messages.saturating_sub(1),
        target_metrics.max_message_bytes,
    )?;
    assert_caption_limit(
        run_caption_create_with_options(&source, under_messages)?,
        "IWA-message max-minus-one",
    )?;

    let under_message_bytes = replace_archive_limit(
        exact,
        u64::try_from(target.len())?,
        target_metrics.max_entry_bytes,
        target_metrics.total_bytes_limit(),
        target_metrics.max_iwa_stream_bytes,
        target_metrics.max_archive_objects,
        target_metrics.max_archive_messages,
        target_metrics.max_message_bytes.saturating_sub(1),
    )?;
    assert_caption_limit(
        run_caption_create_with_options(&source, under_message_bytes)?,
        "message-byte max-minus-one",
    )?;

    let under_semantic_objects = ReadOptions::new(
        exact.archive(),
        SemanticLimits::new(
            target_metrics.total_objects.saturating_sub(1),
            exact.semantic().max_slides(),
            exact.semantic().max_references(),
            exact.semantic().max_text_storages(),
            exact.semantic().max_text_fragments(),
            exact.semantic().max_text_bytes(),
        )?,
    );
    assert_caption_limit(
        run_caption_create_with_options(&source, under_semantic_objects)?,
        "semantic-object max-minus-one",
    )?;

    let under_references = ReadOptions::new(
        exact.archive(),
        SemanticLimits::new(
            exact.semantic().max_objects(),
            exact.semantic().max_slides(),
            4,
            exact.semantic().max_text_storages(),
            exact.semantic().max_text_fragments(),
            exact.semantic().max_text_bytes(),
        )?,
    );
    assert_caption_limit(
        run_caption_create_with_options(&source, under_references)?,
        "semantic-reference max-minus-one",
    )?;

    let under_input = replace_archive_limit(
        exact,
        u64::try_from(source.len() - 1)?,
        target_metrics.max_entry_bytes,
        target_metrics.total_bytes_limit(),
        target_metrics.max_iwa_stream_bytes,
        target_metrics.max_archive_objects,
        target_metrics.max_archive_messages,
        target_metrics.max_message_bytes,
    )?;
    assert!(Package::from_bytes_with_options(&source, under_input).is_err());
    Ok(())
}
