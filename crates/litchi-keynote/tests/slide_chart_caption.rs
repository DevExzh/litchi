use std::io;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::wire::{WireView, append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsch, tsd, tsk, tsp, tswp};
use litchi_keynote::{ChartCaptionError, ChartSelector, Package, Position, SlideSelector};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const METADATA_OBJECT: u64 = 300;
const METADATA_LAST_IDENTIFIER: u64 = 1_000;
const DOCUMENT_COMPONENT: u64 = 1;
const UNRELATED_COMPONENT: u64 = 2;
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

fn metadata_component_payload(
    identifier: u64,
    preferred_locator: &str,
    locator: Option<&str>,
    save_token: u64,
    object_identifiers: &[u64],
) -> TestResult<Vec<u8>> {
    let mut payload = tsp::ComponentInfo {
        identifier,
        preferred_locator: preferred_locator.to_owned(),
        locator: locator.map(str::to_owned),
        save_token: Some(save_token),
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
        .push(object(80, 9_001, caption_theme_payload()?)?);
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

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
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

#[test]
fn existing_caption_replacement_is_exact_reversible_and_local() -> TestResult<()> {
    let source = synthetic_package()?;
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
    assert_eq!(commit.diagnostics().touched_components(), 1);
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
    let source = synthetic_package()?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_slide_chart_caption(0usize, 0usize)?
        .set("private chart caption")?
        .commit()?;
    let other = Package::from_bytes(&synthetic_package_with_captions([
        Some("Different"),
        Some("South"),
    ])?)?;
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
    let source = synthetic_package()?;
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
fn metadata_missing_ambiguous_malformed_overflow_and_uuid_collision_fail_atomically()
-> TestResult<()> {
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

    let collision = synthetic_metadata_package_with_captions(
        [None, Some("South")],
        METADATA_LAST_IDENTIFIER,
        Some(METADATA_LAST_IDENTIFIER + 1),
    )?;
    assert_graph_change_rejected(&collision)?;
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
