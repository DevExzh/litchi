use std::io;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::wire::{append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsch, tsd, tsk, tsp, tswp};
use litchi_keynote::{ChartCaptionError, ChartSelector, Package, Position, SlideSelector};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
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
fn no_op_is_exact_and_changed_graph_creation_or_clear_is_refused() -> TestResult<()> {
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
        Err(ChartCaptionError::UnsupportedDependency)
    ));
    assert!(matches!(
        package
            .edit_slide_chart_caption(0usize, 0usize)?
            .clear()?
            .commit(),
        Err(ChartCaptionError::UnsupportedDependency)
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
