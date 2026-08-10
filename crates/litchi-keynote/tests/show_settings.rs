use std::io;
use std::sync::Arc;

use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_common::{
    decode_varint_from_bytes,
    wire::{WireView, append_varint_field},
};
use litchi_iwa_core::{
    Archive, ArchiveInfo, ArchiveObject, FieldInfo, FieldPath, RawMessage, SnappyStream,
};
use litchi_iwa_protos::{kn, tsa, tsk, tsp};
use litchi_keynote::{
    Limits, Package, ReadError, ReadOptions, Seconds, SemanticLimits,
    show::{
        Commit as ShowSettingsCommit, Diagnostics as ShowSettingsDiagnostics,
        Edit as ShowSettingsEdit, Error as ShowSettingsError, LimitKind as ShowSettingsLimitKind,
        Mode, Patch as ShowSettingsPatch, Settings, Size,
    },
};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const SLIDE_NODE_MEMBER: &str = "Index/SlideNodes.iwa";
const PREVIEW_MEMBERS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];
const DOCUMENT_IDENTIFIER: u64 = 1;
const SHOW_IDENTIFIER: u64 = 2;
const SLIDE_NODE_IDENTIFIER: u64 = 3;
const SLIDE_IDENTIFIER: u64 = 4;
const SIBLING_IDENTIFIER: u64 = 99;
const PRIVATE_NATIVE_IDENTIFIER: u64 = 7_777_777_777_777_779;
const OPTIONAL_SETTING_FIELDS: [u32; 8] = [6, 8, 9, 10, 11, 15, 16, 18];
const SETTING_FIELDS: [u32; 9] = [4, 6, 8, 9, 10, 11, 15, 16, 18];

type TestResult<T> = Result<T, Box<dyn std::error::Error>>;

fn native_fixture_path() -> std::path::PathBuf {
    std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/keynote/basic.key")
}

#[derive(Debug, Clone, Copy)]
enum Malformation {
    None,
    DuplicateSlideNumbers,
    WrongSlideNumbersWireType,
}

#[allow(
    clippy::struct_excessive_bools,
    reason = "the fixture intentionally mirrors all independent native boolean settings"
)]
#[derive(Debug, Clone, Copy)]
struct NativeSettings {
    width: f32,
    height: f32,
    slide_numbers_visible: Option<bool>,
    loop_presentation: Option<bool>,
    mode: Option<i32>,
    autoplay_transition_delay: Option<f64>,
    autoplay_build_delay: Option<f64>,
    idle_timer_active: Option<bool>,
    idle_timer_delay: Option<f64>,
    automatically_plays_upon_open: Option<bool>,
}

impl NativeSettings {
    const fn absent() -> Self {
        Self {
            width: 1_024.0,
            height: 768.0,
            slide_numbers_visible: None,
            loop_presentation: None,
            mode: None,
            autoplay_transition_delay: None,
            autoplay_build_delay: None,
            idle_timer_active: None,
            idle_timer_delay: None,
            automatically_plays_upon_open: None,
        }
    }

    const fn present() -> Self {
        Self {
            width: 1_920.0,
            height: 1_080.0,
            slide_numbers_visible: Some(false),
            loop_presentation: Some(true),
            mode: Some(-7),
            autoplay_transition_delay: Some(1.25),
            autoplay_build_delay: Some(2.5),
            idle_timer_active: Some(false),
            idle_timer_delay: Some(30.0),
            automatically_plays_upon_open: Some(true),
        }
    }
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

fn component(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

#[allow(
    clippy::cast_possible_truncation,
    reason = "each emitted byte intentionally retains only the low seven varint bits"
)]
fn push_varint(mut value: u64, output: &mut Vec<u8>) {
    while value >= 0x80 {
        output.push(((value as u8) & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn length_delimited_field(number: u32, payload: &[u8]) -> Vec<u8> {
    let mut field = Vec::with_capacity(payload.len().saturating_add(8));
    push_varint((u64::from(number) << 3) | 2, &mut field);
    push_varint(payload.len() as u64, &mut field);
    field.extend_from_slice(payload);
    field
}

fn fixed32_field(number: u32, value: u32) -> Vec<u8> {
    let mut field = Vec::with_capacity(8);
    push_varint((u64::from(number) << 3) | 5, &mut field);
    field.extend_from_slice(&value.to_le_bytes());
    field
}

fn fixed64_field(number: u32, value: u64) -> Vec<u8> {
    let mut field = Vec::with_capacity(12);
    push_varint((u64::from(number) << 3) | 1, &mut field);
    field.extend_from_slice(&value.to_le_bytes());
    field
}

fn adversarial_size(width: f32, height: f32) -> TestResult<Vec<u8>> {
    let canonical = tsp::Size { width, height }.encode_to_vec();
    let view = WireView::parse(&canonical)?;
    let width_record = view
        .fields()
        .find(|field| field.number() == 1)
        .ok_or_else(|| io::Error::other("synthetic size has no width"))?;
    let height_record = view
        .fields()
        .find(|field| field.number() == 2)
        .ok_or_else(|| io::Error::other("synthetic size has no height"))?;

    let mut output = Vec::with_capacity(canonical.len().saturating_add(40));
    output.extend_from_slice(&fixed64_field(70, 0x1122_3344_5566_7788));
    output.extend_from_slice(width_record.raw());
    output.extend_from_slice(&length_delimited_field(71, b"unknown-size"));
    output.extend_from_slice(height_record.raw());
    append_varint_field(&mut output, 72, 16_384)?;
    Ok(output)
}

fn raw_show(
    settings: NativeSettings,
    malformation: Malformation,
    unknown_sentinel: u64,
) -> TestResult<Vec<u8>> {
    raw_show_with_slide_nodes(settings, malformation, unknown_sentinel, &[])
}

fn raw_show_with_slide_nodes(
    settings: NativeSettings,
    malformation: Malformation,
    unknown_sentinel: u64,
    slide_nodes: &[u64],
) -> TestResult<Vec<u8>> {
    let canonical = kn::ShowArchive {
        ui_state: Some(reference(PRIVATE_NATIVE_IDENTIFIER)),
        theme: reference(80),
        slide_tree: kn::SlideTreeArchive {
            slides: slide_nodes.iter().copied().map(reference).collect(),
            ..Default::default()
        },
        size: tsp::Size {
            width: settings.width,
            height: settings.height,
        },
        stylesheet: reference(81),
        slide_numbers_visible: settings.slide_numbers_visible,
        recording: Some(reference(82)),
        loop_presentation: settings.loop_presentation,
        mode: settings.mode,
        autoplay_transition_delay: settings.autoplay_transition_delay,
        autoplay_build_delay: settings.autoplay_build_delay,
        idle_timer_active: settings.idle_timer_active,
        idle_timer_delay: settings.idle_timer_delay,
        soundtrack: Some(reference(83)),
        automatically_plays_upon_open: settings.automatically_plays_upon_open,
        ..Default::default()
    }
    .encode_to_vec();
    let view = WireView::parse(&canonical)?;
    let mut output = Vec::with_capacity(canonical.len().saturating_add(100));
    append_varint_field(&mut output, 90, unknown_sentinel)?;
    for field in view.fields() {
        match (field.number(), malformation) {
            (4, _) => output.extend_from_slice(&length_delimited_field(
                4,
                &adversarial_size(settings.width, settings.height)?,
            )),
            (6, Malformation::DuplicateSlideNumbers) => {
                output.extend_from_slice(field.raw());
                output.extend_from_slice(field.raw());
            },
            (6, Malformation::WrongSlideNumbersWireType) => {
                output.extend_from_slice(&length_delimited_field(6, &[0]));
            },
            _ => output.extend_from_slice(field.raw()),
        }
        if field.number() == 5 {
            output.extend_from_slice(&fixed32_field(91, 0xaabb_ccdd));
            output.extend_from_slice(&length_delimited_field(92, b"unknown-show"));
        }
    }
    output.extend_from_slice(&fixed64_field(93, 0x8877_6655_4433_2211));
    Ok(output)
}

fn package_bytes(
    settings: NativeSettings,
    malformation: Malformation,
    unknown_sentinel: u64,
) -> TestResult<Vec<u8>> {
    let document = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..Default::default()
        },
        show: reference(2),
        ..Default::default()
    };
    let show = ArchiveObject::new(
        SHOW_IDENTIFIER,
        vec![
            RawMessage {
                type_: 777,
                data: b"before-show-sentinel".to_vec(),
            },
            RawMessage {
                type_: 2,
                data: raw_show(settings, malformation, unknown_sentinel)?,
            },
            RawMessage {
                type_: 778,
                data: b"after-show-sentinel".to_vec(),
            },
        ],
    )?;
    let mut document = object(DOCUMENT_IDENTIFIER, 1, document.encode_to_vec())?;
    document.archive_info.message_infos[0].object_references = vec![SHOW_IDENTIFIER];
    let mut show_field = FieldInfo::new(vec![2]);
    show_field.object_references = vec![SHOW_IDENTIFIER];
    document.archive_info.message_infos[0]
        .field_infos
        .push(show_field);
    let document_component = component(vec![document, show])?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Data/sentinel.bin", b"unrelated opaque sentinel".as_slice()),
            (DOCUMENT_MEMBER, document_component.as_slice()),
        ],
        Limits::default(),
    )?)
}

fn interleaved_adversarial_zip_package_bytes() -> TestResult<Vec<u8>> {
    let source = package_bytes(NativeSettings::absent(), Malformation::None, 160)?;
    let catalog = Catalog::from_bytes(&source)?;
    let document = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("missing synthetic document member"))?
        .data()
        .to_vec();
    let interleaved = litchi_iwa_archive::package::to_bytes(
        [
            ("Data/before.bin", b"retained before show".as_slice()),
            (DOCUMENT_MEMBER, document.as_slice()),
            ("Data/after.bin", b"retained after show".as_slice()),
        ],
        Limits::default(),
    )?;
    adversarialize_selected_zip_metadata(&interleaved)
}

fn adversarialize_selected_zip_metadata(package: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let selected = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("missing synthetic document member"))?;
    let local = selected.raw_record().local_record().to_vec();
    let central = selected.raw_record().central_directory_record().to_vec();
    drop(catalog);

    let local_offset = unique_subslice_offset(package, &local)?;
    let central_offset = unique_subslice_offset(package, &central)?;
    let mut output = package.to_vec();
    let modified_time = 0x5b7d_u16.to_le_bytes();
    let modified_date = 0x5794_u16.to_le_bytes();
    output[local_offset + 10..local_offset + 12].copy_from_slice(&modified_time);
    output[local_offset + 12..local_offset + 14].copy_from_slice(&modified_date);
    output[central_offset + 4..central_offset + 6].copy_from_slice(&0x031e_u16.to_le_bytes());
    output[central_offset + 12..central_offset + 14].copy_from_slice(&modified_time);
    output[central_offset + 14..central_offset + 16].copy_from_slice(&modified_date);
    output[central_offset + 36..central_offset + 38].copy_from_slice(&0xa55a_u16.to_le_bytes());
    output[central_offset + 38..central_offset + 42]
        .copy_from_slice(&0x81a4_0000_u32.to_le_bytes());
    Catalog::from_bytes(&output)?;
    Ok(output)
}

fn unique_subslice_offset(haystack: &[u8], needle: &[u8]) -> TestResult<usize> {
    let mut matches = haystack
        .windows(needle.len())
        .enumerate()
        .filter_map(|(offset, candidate)| (candidate == needle).then_some(offset));
    let offset = matches
        .next()
        .ok_or_else(|| io::Error::other("raw ZIP record is missing"))?;
    if matches.next().is_some() {
        return Err(io::Error::other("raw ZIP record is not unique").into());
    }
    Ok(offset)
}

fn cache_package_bytes(split_components: bool) -> TestResult<Vec<u8>> {
    let document_value = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..Default::default()
        },
        show: reference(SHOW_IDENTIFIER),
        ..Default::default()
    };
    let mut document = object(DOCUMENT_IDENTIFIER, 1, document_value.encode_to_vec())?;
    document.archive_info.message_infos[0].object_references = vec![SHOW_IDENTIFIER];
    let mut show_field = FieldInfo::new(vec![2]);
    show_field.object_references = vec![SHOW_IDENTIFIER];
    document.archive_info.message_infos[0]
        .field_infos
        .push(show_field);

    let mut show = object(
        SHOW_IDENTIFIER,
        2,
        raw_show_with_slide_nodes(
            NativeSettings::absent(),
            Malformation::None,
            151,
            &[SLIDE_NODE_IDENTIFIER],
        )?,
    )?;
    show.archive_info.message_infos[0].object_references = vec![SLIDE_NODE_IDENTIFIER];
    let mut slide_tree_field = FieldInfo::new(vec![3, 2]);
    slide_tree_field.object_references = vec![SLIDE_NODE_IDENTIFIER];
    show.archive_info.message_infos[0]
        .field_infos
        .push(slide_tree_field);

    #[allow(
        deprecated,
        reason = "the fixture exercises legacy Keynote thumbnail caches"
    )]
    let node_value = kn::SlideNodeArchive {
        slide: Some(reference(SLIDE_IDENTIFIER)),
        thumbnails: vec![
            tsp::DataReference { identifier: 7_001 },
            tsp::DataReference { identifier: 7_002 },
        ],
        thumbnail_sizes: vec![
            tsp::Size {
                width: 320.0,
                height: 240.0,
            },
            tsp::Size {
                width: 160.0,
                height: 120.0,
            },
        ],
        thumbnails_are_dirty: Some(false),
        digests_for_datas_needing_download_for_thumbnail: vec!["stale-digest".to_owned()],
        is_skipped: false,
        has_builds: false,
        has_transition: false,
        ..Default::default()
    };
    let mut node = object(SLIDE_NODE_IDENTIFIER, 4, node_value.encode_to_vec())?;
    node.archive_info.message_infos[0].object_references = vec![SLIDE_IDENTIFIER];
    node.archive_info.message_infos[0].data_references = vec![7_001, 7_002];
    let mut slide_field = FieldInfo::new(vec![2]);
    slide_field.object_references = vec![SLIDE_IDENTIFIER];
    node.archive_info.message_infos[0]
        .field_infos
        .push(slide_field);
    let mut thumbnails_field = FieldInfo::new(vec![16]);
    thumbnails_field.data_references = vec![7_001, 7_002];
    node.archive_info.message_infos[0]
        .field_infos
        .push(thumbnails_field);

    let slide = object(
        SLIDE_IDENTIFIER,
        5,
        kn::SlideArchive {
            style: reference(1_004),
            transition: kn::TransitionArchive::default(),
            name: Some("Cached slide".to_owned()),
            in_document: true,
            ..Default::default()
        }
        .encode_to_vec(),
    )?;

    let mut document_objects = vec![document, show];
    let mut components = Vec::<(&str, Vec<u8>)>::new();
    if split_components {
        components.push((SLIDE_NODE_MEMBER, component(vec![node, slide])?));
    } else {
        document_objects.extend([node, slide]);
    }
    components.push((DOCUMENT_MEMBER, component(document_objects)?));
    components.push((
        "Index/Unrelated.iwa",
        component(vec![object(900, 999, b"unrelated component".to_vec())?])?,
    ));

    let mut entries = vec![("Data/sentinel.bin", b"unrelated ZIP sentinel".as_slice())];
    for (name, data) in &components {
        entries.push((*name, data.as_slice()));
    }
    entries.extend([
        (PREVIEW_MEMBERS[0], b"full preview".as_slice()),
        (PREVIEW_MEMBERS[1], b"micro preview".as_slice()),
        (PREVIEW_MEMBERS[2], b"web preview".as_slice()),
    ]);
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        Limits::default(),
    )?)
}

fn empty_show_package_bytes() -> TestResult<Vec<u8>> {
    let document = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..Default::default()
        },
        show: reference(0),
        ..Default::default()
    };
    let document_component = component(vec![object(1, 1, document.encode_to_vec())?])?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [(DOCUMENT_MEMBER, document_component.as_slice())],
        Limits::default(),
    )?)
}

fn legacy_package_bytes(flat: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(flat)?;
    let inner_entries = catalog
        .iter()
        .filter(|entry| {
            std::path::Path::new(entry.name())
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("iwa"))
        })
        .map(|entry| (entry.name(), entry.data()))
        .collect::<Vec<_>>();
    let inner =
        litchi_iwa_archive::package::to_bytes(inner_entries.iter().copied(), Limits::default())?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("legacy.key/Index.zip", inner.as_slice()),
            (
                "legacy.key/Data/sentinel.bin",
                b"legacy outer sentinel".as_slice(),
            ),
        ],
        Limits::default(),
    )?)
}

fn document_stream(package: &[u8]) -> TestResult<Vec<u8>> {
    component_stream(package, DOCUMENT_MEMBER)
}

fn component_stream(package: &[u8], component: &str) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == component)
        .ok_or_else(|| io::Error::other("missing synthetic document member"))?;
    Ok(SnappyStream::decompress(entry.data())?.into_bytes())
}

fn rewrite_component(
    package: &[u8],
    component: &str,
    mutate: impl FnOnce(&mut Archive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let stream = component_stream(package, component)?;
    let mut archive = Archive::parse(&stream)?;
    mutate(&mut archive)?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    let catalog = Catalog::from_bytes(package)?;
    Ok(
        catalog
            .reassemble_to_bytes(&[EntryEdit::new(component, &compressed)], Limits::default())?,
    )
}

fn adversarialize_show_archive_header(package: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let mut entries = Vec::with_capacity(catalog.len());
    for entry in catalog.iter() {
        let data = if entry.name() == DOCUMENT_MEMBER {
            let stream = SnappyStream::decompress(entry.data())?.into_bytes();
            let stream = rewrite_object_header(&stream, DOCUMENT_IDENTIFIER, 0)?;
            SnappyStream::compress(&rewrite_object_header(&stream, SHOW_IDENTIFIER, 1)?)?
        } else {
            entry.data().to_vec()
        };
        entries.push((entry.name().to_owned(), data));
    }
    Ok(litchi_iwa_archive::package::to_bytes(
        entries
            .iter()
            .map(|(name, data)| (name.as_str(), data.as_slice())),
        Limits::default(),
    )?)
}

fn with_overlong_object_length_prefix(
    package: &[u8],
    component: &str,
    identifier: u64,
) -> TestResult<Vec<u8>> {
    let mut stream = component_stream(package, component)?;
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
    let catalog = Catalog::from_bytes(package)?;
    Ok(
        catalog
            .reassemble_to_bytes(&[EntryEdit::new(component, &compressed)], Limits::default())?,
    )
}

#[derive(Debug, Clone, Copy)]
enum MetadataGuard {
    ShouldMerge,
    Base,
    MergeVersion,
    DiffPath,
    FieldsToRemove,
    ReadVersion,
}

fn with_metadata_guard(
    package: &[u8],
    identifier: u64,
    message_type: u32,
    guard: MetadataGuard,
) -> TestResult<Vec<u8>> {
    rewrite_component(package, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(identifier)
            .ok_or_else(|| io::Error::other("missing guarded object"))?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == message_type)
            .ok_or_else(|| io::Error::other("missing guarded message"))?;
        match guard {
            MetadataGuard::ShouldMerge => object.archive_info.should_merge = Some(true),
            MetadataGuard::Base => {
                object.archive_info.message_infos[index].base_message_index = Some(0);
            },
            MetadataGuard::MergeVersion => {
                object.archive_info.message_infos[index].diff_merge_version = vec![1];
            },
            MetadataGuard::DiffPath => {
                object.archive_info.message_infos[index].diff_field_path =
                    Some(FieldPath::new(vec![4]));
            },
            MetadataGuard::FieldsToRemove => object.archive_info.message_infos[index]
                .fields_to_remove
                .push(FieldPath::new(vec![4])),
            MetadataGuard::ReadVersion => {
                object.archive_info.message_infos[index].diff_read_version = vec![1];
            },
        }
        Ok(())
    })
}

fn rewrite_object_header(
    stream: &[u8],
    object_identifier: u64,
    target_message_index: usize,
) -> TestResult<Vec<u8>> {
    let (object_start, prefix_length, header_length) =
        object_header_location(stream, object_identifier)?;
    let header_start = object_start
        .checked_add(prefix_length)
        .ok_or_else(|| io::Error::other("synthetic header offset overflow"))?;
    let header_end = header_start
        .checked_add(header_length)
        .ok_or_else(|| io::Error::other("synthetic header range overflow"))?;
    let header = stream
        .get(header_start..header_end)
        .ok_or_else(|| io::Error::other("synthetic header is truncated"))?;
    let view = WireView::parse(header)?;
    let mut rewritten = Vec::with_capacity(header.len().saturating_add(64));
    rewritten.extend_from_slice(&fixed64_field(90, 0x1020_3040_5060_7080));
    let mut message_index = 0usize;
    for field in view.fields() {
        match (field.number(), field.wire_type()) {
            (1, 0) => {
                push_varint_width(u64::from(1_u32) << 3, 2, &mut rewritten);
                push_varint_width(object_identifier, 2, &mut rewritten);
            },
            (2, 2) if message_index == target_message_index => {
                let message = adversarial_message_info(field.payload())?;
                push_varint_width((u64::from(2_u32) << 3) | 2, 2, &mut rewritten);
                push_varint_width(u64::try_from(message.len())?, 3, &mut rewritten);
                rewritten.extend_from_slice(&message);
                message_index += 1;
            },
            (2, 2) => {
                rewritten.extend_from_slice(field.raw());
                message_index += 1;
            },
            _ => rewritten.extend_from_slice(field.raw()),
        }
    }
    rewritten.extend_from_slice(&length_delimited_field(91, b"archive-header-sentinel"));
    push_varint_width(u64::from(92_u32) << 3, 2, &mut rewritten);
    push_varint_width(17, 2, &mut rewritten);

    let output_length = stream
        .len()
        .checked_sub(prefix_length)
        .and_then(|length| length.checked_sub(header_length))
        .and_then(|length| length.checked_add(rewritten.len()))
        .and_then(|length| length.checked_add(10))
        .ok_or_else(|| io::Error::other("synthetic stream rewrite overflow"))?;
    let mut output = Vec::with_capacity(output_length);
    output.extend_from_slice(&stream[..object_start]);
    push_varint(u64::try_from(rewritten.len())?, &mut output);
    output.extend_from_slice(&rewritten);
    output.extend_from_slice(&stream[header_end..]);
    Ok(output)
}

fn adversarial_message_info(source: &[u8]) -> TestResult<Vec<u8>> {
    let view = WireView::parse(source)?;
    let mut output = Vec::with_capacity(source.len().saturating_add(32));
    output.extend_from_slice(&fixed32_field(90, 0xa1b2_c3d4));
    for field in view.fields() {
        if field.number() == 3 && field.wire_type() == 0 {
            let (length, consumed) = decode_varint_from_bytes(field.payload())?;
            if consumed != field.payload().len() {
                return Err(io::Error::other("synthetic MessageInfo length is malformed").into());
            }
            push_varint_width(u64::from(3_u32) << 3, 2, &mut output);
            push_varint_width(length, 4, &mut output);
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    output.extend_from_slice(&length_delimited_field(91, b"message-info-sentinel"));
    Ok(output)
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
            });
        object_start = header_end
            .checked_add(payload_length.ok_or_else(|| io::Error::other("payload length overflow"))?)
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

fn normalize_message_length(header: &[u8], target_index: usize) -> TestResult<Vec<u8>> {
    let view = WireView::parse(header)?;
    let mut output = Vec::with_capacity(header.len());
    let mut message_index = 0usize;
    for field in view.fields() {
        if field.number() != 2 || field.wire_type() != 2 {
            output.extend_from_slice(field.raw());
            continue;
        }
        if message_index != target_index {
            output.extend_from_slice(field.raw());
            message_index += 1;
            continue;
        }

        output.extend_from_slice(field.key());
        output.extend_from_slice(b"<message-info>");
        let message = WireView::parse(field.payload())?;
        let effective_length = message
            .fields()
            .enumerate()
            .filter_map(|(index, nested)| {
                (nested.number() == 3 && nested.wire_type() == 0).then_some(index)
            })
            .last()
            .ok_or_else(|| io::Error::other("synthetic MessageInfo has no length"))?;
        for (index, nested) in message.fields().enumerate() {
            if index == effective_length {
                output.extend_from_slice(nested.key());
                output.extend_from_slice(b"<payload-length>");
            } else {
                output.extend_from_slice(nested.raw());
            }
        }
        output.extend_from_slice(b"</message-info>");
        message_index += 1;
    }
    Ok(output)
}

fn show_messages(package: &[u8]) -> TestResult<Vec<(u32, Vec<u8>)>> {
    let stream = document_stream(package)?;
    let archive = Archive::parse(&stream)?;
    Ok(archive
        .object(2)
        .ok_or_else(|| io::Error::other("missing synthetic show object"))?
        .messages
        .iter()
        .map(|message| (message.type_, message.data.clone()))
        .collect())
}

fn show_records(package: &[u8]) -> TestResult<Vec<(u32, Vec<u8>)>> {
    let messages = show_messages(package)?;
    let payload = messages
        .iter()
        .find_map(|(type_, data)| (*type_ == 2).then_some(data.as_slice()))
        .ok_or_else(|| io::Error::other("missing synthetic show message"))?;
    Ok(WireView::parse(payload)?
        .fields()
        .map(|field| (field.number(), field.raw().to_vec()))
        .collect())
}

fn size_records(package: &[u8]) -> TestResult<Vec<(u32, Vec<u8>)>> {
    let records = show_records(package)?;
    let size = records
        .iter()
        .find_map(|(number, raw)| (*number == 4).then_some(raw.as_slice()))
        .ok_or_else(|| io::Error::other("missing synthetic show size"))?;
    let size_field = WireView::parse(size)?
        .fields()
        .next()
        .ok_or_else(|| io::Error::other("malformed synthetic show size record"))?;
    Ok(WireView::parse(size_field.payload())?
        .fields()
        .map(|field| (field.number(), field.raw().to_vec()))
        .collect())
}

fn rewrite_root(
    package: &[u8],
    mutate: impl FnOnce(&mut ArchiveObject, usize) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    rewrite_component(package, DOCUMENT_MEMBER, |archive| {
        let root = archive
            .object_mut(DOCUMENT_IDENTIFIER)
            .ok_or_else(|| io::Error::other("missing Keynote document root"))?;
        let index = root
            .messages
            .iter()
            .position(|message| message.type_ == 1)
            .ok_or_else(|| io::Error::other("missing Keynote document message"))?;
        mutate(root, index)
    })
}

fn rewrite_root_reference(
    package: &[u8],
    mutate: impl Fn(&mut Vec<u8>) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    rewrite_root(package, |root, index| {
        let document = WireView::parse(&root.messages[index].data)?;
        let mut rewritten = Vec::new();
        let mut selected = false;
        for field in document.fields() {
            if field.number() != 2 {
                rewritten.extend_from_slice(field.raw());
                continue;
            }
            if std::mem::replace(&mut selected, true) || field.wire_type() != 2 {
                return Err(io::Error::other("ambiguous root show reference").into());
            }
            let mut reference = field.payload().to_vec();
            mutate(&mut reference)?;
            litchi_iwa_common::wire::append_length_delimited_field(&mut rewritten, 2, &reference)?;
        }
        if !selected {
            return Err(io::Error::other("missing root show reference").into());
        }
        root.replace_message_preserving_header(
            index,
            RawMessage {
                type_: 1,
                data: rewritten,
            },
        )?;
        Ok(())
    })
}

fn assert_noop_then_changed_refused(bytes: &[u8]) -> TestResult<()> {
    let package = Package::from_bytes(bytes)?;
    let before = package.show_settings()?;
    let noop = package.edit_show_settings()?.commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(written(noop.package())?, bytes);

    let mut after = before;
    after.set_loop_presentation(Some(!before.loop_presentation().unwrap_or(false)));
    let changed = package.edit_show_settings()?.set(after);
    assert!(matches!(
        changed.commit(),
        Err(ShowSettingsError::InvalidSource)
    ));
    assert_eq!(written(&package)?, bytes);
    Ok(())
}

fn entry_bytes(package: &[u8], name: &str) -> TestResult<Vec<u8>> {
    Ok(Catalog::from_bytes(package)?
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or_else(|| io::Error::other("missing synthetic package entry"))?
        .data()
        .to_vec())
}

fn written(package: &Package) -> TestResult<Vec<u8>> {
    let mut output = Vec::new();
    package.write_to(&mut output)?;
    Ok(output)
}

fn object_bytes(package: &[u8], component: &str, identifier: u64) -> TestResult<Vec<u8>> {
    let stream = component_stream(package, component)?;
    let archive = Archive::parse(&stream)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing synthetic object"))?;
    let start = usize::try_from(object.header_offset)?;
    let end = usize::try_from(
        object
            .data_offset
            .checked_add(object.data_length)
            .ok_or_else(|| io::Error::other("synthetic object end overflow"))?,
    )?;
    Ok(stream[start..end].to_vec())
}

fn assert_previews_absent(package: &[u8]) -> TestResult<()> {
    let catalog = Catalog::from_bytes(package)?;
    for preview in PREVIEW_MEMBERS {
        assert!(catalog.iter().all(|entry| entry.name() != preview));
    }
    Ok(())
}

fn assert_entry_preserved(before: &[u8], after: &[u8], name: &str) -> TestResult<()> {
    let before = Catalog::from_bytes(before)?;
    let after = Catalog::from_bytes(after)?;
    let source = before
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or_else(|| io::Error::other("missing source entry"))?;
    let target = after
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or_else(|| io::Error::other("missing target entry"))?;
    assert_eq!(target.raw_name(), source.raw_name());
    assert_eq!(target.data(), source.data());
    assert_eq!(target.metadata(), source.metadata());
    assert_eq!(
        target.raw_record().local_record(),
        source.raw_record().local_record()
    );
    let source_central = source.raw_record().central_directory_record();
    let target_central = target.raw_record().central_directory_record();
    assert_eq!(target_central.len(), source_central.len());
    assert!(source_central.len() >= 46);
    assert_eq!(&target_central[..42], &source_central[..42]);
    assert_eq!(&target_central[46..], &source_central[46..]);
    Ok(())
}

fn present_settings() -> TestResult<Settings> {
    let mut settings = Settings::new(Size::new(1_920.0, 1_080.0)?);
    settings.set_slide_numbers_visible(Some(false));
    settings.set_loop_presentation(Some(true));
    settings.set_mode(Some(Mode::unknown(-7)?))?;
    settings.set_autoplay_transition_delay(Some(Seconds::new(1.25)?));
    settings.set_autoplay_build_delay(Some(Seconds::new(2.5)?));
    settings.set_idle_timer_active(Some(false));
    settings.set_idle_timer_delay(Some(Seconds::new(30.0)?));
    settings.set_automatically_plays_upon_open(Some(true));
    Ok(settings)
}

fn clear_all_settings(settings: &mut Settings) -> TestResult<()> {
    settings.set_size(Size::new(640.0, 480.0)?);
    settings.set_slide_numbers_visible(None);
    settings.set_loop_presentation(None);
    settings.set_mode(None)?;
    settings.set_autoplay_transition_delay(None);
    settings.set_autoplay_build_delay(None);
    settings.set_idle_timer_active(None);
    settings.set_idle_timer_delay(None);
    settings.set_automatically_plays_upon_open(None);
    Ok(())
}

fn assert_optional_fields(package: &[u8], expected_count: usize) -> TestResult<()> {
    let records = show_records(package)?;
    for number in OPTIONAL_SETTING_FIELDS {
        assert_eq!(
            records
                .iter()
                .filter(|(actual, _raw)| *actual == number)
                .count(),
            expected_count,
            "unexpected presence count for field {number}"
        );
    }
    Ok(())
}

fn tight_limits_for(artifacts: &[&[u8]]) -> TestResult<Limits> {
    let mut max_input_bytes = 0u64;
    let mut max_entries = 0usize;
    let mut max_entry_bytes = 0u64;
    let mut max_total_bytes = 0u64;
    let mut max_iwa_stream_bytes = 0usize;

    for artifact in artifacts {
        max_input_bytes = max_input_bytes.max(u64::try_from(artifact.len())?);
        let catalog = Catalog::from_bytes(artifact)?;
        max_entries = max_entries.max(catalog.len());
        let mut total_bytes = 0u64;
        for entry in catalog.iter() {
            let entry_bytes = u64::try_from(entry.data().len())?;
            max_entry_bytes = max_entry_bytes.max(entry_bytes);
            total_bytes = total_bytes.checked_add(entry_bytes).ok_or_else(|| {
                io::Error::other("synthetic package entry total does not fit u64")
            })?;
            if std::path::Path::new(entry.name())
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("iwa"))
            {
                let stream_bytes = SnappyStream::decompress(entry.data())?.as_bytes().len();
                // Replacing one Archive message briefly retains a slightly larger
                // encoded header before the final stream is serialized.
                max_iwa_stream_bytes = max_iwa_stream_bytes.max(stream_bytes.saturating_add(16));
            }
        }
        max_total_bytes = max_total_bytes.max(total_bytes);
    }
    max_total_bytes = max_total_bytes.max(u64::try_from(max_iwa_stream_bytes)?);

    Ok(Limits::new(
        max_input_bytes,
        max_entries,
        max_entry_bytes,
        max_total_bytes,
        max_iwa_stream_bytes,
    )?)
}

#[test]
fn every_setting_round_trips_absent_present_and_present_absent() -> TestResult<()> {
    let bytes = package_bytes(NativeSettings::absent(), Malformation::None, 41)?;
    let package = Package::from_bytes(&bytes)?;
    let absent = Settings::new(Size::new(1_024.0, 768.0)?);
    let present = present_settings()?;
    assert_eq!(package.show_settings()?, absent);
    assert_eq!(*package.show()?.settings(), absent);
    assert!(package.show()?.is_empty());
    assert!(package.slides()?.is_empty());
    assert_eq!(package.text()?, "");
    assert_optional_fields(&bytes, 0)?;

    let add = package.edit_show_settings()?;
    assert_eq!(add.settings(), absent);
    let added = add.set(present).commit()?;
    assert_eq!(*added.package().show()?.settings(), present);
    assert_eq!(added.patch().before(), absent);
    assert_eq!(added.patch().after(), present);
    let added_bytes = written(added.package())?;
    assert_optional_fields(&added_bytes, 1)?;

    let mode_record = show_records(&added_bytes)?
        .into_iter()
        .find_map(|(number, raw)| (number == 9).then_some(raw))
        .ok_or_else(|| io::Error::other("committed show has no mode record"))?;
    let mut canonical_negative = Vec::new();
    let negative_mode = u64::from_ne_bytes(i64::from(-7).to_ne_bytes());
    append_varint_field(&mut canonical_negative, 9, negative_mode)?;
    assert_eq!(mode_record, canonical_negative);
    assert_eq!(mode_record.len(), 11);

    let clear = added.package().edit_show_settings()?;
    let mut cleared_settings = clear.settings();
    clear_all_settings(&mut cleared_settings)?;
    let cleared = clear.set(cleared_settings).commit()?;
    assert_eq!(cleared_settings, Settings::new(Size::new(640.0, 480.0)?));
    assert_eq!(*cleared.package().show()?.settings(), cleared_settings);
    assert_eq!(cleared.patch().before(), present);
    assert_eq!(cleared.patch().after(), cleared_settings);
    assert_optional_fields(&written(cleared.package())?, 0)?;
    Ok(())
}

#[test]
fn exact_noop_preserves_bytes_and_has_zero_change_diagnostics() -> TestResult<()> {
    let bytes = package_bytes(NativeSettings::present(), Malformation::None, 42)?;
    let package = Package::from_bytes(&bytes)?;
    let source_settings = *package.show()?.settings();
    let commit = package
        .edit_show_settings()?
        .set(source_settings)
        .commit()?;

    assert_eq!(written(commit.package())?, bytes);
    assert!(commit.patch().is_noop());
    assert_eq!(commit.patch().before(), source_settings);
    assert_eq!(commit.patch().after(), source_settings);
    assert_eq!(
        commit.patch().source_fingerprint(),
        commit.patch().target_fingerprint()
    );
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 0);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    assert!(!commit.diagnostics().full_reparse_performed());
    Ok(())
}

#[test]
fn changed_commit_preserves_unknowns_and_the_immutable_source() -> TestResult<()> {
    let bytes = package_bytes(NativeSettings::absent(), Malformation::None, 43)?;
    let package = Package::from_bytes(&bytes)?;
    let source_copy = written(&package)?;
    let before_messages = show_messages(&bytes)?;
    let before_nonsettings = show_records(&bytes)?
        .into_iter()
        .filter(|(number, _raw)| !SETTING_FIELDS.contains(number))
        .collect::<Vec<_>>();
    let before_size_unknowns = size_records(&bytes)?
        .into_iter()
        .filter(|(number, _raw)| ![1, 2].contains(number))
        .collect::<Vec<_>>();

    let commit = package
        .edit_show_settings()?
        .set(present_settings()?)
        .commit()?;
    let after_bytes = written(commit.package())?;
    let after_messages = show_messages(&after_bytes)?;
    let after_nonsettings = show_records(&after_bytes)?
        .into_iter()
        .filter(|(number, _raw)| !SETTING_FIELDS.contains(number))
        .collect::<Vec<_>>();
    let after_size_unknowns = size_records(&after_bytes)?
        .into_iter()
        .filter(|(number, _raw)| ![1, 2].contains(number))
        .collect::<Vec<_>>();

    assert_eq!(written(&package)?, source_copy);
    assert_ne!(after_bytes, source_copy);
    assert_eq!(after_nonsettings, before_nonsettings);
    assert_eq!(after_size_unknowns, before_size_unknowns);
    assert_eq!(after_messages[0], before_messages[0]);
    assert_eq!(after_messages[2], before_messages[2]);
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert!(commit.diagnostics().full_reparse_performed());

    let before_catalog = Catalog::from_bytes(&bytes)?;
    let after_catalog = Catalog::from_bytes(&after_bytes)?;
    let mut changed_entries = 0usize;
    for (before, after) in before_catalog.iter().zip(after_catalog.iter()) {
        assert_eq!(before.name(), after.name());
        if before.data() == after.data() {
            assert_eq!(
                before.raw_record().local_record(),
                after.raw_record().local_record()
            );
            assert_eq!(
                before.raw_record().central_directory_record(),
                after.raw_record().central_directory_record()
            );
        } else {
            changed_entries += 1;
            assert_eq!(before.name(), DOCUMENT_MEMBER);
        }
    }
    assert_eq!(changed_entries, 1);
    Ok(())
}

#[test]
fn changed_commit_preserves_selected_zip_metadata_and_interleaved_members() -> TestResult<()> {
    let bytes = interleaved_adversarial_zip_package_bytes()?;
    let before_catalog = Catalog::from_bytes(&bytes)?;
    let before_order = before_catalog
        .iter()
        .map(|entry| entry.name().to_owned())
        .collect::<Vec<_>>();
    assert_eq!(
        before_order,
        ["Data/before.bin", DOCUMENT_MEMBER, "Data/after.bin"]
    );
    let before_show = before_catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("missing synthetic document member"))?;
    assert_eq!(
        before_show.metadata().local().last_modified().time(),
        0x5b7d
    );
    assert_eq!(
        before_show.metadata().local().last_modified().date(),
        0x5794
    );
    assert_eq!(
        &before_show.raw_record().central_directory_record()[4..6],
        &0x031e_u16.to_le_bytes()
    );
    assert_eq!(
        &before_show.raw_record().central_directory_record()[36..38],
        &0xa55a_u16.to_le_bytes()
    );
    assert_eq!(
        &before_show.raw_record().central_directory_record()[38..42],
        &0x81a4_0000_u32.to_le_bytes()
    );

    let package = Package::from_bytes(&bytes)?;
    let changed = package
        .edit_show_settings()?
        .set(present_settings()?)
        .commit()?;
    let changed_bytes = written(changed.package())?;
    let after_catalog = Catalog::from_bytes(&changed_bytes)?;
    let after_order = after_catalog
        .iter()
        .map(|entry| entry.name().to_owned())
        .collect::<Vec<_>>();
    assert_eq!(after_order, before_order);
    assert_entry_preserved(&bytes, &changed_bytes, "Data/before.bin")?;
    assert_entry_preserved(&bytes, &changed_bytes, "Data/after.bin")?;

    let after_show = after_catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("missing rewritten document member"))?;
    assert_ne!(after_show.data(), before_show.data());
    assert_eq!(after_show.raw_name(), before_show.raw_name());
    assert_eq!(after_show.is_opaque(), before_show.is_opaque());
    assert_eq!(
        after_show.metadata().local(),
        before_show.metadata().local()
    );
    assert_eq!(
        after_show.metadata().central(),
        before_show.metadata().central()
    );
    let before_show_central = before_show.raw_record().central_directory_record();
    let after_show_central = after_show.raw_record().central_directory_record();
    assert_eq!(after_show_central.len(), before_show_central.len());
    assert_eq!(&after_show_central[..16], &before_show_central[..16]);
    assert_eq!(&after_show_central[28..42], &before_show_central[28..42]);
    assert_eq!(&after_show_central[46..], &before_show_central[46..]);

    let before_after = before_catalog
        .iter()
        .find(|entry| entry.name() == "Data/after.bin")
        .ok_or_else(|| io::Error::other("missing trailing retained member"))?;
    let after_after = after_catalog
        .iter()
        .find(|entry| entry.name() == "Data/after.bin")
        .ok_or_else(|| io::Error::other("missing rewritten trailing member"))?;
    assert_ne!(
        &after_after.raw_record().central_directory_record()[42..46],
        &before_after.raw_record().central_directory_record()[42..46]
    );

    let restored = changed
        .package()
        .apply_show_settings(&changed.patch().inverse())?;
    assert_eq!(written(restored.package())?, bytes);
    Ok(())
}

#[test]
fn length_changing_commit_preserves_adversarial_archive_info_header() -> TestResult<()> {
    let canonical = package_bytes(NativeSettings::absent(), Malformation::None, 143)?;
    let bytes = adversarialize_show_archive_header(&canonical)?;
    let root_header = object_header(&bytes, DOCUMENT_IDENTIFIER)?;
    let before_header = object_header(&bytes, SHOW_IDENTIFIER)?;
    assert!(
        before_header
            .windows(b"archive-header-sentinel".len())
            .any(|window| window == b"archive-header-sentinel")
    );
    assert!(
        before_header
            .windows(b"message-info-sentinel".len())
            .any(|window| window == b"message-info-sentinel")
    );

    let package = Package::from_bytes(&bytes)?;
    let commit = package
        .edit_show_settings()?
        .set(present_settings()?)
        .commit()?;
    let committed_bytes = written(commit.package())?;
    let after_header = object_header(&committed_bytes, SHOW_IDENTIFIER)?;

    assert_eq!(
        object_header(&committed_bytes, DOCUMENT_IDENTIFIER)?,
        root_header,
        "the untouched root header, including unknown fields, must remain exact"
    );

    assert_ne!(after_header, before_header);
    assert_eq!(
        normalize_message_length(&after_header, 1)?,
        normalize_message_length(&before_header, 1)?,
        "the effective payload length is the only permitted ArchiveInfo change"
    );
    // The deliberately overlong root identifier, MessageInfo wrapper key and
    // effective length key remain non-canonical after the edit.
    for marker in [
        &[0x88, 0x00, 0x82, 0x00][..],
        &[0x92, 0x00][..],
        &[0x98, 0x00][..],
    ] {
        assert!(
            after_header
                .windows(marker.len())
                .any(|window| window == marker),
            "non-canonical header marker {marker:x?} was not preserved"
        );
    }

    let inverse = commit.patch().inverse();
    let restored = commit.package().apply_show_settings(&inverse)?;
    let restored_bytes = written(restored.package())?;
    assert_eq!(restored_bytes, bytes);
    assert_eq!(
        object_header(&restored_bytes, SHOW_IDENTIFIER)?,
        before_header
    );
    Ok(())
}

#[test]
fn patch_is_exact_conflict_checked_and_byte_reversible() -> TestResult<()> {
    let bytes = package_bytes(NativeSettings::absent(), Malformation::None, 44)?;
    let package = Package::from_bytes(&bytes)?;
    let before = *package.show()?.settings();
    let after = present_settings()?;
    let commit = package.edit_show_settings()?.set(after).commit()?;
    let patch = commit.patch();

    assert_eq!(patch.before(), before);
    assert_eq!(patch.after(), after);
    assert_ne!(patch.source_fingerprint(), patch.target_fingerprint());
    assert!(!patch.is_noop());
    let applied = package.apply_show_settings(patch)?;
    assert_eq!(written(applied.package())?, written(commit.package())?);
    assert_eq!(applied.patch(), patch);

    let inverse = patch.inverse();
    assert_eq!(inverse.before(), after);
    assert_eq!(inverse.after(), before);
    assert_eq!(inverse.source_fingerprint(), patch.target_fingerprint());
    assert_eq!(inverse.target_fingerprint(), patch.source_fingerprint());
    assert_eq!(inverse.inverse(), patch.clone());
    let restored = commit.package().apply_show_settings(&inverse)?;
    assert_eq!(written(restored.package())?, bytes);
    let replayed = restored.package().apply_show_settings(patch)?;
    assert_eq!(written(replayed.package())?, written(commit.package())?);

    assert!(matches!(
        commit.package().apply_show_settings(patch),
        Err(ShowSettingsError::PatchConflict)
    ));
    assert!(matches!(
        package.apply_show_settings(&inverse),
        Err(ShowSettingsError::PatchConflict)
    ));

    let catalog = Catalog::from_bytes(&bytes)?;
    let tampered_bytes = catalog.reassemble_to_bytes(
        &[EntryEdit::new(
            "Data/sentinel.bin",
            b"tampered unrelated sentinel",
        )],
        Limits::default(),
    )?;
    let tampered = Package::from_bytes(&tampered_bytes)?;
    assert_eq!(tampered.show_settings()?, before);
    assert!(matches!(
        tampered.apply_show_settings(patch),
        Err(ShowSettingsError::PatchConflict)
    ));

    let unrelated_bytes = package_bytes(NativeSettings::absent(), Malformation::None, 45)?;
    let unrelated = Package::from_bytes(&unrelated_bytes)?;
    assert!(matches!(
        unrelated.apply_show_settings(patch),
        Err(ShowSettingsError::PatchConflict)
    ));
    Ok(())
}

#[test]
fn nondefault_tight_limits_survive_read_edit_commit_and_apply() -> TestResult<()> {
    let bytes = package_bytes(NativeSettings::absent(), Malformation::None, 49)?;
    let baseline = Package::from_bytes(&bytes)?;
    let baseline_commit = baseline
        .edit_show_settings()?
        .set(present_settings()?)
        .commit()?;
    let target_bytes = written(baseline_commit.package())?;

    let physical = tight_limits_for(&[&bytes, &target_bytes])?;
    let semantic = SemanticLimits::new(2, 1, 16, 1, 1, 64)?;
    let options = ReadOptions::new(physical, semantic);
    assert_ne!(options, ReadOptions::default());
    assert_eq!(physical.max_entries(), 2);
    let expected_max_input = u64::try_from(bytes.len().max(target_bytes.len()))?;
    assert_eq!(physical.max_input_bytes(), expected_max_input);

    let package = Package::from_bytes_with_options(&bytes, options)?;
    assert_eq!(package.read_options(), options);
    assert_eq!(*package.show()?.settings(), Settings::default());

    let edit = package.edit_show_settings()?;
    assert_eq!(edit.settings(), Settings::default());
    let commit = edit.set(present_settings()?).commit()?;
    assert_eq!(commit.package().read_options(), options);
    assert_eq!(written(commit.package())?, target_bytes);

    let applied = package.apply_show_settings(commit.patch())?;
    assert_eq!(applied.package().read_options(), options);
    assert_eq!(written(applied.package())?, target_bytes);

    let restored = commit
        .package()
        .apply_show_settings(&commit.patch().inverse())?;
    assert_eq!(restored.package().read_options(), options);
    assert_eq!(written(restored.package())?, bytes);
    Ok(())
}

#[test]
fn patch_debug_redacts_physical_identity_and_exact_bytes() -> TestResult<()> {
    let bytes = package_bytes(NativeSettings::absent(), Malformation::None, 50)?;
    let package = Package::from_bytes(&bytes)?;
    let commit = package
        .edit_show_settings()?
        .set(present_settings()?)
        .commit()?;
    let patch = commit.patch();
    let debug = format!("{patch:?}");
    let edit_debug = format!("{:?}", package.edit_show_settings()?);
    let commit_debug = format!("{commit:?}");
    let error_debug = format!("{:?}", ShowSettingsError::InvalidSource);
    let error_display = ShowSettingsError::InvalidSource.to_string();

    assert!(debug.starts_with("Patch"));
    assert!(debug.contains("before"));
    assert!(debug.contains("after"));
    for private in [
        "Index/Document.iwa",
        "Data/sentinel.bin",
        "before-show-sentinel",
        "after-show-sentinel",
        "unknown-show",
        "unrelated opaque sentinel",
        "source_bytes",
        "target_bytes",
        "fingerprint",
        "identifier",
    ] {
        for rendered in [
            &debug,
            &edit_debug,
            &commit_debug,
            &error_debug,
            &error_display,
        ] {
            assert!(
                !rendered.contains(private),
                "public formatting leaked private marker {private:?}: {rendered}"
            );
        }
    }
    for fingerprint in [patch.source_fingerprint(), patch.target_fingerprint()] {
        assert!(!debug.contains(&fingerprint.to_string()));
        assert!(!debug.contains(&format!("{fingerprint:x}")));
    }
    assert!(!debug.contains(&PRIVATE_NATIVE_IDENTIFIER.to_string()));
    Ok(())
}

#[test]
fn root_reference_and_ownership_metadata_fail_closed() -> TestResult<()> {
    let bytes = package_bytes(NativeSettings::absent(), Malformation::None, 152)?;
    let external = rewrite_root(&bytes, |root, index| {
        let mut document = kn::DocumentArchive::decode(root.messages[index].data.as_slice())?;
        document.show.deprecated_is_external = Some(true);
        root.replace_message_preserving_header(
            index,
            RawMessage {
                type_: 1,
                data: document.encode_to_vec(),
            },
        )?;
        Ok(())
    })?;
    let external_package = Package::from_bytes(&external)?;
    assert!(matches!(
        external_package.show_settings(),
        Err(ShowSettingsError::InvalidSource)
    ));
    assert!(matches!(
        external_package.edit_show_settings(),
        Err(ShowSettingsError::InvalidSource)
    ));
    assert_eq!(written(&external_package)?, external);

    let duplicate_external = rewrite_root_reference(&bytes, |reference| {
        append_varint_field(reference, 3, 1)?;
        Ok(())
    })?;
    let duplicate_type = rewrite_root_reference(&bytes, |reference| {
        append_varint_field(reference, 2, 7)?;
        Ok(())
    })?;
    let wrong_wire = rewrite_root_reference(&bytes, |reference| {
        let view = WireView::parse(reference)?;
        let mut rewritten = Vec::new();
        for field in view.fields().filter(|field| field.number() != 3) {
            rewritten.extend_from_slice(field.raw());
        }
        litchi_iwa_common::wire::append_length_delimited_field(&mut rewritten, 3, &[0])?;
        *reference = rewritten;
        Ok(())
    })?;
    let noncanonical = rewrite_root_reference(&bytes, |reference| {
        let view = WireView::parse(reference)?;
        let mut rewritten = Vec::new();
        for field in view.fields().filter(|field| field.number() != 3) {
            rewritten.extend_from_slice(field.raw());
        }
        rewritten.extend_from_slice(&[0x18, 0x80, 0x00]);
        *reference = rewritten;
        Ok(())
    })?;
    for malformed in [duplicate_external, duplicate_type, wrong_wire, noncanonical] {
        match Package::from_bytes(&malformed) {
            Err(_strict_ingress_rejection) => {},
            Ok(package) => {
                assert!(matches!(
                    package.show_settings(),
                    Err(ShowSettingsError::InvalidSource)
                ));
                assert!(matches!(
                    package.edit_show_settings(),
                    Err(ShowSettingsError::InvalidSource)
                ));
                assert_eq!(written(&package)?, malformed);
            },
        }
    }

    let missing = rewrite_root(&bytes, |root, index| {
        root.archive_info.message_infos[index]
            .object_references
            .clear();
        Ok(())
    })?;
    let duplicate = rewrite_root(&bytes, |root, index| {
        root.archive_info.message_infos[index]
            .object_references
            .push(SHOW_IDENTIFIER);
        Ok(())
    })?;
    let wrong_path = rewrite_root(&bytes, |root, index| {
        root.archive_info.message_infos[index].field_infos[0].path = FieldPath::new(vec![3]);
        Ok(())
    })?;
    for malformed in [missing, duplicate, wrong_path] {
        let package = Package::from_bytes(&malformed)?;
        assert!(matches!(
            package.show_settings(),
            Err(ShowSettingsError::InvalidSource)
        ));
        assert!(matches!(
            package.edit_show_settings(),
            Err(ShowSettingsError::InvalidSource)
        ));
        assert_eq!(written(&package)?, malformed);
    }

    let unrelated = rewrite_component(&bytes, DOCUMENT_MEMBER, |archive| {
        let root = archive
            .object_mut(DOCUMENT_IDENTIFIER)
            .ok_or_else(|| io::Error::other("missing Keynote document root"))?;
        root.archive_info.message_infos[0]
            .object_references
            .push(SIBLING_IDENTIFIER);
        let mut unrelated = FieldInfo::new(vec![77]);
        unrelated.object_references = vec![SIBLING_IDENTIFIER];
        root.archive_info.message_infos[0]
            .field_infos
            .push(unrelated);
        archive.objects.push(object(
            SIBLING_IDENTIFIER,
            999,
            b"unrelated root reference".to_vec(),
        )?);
        Ok(())
    })?;
    let root_bytes = object_bytes(&unrelated, DOCUMENT_MEMBER, DOCUMENT_IDENTIFIER)?;
    let package = Package::from_bytes(&unrelated)?;
    let mut after = package.show_settings()?;
    after.set_loop_presentation(Some(true));
    let changed = package.edit_show_settings()?.set(after).commit()?;
    let changed_bytes = written(changed.package())?;
    assert_eq!(
        object_bytes(&changed_bytes, DOCUMENT_MEMBER, DOCUMENT_IDENTIFIER)?,
        root_bytes
    );
    Ok(())
}

#[test]
fn merge_diff_and_outer_prefix_admission_is_lazy_for_noop_but_strict_for_changes() -> TestResult<()>
{
    let bytes = package_bytes(NativeSettings::absent(), Malformation::None, 153)?;
    let with_sibling = rewrite_component(&bytes, DOCUMENT_MEMBER, |archive| {
        archive.objects.push(object(
            SIBLING_IDENTIFIER,
            999,
            b"untouched sibling".to_vec(),
        )?);
        Ok(())
    })?;
    let overlong_show =
        with_overlong_object_length_prefix(&with_sibling, DOCUMENT_MEMBER, SHOW_IDENTIFIER)?;
    let overlong_sibling =
        with_overlong_object_length_prefix(&with_sibling, DOCUMENT_MEMBER, SIBLING_IDENTIFIER)?;
    for adversarial in [overlong_show, overlong_sibling] {
        assert_noop_then_changed_refused(&adversarial)?;
    }

    for identifier in [DOCUMENT_IDENTIFIER, SHOW_IDENTIFIER] {
        let message_type = if identifier == DOCUMENT_IDENTIFIER {
            1
        } else {
            2
        };
        for guard in [
            MetadataGuard::ShouldMerge,
            MetadataGuard::Base,
            MetadataGuard::MergeVersion,
            MetadataGuard::DiffPath,
            MetadataGuard::FieldsToRemove,
            MetadataGuard::ReadVersion,
        ] {
            let adversarial = with_metadata_guard(&bytes, identifier, message_type, guard)?;
            assert_noop_then_changed_refused(&adversarial)?;
        }
    }
    Ok(())
}

#[test]
fn checked_native_fixture_noop_playback_and_rendering_changes_follow_cache_policy() -> TestResult<()>
{
    let bytes = std::fs::read(native_fixture_path())?;
    let package = Package::from_bytes(&bytes)?;
    let before = package.show_settings()?;
    let catalog = Catalog::from_bytes(&bytes)?;
    let slide_components = catalog
        .iter()
        .filter_map(|entry| {
            let basename = entry.name().rsplit('/').next()?;
            (basename.starts_with("Slide") && basename.ends_with(".iwa"))
                .then(|| entry.name().to_owned())
        })
        .collect::<Vec<_>>();
    assert!(!slide_components.is_empty());
    for preview in PREVIEW_MEMBERS {
        assert!(catalog.iter().any(|entry| entry.name() == preview));
    }

    let noop = package.edit_show_settings()?.set(before).commit()?;
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert_eq!(noop.diagnostics().deleted_previews(), 0);
    assert!(!noop.diagnostics().full_reparse_performed());
    assert_eq!(written(noop.package())?, bytes);

    let mut playback = before;
    playback.set_loop_presentation(Some(!before.loop_presentation().unwrap_or(false)));
    let playback = package.edit_show_settings()?.set(playback).commit()?;
    let playback_bytes = written(playback.package())?;
    assert_eq!(playback.diagnostics().deleted_previews(), 0);
    for name in slide_components.iter().map(String::as_str) {
        assert_entry_preserved(&bytes, &playback_bytes, name)?;
    }
    for preview in PREVIEW_MEMBERS {
        assert_entry_preserved(&bytes, &playback_bytes, preview)?;
    }
    assert_eq!(
        written(
            playback
                .package()
                .apply_show_settings(&playback.patch().inverse())?
                .package(),
        )?,
        bytes
    );

    let mut resized = before;
    resized.set_size(Size::new(
        before.size().width() + 1.0,
        before.size().height(),
    )?);
    let mut numbered = before;
    numbered.set_slide_numbers_visible(Some(!before.slide_numbers_visible().unwrap_or(false)));
    for invalidating in [resized, numbered] {
        let changed = package.edit_show_settings()?.set(invalidating).commit()?;
        let changed_bytes = written(changed.package())?;
        assert_eq!(changed.diagnostics().deleted_previews(), 3);
        assert_previews_absent(&changed_bytes)?;
        for name in slide_components.iter().map(String::as_str) {
            assert_entry_preserved(&bytes, &changed_bytes, name)?;
        }
        assert_eq!(
            written(
                changed
                    .package()
                    .apply_show_settings(&changed.patch().inverse())?
                    .package(),
            )?,
            bytes
        );
    }
    Ok(())
}

#[test]
fn size_change_deletes_root_previews_but_preserves_slide_caches_in_same_and_split_components()
-> TestResult<()> {
    for split in [false, true] {
        let bytes = cache_package_bytes(split)?;
        let package = Package::from_bytes(&bytes)?;
        let before = package.show_settings()?;
        let node_component = if split {
            SLIDE_NODE_MEMBER
        } else {
            DOCUMENT_MEMBER
        };
        let node_bytes = object_bytes(&bytes, node_component, SLIDE_NODE_IDENTIFIER)?;
        let mut after = before;
        after.set_size(Size::new(1_280.0, 720.0)?);
        let changed = package.edit_show_settings()?.set(after).commit()?;
        let changed_bytes = written(changed.package())?;
        assert_eq!(changed.package().show_settings()?, after);
        assert!(changed.diagnostics().changed());
        assert_eq!(changed.diagnostics().touched_components(), 1);
        assert_eq!(changed.diagnostics().deleted_previews(), 3);
        assert!(changed.diagnostics().full_reparse_performed());
        assert_eq!(
            object_bytes(&changed_bytes, node_component, SLIDE_NODE_IDENTIFIER,)?,
            node_bytes
        );
        assert_previews_absent(&changed_bytes)?;
        assert_eq!(
            entry_bytes(&changed_bytes, "Data/sentinel.bin")?,
            entry_bytes(&bytes, "Data/sentinel.bin")?
        );
        assert_eq!(
            entry_bytes(&changed_bytes, "Index/Unrelated.iwa")?,
            entry_bytes(&bytes, "Index/Unrelated.iwa")?
        );

        let applied = package.apply_show_settings(changed.patch())?;
        assert_eq!(written(applied.package())?, changed_bytes);
        let restored = changed
            .package()
            .apply_show_settings(&changed.patch().inverse())?;
        assert_eq!(written(restored.package())?, bytes);
    }
    Ok(())
}

#[test]
fn slide_number_visibility_change_deletes_previews_but_preserves_split_slide_component()
-> TestResult<()> {
    let bytes = cache_package_bytes(true)?;
    let package = Package::from_bytes(&bytes)?;
    let slide_component = entry_bytes(&bytes, SLIDE_NODE_MEMBER)?;
    let mut after = package.show_settings()?;
    after.set_slide_numbers_visible(Some(true));
    let changed = package.edit_show_settings()?.set(after).commit()?;
    let changed_bytes = written(changed.package())?;
    assert_eq!(changed.diagnostics().touched_components(), 1);
    assert_eq!(changed.diagnostics().deleted_previews(), 3);
    assert_previews_absent(&changed_bytes)?;
    assert_eq!(
        entry_bytes(&changed_bytes, SLIDE_NODE_MEMBER)?,
        slide_component
    );
    let restored = changed
        .package()
        .apply_show_settings(&changed.patch().inverse())?;
    assert_eq!(written(restored.package())?, bytes);
    Ok(())
}

#[test]
fn playback_only_change_preserves_thumbnail_and_preview_caches_exactly() -> TestResult<()> {
    let bytes = cache_package_bytes(true)?;
    let package = Package::from_bytes(&bytes)?;
    let node_component = entry_bytes(&bytes, SLIDE_NODE_MEMBER)?;
    let previews = PREVIEW_MEMBERS
        .iter()
        .map(|name| entry_bytes(&bytes, name))
        .collect::<TestResult<Vec<_>>>()?;
    let edit = package.edit_show_settings()?;
    let mut settings = edit.settings();
    settings.set_loop_presentation(Some(true));
    settings.set_mode(Some(Mode::SelfPlaying))?;
    settings.set_autoplay_transition_delay(Some(Seconds::new(3.25)?));
    settings.set_autoplay_build_delay(Some(Seconds::new(4.5)?));
    settings.set_idle_timer_active(Some(true));
    settings.set_idle_timer_delay(Some(Seconds::new(120.0)?));
    settings.set_automatically_plays_upon_open(Some(true));
    let changed = edit.set(settings).commit()?;
    let changed_bytes = written(changed.package())?;
    assert_eq!(changed.diagnostics().touched_components(), 1);
    assert_eq!(changed.diagnostics().deleted_previews(), 0);
    assert_eq!(
        entry_bytes(&changed_bytes, SLIDE_NODE_MEMBER)?,
        node_component
    );
    for (name, expected) in PREVIEW_MEMBERS.iter().zip(previews) {
        assert_eq!(entry_bytes(&changed_bytes, name)?, expected);
        assert_entry_preserved(&bytes, &changed_bytes, name)?;
    }
    assert_eq!(written(&package)?, bytes);
    Ok(())
}

#[test]
fn huge_nested_size_opaque_field_hits_byte_work_limit_before_publication() -> TestResult<()> {
    let bytes = package_bytes(NativeSettings::absent(), Malformation::None, 154)?;
    let huge = rewrite_component(&bytes, DOCUMENT_MEMBER, |archive| {
        let show = archive
            .object_mut(SHOW_IDENTIFIER)
            .ok_or_else(|| io::Error::other("missing show object"))?;
        let index = show
            .messages
            .iter()
            .position(|message| message.type_ == 2)
            .ok_or_else(|| io::Error::other("missing show message"))?;
        let source = WireView::parse(&show.messages[index].data)?;
        let mut payload = Vec::new();
        let mut selected_size = false;
        for field in source.fields() {
            if field.number() != 4 {
                payload.extend_from_slice(field.raw());
                continue;
            }
            if std::mem::replace(&mut selected_size, true) || field.wire_type() != 2 {
                return Err(io::Error::other("ambiguous show size field").into());
            }
            let mut size = field.payload().to_vec();
            litchi_iwa_common::wire::append_length_delimited_field(
                &mut size,
                199,
                &vec![0xa5; 16_100_000],
            )?;
            litchi_iwa_common::wire::append_length_delimited_field(&mut payload, 4, &size)?;
        }
        if !selected_size {
            return Err(io::Error::other("missing show size field").into());
        }
        show.replace_message_preserving_header(
            index,
            RawMessage {
                type_: 2,
                data: payload,
            },
        )?;
        Ok(())
    })?;
    let package = Package::from_bytes(&huge)?;
    let field_count = show_records(&huge)?.len();
    assert!(field_count < 100);
    assert!(matches!(
        package.edit_show_settings(),
        Err(ShowSettingsError::LimitExceeded {
            kind: ShowSettingsLimitKind::WireWork,
            ..
        })
    ));
    assert_eq!(written(&package)?, huge);
    Ok(())
}

#[test]
fn legacy_source_accepts_exact_noop_and_refuses_changed_reassembly() -> TestResult<()> {
    let flat = package_bytes(NativeSettings::absent(), Malformation::None, 46)?;
    let legacy = legacy_package_bytes(&flat)?;
    let package = Package::from_bytes(&legacy)?;
    let settings = package.show_settings()?;

    let noop_commit = package.edit_show_settings()?.set(settings).commit()?;
    assert!(noop_commit.patch().is_noop());
    assert_eq!(written(noop_commit.package())?, legacy);
    let applied = package.apply_show_settings(noop_commit.patch())?;
    assert_eq!(written(applied.package())?, legacy);

    let changed = package.edit_show_settings()?;
    let mut settings = changed.settings();
    settings.set_automatically_plays_upon_open(Some(true));
    assert!(matches!(
        changed.set(settings).commit(),
        Err(ShowSettingsError::UnsupportedSource)
    ));
    Ok(())
}

#[test]
fn null_show_reader_matches_full_semantics_and_only_exact_noop_is_editable() -> TestResult<()> {
    let bytes = empty_show_package_bytes()?;
    let package = Package::from_bytes(&bytes)?;
    assert_eq!(package.show_settings()?, Settings::default());
    assert_eq!(*package.show()?.settings(), Settings::default());
    assert!(package.show()?.is_empty());
    assert!(package.slides()?.is_empty());
    assert!(package.semantic_snapshot()?.slides().is_empty());
    assert_eq!(package.text()?, "");

    let noop = package.edit_show_settings()?.commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(written(noop.package())?, bytes);
    assert_eq!(
        written(package.apply_show_settings(noop.patch())?.package())?,
        bytes
    );

    let changed = package.edit_show_settings()?;
    let mut settings = changed.settings();
    settings.set_loop_presentation(Some(true));
    assert!(matches!(
        changed.set(settings).commit(),
        Err(ShowSettingsError::UnsupportedSource)
    ));
    Ok(())
}

#[test]
fn invalid_semantics_and_malformed_settings_wire_are_rejected() -> TestResult<()> {
    let mut invalid_size = NativeSettings::absent();
    invalid_size.width = 0.0;
    let mut invalid_delay = NativeSettings::present();
    invalid_delay.autoplay_build_delay = Some(-0.25);
    for settings in [invalid_size, invalid_delay] {
        let bytes = package_bytes(settings, Malformation::None, 47)?;
        let package = Package::from_bytes(&bytes)?;
        assert!(matches!(package.show(), Err(ReadError::Decode(_))));
        assert!(matches!(
            package.show_settings(),
            Err(ShowSettingsError::InvalidSource)
        ));
        assert!(matches!(
            package.edit_show_settings(),
            Err(ShowSettingsError::InvalidSource)
        ));
    }

    for malformation in [
        Malformation::DuplicateSlideNumbers,
        Malformation::WrongSlideNumbersWireType,
    ] {
        let bytes = package_bytes(NativeSettings::present(), malformation, 48)?;
        let package = Package::from_bytes(&bytes)?;
        assert!(matches!(
            package.show(),
            Err(ReadError::InvalidFormat(_) | ReadError::Decode(_))
        ));
        assert!(matches!(
            package.show_settings(),
            Err(ShowSettingsError::InvalidSource)
        ));
        assert!(matches!(
            package.edit_show_settings(),
            Err(ShowSettingsError::InvalidSource)
        ));
    }

    assert!(Size::new(f32::NAN, 768.0).is_err());
    assert!(Seconds::new(-1.0).is_err());
    let mut canonical = Settings::default();
    assert!(canonical.set_mode(Some(Mode::Unknown(1))).is_err());
    assert_eq!(canonical.mode(), None);
    Ok(())
}

#[test]
fn public_show_settings_transaction_values_are_send_sync() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<Settings>();
    assert_send_sync::<Mode>();
    assert_send_sync::<Size>();
    assert_send_sync::<ShowSettingsEdit<'static>>();
    assert_send_sync::<ShowSettingsCommit>();
    assert_send_sync::<ShowSettingsPatch>();
    assert_send_sync::<ShowSettingsDiagnostics>();
    assert_send_sync::<ShowSettingsError>();
    assert_send_sync::<ShowSettingsLimitKind>();
    assert_send_sync::<Arc<[u8]>>();
}
