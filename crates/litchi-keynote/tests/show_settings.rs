use std::io;
use std::sync::Arc;

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::{
    decode_varint_from_bytes,
    wire::{WireView, append_varint_field},
};
use litchi_iwa_core::{Archive, ArchiveInfo, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsk, tsp};
use litchi_keynote::{
    Limits, Mode, Package, ReadError, ReadOptions, Seconds, SemanticLimits, Settings,
    ShowSettingsError, Size,
};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const PRIVATE_NATIVE_IDENTIFIER: u64 = 7_777_777_777_777_779;
const OPTIONAL_SETTING_FIELDS: [u32; 8] = [6, 8, 9, 10, 11, 15, 16, 18];
const SETTING_FIELDS: [u32; 9] = [4, 6, 8, 9, 10, 11, 15, 16, 18];

type TestResult<T> = Result<T, Box<dyn std::error::Error>>;

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
    let canonical = kn::ShowArchive {
        ui_state: Some(reference(PRIVATE_NATIVE_IDENTIFIER)),
        theme: reference(80),
        slide_tree: kn::SlideTreeArchive::default(),
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
        2,
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
    let document_component = component(vec![object(1, 1, document.encode_to_vec())?, show])?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Data/sentinel.bin", b"unrelated opaque sentinel".as_slice()),
            (DOCUMENT_MEMBER, document_component.as_slice()),
        ],
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
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("missing synthetic document member"))?;
    Ok(SnappyStream::decompress(entry.data())?.into_bytes())
}

fn adversarialize_show_archive_header(package: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let mut entries = Vec::with_capacity(catalog.len());
    for entry in catalog.iter() {
        let data = if entry.name() == DOCUMENT_MEMBER {
            let stream = SnappyStream::decompress(entry.data())?.into_bytes();
            SnappyStream::compress(&rewrite_object_header(&stream, 2, 1)?)?
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

    let mut add = package.edit_show_settings()?;
    assert_eq!(*add.settings(), absent);
    add.set_settings(present)?;
    let added = add.commit()?;
    assert_eq!(*added.package().show()?.settings(), present);
    assert_eq!(added.patch().before(), absent);
    assert_eq!(added.patch().after(), present);
    assert_optional_fields(added.package().source_bytes(), 1)?;

    let mode_record = show_records(added.package().source_bytes())?
        .into_iter()
        .find_map(|(number, raw)| (number == 9).then_some(raw))
        .ok_or_else(|| io::Error::other("committed show has no mode record"))?;
    let mut canonical_negative = Vec::new();
    let negative_mode = u64::from_ne_bytes(i64::from(-7).to_ne_bytes());
    append_varint_field(&mut canonical_negative, 9, negative_mode)?;
    assert_eq!(mode_record, canonical_negative);
    assert_eq!(mode_record.len(), 11);

    let mut clear = added.package().edit_show_settings()?;
    clear_all_settings(clear.settings_mut())?;
    let cleared_settings = *clear.settings();
    let cleared = clear.commit()?;
    assert_eq!(cleared_settings, Settings::new(Size::new(640.0, 480.0)?));
    assert_eq!(*cleared.package().show()?.settings(), cleared_settings);
    assert_eq!(cleared.patch().before(), present);
    assert_eq!(cleared.patch().after(), cleared_settings);
    assert_optional_fields(cleared.package().source_bytes(), 0)?;
    Ok(())
}

#[test]
fn exact_noop_reuses_source_arc_and_has_zero_change_diagnostics() -> TestResult<()> {
    let bytes = package_bytes(NativeSettings::present(), Malformation::None, 42)?;
    let package = Package::from_bytes(&bytes)?;
    let source_pointer = package.source_bytes().as_ptr();
    let source_settings = *package.show()?.settings();
    let mut edit = package.edit_show_settings()?;
    edit.set_settings(source_settings)?;
    let commit = edit.commit()?;

    assert_eq!(commit.package().source_bytes().as_ptr(), source_pointer);
    assert_eq!(commit.package().source_bytes(), bytes);
    assert!(commit.patch().is_noop());
    assert_eq!(commit.patch().before(), source_settings);
    assert_eq!(commit.patch().after(), source_settings);
    assert_eq!(
        commit.patch().source_fingerprint(),
        commit.patch().target_fingerprint()
    );
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 0);
    assert!(!commit.diagnostics().full_reparse_performed());
    Ok(())
}

#[test]
fn changed_commit_preserves_unknowns_and_the_immutable_source() -> TestResult<()> {
    let bytes = package_bytes(NativeSettings::absent(), Malformation::None, 43)?;
    let package = Package::from_bytes(&bytes)?;
    let source_pointer = package.source_bytes().as_ptr();
    let source_copy = package.source_bytes().to_vec();
    let before_messages = show_messages(&bytes)?;
    let before_nonsettings = show_records(&bytes)?
        .into_iter()
        .filter(|(number, _raw)| !SETTING_FIELDS.contains(number))
        .collect::<Vec<_>>();
    let before_size_unknowns = size_records(&bytes)?
        .into_iter()
        .filter(|(number, _raw)| ![1, 2].contains(number))
        .collect::<Vec<_>>();

    let mut edit = package.edit_show_settings()?;
    edit.set_settings(present_settings()?)?;
    let commit = edit.commit()?;
    let after_bytes = commit.package().source_bytes();
    let after_messages = show_messages(after_bytes)?;
    let after_nonsettings = show_records(after_bytes)?
        .into_iter()
        .filter(|(number, _raw)| !SETTING_FIELDS.contains(number))
        .collect::<Vec<_>>();
    let after_size_unknowns = size_records(after_bytes)?
        .into_iter()
        .filter(|(number, _raw)| ![1, 2].contains(number))
        .collect::<Vec<_>>();

    assert_eq!(package.source_bytes().as_ptr(), source_pointer);
    assert_eq!(package.source_bytes(), source_copy);
    assert_ne!(after_bytes, source_copy);
    assert_eq!(after_nonsettings, before_nonsettings);
    assert_eq!(after_size_unknowns, before_size_unknowns);
    assert_eq!(after_messages[0], before_messages[0]);
    assert_eq!(after_messages[2], before_messages[2]);
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert!(commit.diagnostics().full_reparse_performed());

    let before_catalog = Catalog::from_bytes(&bytes)?;
    let after_catalog = Catalog::from_bytes(after_bytes)?;
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
fn length_changing_commit_preserves_adversarial_archive_info_header() -> TestResult<()> {
    let canonical = package_bytes(NativeSettings::absent(), Malformation::None, 143)?;
    let bytes = adversarialize_show_archive_header(&canonical)?;
    let before_header = object_header(&bytes, 2)?;
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
    let mut edit = package.edit_show_settings()?;
    edit.set_settings(present_settings()?)?;
    let commit = edit.commit()?;
    let after_header = object_header(commit.package().source_bytes(), 2)?;

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
    assert_eq!(restored.package().source_bytes(), bytes);
    assert_eq!(
        object_header(restored.package().source_bytes(), 2)?,
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
    let mut edit = package.edit_show_settings()?;
    edit.set_settings(after)?;
    let commit = edit.commit()?;
    let patch = commit.patch();

    assert_eq!(patch.before(), before);
    assert_eq!(patch.after(), after);
    assert_ne!(patch.source_fingerprint(), patch.target_fingerprint());
    assert!(!patch.is_noop());
    let applied = package.apply_show_settings(patch)?;
    assert_eq!(
        applied.package().source_bytes(),
        commit.package().source_bytes()
    );
    assert_eq!(applied.patch(), patch);

    let inverse = patch.inverse();
    assert_eq!(inverse.before(), after);
    assert_eq!(inverse.after(), before);
    assert_eq!(inverse.source_fingerprint(), patch.target_fingerprint());
    assert_eq!(inverse.target_fingerprint(), patch.source_fingerprint());
    assert_eq!(inverse.inverse(), patch.clone());
    let restored = commit.package().apply_show_settings(&inverse)?;
    assert_eq!(restored.package().source_bytes(), bytes);
    let replayed = restored.package().apply_show_settings(patch)?;
    assert_eq!(
        replayed.package().source_bytes(),
        commit.package().source_bytes()
    );

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
    let mut baseline_edit = baseline.edit_show_settings()?;
    baseline_edit.set_settings(present_settings()?)?;
    let baseline_commit = baseline_edit.commit()?;
    let target_bytes = baseline_commit.package().source_bytes();

    let physical = tight_limits_for(&[&bytes, target_bytes])?;
    let semantic = SemanticLimits::new(2, 1, 16, 1, 1, 64)?;
    let options = ReadOptions::new(physical, semantic);
    assert_ne!(options, ReadOptions::default());
    assert_eq!(physical.max_entries(), 2);
    let expected_max_input = u64::try_from(bytes.len().max(target_bytes.len()))?;
    assert_eq!(physical.max_input_bytes(), expected_max_input);

    let package = Package::from_bytes_with_options(&bytes, options)?;
    assert_eq!(package.read_options(), options);
    assert_eq!(*package.show()?.settings(), Settings::default());

    let mut edit = package.edit_show_settings()?;
    assert_eq!(*edit.settings(), Settings::default());
    edit.set_settings(present_settings()?)?;
    let commit = edit.commit()?;
    assert_eq!(commit.package().read_options(), options);
    assert_eq!(commit.package().source_bytes(), target_bytes);

    let applied = package.apply_show_settings(commit.patch())?;
    assert_eq!(applied.package().read_options(), options);
    assert_eq!(applied.package().source_bytes(), target_bytes);

    let restored = commit
        .package()
        .apply_show_settings(&commit.patch().inverse())?;
    assert_eq!(restored.package().read_options(), options);
    assert_eq!(restored.package().source_bytes(), bytes);
    Ok(())
}

#[test]
fn patch_debug_redacts_physical_identity_and_exact_bytes() -> TestResult<()> {
    let bytes = package_bytes(NativeSettings::absent(), Malformation::None, 50)?;
    let package = Package::from_bytes(&bytes)?;
    let mut edit = package.edit_show_settings()?;
    edit.set_settings(present_settings()?)?;
    let commit = edit.commit()?;
    let patch = commit.patch();
    let debug = format!("{patch:?}");

    assert!(debug.starts_with("ShowSettingsPatch"));
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
        assert!(
            !debug.contains(private),
            "patch Debug leaked private marker {private:?}: {debug}"
        );
    }
    for fingerprint in [patch.source_fingerprint(), patch.target_fingerprint()] {
        assert!(!debug.contains(&fingerprint.to_string()));
        assert!(!debug.contains(&format!("{fingerprint:x}")));
    }
    assert!(!debug.contains(&PRIVATE_NATIVE_IDENTIFIER.to_string()));
    Ok(())
}

#[test]
fn legacy_source_accepts_exact_noop_and_refuses_changed_reassembly() -> TestResult<()> {
    let flat = package_bytes(NativeSettings::absent(), Malformation::None, 46)?;
    let legacy = legacy_package_bytes(&flat)?;
    let package = Package::from_bytes(&legacy)?;
    let settings = package.show_settings()?;

    let mut noop = package.edit_show_settings()?;
    noop.set_settings(settings)?;
    let noop_commit = noop.commit()?;
    assert!(noop_commit.patch().is_noop());
    assert_eq!(noop_commit.package().source_bytes(), legacy);
    let applied = package.apply_show_settings(noop_commit.patch())?;
    assert_eq!(applied.package().source_bytes(), legacy);

    let mut changed = package.edit_show_settings()?;
    changed
        .settings_mut()
        .set_automatically_plays_upon_open(Some(true));
    assert!(matches!(
        changed.commit(),
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
    assert_eq!(noop.package().source_bytes(), bytes);
    assert_eq!(
        package
            .apply_show_settings(noop.patch())?
            .package()
            .source_bytes(),
        bytes
    );

    let mut changed = package.edit_show_settings()?;
    changed.settings_mut().set_loop_presentation(Some(true));
    assert!(matches!(
        changed.commit(),
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
    assert_send_sync::<litchi_keynote::ShowSettingsEdit<'static>>();
    assert_send_sync::<litchi_keynote::ShowSettingsCommit>();
    assert_send_sync::<litchi_keynote::ShowSettingsPatch>();
    assert_send_sync::<litchi_keynote::ShowSettingsDiagnostics>();
    assert_send_sync::<ShowSettingsError>();
    assert_send_sync::<litchi_keynote::ShowSettingsLimitKind>();
    assert_send_sync::<Arc<[u8]>>();
}
