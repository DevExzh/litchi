//! Strict, source-preserving Keynote chart-axis-title integration coverage.
//!
//! The fixture below follows the producer graph used by source-built iWork
//! charts: a chart drawable owns a title stand-in, chart non-style, primary
//! and secondary axis style/non-style objects, and records the same private
//! graph in its IWA message metadata.  The tests deliberately mutate only
//! this test-local fixture; no production source is involved.

use std::io;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::wire::{WireView, append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{
    Archive, ArchiveObject, FieldInfo, FieldPath, FieldType, RawMessage, SnappyStream,
};
use litchi_iwa_protos::{kn, tsa, tsch, tsd, tsk, tsp, tss};
use litchi_keynote::{
    Axis, ChartAxisTitleError, ChartAxisTitleLimitKind, ChartSelector, Package, Position,
    ReadOptions, SemanticLimits, SlideSelector,
};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const FOREIGN_MEMBER: &str = "Index/Foreign.iwa";
const STYLESHEET_MEMBER: &str = "Index/DocumentStylesheet.iwa";
const METADATA_OBJECT: u64 = 300;
const DOCUMENT_COMPONENT: u64 = 1;
const UNRELATED_COMPONENT: u64 = 2;
const FOREIGN_COMPONENT: u64 = 3;
const STYLESHEET_COMPONENT: u64 = 4;
const SLIDE_NODE: u64 = 3;
const SLIDE: u64 = 4;
const CHARTS: [u64; 2] = [100, 101];
const TITLES: [u64; 2] = [110, 111];
const CHART_NON_STYLES: [u64; 2] = [120, 121];
const CATEGORY_STYLES: [u64; 2] = [130, 131];
const CATEGORY_NON_STYLES: [u64; 2] = [140, 141];
const VALUE_STYLES: [u64; 2] = [150, 151];
const VALUE_NON_STYLES: [u64; 2] = [160, 161];
const SECONDARY_VALUE_STYLES: [u64; 2] = [170, 171];
const SECONDARY_VALUE_NON_STYLES: [u64; 2] = [180, 181];
const FOREIGN_OBJECT: u64 = 900;
const STYLESHEET_OBJECT: u64 = 910;
const CHART_MESSAGE_TYPE: u32 = 5_021;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
const CHART_NON_STYLE_MESSAGE_TYPE: u32 = 5_023;
const AXIS_STYLE_MESSAGE_TYPE: u32 = 5_026;
const AXIS_NON_STYLE_MESSAGE_TYPE: u32 = 5_027;
const METADATA_MESSAGE_TYPE: u32 = 11_006;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const GENERATED_EXTENSION_FIELD: u32 = 10_000;
const UNKNOWN_OUTER_FIELD: u32 = 4_000;
const UNKNOWN_GENERATED_FIELD: u32 = 4_001;
const METADATA_UNKNOWN_FIELD: u32 = 4_002;
const SUPPORTS_PRIMARY_FEATURE_FIELD: u32 = 10_001;
const SUPPORTS_SECONDARY_FEATURE_FIELD: u32 = 10_002;
const METADATA_UNKNOWN_MARKER: &[u8] = b"chart-axis-title metadata extension";
const PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

type TestResult<T> = Result<T, Box<dyn std::error::Error>>;

#[derive(Clone, Copy)]
struct AxisState<'a> {
    visible: Option<bool>,
    title: Option<&'a [u8]>,
}

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
    let mut value = object(identifier, type_, data)?;
    value.archive_info.message_infos[0].object_references = references;
    Ok(value)
}

fn chart_payload(
    chart: usize,
    category_non_style: Option<u64>,
    value_non_style: Option<u64>,
    secondary_value_non_style: Option<u64>,
) -> TestResult<Vec<u8>> {
    let drawable = tsd::DrawableArchive {
        parent: Some(reference(SLIDE)),
        title: Some(reference(TITLES[chart])),
        ..tsd::DrawableArchive::default()
    };
    let chart_data = tsch::ChartArchive {
        chart_type: Some(1),
        series_direction: Some(1),
        chart_non_style: Some(reference(CHART_NON_STYLES[chart])),
        value_axis_styles: secondary_value_non_style
            .map(|_| {
                vec![
                    reference(VALUE_STYLES[chart]),
                    reference(SECONDARY_VALUE_STYLES[chart]),
                ]
            })
            .unwrap_or_else(|| vec![reference(VALUE_STYLES[chart])]),
        value_axis_nonstyles: value_non_style
            .into_iter()
            .chain(secondary_value_non_style)
            .map(reference)
            .collect(),
        category_axis_styles: category_non_style
            .map(|_| vec![reference(CATEGORY_STYLES[chart])])
            .unwrap_or_default(),
        category_axis_nonstyles: category_non_style
            .map(|identifier| vec![reference(identifier)])
            .unwrap_or_default(),
        ..tsch::ChartArchive::default()
    };
    let mut payload = Vec::new();
    append_length_delimited_field(&mut payload, 1, &drawable.encode_to_vec())?;
    append_length_delimited_field(
        &mut payload,
        GENERATED_EXTENSION_FIELD,
        &chart_data.encode_to_vec(),
    )?;
    Ok(payload)
}

fn axis_payload(state: AxisState<'_>, axis: Axis, unknown: bool) -> TestResult<Vec<u8>> {
    let mut generated = Vec::new();
    if unknown {
        append_varint_field(&mut generated, UNKNOWN_GENERATED_FIELD, 73)?;
    }
    let (visible_field, title_field) = match axis {
        Axis::Category => (13, 15),
        Axis::Value => (14, 16),
    };
    if let Some(visible) = state.visible {
        append_varint_field(&mut generated, visible_field, u64::from(visible))?;
    }
    if let Some(title) = state.title {
        append_length_delimited_field(&mut generated, title_field, title)?;
    }
    let mut payload = tsch::ChartAxisNonStyleArchive {
        super_: Some(tss::StyleArchive {
            stylesheet: Some(reference(81)),
            ..tss::StyleArchive::default()
        }),
    }
    .encode_to_vec();
    append_length_delimited_field(&mut payload, GENERATED_EXTENSION_FIELD, &generated)?;
    append_varint_field(&mut payload, SUPPORTS_PRIMARY_FEATURE_FIELD, 1)?;
    append_varint_field(&mut payload, SUPPORTS_SECONDARY_FEATURE_FIELD, 1)?;
    if unknown {
        append_length_delimited_field(&mut payload, UNKNOWN_OUTER_FIELD, b"opaque axis bytes")?;
    }
    Ok(payload)
}

fn axis_style_payload(unknown: bool) -> TestResult<Vec<u8>> {
    let mut payload = tsch::ChartAxisStyleArchive {
        super_: Some(tss::StyleArchive {
            stylesheet: Some(reference(81)),
            ..tss::StyleArchive::default()
        }),
    }
    .encode_to_vec();
    let mut generated = tsch::generated::ChartAxisStyleArchive {
        tschchartaxisvalueshowaxis: Some(true),
        ..tsch::generated::ChartAxisStyleArchive::default()
    }
    .encode_to_vec();
    if unknown {
        append_varint_field(&mut generated, UNKNOWN_GENERATED_FIELD, 91)?;
    }
    append_length_delimited_field(&mut payload, GENERATED_EXTENSION_FIELD, &generated)?;
    append_varint_field(&mut payload, SUPPORTS_PRIMARY_FEATURE_FIELD, 1)?;
    if unknown {
        append_length_delimited_field(&mut payload, UNKNOWN_OUTER_FIELD, b"opaque axis style")?;
    }
    Ok(payload)
}

fn chart_non_style_payload(title: &str) -> TestResult<Vec<u8>> {
    let mut generated = Vec::new();
    append_varint_field(&mut generated, 21, 1)?;
    append_length_delimited_field(&mut generated, 23, title.as_bytes())?;
    let mut payload = tsch::ChartNonStyleArchive {
        super_: Some(tss::StyleArchive {
            stylesheet: Some(reference(81)),
            ..tss::StyleArchive::default()
        }),
    }
    .encode_to_vec();
    append_length_delimited_field(&mut payload, GENERATED_EXTENSION_FIELD, &generated)?;
    Ok(payload)
}

fn component(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

fn chart_object_references(chart: usize) -> Vec<u64> {
    vec![
        TITLES[chart],
        CHART_NON_STYLES[chart],
        CATEGORY_STYLES[chart],
        CATEGORY_NON_STYLES[chart],
        VALUE_STYLES[chart],
        VALUE_NON_STYLES[chart],
        SECONDARY_VALUE_STYLES[chart],
        SECONDARY_VALUE_NON_STYLES[chart],
    ]
}

fn all_document_object_ids() -> Vec<u64> {
    let mut identifiers = vec![1, 2, SLIDE_NODE, SLIDE];
    for chart in 0..CHARTS.len() {
        identifiers.extend([
            CHARTS[chart],
            TITLES[chart],
            CHART_NON_STYLES[chart],
            CATEGORY_STYLES[chart],
            CATEGORY_NON_STYLES[chart],
            VALUE_STYLES[chart],
            VALUE_NON_STYLES[chart],
            SECONDARY_VALUE_STYLES[chart],
            SECONDARY_VALUE_NON_STYLES[chart],
        ]);
    }
    identifiers
}

fn synthetic_package_with_states(
    category: [AxisState<'_>; 2],
    value: [AxisState<'_>; 2],
    secondary_value: [AxisState<'_>; 2],
) -> TestResult<Vec<u8>> {
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
        object_with_references(SLIDE, 5, slide.encode_to_vec(), CHARTS.to_vec())?,
        object(80, 10, Vec::new())?,
        object(81, 9_002, Vec::new())?,
        object(90, 9_003, Vec::new())?,
    ];
    for chart in 0..CHARTS.len() {
        objects.push(object_with_references(
            CHARTS[chart],
            CHART_MESSAGE_TYPE,
            chart_payload(
                chart,
                Some(CATEGORY_NON_STYLES[chart]),
                Some(VALUE_NON_STYLES[chart]),
                Some(SECONDARY_VALUE_NON_STYLES[chart]),
            )?,
            chart_object_references(chart),
        )?);
        objects.push(object(
            TITLES[chart],
            STANDIN_MESSAGE_TYPE,
            tsd::StandinCaptionArchive::default().encode_to_vec(),
        )?);
        objects.push(object(
            CHART_NON_STYLES[chart],
            CHART_NON_STYLE_MESSAGE_TYPE,
            chart_non_style_payload(if chart == 0 { "Revenue" } else { "Costs" })?,
        )?);
        objects.push(object(
            CATEGORY_STYLES[chart],
            AXIS_STYLE_MESSAGE_TYPE,
            axis_style_payload(true)?,
        )?);
        objects.push(object(
            CATEGORY_NON_STYLES[chart],
            AXIS_NON_STYLE_MESSAGE_TYPE,
            axis_payload(category[chart], Axis::Category, true)?,
        )?);
        objects.push(object(
            VALUE_STYLES[chart],
            AXIS_STYLE_MESSAGE_TYPE,
            axis_style_payload(true)?,
        )?);
        objects.push(object(
            VALUE_NON_STYLES[chart],
            AXIS_NON_STYLE_MESSAGE_TYPE,
            axis_payload(value[chart], Axis::Value, true)?,
        )?);
        objects.push(object(
            SECONDARY_VALUE_STYLES[chart],
            AXIS_STYLE_MESSAGE_TYPE,
            axis_style_payload(true)?,
        )?);
        objects.push(object(
            SECONDARY_VALUE_NON_STYLES[chart],
            AXIS_NON_STYLE_MESSAGE_TYPE,
            axis_payload(secondary_value[chart], Axis::Value, true)?,
        )?);
    }
    let document_component = component(objects)?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Data/sentinel.bin", b"unrelated ZIP sentinel".as_slice()),
            (PREVIEWS[0], b"large preview".as_slice()),
            (PREVIEWS[1], b"micro preview".as_slice()),
            (PREVIEWS[2], b"web preview".as_slice()),
            (DOCUMENT_MEMBER, document_component.as_slice()),
        ],
        Limits::default(),
    )?)
}

fn synthetic_package() -> TestResult<Vec<u8>> {
    synthetic_package_with_states(
        [
            AxisState {
                visible: Some(true),
                title: Some(b"Month"),
            },
            AxisState {
                visible: Some(true),
                title: None,
            },
        ],
        [
            AxisState {
                visible: Some(true),
                title: Some(b"Revenue"),
            },
            AxisState {
                visible: Some(false),
                title: Some(b"stale value"),
            },
        ],
        [
            AxisState {
                visible: Some(true),
                title: Some(b"Secondary value"),
            },
            AxisState {
                visible: Some(false),
                title: Some(b"secondary stale"),
            },
        ],
    )
}

fn synthetic_package_without_category_extension() -> TestResult<Vec<u8>> {
    let source = synthetic_package_with_states(
        [
            AxisState {
                visible: Some(true),
                title: Some(b"Month"),
            },
            AxisState {
                visible: Some(true),
                title: Some(b"Other category"),
            },
        ],
        [
            AxisState {
                visible: Some(true),
                title: Some(b"Revenue"),
            },
            AxisState {
                visible: Some(false),
                title: Some(b"stale value"),
            },
        ],
        [
            AxisState {
                visible: Some(true),
                title: Some(b"Secondary value"),
            },
            AxisState {
                visible: Some(false),
                title: Some(b"secondary stale"),
            },
        ],
    )?;
    with_axis_payload(
        &source,
        CATEGORY_NON_STYLES[0],
        axis_payload_without_extension(true)?,
    )
}

fn axis_payload_without_extension(unknown: bool) -> TestResult<Vec<u8>> {
    let mut payload = tsch::ChartAxisNonStyleArchive {
        super_: Some(tss::StyleArchive {
            stylesheet: Some(reference(81)),
            ..tss::StyleArchive::default()
        }),
    }
    .encode_to_vec();
    append_varint_field(&mut payload, SUPPORTS_PRIMARY_FEATURE_FIELD, 1)?;
    append_varint_field(&mut payload, SUPPORTS_SECONDARY_FEATURE_FIELD, 1)?;
    if unknown {
        append_length_delimited_field(&mut payload, UNKNOWN_OUTER_FIELD, b"opaque axis bytes")?;
    }
    Ok(payload)
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
    locator: &str,
    save_token: u64,
    object_identifiers: &[u64],
) -> TestResult<Vec<u8>> {
    let mut payload = tsp::ComponentInfo {
        identifier,
        preferred_locator: locator.to_owned(),
        locator: Some(locator.to_owned()),
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
        METADATA_UNKNOWN_FIELD,
        METADATA_UNKNOWN_MARKER,
    )?;
    Ok(payload)
}

fn metadata_payload(last_identifier: u64, include_foreign: bool) -> TestResult<Vec<u8>> {
    let document = metadata_component_payload(
        DOCUMENT_COMPONENT,
        "Document",
        10,
        &all_document_object_ids(),
    )?;
    let unrelated = metadata_component_payload(UNRELATED_COMPONENT, "Unrelated", 7, &[901])?;
    let mut payload = Vec::new();
    append_varint_field(&mut payload, 1, last_identifier)?;
    append_length_delimited_field(&mut payload, 3, &document)?;
    append_length_delimited_field(&mut payload, 3, &unrelated)?;
    if include_foreign {
        let foreign =
            metadata_component_payload(FOREIGN_COMPONENT, "Foreign", 8, &[FOREIGN_OBJECT])?;
        append_length_delimited_field(&mut payload, 3, &foreign)?;
    }
    append_varint_field(&mut payload, 8, 10)?;
    let versioned = tsp::ComponentInfo {
        identifier: DOCUMENT_COMPONENT,
        preferred_locator: "Document".to_owned(),
        locator: Some("Document".to_owned()),
        save_token: Some(3),
        object_uuid_map_entries: vec![metadata_uuid_entry(902)],
        ..tsp::ComponentInfo::default()
    }
    .encode_to_vec();
    append_length_delimited_field(&mut payload, 11, &versioned)?;
    append_length_delimited_field(
        &mut payload,
        METADATA_UNKNOWN_FIELD,
        METADATA_UNKNOWN_MARKER,
    )?;
    Ok(payload)
}

fn synthetic_metadata_package(include_foreign: bool) -> TestResult<Vec<u8>> {
    let source = synthetic_package()?;
    let metadata = component(vec![object(
        METADATA_OBJECT,
        METADATA_MESSAGE_TYPE,
        metadata_payload(2_000, include_foreign)?,
    )?])?;
    let foreign = if include_foreign {
        Some(component(vec![object(FOREIGN_OBJECT, 9_000, Vec::new())?])?)
    } else {
        None
    };
    let catalog = Catalog::from_bytes(&source)?;
    let mut entries = catalog
        .iter()
        .map(|entry| (entry.name(), entry.data()))
        .collect::<Vec<_>>();
    entries.push((METADATA_MEMBER, metadata.as_slice()));
    if let Some(foreign) = foreign.as_deref() {
        entries.push((FOREIGN_MEMBER, foreign));
    }
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

fn component_stream(package: &[u8], member: &str) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or_else(|| io::Error::other("missing synthetic component"))?;
    Ok(SnappyStream::decompress(entry.data())?.into_bytes())
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

fn replace_component_stream(source: &[u8], member: &str, archive: Archive) -> TestResult<Vec<u8>> {
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(Catalog::from_bytes(source)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            member,
            &compressed,
        )],
        Limits::default(),
    )?)
}

fn message_payload(package: &[u8], identifier: u64, type_: u32) -> TestResult<Vec<u8>> {
    let archive = Archive::parse(&document_stream(package)?)?;
    archive
        .object(identifier)
        .and_then(|object| {
            object
                .messages
                .iter()
                .find(|message| message.type_ == type_)
        })
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("missing synthetic message").into())
}

fn nested_field_raw(payload: &[u8], outer: u32, nested: u32) -> TestResult<Vec<u8>> {
    let extension = WireView::parse(payload)?
        .fields()
        .find(|field| field.number() == outer)
        .ok_or_else(|| io::Error::other("missing synthetic nested extension"))?;
    WireView::parse(extension.payload())?
        .fields()
        .find(|field| field.number() == nested)
        .map(|field| field.raw().to_vec())
        .ok_or_else(|| io::Error::other("missing synthetic nested field").into())
}

fn raw_fields(payload: &[u8], number: u32) -> TestResult<Vec<Vec<u8>>> {
    Ok(WireView::parse(payload)?
        .fields()
        .filter(|field| field.number() == number)
        .map(|field| field.raw().to_vec())
        .collect())
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
        .ok_or_else(|| io::Error::other("missing synthetic object"))?;
    let message = object
        .messages
        .iter_mut()
        .find(|message| message.type_ == type_)
        .ok_or_else(|| io::Error::other("missing synthetic message"))?;
    message.data = data;
    replace_document_stream(source, archive)
}

fn with_axis_payload(source: &[u8], identifier: u64, data: Vec<u8>) -> TestResult<Vec<u8>> {
    with_document_message_payload(source, identifier, AXIS_NON_STYLE_MESSAGE_TYPE, data)
}

fn with_chart_payload(source: &[u8], chart: usize, data: Vec<u8>) -> TestResult<Vec<u8>> {
    with_document_message_payload(source, CHARTS[chart], CHART_MESSAGE_TYPE, data)
}

fn with_axis_role_alias(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut archive = Archive::parse(&document_stream(source)?)?;
    let axis = archive
        .object_mut(CATEGORY_NON_STYLES[0])
        .ok_or_else(|| io::Error::other("missing category axis"))?;
    let data = axis.messages[0].data.clone();
    axis.push_message(RawMessage {
        type_: AXIS_STYLE_MESSAGE_TYPE,
        data,
    })?;
    replace_document_stream(source, archive)
}

fn with_foreign_inbound(source: &[u8], target: u64, mode: ForeignInbound) -> TestResult<Vec<u8>> {
    let mut archive = Archive::parse(&component_stream(source, FOREIGN_MEMBER)?)?;
    let foreign = archive
        .object_mut(FOREIGN_OBJECT)
        .ok_or_else(|| io::Error::other("missing foreign object"))?;
    let info = &mut foreign.archive_info.message_infos[0];
    match mode {
        ForeignInbound::Object => info.object_references.push(target),
        ForeignInbound::Data => info.data_references.push(target),
        ForeignInbound::Field => info.field_infos.push(FieldInfo {
            path: FieldPath::new(vec![99]),
            r#type: Some(FieldType::ObjectReference),
            object_references: vec![target],
            ..FieldInfo::default()
        }),
        ForeignInbound::FieldData => info.field_infos.push(FieldInfo {
            path: FieldPath::new(vec![100]),
            r#type: Some(FieldType::DataReference),
            data_references: vec![target],
            ..FieldInfo::default()
        }),
    }
    replace_component_stream(source, FOREIGN_MEMBER, archive)
}

fn with_stylesheet_registration(source: &[u8], target: u64) -> TestResult<Vec<u8>> {
    let source = with_metadata(source, |metadata| {
        metadata.components.push(tsp::ComponentInfo {
            identifier: STYLESHEET_COMPONENT,
            preferred_locator: "DocumentStylesheet".to_owned(),
            locator: Some("DocumentStylesheet".to_owned()),
            external_references: vec![tsp::ComponentExternalReference {
                component_identifier: DOCUMENT_COMPONENT,
                object_identifier: Some(target),
                is_weak: None,
            }],
            object_uuid_map_entries: vec![metadata_uuid_entry(STYLESHEET_OBJECT)],
            ..tsp::ComponentInfo::default()
        });
    })?;
    let stylesheet = component(vec![object_with_references(
        STYLESHEET_OBJECT,
        STYLESHEET_MESSAGE_TYPE,
        Vec::new(),
        vec![target],
    )?])?;
    let catalog = Catalog::from_bytes(&source)?;
    let mut entries = catalog
        .iter()
        .map(|entry| (entry.name(), entry.data()))
        .collect::<Vec<_>>();
    entries.push((STYLESHEET_MEMBER, stylesheet.as_slice()));
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        Limits::default(),
    )?)
}

fn with_native_stylesheet_layout(source: &[u8]) -> TestResult<Vec<u8>> {
    let moved_identifiers = [CHART_NON_STYLES[0], VALUE_NON_STYLES[0]];
    let source = with_metadata(source, |metadata| {
        let document = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == DOCUMENT_COMPONENT)
            .expect("document metadata component");
        document
            .object_uuid_map_entries
            .retain(|entry| !moved_identifiers.contains(&entry.identifier));
        document
            .external_references
            .extend(moved_identifiers.iter().copied().map(|identifier| {
                tsp::ComponentExternalReference {
                    component_identifier: STYLESHEET_COMPONENT,
                    object_identifier: Some(identifier),
                    is_weak: None,
                }
            }));
        metadata.components.push(tsp::ComponentInfo {
            identifier: STYLESHEET_COMPONENT,
            preferred_locator: "DocumentStylesheet".to_owned(),
            locator: Some("DocumentStylesheet".to_owned()),
            object_uuid_map_entries: moved_identifiers
                .iter()
                .copied()
                .chain([STYLESHEET_OBJECT])
                .map(metadata_uuid_entry)
                .collect(),
            ..tsp::ComponentInfo::default()
        });
    })?;

    let mut document = Archive::parse(&document_stream(&source)?)?;
    let mut moved = Vec::new();
    for identifier in moved_identifiers {
        let position = document
            .objects
            .iter()
            .position(|object| object.archive_info.identifier == Some(identifier))
            .ok_or_else(|| io::Error::other("missing object for native stylesheet layout"))?;
        moved.push(document.objects.remove(position));
    }
    let source = replace_document_stream(&source, document)?;
    moved.push(object_with_references(
        STYLESHEET_OBJECT,
        STYLESHEET_MESSAGE_TYPE,
        Vec::new(),
        vec![VALUE_NON_STYLES[0]],
    )?);
    let stylesheet = component(moved)?;
    let catalog = Catalog::from_bytes(&source)?;
    let mut entries = catalog
        .iter()
        .map(|entry| (entry.name(), entry.data()))
        .collect::<Vec<_>>();
    entries.push((STYLESHEET_MEMBER, stylesheet.as_slice()));
    Ok(litchi_iwa_archive::package::to_bytes(
        entries,
        Limits::default(),
    )?)
}

#[derive(Clone, Copy)]
enum ForeignInbound {
    Object,
    Data,
    Field,
    FieldData,
}

fn with_metadata(
    source: &[u8],
    rewrite: impl FnOnce(&mut tsp::PackageMetadata),
) -> TestResult<Vec<u8>> {
    let mut archive = Archive::parse(&metadata_stream(source)?)?;
    let object = archive
        .object_mut(METADATA_OBJECT)
        .ok_or_else(|| io::Error::other("missing metadata object"))?;
    let message = object
        .messages
        .iter_mut()
        .find(|message| message.type_ == METADATA_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("missing metadata message"))?;
    let mut metadata = tsp::PackageMetadata::decode(message.data.as_slice())?;
    rewrite(&mut metadata);
    message.data = metadata.encode_to_vec();
    replace_metadata_stream(source, archive)
}

fn append_duplicate_field(payload: &[u8], number: u32) -> TestResult<Vec<u8>> {
    let duplicate = WireView::parse(payload)?
        .fields()
        .find(|field| field.number() == number)
        .map(|field| field.raw().to_vec())
        .ok_or_else(|| io::Error::other("missing field to duplicate"))?;
    let mut output = payload.to_vec();
    output.extend_from_slice(&duplicate);
    Ok(output)
}

fn replace_first_field(payload: &[u8], number: u32, replacement: &[u8]) -> TestResult<Vec<u8>> {
    let fields = WireView::parse(payload)?;
    let mut output = Vec::with_capacity(payload.len() + replacement.len());
    let mut replaced = false;
    for field in fields.fields() {
        if !replaced && field.number() == number {
            output.extend_from_slice(replacement);
            replaced = true;
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    if !replaced {
        return Err(io::Error::other("missing field to replace").into());
    }
    Ok(output)
}

fn length_delimited_field_with_key_width(number: u32, payload: &[u8], key_width: usize) -> Vec<u8> {
    let mut output = Vec::new();
    push_varint_width((u64::from(number) << 3) | 2, key_width, &mut output);
    push_varint_width(payload.len() as u64, 1, &mut output);
    output.extend_from_slice(payload);
    output
}

fn varint_field(number: u32, value: u64) -> Vec<u8> {
    let mut output = Vec::new();
    append_varint_field(&mut output, number, value).expect("small hostile varint fits");
    output
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

fn with_axis_extension_variant(source: &[u8], variant: AxisWireVariant) -> TestResult<Vec<u8>> {
    let original = message_payload(source, CATEGORY_NON_STYLES[0], AXIS_NON_STYLE_MESSAGE_TYPE)?;
    let extension = WireView::parse(&original)?
        .fields()
        .find(|field| field.number() == GENERATED_EXTENSION_FIELD)
        .ok_or_else(|| io::Error::other("missing generated axis extension"))?;
    let replacement = match variant {
        AxisWireVariant::DuplicateOuter => {
            return with_axis_payload(
                source,
                CATEGORY_NON_STYLES[0],
                append_duplicate_field(&original, GENERATED_EXTENSION_FIELD)?,
            );
        },
        AxisWireVariant::WrongOuterWire => varint_field(GENERATED_EXTENSION_FIELD, 1),
        AxisWireVariant::NonCanonicalOuter => {
            length_delimited_field_with_key_width(GENERATED_EXTENSION_FIELD, extension.payload(), 4)
        },
        AxisWireVariant::DuplicateTitle => {
            let generated = append_duplicate_field(extension.payload(), 15)?;
            length_delimited_field_with_key_width(GENERATED_EXTENSION_FIELD, &generated, 3)
        },
        AxisWireVariant::InvalidUtf8 => {
            let mut generated = Vec::new();
            append_varint_field(&mut generated, 13, 1)?;
            append_length_delimited_field(&mut generated, 15, &[0xff])?;
            length_delimited_field_with_key_width(GENERATED_EXTENSION_FIELD, &generated, 3)
        },
    };
    let payload = replace_first_field(&original, GENERATED_EXTENSION_FIELD, &replacement)?;
    with_axis_payload(source, CATEGORY_NON_STYLES[0], payload)
}

#[derive(Clone, Copy)]
enum AxisWireVariant {
    DuplicateOuter,
    WrongOuterWire,
    NonCanonicalOuter,
    DuplicateTitle,
    InvalidUtf8,
}

fn with_chart_axis_variant(source: &[u8], variant: ChartAxisVariant) -> TestResult<Vec<u8>> {
    let original = message_payload(source, CHARTS[0], CHART_MESSAGE_TYPE)?;
    let chart_extension = WireView::parse(&original)?
        .fields()
        .find(|field| field.number() == GENERATED_EXTENSION_FIELD)
        .ok_or_else(|| io::Error::other("missing chart extension"))?;
    let mut chart_payload = chart_extension.payload().to_vec();
    match variant {
        ChartAxisVariant::MissingCategory => {
            let chart = tsch::ChartArchive::decode(chart_payload.as_slice())?;
            chart_payload = tsch::ChartArchive {
                category_axis_styles: Vec::new(),
                category_axis_nonstyles: Vec::new(),
                ..chart
            }
            .encode_to_vec();
        },
        ChartAxisVariant::MissingValue => {
            let chart = tsch::ChartArchive::decode(chart_payload.as_slice())?;
            chart_payload = tsch::ChartArchive {
                value_axis_styles: Vec::new(),
                value_axis_nonstyles: Vec::new(),
                ..chart
            }
            .encode_to_vec();
        },
        ChartAxisVariant::DuplicateCategory => {
            append_length_delimited_field(
                &mut chart_payload,
                16,
                &reference(CATEGORY_NON_STYLES[0]).encode_to_vec(),
            )?;
        },
        ChartAxisVariant::AliasedRoles => {
            let chart = tsch::ChartArchive {
                category_axis_styles: vec![reference(CATEGORY_STYLES[0])],
                category_axis_nonstyles: vec![reference(VALUE_NON_STYLES[0])],
                value_axis_styles: vec![reference(VALUE_STYLES[0])],
                value_axis_nonstyles: vec![reference(VALUE_NON_STYLES[0])],
                chart_non_style: Some(reference(CHART_NON_STYLES[0])),
                ..tsch::ChartArchive::default()
            };
            chart_payload = chart.encode_to_vec();
        },
    }
    let replacement = length_delimited_field_with_key_width(
        GENERATED_EXTENSION_FIELD,
        &chart_payload,
        chart_extension.key().len(),
    );
    let payload = replace_first_field(&original, GENERATED_EXTENSION_FIELD, &replacement)?;
    with_chart_payload(source, 0, payload)
}

#[derive(Clone, Copy)]
enum ChartAxisVariant {
    MissingCategory,
    MissingValue,
    DuplicateCategory,
    AliasedRoles,
}

fn assert_axis_rejected_atomically(source: &[u8], axis: Axis) -> TestResult<()> {
    // Every hostile fixture below is physically valid and intentionally keeps
    // package ingress lazy.  Requiring the package to open here prevents a
    // malformed-input regression from making this helper vacuously pass before
    // the selector/edit boundary is exercised.
    let package = Package::from_bytes(source)?;
    let before = exact_bytes(&package)?;
    let result = package
        .edit_slide_chart_axis_title(0usize, 0usize, axis)
        .and_then(|edit| edit.set("hostile replacement"))
        .and_then(|edit| edit.commit());
    assert!(result.is_err());
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

fn assert_locality(before: &[u8], after: &[u8]) -> TestResult<()> {
    let source = Catalog::from_bytes(before)?;
    let target = Catalog::from_bytes(after)?;
    let source_document = source
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("missing source document"))?;
    let target_document = target
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("missing target document"))?;
    assert_ne!(source_document.data(), target_document.data());
    for name in ["Data/sentinel.bin", METADATA_MEMBER, FOREIGN_MEMBER] {
        let source_entry = source.iter().find(|entry| entry.name() == name);
        let target_entry = target.iter().find(|entry| entry.name() == name);
        assert_eq!(
            source_entry.map(|entry| entry.data()),
            target_entry.map(|entry| entry.data()),
            "unrelated component changed: {name}"
        );
    }
    for preview in PREVIEWS {
        assert!(
            target.iter().all(|entry| entry.name() != preview),
            "preview survived changed axis title: {preview}"
        );
    }
    Ok(())
}

#[test]
fn primary_category_and_value_titles_use_semantic_selectors() -> TestResult<()> {
    let source = synthetic_metadata_package(false)?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.slide_chart_axis_title("Charts", "Revenue", Axis::Category)?,
        Some("Month".to_owned())
    );
    assert_eq!(
        package.slide_chart_axis_title(
            SlideSelector::index(0),
            ChartSelector::index(0),
            Axis::Value
        )?,
        Some("Revenue".to_owned())
    );

    let noop = package
        .edit_slide_chart_axis_title(0usize, 0usize, Axis::Category)?
        .set("Month")?
        .commit()?;
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(exact_bytes(noop.package())?, source);
    let reapplied_noop = package.apply_slide_chart_axis_title(noop.patch())?;
    assert!(reapplied_noop.patch().is_noop());
    assert!(!reapplied_noop.diagnostics().changed());
    assert_eq!(exact_bytes(reapplied_noop.package())?, source);

    let category = package.edit_slide_chart_axis_title("Charts", "Revenue", Axis::Category)?;
    assert_eq!(category.slide_position(), Position::new(0));
    assert_eq!(category.chart_position(), Position::new(0));
    assert_eq!(category.before(), Some("Month"));
    assert_eq!(category.after(), Some("Month"));
    let category = category.set("Month name")?.commit()?;
    assert_eq!(
        category
            .package()
            .slide_chart_axis_title(0usize, 0usize, Axis::Category)?,
        Some("Month name".to_owned())
    );
    assert_eq!(
        category
            .package()
            .slide_chart_axis_title(0usize, 0usize, Axis::Value)?,
        Some("Revenue".to_owned())
    );
    assert_locality(&source, &exact_bytes(category.package())?)?;
    assert!(category.diagnostics().changed());
    assert!(category.diagnostics().full_reparse_performed());
    assert!(category.diagnostics().touched_components() >= 1);
    assert_eq!(category.diagnostics().deleted_previews(), PREVIEWS.len());

    let value = package
        .edit_slide_chart_axis_title("Charts", "Revenue", Axis::Value)?
        .set("Gross revenue")?
        .commit()?;
    assert_eq!(
        value
            .package()
            .slide_chart_axis_title(0usize, 0usize, Axis::Value)?,
        Some("Gross revenue".to_owned())
    );
    assert_eq!(
        value
            .package()
            .slide_chart_axis_title(0usize, 0usize, Axis::Category)?,
        Some("Month".to_owned())
    );
    Ok(())
}

#[test]
fn public_axis_title_handles_are_semantic_and_redact_native_identity() -> TestResult<()> {
    let package = Package::from_bytes(&synthetic_metadata_package(false)?)?;
    let edit = package.edit_slide_chart_axis_title(
        SlideSelector::name("Charts"),
        ChartSelector::name("Revenue"),
        Axis::Category,
    )?;
    assert_eq!(edit.slide_position(), Position::new(0));
    assert_eq!(edit.chart_position(), Position::new(0));

    // The public transaction handle intentionally exposes only semantic
    // positions. Native chart, stand-in, and axis object identifiers must not
    // become an accidental debugging or logging surface.
    let debug = format!("{edit:?}");
    for identifier in [
        CHARTS[0],
        TITLES[0],
        CHART_NON_STYLES[0],
        CATEGORY_STYLES[0],
        CATEGORY_NON_STYLES[0],
        VALUE_STYLES[0],
        VALUE_NON_STYLES[0],
    ] {
        assert!(
            !debug.contains(&identifier.to_string()),
            "native identifier leaked from semantic edit debug output: {debug}"
        );
    }

    let commit = edit.set("typed selector title")?.commit()?;
    let patch_debug = format!("{:?}", commit.patch());
    for identifier in [
        CHARTS[0],
        TITLES[0],
        CHART_NON_STYLES[0],
        CATEGORY_STYLES[0],
        CATEGORY_NON_STYLES[0],
    ] {
        assert!(
            !patch_debug.contains(&identifier.to_string()),
            "native identifier leaked from semantic patch debug output: {patch_debug}"
        );
    }
    Ok(())
}

#[test]
fn visible_without_text_is_empty_and_hidden_stale_text_is_ignored() -> TestResult<()> {
    let package = Package::from_bytes(&synthetic_package()?)?;
    assert_eq!(
        package.slide_chart_axis_title(0usize, 1usize, Axis::Category)?,
        Some(String::new())
    );
    assert_eq!(
        package.slide_chart_axis_title(0usize, 1usize, Axis::Value)?,
        None
    );
    assert_eq!(
        package.slide_chart_axis_title(0usize, 1usize, Axis::Value)?,
        None
    );
    Ok(())
}

#[test]
fn absent_clear_is_exact_noop_and_set_clear_inverse_reopens_exactly() -> TestResult<()> {
    let source = synthetic_package_without_category_extension()?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.slide_chart_axis_title(0usize, 0usize, Axis::Category)?,
        None
    );
    let absent_clear = package
        .edit_slide_chart_axis_title(0usize, 0usize, Axis::Category)?
        .clear()?
        .commit()?;
    assert!(absent_clear.patch().is_noop());
    assert!(!absent_clear.diagnostics().changed());
    assert_eq!(exact_bytes(absent_clear.package())?, source);

    let created = package
        .edit_slide_chart_axis_title(0usize, 0usize, Axis::Category)?
        .set("Created category")?
        .commit()?;
    let target = exact_bytes(created.package())?;
    let reopened = Package::from_bytes(&target)?;
    assert_eq!(
        reopened.slide_chart_axis_title(0usize, 0usize, Axis::Category)?,
        Some("Created category".to_owned())
    );
    let cleared = reopened
        .edit_slide_chart_axis_title(0usize, 0usize, Axis::Category)?
        .clear()?
        .commit()?;
    assert_eq!(
        cleared
            .package()
            .slide_chart_axis_title(0usize, 0usize, Axis::Category)?,
        None
    );
    let restored = cleared
        .package()
        .apply_slide_chart_axis_title(&cleared.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, target);
    let restored_source = created
        .package()
        .apply_slide_chart_axis_title(&created.patch().inverse())?;
    assert_eq!(exact_bytes(restored_source.package())?, source);
    Ok(())
}

#[test]
fn empty_and_unicode_titles_round_trip_with_explicit_presence() -> TestResult<()> {
    let source = synthetic_package_without_category_extension()?;
    let package = Package::from_bytes(&source)?;

    let empty = package
        .edit_slide_chart_axis_title(
            SlideSelector::position(Position::new(0)),
            ChartSelector::position(Position::new(0)),
            Axis::Category,
        )?
        .set("")?;
    assert_eq!(empty.before(), None);
    assert_eq!(empty.after(), Some(""));
    let empty = empty.commit()?;
    assert_eq!(
        empty.package().slide_chart_axis_title(
            Position::new(0),
            Position::new(0),
            Axis::Category
        )?,
        Some(String::new())
    );

    let empty_bytes = exact_bytes(empty.package())?;
    let reopened = Package::from_bytes(&empty_bytes)?;
    let unicode = reopened
        .edit_slide_chart_axis_title(0usize, 0usize, Axis::Category)?
        .set("売上 📈")?
        .commit()?;
    let _unicode_bytes = exact_bytes(unicode.package())?;
    assert_eq!(
        unicode
            .package()
            .slide_chart_axis_title(0usize, 0usize, Axis::Category)?,
        Some("売上 📈".to_owned())
    );

    let restored_empty = unicode
        .package()
        .apply_slide_chart_axis_title(&unicode.patch().inverse())?;
    assert_eq!(exact_bytes(restored_empty.package())?, empty_bytes);

    let cleared = reopened
        .edit_slide_chart_axis_title(0usize, 0usize, Axis::Category)?
        .clear()?
        .commit()?;
    assert_eq!(
        cleared
            .package()
            .slide_chart_axis_title(0usize, 0usize, Axis::Category)?,
        None
    );
    let restored_source = cleared
        .package()
        .apply_slide_chart_axis_title(&cleared.patch().inverse())?;
    assert_eq!(exact_bytes(restored_source.package())?, empty_bytes);
    Ok(())
}

#[test]
fn secondary_axis_and_unknown_wire_spans_are_untouched() -> TestResult<()> {
    let source = synthetic_metadata_package(false)?;
    let selected_before =
        message_payload(&source, VALUE_NON_STYLES[0], AXIS_NON_STYLE_MESSAGE_TYPE)?;
    let selected_style_before = message_payload(&source, VALUE_STYLES[0], AXIS_STYLE_MESSAGE_TYPE)?;
    let secondary_before = message_payload(
        &source,
        SECONDARY_VALUE_NON_STYLES[0],
        AXIS_NON_STYLE_MESSAGE_TYPE,
    )?;
    let opposite_before =
        message_payload(&source, CATEGORY_NON_STYLES[0], AXIS_NON_STYLE_MESSAGE_TYPE)?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_slide_chart_axis_title(0usize, 0usize, Axis::Value)?
        .set("Primary changed")?
        .commit()?;
    let target = exact_bytes(commit.package())?;
    let selected_after =
        message_payload(&target, VALUE_NON_STYLES[0], AXIS_NON_STYLE_MESSAGE_TYPE)?;
    let selected_style_after = message_payload(&target, VALUE_STYLES[0], AXIS_STYLE_MESSAGE_TYPE)?;
    let secondary_after = message_payload(
        &target,
        SECONDARY_VALUE_NON_STYLES[0],
        AXIS_NON_STYLE_MESSAGE_TYPE,
    )?;
    let opposite_after =
        message_payload(&target, CATEGORY_NON_STYLES[0], AXIS_NON_STYLE_MESSAGE_TYPE)?;
    assert_ne!(selected_before, selected_after);
    assert_eq!(selected_style_before, selected_style_after);
    assert_eq!(secondary_before, secondary_after);
    assert_eq!(opposite_before, opposite_after);
    assert!(
        selected_after
            .windows(b"opaque axis bytes".len())
            .any(|window| window == b"opaque axis bytes")
    );
    assert!(
        selected_style_after
            .windows(b"opaque axis style".len())
            .any(|window| window == b"opaque axis style")
    );
    assert_eq!(
        nested_field_raw(
            &selected_before,
            GENERATED_EXTENSION_FIELD,
            UNKNOWN_GENERATED_FIELD
        )?,
        nested_field_raw(
            &selected_after,
            GENERATED_EXTENSION_FIELD,
            UNKNOWN_GENERATED_FIELD
        )?,
    );

    let category_before =
        message_payload(&source, CATEGORY_NON_STYLES[0], AXIS_NON_STYLE_MESSAGE_TYPE)?;
    let cleared = package
        .edit_slide_chart_axis_title(0usize, 0usize, Axis::Category)?
        .clear()?
        .commit()?;
    let category_after = message_payload(
        &exact_bytes(cleared.package())?,
        CATEGORY_NON_STYLES[0],
        AXIS_NON_STYLE_MESSAGE_TYPE,
    )?;
    for field in [
        UNKNOWN_OUTER_FIELD,
        SUPPORTS_PRIMARY_FEATURE_FIELD,
        SUPPORTS_SECONDARY_FEATURE_FIELD,
    ] {
        assert_eq!(
            raw_fields(&category_before, field)?,
            raw_fields(&category_after, field)?,
            "unknown or capability field changed while clearing the category title: {field}"
        );
    }
    assert_eq!(
        nested_field_raw(
            &category_before,
            GENERATED_EXTENSION_FIELD,
            UNKNOWN_GENERATED_FIELD,
        )?,
        nested_field_raw(
            &category_after,
            GENERATED_EXTENSION_FIELD,
            UNKNOWN_GENERATED_FIELD,
        )?,
    );
    assert_eq!(
        metadata_stream(&source)?,
        metadata_stream(&exact_bytes(cleared.package())?)?
    );
    Ok(())
}

#[test]
fn stale_and_foreign_patches_conflict_without_source_changes() -> TestResult<()> {
    let source = synthetic_metadata_package(false)?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_slide_chart_axis_title(0usize, 0usize, Axis::Category)?
        .set("new category")?
        .commit()?;
    let applied = package.apply_slide_chart_axis_title(commit.patch())?;
    assert_eq!(
        exact_bytes(applied.package())?,
        exact_bytes(commit.package())?
    );
    assert!(matches!(
        applied
            .package()
            .apply_slide_chart_axis_title(commit.patch()),
        Err(ChartAxisTitleError::PatchConflict)
    ));
    assert_eq!(exact_bytes(&package)?, source);

    let stale_bytes = with_axis_payload(
        &source,
        CATEGORY_NON_STYLES[0],
        axis_payload(
            AxisState {
                visible: Some(true),
                title: Some(b"stale source"),
            },
            Axis::Category,
            true,
        )?,
    )?;
    let stale = Package::from_bytes(&stale_bytes)?;
    assert!(matches!(
        stale.apply_slide_chart_axis_title(commit.patch()),
        Err(ChartAxisTitleError::PatchConflict)
    ));
    assert_eq!(exact_bytes(&stale)?, stale_bytes);

    let foreign_source = with_axis_payload(
        &source,
        CATEGORY_NON_STYLES[0],
        axis_payload(
            AxisState {
                visible: Some(true),
                title: Some(b"foreign source"),
            },
            Axis::Category,
            true,
        )?,
    )?;
    let foreign = Package::from_bytes(&foreign_source)?;
    assert!(matches!(
        foreign.apply_slide_chart_axis_title(commit.patch()),
        Err(ChartAxisTitleError::PatchConflict)
    ));
    assert_eq!(exact_bytes(&foreign)?, foreign_source);
    Ok(())
}

#[test]
fn selectors_report_missing_and_ambiguous_chart_names() -> TestResult<()> {
    let source = synthetic_metadata_package(false)?;
    let package = Package::from_bytes(&source)?;
    assert!(matches!(
        package.slide_chart_axis_title(0usize, ChartSelector::name("missing"), Axis::Category),
        Err(ChartAxisTitleError::ChartNameNotFound)
    ));
    assert!(matches!(
        package.slide_chart_axis_title(SlideSelector::name("missing"), 0usize, Axis::Category),
        Err(ChartAxisTitleError::SlideNameNotFound)
    ));
    assert!(matches!(
        package.slide_chart_axis_title(Position::new(9), 0usize, Axis::Category),
        Err(ChartAxisTitleError::SlidePositionNotFound { .. })
    ));
    assert!(matches!(
        package.slide_chart_axis_title(SlideSelector::name(""), 0usize, Axis::Category),
        Err(ChartAxisTitleError::EmptySlideName)
    ));
    assert!(matches!(
        package.slide_chart_axis_title(0usize, ChartSelector::name(""), Axis::Category),
        Err(ChartAxisTitleError::EmptyChartName)
    ));
    assert!(matches!(
        package.edit_slide_chart_axis_title(0usize, 4usize, Axis::Category),
        Err(ChartAxisTitleError::ChartPositionNotFound { position })
            if position == Position::new(4)
    ));

    let duplicate_chart_title = with_document_message_payload(
        &source,
        CHART_NON_STYLES[1],
        CHART_NON_STYLE_MESSAGE_TYPE,
        chart_non_style_payload("Revenue")?,
    )?;
    let duplicate = Package::from_bytes(&duplicate_chart_title)?;
    assert!(matches!(
        duplicate.slide_chart_axis_title(0usize, ChartSelector::name("Revenue"), Axis::Value),
        Err(ChartAxisTitleError::AmbiguousSelector)
    ));
    Ok(())
}

#[test]
fn locked_or_unsupported_chart_graphs_are_rejected_for_mutation() -> TestResult<()> {
    let source = synthetic_metadata_package(false)?;
    let chart = message_payload(&source, CHARTS[0], CHART_MESSAGE_TYPE)?;
    let drawable = WireView::parse(&chart)?
        .fields()
        .find(|field| field.number() == 1)
        .ok_or_else(|| io::Error::other("missing chart drawable"))?;
    let mut drawable_value = tsd::DrawableArchive::decode(drawable.payload())?;
    drawable_value.locked = Some(true);
    let drawable_replacement = length_delimited_field_with_key_width(
        1,
        &drawable_value.encode_to_vec(),
        drawable.key().len(),
    );
    let locked_chart = replace_first_field(&chart, 1, &drawable_replacement)?;
    let locked =
        with_document_message_payload(&source, CHARTS[0], CHART_MESSAGE_TYPE, locked_chart)?;
    assert_axis_rejected_atomically(&locked, Axis::Category)?;

    let missing_axis = with_chart_axis_variant(&source, ChartAxisVariant::MissingCategory)?;
    assert_axis_rejected_atomically(&missing_axis, Axis::Category)?;
    Ok(())
}

#[test]
fn malformed_axis_extensions_fail_closed_and_preserve_source_bytes() -> TestResult<()> {
    let source = synthetic_metadata_package(false)?;
    for variant in [
        AxisWireVariant::DuplicateOuter,
        AxisWireVariant::WrongOuterWire,
        AxisWireVariant::NonCanonicalOuter,
        AxisWireVariant::DuplicateTitle,
        AxisWireVariant::InvalidUtf8,
    ] {
        let hostile = with_axis_extension_variant(&source, variant)?;
        let package = Package::from_bytes(&hostile)?;
        assert!(matches!(
            package.slide_chart_axis_title(0usize, 0usize, Axis::Category),
            Err(ChartAxisTitleError::InvalidSource)
        ));
        assert_eq!(exact_bytes(&package)?, hostile);
        assert_axis_rejected_atomically(&hostile, Axis::Category)?;
    }
    Ok(())
}

#[test]
fn missing_duplicate_and_role_aliased_axis_refs_fail_closed() -> TestResult<()> {
    let source = synthetic_metadata_package(false)?;
    for variant in [
        ChartAxisVariant::MissingCategory,
        ChartAxisVariant::MissingValue,
        ChartAxisVariant::DuplicateCategory,
        ChartAxisVariant::AliasedRoles,
    ] {
        let hostile = with_chart_axis_variant(&source, variant)?;
        let axis = if matches!(variant, ChartAxisVariant::MissingValue) {
            Axis::Value
        } else {
            Axis::Category
        };
        let package = Package::from_bytes(&hostile)?;
        assert!(matches!(
            package.slide_chart_axis_title(0usize, 0usize, axis),
            Err(ChartAxisTitleError::InvalidSource)
        ));
        assert_eq!(exact_bytes(&package)?, hostile);
        assert_axis_rejected_atomically(&hostile, axis)?;
    }
    assert_axis_rejected_atomically(&with_axis_role_alias(&source)?, Axis::Category)?;
    Ok(())
}

#[test]
fn foreign_object_and_colliding_data_inbound_edges_are_rejected() -> TestResult<()> {
    let source = synthetic_metadata_package(true)?;
    for mode in [ForeignInbound::Object, ForeignInbound::Field] {
        let hostile = with_foreign_inbound(&source, CATEGORY_NON_STYLES[0], mode)?;
        assert_axis_rejected_atomically(&hostile, Axis::Category)?;
    }
    for mode in [ForeignInbound::Data, ForeignInbound::FieldData] {
        let hostile = with_foreign_inbound(&source, CATEGORY_NON_STYLES[0], mode)?;
        assert_axis_rejected_atomically(&hostile, Axis::Category)?;
    }
    Ok(())
}

#[test]
fn unrelated_data_edges_and_stylesheet_registration_remain_editable() -> TestResult<()> {
    let source = synthetic_metadata_package(true)?;
    for mode in [ForeignInbound::Data, ForeignInbound::FieldData] {
        let unrelated = with_foreign_inbound(&source, 2_002, mode)?;
        let commit = Package::from_bytes(&unrelated)?
            .edit_slide_chart_axis_title(0usize, 0usize, Axis::Category)?
            .set("Registered axis")?
            .commit()?;
        assert_eq!(
            commit
                .package()
                .slide_chart_axis_title(0usize, 0usize, Axis::Category)?,
            Some("Registered axis".to_owned())
        );
    }

    let registered =
        with_stylesheet_registration(&synthetic_metadata_package(false)?, VALUE_NON_STYLES[0])?;
    let commit = Package::from_bytes(&registered)?
        .edit_slide_chart_axis_title(0usize, 0usize, Axis::Value)?
        .set("Native registry")?
        .commit()?;
    assert_eq!(
        commit
            .package()
            .slide_chart_axis_title(0usize, 0usize, Axis::Value)?,
        Some("Native registry".to_owned())
    );
    Ok(())
}

#[test]
fn native_stylesheet_owned_chart_graph_remains_selector_first_and_reversible() -> TestResult<()> {
    let source = with_native_stylesheet_layout(&synthetic_metadata_package(false)?)?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.slide_chart_title(0usize, 0usize)?,
        Some("Revenue".to_owned())
    );
    assert_eq!(
        package.slide_chart_axis_title(0usize, 0usize, Axis::Value)?,
        Some("Revenue".to_owned())
    );

    let commit = package
        .edit_slide_chart_axis_title(0usize, ChartSelector::name("Revenue"), Axis::Value)?
        .set("Native registry")?
        .commit()?;
    assert_eq!(
        commit
            .package()
            .slide_chart_axis_title(0usize, 0usize, Axis::Value)?,
        Some("Native registry".to_owned())
    );
    let inverse = commit
        .package()
        .apply_slide_chart_axis_title(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(inverse.package())?, source);

    let title_commit = package
        .edit_slide_chart_title(0usize, 0usize)?
        .set("Native chart title")?
        .commit()?;
    assert_eq!(
        title_commit.package().slide_chart_title(0usize, 0usize)?,
        Some("Native chart title".to_owned())
    );
    Ok(())
}

#[test]
fn metadata_identity_and_authority_namespaces_are_rejected() -> TestResult<()> {
    let source = synthetic_metadata_package(false)?;
    let variants = [
        with_metadata(&source, |metadata| {
            let document = metadata
                .components
                .iter_mut()
                .find(|component| component.identifier == DOCUMENT_COMPONENT)
                .expect("document component");
            document
                .object_uuid_map_entries
                .retain(|entry| entry.identifier != CATEGORY_NON_STYLES[0]);
        })?,
        with_metadata(&source, |metadata| {
            let entry = metadata
                .components
                .iter()
                .find(|component| component.identifier == DOCUMENT_COMPONENT)
                .and_then(|component| {
                    component
                        .object_uuid_map_entries
                        .iter()
                        .find(|entry| entry.identifier == CATEGORY_NON_STYLES[0])
                })
                .cloned()
                .expect("category axis uuid");
            metadata
                .components
                .iter_mut()
                .find(|component| component.identifier == DOCUMENT_COMPONENT)
                .expect("document component")
                .object_uuid_map_entries
                .retain(|candidate| candidate.identifier != CATEGORY_NON_STYLES[0]);
            metadata
                .components
                .iter_mut()
                .find(|component| component.identifier == UNRELATED_COMPONENT)
                .expect("unrelated component")
                .object_uuid_map_entries
                .push(entry);
        })?,
        with_metadata(&source, |metadata| {
            metadata
                .components
                .iter_mut()
                .find(|component| component.identifier == DOCUMENT_COMPONENT)
                .expect("document component")
                .object_uuid_map_entries
                .push(metadata_uuid_entry(CATEGORY_NON_STYLES[0]));
        })?,
        with_metadata(&source, |metadata| {
            metadata.versioned_components.push(tsp::ComponentInfo {
                identifier: DOCUMENT_COMPONENT,
                preferred_locator: "Document".to_owned(),
                locator: Some("Document".to_owned()),
                object_uuid_map_entries: vec![metadata_uuid_entry(CATEGORY_NON_STYLES[0])],
                ..tsp::ComponentInfo::default()
            });
        })?,
        with_metadata(&source, |metadata| {
            metadata
                .components
                .iter_mut()
                .find(|component| component.identifier == DOCUMENT_COMPONENT)
                .expect("document component")
                .ambiguous_object_identifiers
                .push(CATEGORY_NON_STYLES[0]);
        })?,
        with_metadata(&source, |metadata| {
            metadata.data_metadata_map = Some(reference(CATEGORY_NON_STYLES[0]));
        })?,
        with_metadata(&source, |metadata| {
            metadata
                .components
                .iter_mut()
                .find(|component| component.identifier == DOCUMENT_COMPONENT)
                .expect("document component")
                .data_references
                .push(tsp::ComponentDataReference {
                    data_identifier: 2_002,
                    object_reference_list: vec![tsp::component_data_reference::ObjectReference {
                        object_identifier: CATEGORY_NON_STYLES[0],
                        count: 1,
                    }],
                });
        })?,
    ];
    for hostile in variants {
        assert_axis_rejected_atomically(&hostile, Axis::Category)?;
    }
    Ok(())
}

#[test]
fn output_limit_at_max_minus_one_rejects_atomically() -> TestResult<()> {
    let source = synthetic_metadata_package(false)?;
    let long_title = "a".repeat(16 * 1024);
    let baseline = Package::from_bytes(&source)?
        .edit_slide_chart_axis_title(0usize, 0usize, Axis::Category)?
        .set(long_title.clone())?
        .commit()?;
    let target = exact_bytes(baseline.package())?;
    assert!(target.len() > source.len());
    let defaults = Limits::default();
    let limits = Limits::new(
        u64::try_from(target.len() - 1)?,
        defaults.max_entries(),
        defaults.max_entry_bytes(),
        defaults.max_total_bytes(),
        defaults.max_iwa_stream_bytes(),
    )?;
    let package = Package::from_bytes_with_options(
        &source,
        ReadOptions::new(limits, SemanticLimits::default()),
    )?;
    let before = exact_bytes(&package)?;
    let result = package
        .edit_slide_chart_axis_title(0usize, 0usize, Axis::Category)
        .and_then(|edit| edit.set(long_title))
        .and_then(|edit| edit.commit());
    assert!(matches!(
        result,
        Err(ChartAxisTitleError::LimitExceeded {
            kind: ChartAxisTitleLimitKind::OutputBytes,
            ..
        })
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}
