//! Strict, source-preserving Keynote chart-axis-title integration coverage.
//!
//! The fixture below follows the producer graph used by source-built iWork
//! charts: a chart drawable owns a title stand-in, chart non-style, primary
//! and secondary axis style/non-style objects, and records the same private
//! graph in its IWA message metadata.  The tests deliberately mutate only
//! this test-local fixture; no production source is involved.

#![allow(
    dead_code,
    unused_imports,
    reason = "each integration-test crate consumes a different strict subset of this shared fixture"
)]

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

pub(crate) const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
pub(crate) const METADATA_MEMBER: &str = "Index/Metadata.iwa";
pub(crate) const FOREIGN_MEMBER: &str = "Index/Foreign.iwa";
pub(crate) const STYLESHEET_MEMBER: &str = "Index/DocumentStylesheet.iwa";
pub(crate) const METADATA_OBJECT: u64 = 300;
pub(crate) const DOCUMENT_COMPONENT: u64 = 1;
pub(crate) const UNRELATED_COMPONENT: u64 = 2;
pub(crate) const FOREIGN_COMPONENT: u64 = 3;
pub(crate) const STYLESHEET_COMPONENT: u64 = 4;
pub(crate) const SLIDE_NODE: u64 = 3;
pub(crate) const SLIDE: u64 = 4;
pub(crate) const CHARTS: [u64; 2] = [100, 101];
pub(crate) const TITLES: [u64; 2] = [110, 111];
pub(crate) const CHART_NON_STYLES: [u64; 2] = [120, 121];
pub(crate) const CATEGORY_STYLES: [u64; 2] = [130, 131];
pub(crate) const CATEGORY_NON_STYLES: [u64; 2] = [140, 141];
pub(crate) const VALUE_STYLES: [u64; 2] = [150, 151];
pub(crate) const VALUE_NON_STYLES: [u64; 2] = [160, 161];
pub(crate) const SECONDARY_VALUE_STYLES: [u64; 2] = [170, 171];
pub(crate) const SECONDARY_VALUE_NON_STYLES: [u64; 2] = [180, 181];
pub(crate) const FOREIGN_OBJECT: u64 = 900;
pub(crate) const STYLESHEET_OBJECT: u64 = 910;
pub(crate) const CHART_MESSAGE_TYPE: u32 = 5_021;
pub(crate) const STANDIN_MESSAGE_TYPE: u32 = 3_097;
pub(crate) const CHART_NON_STYLE_MESSAGE_TYPE: u32 = 5_023;
pub(crate) const AXIS_STYLE_MESSAGE_TYPE: u32 = 5_026;
pub(crate) const AXIS_NON_STYLE_MESSAGE_TYPE: u32 = 5_027;
pub(crate) const METADATA_MESSAGE_TYPE: u32 = 11_006;
pub(crate) const STYLESHEET_MESSAGE_TYPE: u32 = 401;
pub(crate) const GENERATED_EXTENSION_FIELD: u32 = 10_000;
pub(crate) const UNKNOWN_OUTER_FIELD: u32 = 4_000;
pub(crate) const UNKNOWN_GENERATED_FIELD: u32 = 4_001;
pub(crate) const METADATA_UNKNOWN_FIELD: u32 = 4_002;
pub(crate) const SUPPORTS_PRIMARY_FEATURE_FIELD: u32 = 10_001;
pub(crate) const SUPPORTS_SECONDARY_FEATURE_FIELD: u32 = 10_002;
pub(crate) const METADATA_UNKNOWN_MARKER: &[u8] = b"chart-axis-title metadata extension";
pub(crate) const PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

pub(crate) type TestResult<T> = Result<T, Box<dyn std::error::Error>>;

#[derive(Clone, Copy)]
pub(crate) struct AxisState<'a> {
    pub(crate) visible: Option<bool>,
    pub(crate) title: Option<&'a [u8]>,
}

pub(crate) fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

pub(crate) fn object(identifier: u64, type_: u32, data: Vec<u8>) -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage { type_, data }],
    )?)
}

pub(crate) fn object_with_references(
    identifier: u64,
    type_: u32,
    data: Vec<u8>,
    references: Vec<u64>,
) -> TestResult<ArchiveObject> {
    let mut value = object(identifier, type_, data)?;
    value.archive_info.message_infos[0].object_references = references;
    Ok(value)
}

pub(crate) fn chart_payload(
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

pub(crate) fn axis_payload(state: AxisState<'_>, axis: Axis, unknown: bool) -> TestResult<Vec<u8>> {
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

pub(crate) fn axis_style_payload(unknown: bool) -> TestResult<Vec<u8>> {
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

pub(crate) fn chart_non_style_payload(title: &str) -> TestResult<Vec<u8>> {
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

pub(crate) fn component(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

pub(crate) fn chart_object_references(chart: usize) -> Vec<u64> {
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

pub(crate) fn all_document_object_ids() -> Vec<u64> {
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

pub(crate) fn synthetic_package_with_states(
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

pub(crate) fn synthetic_package() -> TestResult<Vec<u8>> {
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

pub(crate) fn synthetic_package_without_category_extension() -> TestResult<Vec<u8>> {
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

pub(crate) fn axis_payload_without_extension(unknown: bool) -> TestResult<Vec<u8>> {
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

pub(crate) fn metadata_uuid_entry(identifier: u64) -> tsp::ObjectUuidMapEntry {
    tsp::ObjectUuidMapEntry {
        identifier,
        uuid: tsp::Uuid {
            lower: identifier.saturating_add(10_000),
            upper: identifier.saturating_add(20_000),
        },
    }
}

pub(crate) fn metadata_component_payload(
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

pub(crate) fn metadata_payload(last_identifier: u64, include_foreign: bool) -> TestResult<Vec<u8>> {
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

pub(crate) fn synthetic_metadata_package(include_foreign: bool) -> TestResult<Vec<u8>> {
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

pub(crate) fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

pub(crate) fn document_stream(package: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("missing synthetic document component"))?;
    Ok(SnappyStream::decompress(entry.data())?.into_bytes())
}

pub(crate) fn metadata_stream(package: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == METADATA_MEMBER)
        .ok_or_else(|| io::Error::other("missing synthetic metadata component"))?;
    Ok(SnappyStream::decompress(entry.data())?.into_bytes())
}

pub(crate) fn component_stream(package: &[u8], member: &str) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or_else(|| io::Error::other("missing synthetic component"))?;
    Ok(SnappyStream::decompress(entry.data())?.into_bytes())
}

pub(crate) fn replace_document_stream(source: &[u8], archive: Archive) -> TestResult<Vec<u8>> {
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(Catalog::from_bytes(source)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            DOCUMENT_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?)
}

pub(crate) fn replace_metadata_stream(source: &[u8], archive: Archive) -> TestResult<Vec<u8>> {
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(Catalog::from_bytes(source)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            METADATA_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?)
}

pub(crate) fn replace_component_stream(
    source: &[u8],
    member: &str,
    archive: Archive,
) -> TestResult<Vec<u8>> {
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(Catalog::from_bytes(source)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            member,
            &compressed,
        )],
        Limits::default(),
    )?)
}

pub(crate) fn message_payload(package: &[u8], identifier: u64, type_: u32) -> TestResult<Vec<u8>> {
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

pub(crate) fn nested_field_raw(payload: &[u8], outer: u32, nested: u32) -> TestResult<Vec<u8>> {
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

pub(crate) fn raw_fields(payload: &[u8], number: u32) -> TestResult<Vec<Vec<u8>>> {
    Ok(WireView::parse(payload)?
        .fields()
        .filter(|field| field.number() == number)
        .map(|field| field.raw().to_vec())
        .collect())
}

pub(crate) fn with_document_message_payload(
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

pub(crate) fn with_axis_payload(
    source: &[u8],
    identifier: u64,
    data: Vec<u8>,
) -> TestResult<Vec<u8>> {
    with_document_message_payload(source, identifier, AXIS_NON_STYLE_MESSAGE_TYPE, data)
}

pub(crate) fn with_chart_payload(
    source: &[u8],
    chart: usize,
    data: Vec<u8>,
) -> TestResult<Vec<u8>> {
    with_document_message_payload(source, CHARTS[chart], CHART_MESSAGE_TYPE, data)
}

pub(crate) fn with_axis_role_alias(source: &[u8]) -> TestResult<Vec<u8>> {
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

pub(crate) fn with_foreign_inbound(
    source: &[u8],
    target: u64,
    mode: ForeignInbound,
) -> TestResult<Vec<u8>> {
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

pub(crate) fn with_stylesheet_registration(source: &[u8], target: u64) -> TestResult<Vec<u8>> {
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

pub(crate) fn with_native_stylesheet_layout(source: &[u8]) -> TestResult<Vec<u8>> {
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
pub(crate) enum ForeignInbound {
    Object,
    Data,
    Field,
    FieldData,
}

pub(crate) fn with_metadata(
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

pub(crate) fn append_duplicate_field(payload: &[u8], number: u32) -> TestResult<Vec<u8>> {
    let duplicate = WireView::parse(payload)?
        .fields()
        .find(|field| field.number() == number)
        .map(|field| field.raw().to_vec())
        .ok_or_else(|| io::Error::other("missing field to duplicate"))?;
    let mut output = payload.to_vec();
    output.extend_from_slice(&duplicate);
    Ok(output)
}

pub(crate) fn replace_first_field(
    payload: &[u8],
    number: u32,
    replacement: &[u8],
) -> TestResult<Vec<u8>> {
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

pub(crate) fn length_delimited_field_with_key_width(
    number: u32,
    payload: &[u8],
    key_width: usize,
) -> Vec<u8> {
    let mut output = Vec::new();
    push_varint_width((u64::from(number) << 3) | 2, key_width, &mut output);
    push_varint_width(payload.len() as u64, 1, &mut output);
    output.extend_from_slice(payload);
    output
}

pub(crate) fn varint_field(number: u32, value: u64) -> Vec<u8> {
    let mut output = Vec::new();
    append_varint_field(&mut output, number, value).expect("small hostile varint fits");
    output
}

pub(crate) fn push_varint_width(mut value: u64, width: usize, output: &mut Vec<u8>) {
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

pub(crate) fn with_axis_extension_variant(
    source: &[u8],
    variant: AxisWireVariant,
) -> TestResult<Vec<u8>> {
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
pub(crate) enum AxisWireVariant {
    DuplicateOuter,
    WrongOuterWire,
    NonCanonicalOuter,
    DuplicateTitle,
    InvalidUtf8,
}

pub(crate) fn with_chart_axis_variant(
    source: &[u8],
    variant: ChartAxisVariant,
) -> TestResult<Vec<u8>> {
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
pub(crate) enum ChartAxisVariant {
    MissingCategory,
    MissingValue,
    DuplicateCategory,
    AliasedRoles,
}

pub(crate) fn assert_axis_rejected_atomically(source: &[u8], axis: Axis) -> TestResult<()> {
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

pub(crate) fn assert_locality(before: &[u8], after: &[u8]) -> TestResult<()> {
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
