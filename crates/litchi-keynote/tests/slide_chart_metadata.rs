//! Focused Keynote chart metadata readback.

use std::{io, path::PathBuf};

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::wire::append_length_delimited_field;
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsch, tsd, tsk, tsp};
use litchi_keynote::{ChartSelector, Package, SlideChartMetadataError, SlideSelector};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/keynote/chart-titles-native.key")
}

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const SLIDE_NODE: u64 = 3;
const SLIDE: u64 = 4;
const CHART: u64 = 100;
const TITLE: u64 = 110;
const NON_STYLE: u64 = 120;
const CHART_MESSAGE_TYPE: u32 = 5_021;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
const CHART_NON_STYLE_MESSAGE_TYPE: u32 = 5_023;

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

fn chart_payload(non_style: u64) -> TestResult<Vec<u8>> {
    let drawable = tsd::DrawableArchive {
        parent: Some(reference(SLIDE)),
        title: Some(reference(TITLE)),
        ..tsd::DrawableArchive::default()
    };
    let chart = tsch::ChartArchive {
        chart_type: Some(tsch::ChartType::ColumnChartType2D as i32),
        contains_default_data: Some(false),
        grid: Some(tsch::ChartGridArchive {
            row_name: vec!["Region 1".to_owned(), "Region 2".to_owned()],
            column_name: vec![
                "April".to_owned(),
                "May".to_owned(),
                "June".to_owned(),
                "July".to_owned(),
            ],
            grid_row: vec![
                tsch::GridRow {
                    value: vec![
                        tsch::GridValue {
                            numeric_value: Some(1.0),
                            ..tsch::GridValue::default()
                        };
                        4
                    ],
                },
                tsch::GridRow {
                    value: vec![
                        tsch::GridValue {
                            numeric_value: Some(2.0),
                            ..tsch::GridValue::default()
                        };
                        4
                    ],
                },
            ],
            ..tsch::ChartGridArchive::default()
        }),
        chart_non_style: Some(reference(non_style)),
        ..tsch::ChartArchive::default()
    };
    let mut payload = Vec::new();
    append_length_delimited_field(&mut payload, 1, &drawable.encode_to_vec())?;
    append_length_delimited_field(&mut payload, 10_000, &chart.encode_to_vec())?;
    Ok(payload)
}

fn metadata_package() -> TestResult<Vec<u8>> {
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
        owned_drawables: vec![reference(CHART)],
        drawables_z_order: vec![reference(CHART)],
        name: Some("Charts".to_owned()),
        in_document: true,
        ..kn::SlideArchive::default()
    };
    let objects = vec![
        object(1, 1, document.encode_to_vec())?,
        object(2, 2, show.encode_to_vec())?,
        object(SLIDE_NODE, 4, node.encode_to_vec())?,
        object(SLIDE, 5, slide.encode_to_vec())?,
        object(CHART, CHART_MESSAGE_TYPE, chart_payload(NON_STYLE)?)?,
        object(TITLE, STANDIN_MESSAGE_TYPE, Vec::new())?,
        object(
            NON_STYLE,
            CHART_NON_STYLE_MESSAGE_TYPE,
            non_style_payload(Some(true), Some("Revenue"))?,
        )?,
    ];
    let component = SnappyStream::compress(&Archive { objects }.to_bytes()?)?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [(DOCUMENT_MEMBER, component.as_slice())],
        Limits::default(),
    )?)
}

fn non_style_payload(visible: Option<bool>, title: Option<&str>) -> TestResult<Vec<u8>> {
    let mut extension = Vec::new();
    if let Some(visible) = visible {
        litchi_iwa_common::wire::append_varint_field(&mut extension, 21, u64::from(visible))?;
    }
    if let Some(title) = title {
        append_length_delimited_field(&mut extension, 23, title.as_bytes())?;
    }
    let mut payload = Vec::new();
    append_length_delimited_field(&mut payload, 10_000, &extension)?;
    Ok(payload)
}

fn replace_document(source: &[u8], mutate: impl FnOnce(&mut Archive)) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("missing synthetic document component"))?;
    let stream = SnappyStream::decompress(entry.data())?;
    let mut archive = Archive::parse(stream.as_bytes())?;
    mutate(&mut archive);
    let component = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            DOCUMENT_MEMBER,
            &component,
        )],
        Limits::default(),
    )?)
}

#[test]
fn native_chart_metadata_is_selector_first_and_archive_free() -> TestResult {
    let source = std::fs::read(fixture_path())?;
    let package = Package::from_bytes(&source)?;

    let by_position =
        package.slide_chart_metadata(SlideSelector::index(0), ChartSelector::index(0))?;
    let by_title = package.slide_chart_metadata(
        SlideSelector::index(0),
        ChartSelector::name("Native chart title — 北区"),
    )?;

    assert_eq!(by_position, by_title);
    assert_eq!(by_position.kind().native_value(), 1);
    assert_eq!(by_position.title(), Some("Native chart title — 北区"));
    assert_eq!(
        by_position.row_names(),
        &[String::from("Region 1"), String::from("Region 2")]
    );
    assert_eq!(
        by_position.column_names(),
        &[
            String::from("April"),
            String::from("May"),
            String::from("June"),
            String::from("July"),
        ]
    );
    assert_eq!(by_position.series_count(), 2);
    let mut roundtrip = Vec::new();
    package.write_to(&mut roundtrip)?;
    assert_eq!(roundtrip, source);
    Ok(())
}

#[test]
fn chart_metadata_rejects_missing_semantic_chart_without_exposing_native_ids() {
    let source = std::fs::read(fixture_path()).expect("native fixture");
    let package = Package::from_bytes(&source).expect("native package");
    let error = package
        .slide_chart_metadata(SlideSelector::index(0), ChartSelector::index(1))
        .expect_err("the fixture contains one chart");
    assert!(matches!(
        error,
        SlideChartMetadataError::ChartPositionNotFound { .. }
    ));
}

#[test]
fn synthetic_chart_metadata_reads_embedded_grid_labels() -> TestResult {
    let package = Package::from_bytes(&metadata_package()?)?;
    let metadata = package.slide_chart_metadata("Charts", "Revenue")?;
    assert_eq!(metadata.kind().native_value(), 1);
    assert_eq!(metadata.title(), Some("Revenue"));
    assert_eq!(
        metadata.row_names(),
        &[String::from("Region 1"), String::from("Region 2")]
    );
    assert_eq!(metadata.column_names().len(), 4);
    assert_eq!(metadata.series_count(), 2);
    assert!(!metadata.contains_default_data());
    Ok(())
}

#[test]
fn synthetic_chart_metadata_rejects_duplicate_chart_messages() -> TestResult {
    let source = metadata_package()?;
    let malformed = replace_document(&source, |archive| {
        let chart = archive.object_mut(CHART).expect("synthetic chart object");
        let duplicate = chart
            .messages
            .iter()
            .find(|message| message.type_ == CHART_MESSAGE_TYPE)
            .expect("synthetic chart message")
            .clone();
        chart
            .push_message(duplicate)
            .expect("append duplicate chart message");
    })?;
    let package = Package::from_bytes(&malformed)?;
    assert!(matches!(
        package.slide_chart_metadata(0usize, 0usize),
        Err(SlideChartMetadataError::InvalidSource)
    ));
    Ok(())
}

#[test]
fn synthetic_chart_metadata_rejects_foreign_non_style_reference() -> TestResult {
    let source = metadata_package()?;
    let malformed = replace_document(&source, |archive| {
        let chart = archive.object_mut(CHART).expect("synthetic chart object");
        let message = chart
            .messages
            .iter_mut()
            .find(|message| message.type_ == CHART_MESSAGE_TYPE)
            .expect("synthetic chart message");
        message.data = chart_payload(999_999).expect("foreign chart payload");
    })?;
    let package = Package::from_bytes(&malformed)?;
    assert!(matches!(
        package.slide_chart_metadata(0usize, 0usize),
        Err(SlideChartMetadataError::InvalidSource)
    ));
    Ok(())
}
