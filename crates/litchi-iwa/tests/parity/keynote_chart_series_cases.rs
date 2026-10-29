//! Keynote slide-chart data-symbol facade parity (W2-C, record 0401).
//!
//! The shared series graph is covered by record 0375.  These cases exercise
//! the selector-first Keynote package boundary on the same host-authored
//! chart, including direct host-vs-leaf reads, all-series and one-series host
//! writes, exact inverse/no-op behavior, and terminal selector refusals.

use litchi_iwa::charts::{ChartSeriesSymbol, ChartSeriesSymbolShape, ChartSeriesSymbolSize};
use litchi_iwa::keynote;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::chart::data::ChartData;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_common::chart::series::{Symbol, SymbolShape, SymbolSize};
use litchi_iwa_common::varint::encode_varint_into;
use litchi_iwa_protos::chart_style_codec::{self as chart_style_codec, Limits as ChartCodecLimits};
use litchi_keynote::chart::series::{Error, Property, Refusal};
use litchi_keynote::{ChartSelector, OpenOptions, Package, SemanticLimits, SlideSelector};

use crate::chart_style_cases::{Components, Input, input, publish};
use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier};
use crate::leaf_bytes;

const SLIDE: usize = 0;
const LEAF_SLIDE: SlideSelector<'static> = SlideSelector::index(SLIDE);
const LEAF_CHART: ChartSelector<'static> = ChartSelector::index(1);
const CHART_MESSAGE_TYPE: u32 = 5_021;

#[derive(Clone, Debug)]
struct Values(Vec<Option<(SymbolShape, Option<f32>)>>);

impl Values {
    fn leaf(&self) -> TestResult<Vec<Option<Symbol>>> {
        self.0
            .iter()
            .map(|value| {
                value
                    .map(|(shape, points)| {
                        Ok(Symbol {
                            shape,
                            size: points.map(SymbolSize::new).transpose()?,
                        })
                    })
                    .transpose()
            })
            .collect()
    }

    fn host(&self) -> TestResult<Vec<Option<ChartSeriesSymbol>>> {
        self.0
            .iter()
            .map(|value| {
                Ok(match value {
                    None => None,
                    Some((shape, None)) => Some(ChartSeriesSymbol::automatic(host_shape(*shape))),
                    Some((shape, Some(points))) => Some(ChartSeriesSymbol::sized(
                        host_shape(*shape),
                        ChartSeriesSymbolSize::new(*points)?,
                    )),
                })
            })
            .collect()
    }
}

const fn host_shape(shape: SymbolShape) -> ChartSeriesSymbolShape {
    match shape {
        SymbolShape::Circle => ChartSeriesSymbolShape::Circle,
        SymbolShape::Triangle => ChartSeriesSymbolShape::Triangle,
        SymbolShape::InvertedTriangle => ChartSeriesSymbolShape::InvertedTriangle,
        SymbolShape::Diamond => ChartSeriesSymbolShape::Diamond,
        SymbolShape::Square => ChartSeriesSymbolShape::Square,
    }
}

fn host_read(source: &[u8], chart: u64) -> TestResult<Vec<Option<Symbol>>> {
    let editor = keynote::Draft::from_bytes(source)?;
    editor
        .slide_chart_series_symbols(SLIDE, chart)?
        .into_iter()
        .map(|value| match value {
            None => Ok(None),
            Some(value) => Ok(Some(Symbol {
                shape: match value.shape {
                    ChartSeriesSymbolShape::Circle => SymbolShape::Circle,
                    ChartSeriesSymbolShape::Triangle => SymbolShape::Triangle,
                    ChartSeriesSymbolShape::InvertedTriangle => SymbolShape::InvertedTriangle,
                    ChartSeriesSymbolShape::Diamond => SymbolShape::Diamond,
                    ChartSeriesSymbolShape::Square => SymbolShape::Square,
                },
                size: value
                    .size
                    .map(|size| SymbolSize::new(size.points()))
                    .transpose()?,
            })),
        })
        .collect()
}

fn host_symbol_read(source: &[u8], chart: u64, index: usize) -> TestResult<Option<Symbol>> {
    let editor = keynote::Draft::from_bytes(source)?;
    let value = editor.slide_chart_series_symbol(
        SLIDE,
        chart,
        litchi_iwa::charts::Index::from_zero_based(index),
    )?;
    value
        .map(|value| {
            Ok::<Symbol, Box<dyn std::error::Error>>(Symbol {
                shape: match value.shape {
                    ChartSeriesSymbolShape::Circle => SymbolShape::Circle,
                    ChartSeriesSymbolShape::Triangle => SymbolShape::Triangle,
                    ChartSeriesSymbolShape::InvertedTriangle => SymbolShape::InvertedTriangle,
                    ChartSeriesSymbolShape::Diamond => SymbolShape::Diamond,
                    ChartSeriesSymbolShape::Square => SymbolShape::Square,
                },
                size: value
                    .size
                    .map(|size| SymbolSize::new(size.points()))
                    .transpose()?,
            })
        })
        .transpose()
}

fn leaf_symbol_read(source: &[u8], index: usize) -> TestResult<Option<Symbol>> {
    let values = leaf_read(source)?;
    values
        .get(index)
        .cloned()
        .ok_or_else(|| "leaf series index was not present".into())
}

fn direct_symbol_read_proof(source: &[u8], chart: u64, index: usize) -> TestResult {
    assert_eq!(
        host_symbol_read(source, chart, index)?,
        leaf_symbol_read(source, index)?,
        "host single-series and leaf projected readers must agree directly"
    );
    Ok(())
}

fn leaf_read(source: &[u8]) -> TestResult<Vec<Option<Symbol>>> {
    let package = Package::from_bytes(source)?;
    let Property::Symbols(values) = package.slide_chart_series(LEAF_SLIDE, LEAF_CHART)? else {
        return Err("unexpected chart-series property".into());
    };
    Ok(values)
}

fn host_write(
    source: &[u8],
    chart: u64,
    values: &Values,
    one: Option<usize>,
) -> TestResult<Vec<u8>> {
    let mut editor = keynote::Draft::from_bytes(source)?;
    let host_values = values.host()?;
    match one {
        None => editor.set_slide_chart_series_symbols(SLIDE, chart, &host_values)?,
        Some(index) => editor.set_slide_chart_series_symbol(
            SLIDE,
            chart,
            litchi_iwa::charts::Index::from_zero_based(index),
            host_values[index],
        )?,
    }
    Ok(editor.to_bytes()?)
}

fn leaf_write(source: &[u8], values: &Values, one: Option<usize>) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut symbols = match one {
        None => values.leaf()?,
        Some(_index) => leaf_read(source)?,
    };
    if let Some(index) = one {
        symbols[index] = values.leaf()?[index];
    }
    let mut edit = package.edit_slide_chart_series(LEAF_SLIDE, LEAF_CHART)?;
    edit.set(symbols)?;
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_slide_chart_series(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn direct_read_proof(source: &[u8], chart: u64) -> TestResult {
    let host = host_read(source, chart)?;
    let leaf = leaf_read(source)?;
    assert_eq!(host, leaf, "host and leaf readers must agree directly");
    Ok(())
}

fn input_with_values(values: &Values) -> TestResult<(Input, u64)> {
    let mut input = input(Format::Keynote, Kind::Line2d)?;
    let chart = input.chart;
    input.fixture.bytes = host_write(&input.fixture.bytes, chart, values, None)?;
    direct_read_proof(&input.fixture.bytes, chart)?;
    Ok((input, chart))
}

fn chart_payload(source: &[u8], chart: u64) -> TestResult<Vec<u8>> {
    let components = Components::read(source)?;
    let location = components.locate(chart, CHART_MESSAGE_TYPE)?;
    Ok(components.payload(chart, location)?.to_vec())
}

fn rewrite_chart_payload(source: &[u8], chart: u64, payload: Vec<u8>) -> TestResult<Vec<u8>> {
    let components = Components::read(source)?;
    publish(
        source,
        components,
        vec![(chart, CHART_MESSAGE_TYPE, payload)],
    )
}

fn unknown_marker() -> Vec<u8> {
    let mut marker = Vec::new();
    encode_varint_into(&mut marker, 12_345_u64 << 3);
    marker.push(7);
    marker
}

fn source_with_unknown_chart_field(source: &[u8], chart: u64) -> TestResult<(Vec<u8>, Vec<u8>)> {
    let marker = unknown_marker();
    let mut payload = chart_payload(source, chart)?;
    litchi_iwa_common::wire::append_varint_field(&mut payload, 12_345, 7)?;
    Ok((rewrite_chart_payload(source, chart, payload)?, marker))
}

fn two_chart_source() -> TestResult<(Vec<u8>, u64, u64)> {
    let input = input(Format::Keynote, Kind::Line2d)?;
    let mut editor = keynote::Draft::from_bytes(&input.fixture.bytes)?;
    let data = ChartData::new(
        vec!["Region 1".to_owned(), "Region 2".to_owned()],
        vec!["April".to_owned(), "May".to_owned(), "June".to_owned()],
        vec![
            vec![Some(17.0), Some(26.0), Some(53.0)],
            vec![Some(55.0), Some(43.0), Some(70.0)],
        ],
    )?;
    let second = editor.add_slide_chart(
        SLIDE,
        Kind::Line2d,
        data,
        DrawablePoint { x: 300.0, y: 140.0 },
        DrawableSize {
            width: 360.0,
            height: 240.0,
        },
    )?;
    Ok((editor.to_bytes()?, input.chart, second.drawable_object_id))
}

fn share_series_style_between_charts(
    source: &[u8],
    first: u64,
    second: u64,
) -> TestResult<Vec<u8>> {
    let components = Components::read(source)?;
    let first_location = components.locate(first, CHART_MESSAGE_TYPE)?;
    let second_location = components.locate(second, CHART_MESSAGE_TYPE)?;
    let styles = chart_style_codec::references(
        components.payload(first, first_location)?,
        ChartCodecLimits::DEFAULT,
    )
    .map_err(|error| error.to_string())?
    .series_theme_styles()
    .to_vec();
    let rewritten = litchi_iwa_common::wire::transform_length_delimited_fields_at_path(
        components.payload(second, second_location)?,
        &[10_000],
        |extension| {
            let references = styles
                .iter()
                .map(|identifier| {
                    let mut reference = Vec::new();
                    litchi_iwa_common::wire::append_varint_field(&mut reference, 1, *identifier)?;
                    Ok(reference)
                })
                .collect::<litchi_iwa_common::Result<Vec<_>>>()?;
            litchi_iwa_common::wire::rewrite_repeated_length_delimited_fields(
                extension,
                17,
                &references,
            )
        },
    )?;
    publish(
        source,
        components,
        vec![(second, CHART_MESSAGE_TYPE, rewritten)],
    )
}

fn assert_known_locality(source: &[u8], candidate: &[u8]) -> TestResult {
    let source_catalog = Catalog::from_bytes(source, Default::default())?;
    let candidate_catalog = Catalog::from_bytes(candidate, Default::default())?;
    for member in ["Index/Document.iwa", "Index/ViewState.iwa"] {
        assert_eq!(
            source_catalog
                .iter()
                .find(|entry| entry.name() == member)
                .map(|entry| entry.data()),
            candidate_catalog
                .iter()
                .find(|entry| entry.name() == member)
                .map(|entry| entry.data()),
            "unrelated Keynote member {member} changed"
        );
    }
    Ok(())
}

fn run_case(name: &'static str, values: Values, one: Option<usize>) -> TestResult {
    let input = input(Format::Keynote, Kind::Line2d)?;
    let chart = input.chart;
    let values_for_host = values.clone();
    let values_for_leaf = values.clone();
    Case::new(
        name,
        Format::Keynote,
        Tier::B,
        Evidence::Queued,
        &input.fixture,
    )
    .host(move |bytes| host_write(bytes, chart, &values_for_host, one))
    .leaf(move |bytes| leaf_write(bytes, &values_for_leaf, one))
    .read("host/leaf chart-series symbols", move |bytes| {
        host_read(bytes, chart)
    })
    .leaf_invalidates_root_previews()
    .leaf_structural_metadata()
    .run()?;
    Ok(())
}

#[test]
fn keynote_slide_chart_series_symbols_allocate_facade_parity() -> TestResult {
    let _tier = Tier::B;
    let _evidence = Evidence::Queued;
    run_case(
        "keynote_slide_chart_series_symbols_allocate_facade_parity",
        Values(vec![
            Some((SymbolShape::Diamond, Some(9.0))),
            Some((SymbolShape::Square, None)),
        ]),
        None,
    )
}

#[test]
fn keynote_slide_chart_series_symbol_allocate_facade_parity() -> TestResult {
    let _tier = Tier::B;
    let _evidence = Evidence::Queued;
    run_case(
        "keynote_slide_chart_series_symbol_allocate_facade_parity",
        Values(vec![Some((SymbolShape::Diamond, Some(11.0))), None]),
        Some(0),
    )
}

#[test]
fn keynote_slide_chart_series_symbols_read_facade_parity() -> TestResult {
    let _tier = Tier::A;
    let _evidence = Evidence::Queued;
    let values = Values(vec![
        Some((SymbolShape::Diamond, Some(9.0))),
        Some((SymbolShape::Square, None)),
    ]);
    let (input, chart) = input_with_values(&values)?;
    direct_read_proof(&input.fixture.bytes, chart)?;
    assert_eq!(host_read(&input.fixture.bytes, chart)?, values.leaf()?);
    Ok(())
}

#[test]
fn keynote_slide_chart_series_symbol_read_facade_parity() -> TestResult {
    let _tier = Tier::A;
    let _evidence = Evidence::Queued;
    let values = Values(vec![Some((SymbolShape::Triangle, Some(8.0))), None]);
    let (input, chart) = input_with_values(&values)?;
    direct_symbol_read_proof(&input.fixture.bytes, chart, 0)?;
    assert_eq!(
        host_symbol_read(&input.fixture.bytes, chart, 0)?,
        values.leaf()?[0]
    );
    Ok(())
}

#[test]
fn keynote_slide_chart_series_noop_inverse_and_refusal_controls() -> TestResult {
    let values = Values(vec![
        Some((SymbolShape::Diamond, Some(9.0))),
        Some((SymbolShape::Square, None)),
    ]);
    let (input, _chart) = input_with_values(&values)?;
    let package = Package::from_bytes(&input.fixture.bytes)?;
    let current = values.leaf()?;
    let mut edit = package.edit_slide_chart_series(LEAF_SLIDE, LEAF_CHART)?;
    edit.set(current)?;
    let noop = edit.commit()?;
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(leaf_bytes!(noop.package()), input.fixture.bytes);

    let mut changed = package.edit_slide_chart_series(LEAF_SLIDE, LEAF_CHART)?;
    changed.set(Values(vec![None, None]).leaf()?)?;
    let changed = changed.commit()?;
    let forward = package.apply_slide_chart_series(changed.patch())?;
    assert_eq!(
        leaf_bytes!(forward.package()),
        leaf_bytes!(changed.package()),
        "exact forward patch application must reproduce the committed candidate"
    );
    assert!(matches!(
        changed.package().apply_slide_chart_series(changed.patch()),
        Err(Error::PatchConflict)
    ));
    assert_eq!(
        leaf_read(&leaf_bytes!(changed.package()))?,
        vec![None, None]
    );
    let restored = changed
        .package()
        .apply_slide_chart_series(&changed.patch().inverse())?;
    assert_eq!(leaf_bytes!(restored.package()), input.fixture.bytes);

    assert!(matches!(
        package.slide_chart_series(LEAF_SLIDE, ChartSelector::index(99)),
        Err(Error::Refused(Refusal::ChartPositionNotFound { .. }))
    ));
    assert!(matches!(
        package
            .edit_slide_chart_series(LEAF_SLIDE, LEAF_CHART)?
            .set(Vec::new()),
        Err(Error::Refused(Refusal::SeriesCount))
    ));
    let mut mismatch = package.edit_slide_chart_series(LEAF_SLIDE, LEAF_CHART)?;
    mismatch.set(vec![None])?;
    assert!(matches!(
        mismatch.commit(),
        Err(Error::Refused(Refusal::SeriesCount))
    ));
    Ok(())
}

#[test]
fn keynote_slide_chart_series_unknown_and_locality_parity() -> TestResult {
    let values = Values(vec![
        Some((SymbolShape::Diamond, Some(9.0))),
        Some((SymbolShape::Square, None)),
    ]);
    let input = input(Format::Keynote, Kind::Line2d)?;
    let chart = input.chart;
    let (source, marker) = source_with_unknown_chart_field(&input.fixture.bytes, chart)?;
    let host = host_write(&source, chart, &values, None)?;
    let leaf = leaf_write(&source, &values, None)?;
    let leaf_bytes = leaf.output;
    assert_eq!(host_read(&host, chart)?, leaf_read(&leaf_bytes)?);
    assert_known_locality(&source, &leaf_bytes)?;
    let payload = chart_payload(&leaf_bytes, chart)?;
    assert!(
        payload.windows(marker.len()).any(|window| window == marker),
        "selected chart unknown field was not preserved"
    );
    Ok(())
}

#[test]
fn keynote_slide_chart_series_malformed_and_low_budget_refuse() -> TestResult {
    let input = input(Format::Keynote, Kind::Line2d)?;
    let malformed = rewrite_chart_payload(&input.fixture.bytes, input.chart, vec![0xff])?;
    let package = Package::from_bytes(&malformed)?;
    assert!(matches!(
        package.slide_chart_series(LEAF_SLIDE, LEAF_CHART),
        Err(Error::InvalidSource)
    ));

    let limits = SemanticLimits::builder().max_references(1)?.build()?;
    let package = Package::from_bytes_with_options(
        &input.fixture.bytes,
        OpenOptions::new().semantic_limits(limits),
    )?;
    assert!(matches!(
        package.slide_chart_series(LEAF_SLIDE, LEAF_CHART),
        Err(Error::Limit { .. })
    ));

    let object_limited = Package::from_bytes_with_options(
        &input.fixture.bytes,
        OpenOptions::new().semantic_limits(SemanticLimits::builder().max_objects(1)?.build()?),
    );
    match object_limited {
        Err(error) => assert!(
            format!("{error:?}").contains("Objects"),
            "object census must remain typed: {error:?}"
        ),
        Ok(_) => panic!("object census must refuse before the series facade"),
    }
    Ok(())
}

#[test]
fn keynote_slide_chart_series_shared_style_refuses_atomically() -> TestResult {
    let (source, first, second) = two_chart_source()?;
    let source = share_series_style_between_charts(&source, first, second)?;
    let package = Package::from_bytes(&source)?;
    let before = leaf_bytes!(package.snapshot());
    let mut edit = package.edit_slide_chart_series(LEAF_SLIDE, ChartSelector::index(2))?;
    edit.set(
        Values(vec![
            Some((SymbolShape::Diamond, Some(12.0))),
            Some((SymbolShape::Square, None)),
        ])
        .leaf()?,
    )?;
    assert!(matches!(
        edit.commit(),
        Err(Error::Refused(Refusal::UnsupportedDependency))
    ));
    assert_eq!(leaf_bytes!(package.snapshot()), before);
    Ok(())
}
