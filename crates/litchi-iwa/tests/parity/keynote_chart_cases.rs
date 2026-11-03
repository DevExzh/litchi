//! Keynote chart-lifecycle parity cases (wave W1-C, record 0372): kind,
//! series direction and frame through `chart::lifecycle`.

use litchi_iwa::keynote;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_keynote::chart::lifecycle::{Direction, Kind};
use litchi_keynote::slide::drawable::Geometry;
use litchi_keynote::{ChartSelector, Package, SlideSelector};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

const SLIDE: usize = 0;

/// `basic.key` with one host-created column chart on slide 0: the host
/// edits only charts whose styles live in the slide component (native
/// charts keep theirs in the document stylesheet, which the host refuses).
fn chart_fixture() -> TestResult<Fixture> {
    let base = fixture("iwork/keynote/basic.key")?;
    let mut editor = keynote::Draft::from_bytes(&base.bytes)?;
    let data = litchi_iwa_common::chart::data::ChartData::new(
        vec!["2025".into(), "2026".into()],
        vec!["North".into(), "South".into(), "West".into()],
        vec![
            vec![Some(1.0), Some(2.0), Some(3.0)],
            vec![Some(4.0), None, Some(6.0)],
        ],
    )?;
    editor.add_slide_chart(
        SLIDE,
        Kind::Column2d,
        data,
        DrawablePoint { x: 80.0, y: 160.0 },
        DrawableSize {
            width: 520.0,
            height: 320.0,
        },
    )?;
    Ok(Fixture {
        path: "iwork/keynote/basic.key + host add_slide_chart(column 2D, 2x3)",
        bytes: editor.to_bytes()?,
    })
}

fn host_charts(bytes: &[u8]) -> TestResult<Vec<String>> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    Ok(editor
        .slide_charts(SLIDE)?
        .into_iter()
        .map(|chart| {
            format!(
                "{:?} {:?} {:?}",
                chart.kind, chart.direction, chart.geometry
            )
        })
        .collect())
}

fn leaf_charts(bytes: &[u8]) -> TestResult<Vec<String>> {
    let package = Package::from_bytes(bytes)?;
    let count = package
        .slide_chart_catalog(SlideSelector::index(SLIDE))?
        .len();
    let mut charts = Vec::with_capacity(count);
    for index in 0..count {
        charts.push(format!(
            "{:?} {:?}",
            package
                .slide_chart_metadata(SlideSelector::index(SLIDE), ChartSelector::index(index))?
                .kind(),
            package
                .slide_chart_direction(SlideSelector::index(SLIDE), ChartSelector::index(index))?
        ));
    }
    Ok(charts)
}

fn leaf_run(
    source: &[u8],
    stage: impl FnOnce(&mut litchi_keynote::chart::lifecycle::Edit<'_>) -> TestResult,
) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_slide_chart_lifecycle();
    stage(&mut edit)?;
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_slide_chart_lifecycle(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn run(
    name: &'static str,
    tier: Tier,
    evidence: Evidence,
    host: impl FnOnce(&mut keynote::Draft) -> TestResult,
    leaf: impl FnOnce(&[u8]) -> TestResult<LeafRun>,
) -> TestResult {
    let input = chart_fixture()?;
    Case::new(name, Format::Keynote, tier, evidence, &input)
        .host(|source| {
            let mut editor = keynote::Draft::from_bytes(source)?;
            host(&mut editor)?;
            Ok(editor.to_bytes()?)
        })
        .leaf(leaf)
        .leaf_invalidates_root_previews()
        .leaf_structural_metadata()
        .read("host charts", host_charts)
        .read("leaf charts", leaf_charts)
        .run()?;
    Ok(())
}

#[test]
fn keynote_slide_chart_kind_parity() -> TestResult {
    run(
        "keynote_slide_chart_kind_parity",
        Tier::B,
        Evidence::Queued,
        |editor| {
            editor.set_slide_chart_kind(SLIDE, ChartSelector::index(0), Kind::Line2d)?;
            Ok(())
        },
        |source| {
            leaf_run(source, |edit| {
                edit.set_kind(
                    SlideSelector::index(SLIDE),
                    ChartSelector::index(0),
                    Kind::Line2d,
                )?;
                Ok(())
            })
        },
    )
}

#[test]
fn keynote_slide_chart_direction_parity() -> TestResult {
    run(
        "keynote_slide_chart_direction_parity",
        Tier::B,
        Evidence::Queued,
        |editor| {
            editor.set_slide_chart_direction(SLIDE, ChartSelector::index(0), Direction::Columns)?;
            Ok(())
        },
        |source| {
            leaf_run(source, |edit| {
                edit.set_direction(
                    SlideSelector::index(SLIDE),
                    ChartSelector::index(0),
                    Direction::Columns,
                )?;
                Ok(())
            })
        },
    )
}

#[test]
fn keynote_slide_chart_geometry_parity() -> TestResult {
    let frame = Geometry::new(150.0, 120.0, 640.0, 360.0);
    run(
        "keynote_slide_chart_geometry_parity",
        Tier::B,
        Evidence::Queued,
        |editor| {
            let mut geometry = editor.slide_charts(SLIDE)?[0].geometry;
            geometry.position = Some(DrawablePoint {
                x: frame.x,
                y: frame.y,
            });
            geometry.size = Some(DrawableSize {
                width: frame.width,
                height: frame.height,
            });
            editor.set_slide_chart_geometry(SLIDE, ChartSelector::index(0), geometry)?;
            Ok(())
        },
        |source| {
            leaf_run(source, |edit| {
                edit.set_geometry(SlideSelector::index(SLIDE), ChartSelector::index(0), frame)?;
                Ok(())
            })
        },
    )
}

#[test]
fn keynote_slide_chart_remove_parity() -> TestResult {
    run(
        "keynote_slide_chart_remove_parity",
        Tier::B,
        Evidence::Queued,
        |editor| {
            editor.remove_slide_chart(SLIDE, ChartSelector::index(0))?;
            Ok(())
        },
        |source| {
            leaf_run(source, |edit| {
                edit.remove(SlideSelector::index(SLIDE), ChartSelector::index(0))?;
                Ok(())
            })
        },
    )
}

#[test]
fn keynote_slide_chart_duplicate_parity() -> TestResult {
    run(
        "keynote_slide_chart_duplicate_parity",
        Tier::B,
        Evidence::Queued,
        |editor| {
            editor.duplicate_slide_chart(SLIDE, ChartSelector::index(0))?;
            Ok(())
        },
        |source| {
            leaf_run(source, |edit| {
                edit.duplicate(SlideSelector::index(SLIDE), ChartSelector::index(0))?;
                Ok(())
            })
        },
    )
}

/// A complete replacement (new labels and shape) through
/// `chart::lifecycle::Edit::set_data`, and a same-shape update, which the
/// host delegates to `edit_slide_chart_data`.
#[test]
fn keynote_slide_chart_data_set_parity() -> TestResult {
    use litchi_iwa_common::chart::data::ChartData;
    let reshaped = ChartData::new(
        vec!["Q1".into(), "Q2".into(), "Q3".into()],
        vec!["Revenue".into(), "Cost".into()],
        vec![
            vec![Some(10.0), Some(4.0)],
            vec![Some(12.5), None],
            vec![Some(9.0), Some(3.5)],
        ],
    )?;
    let reshaped_leaf = reshaped.clone();
    run(
        "keynote_slide_chart_data_set_parity",
        Tier::B,
        Evidence::Queued,
        |editor| {
            editor.set_slide_chart_data(SLIDE, ChartSelector::index(0), reshaped)?;
            Ok(())
        },
        |source| {
            leaf_run(source, |edit| {
                edit.set_data(
                    SlideSelector::index(SLIDE),
                    ChartSelector::index(0),
                    reshaped_leaf,
                )?;
                Ok(())
            })
        },
    )?;
    let same_shape = ChartData::new(
        vec!["2025".into(), "2026".into()],
        vec!["North".into(), "South".into(), "West".into()],
        vec![
            vec![Some(7.0), Some(8.0), Some(9.0)],
            vec![None, Some(5.0), Some(6.0)],
        ],
    )?;
    let same_shape_leaf = same_shape.clone();
    run(
        "keynote_slide_chart_data_set_parity",
        Tier::B,
        Evidence::Queued,
        |editor| {
            editor.set_slide_chart_data(SLIDE, ChartSelector::index(0), same_shape)?;
            Ok(())
        },
        |source| {
            let package = Package::from_bytes(source)?;
            let commit = package
                .edit_slide_chart_data(SlideSelector::index(SLIDE), ChartSelector::index(0))?
                .set(same_shape_leaf)
                .commit()?;
            let restored = commit
                .package()
                .apply_slide_chart_data(&commit.patch().inverse())?;
            Ok(LeafRun::with_inverse(
                leaf_bytes!(commit.package()),
                leaf_bytes!(restored.package()),
            ))
        },
    )
}

/// Host `slide_charts` against the leaf chart reads (metadata kind,
/// direction, data, frame) on the host-created chart and native charts.
fn charts_agree(bytes: &[u8]) -> TestResult<usize> {
    let editor = keynote::Draft::from_bytes(bytes)?;
    let host = editor.slide_charts(SLIDE)?;
    let package = Package::from_bytes(bytes)?;
    let slide = SlideSelector::index(SLIDE);
    let drawables = package.slide_drawable_list(slide)?;
    let frames = drawables
        .iter()
        .filter(|drawable| drawable.kind() == litchi_keynote::slide::drawable::Kind::Chart)
        .collect::<Vec<_>>();
    if host.len() != package.slide_chart_catalog(slide)?.len() {
        return Err("chart count differs".into());
    }
    for (index, chart) in host.iter().enumerate() {
        let chart_selector = ChartSelector::index(index);
        let metadata = package.slide_chart_metadata(slide, chart_selector)?;
        let data = package.slide_chart_data(slide, chart_selector)?;
        let direction = package.slide_chart_direction(slide, chart_selector)?;
        if metadata.kind() != chart.kind || direction != chart.direction || data != chart.data {
            return Err(format!("chart {index}: host {chart:?}").into());
        }
        let _ = &frames;
    }
    Ok(host.len())
}

#[test]
fn keynote_slide_chart_read_parity() -> TestResult {
    let input = chart_fixture()?;
    Case::new(
        "keynote_slide_chart_read_parity",
        Format::Keynote,
        Tier::A,
        Evidence::Queued,
        &input,
    )
    .read_only()
    .host(|source| Ok(keynote::Draft::from_bytes(source)?.to_bytes()?))
    .leaf(|source| Ok(LeafRun::output(leaf_bytes!(Package::from_bytes(source)?))))
    .read("host and leaf charts agree", charts_agree)
    .run()?;
    Ok(())
}

/// Source-free chart creation on `basic.key` (every concrete kind family is
/// exercised by the host's own tests; this case covers a two-series line
/// chart and a scatter chart).
#[test]
fn keynote_slide_chart_add_parity() -> TestResult {
    let input = fixture("iwork/keynote/basic.key")?;
    let data = || {
        litchi_iwa_common::chart::data::ChartData::new(
            vec!["X".into(), "Y".into()],
            vec!["Q1".into(), "Q2".into(), "Q3".into()],
            vec![
                vec![Some(1.0), Some(2.0), Some(3.0)],
                vec![Some(5.0), None, Some(7.5)],
            ],
        )
    };
    let (line, scatter) = (data()?, data()?);
    let (host_line, host_scatter) = (line.clone(), scatter.clone());
    Case::new(
        "keynote_slide_chart_add_parity",
        Format::Keynote,
        Tier::B,
        Evidence::Queued,
        &input,
    )
    .host(|source| {
        let mut editor = keynote::Draft::from_bytes(source)?;
        editor.add_slide_chart(
            SLIDE,
            Kind::Line2d,
            host_line,
            DrawablePoint { x: 60.0, y: 120.0 },
            DrawableSize {
                width: 480.0,
                height: 300.0,
            },
        )?;
        editor.add_slide_chart(
            SLIDE,
            Kind::Scatter2d,
            host_scatter,
            DrawablePoint { x: 400.0, y: 200.0 },
            DrawableSize {
                width: 300.0,
                height: 240.0,
            },
        )?;
        Ok(editor.to_bytes()?)
    })
    .leaf(|source| {
        leaf_run(source, |edit| {
            edit.add(
                SlideSelector::index(SLIDE),
                Kind::Line2d,
                line,
                Geometry::new(60.0, 120.0, 480.0, 300.0),
            )?;
            edit.add(
                SlideSelector::index(SLIDE),
                Kind::Scatter2d,
                scatter,
                Geometry::new(400.0, 200.0, 300.0, 240.0),
            )?;
            Ok(())
        })
    })
    .leaf_invalidates_root_previews()
    .leaf_structural_metadata()
    .read("host charts", host_charts)
    .read("leaf charts", leaf_charts)
    .run()?;
    Ok(())
}

/// The host's `slide_chart_catalog` (which already reopens the package
/// through the leaf) and the leaf catalog agree on a native chart fixture
/// with a host-added chart, on every slide (record 0419).
#[test]
fn keynote_slide_chart_catalog_read_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    for bytes in [
        chart_fixture()?.bytes,
        fixture("iwork/keynote/chart-data-native.key")?.bytes,
    ] {
        let editor = keynote::Draft::from_bytes(&bytes)?;
        let package = Package::from_bytes(&bytes)?;
        for slide in 0..package.slides()?.len() {
            assert_eq!(
                editor.slide_chart_catalog(slide)?,
                package.slide_chart_catalog(SlideSelector::index(slide))?,
                "slide {slide}"
            );
        }
    }
    Ok(())
}
