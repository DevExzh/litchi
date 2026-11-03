//! Value-axis bounds, step counts and scale (wave W4-I, record 0422): the
//! host's `set_{body,sheet}_chart_value_axis_{bounds,steps,scale}` and the
//! Pages and Numbers facades over `litchi_iwa_style::value_axis` (the rules
//! lifted from Keynote's `slide_chart_value_axis`), on a column chart the
//! host added to the native-saved chart fixture.
//!
//! Every write is set through the host and through the facade after a host
//! write of the starting state, with the facade's inverse replayed; the
//! reads of every state are compared too.

use litchi_iwa::{numbers, pages};
use litchi_iwa_common::chart::axis::{
    Bound, Bounds, MajorStepCount, MinorStepCount, Scale, Steps, ValueAxisSettings,
};
use litchi_iwa_common::chart::kind::Kind;

use crate::chart_style_cases::{Target, input};
use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier};

/// One host value-axis write.
#[derive(Debug, Clone, Copy)]
enum Write {
    Bounds(Bounds),
    Steps(Steps),
    Scale(Scale),
}

fn host_write(source: &[u8], target: Target, chart: u64, writes: &[Write]) -> TestResult<Vec<u8>> {
    Ok(match target {
        Target::Numbers { sheet } => {
            let mut editor = numbers::Draft::from_bytes(source)?;
            for write in writes {
                match *write {
                    Write::Bounds(bounds) => {
                        editor.set_sheet_chart_value_axis_bounds(sheet, chart, bounds)?;
                    },
                    Write::Steps(steps) => {
                        editor.set_sheet_chart_value_axis_steps(sheet, chart, steps)?;
                    },
                    Write::Scale(scale) => {
                        editor.set_sheet_chart_value_axis_scale(sheet, chart, scale)?;
                    },
                }
            }
            editor.to_bytes()?
        },
        Target::Pages => {
            let mut editor = pages::Draft::from_bytes(source)?;
            for write in writes {
                match *write {
                    Write::Bounds(bounds) => {
                        editor.set_body_chart_value_axis_bounds(chart, bounds)?
                    },
                    Write::Steps(steps) => editor.set_body_chart_value_axis_steps(chart, steps)?,
                    Write::Scale(scale) => editor.set_body_chart_value_axis_scale(chart, scale)?,
                }
            }
            editor.to_bytes()?
        },
        Target::Keynote { .. } => return Err("Keynote has no host value-axis route".into()),
    })
}

/// The host's three reads, in one rendering.
fn host_read(bytes: &[u8], target: Target, chart: u64) -> TestResult<String> {
    Ok(match target {
        Target::Numbers { sheet } => {
            let editor = numbers::Draft::from_bytes(bytes)?;
            format!(
                "{:?} {:?} {:?}",
                editor.sheet_chart_value_axis_bounds(sheet, chart)?,
                editor.sheet_chart_value_axis_steps(sheet, chart)?,
                editor.sheet_chart_value_axis_scale(sheet, chart)?,
            )
        },
        Target::Pages => {
            let editor = pages::Draft::from_bytes(bytes)?;
            format!(
                "{:?} {:?} {:?}",
                editor.body_chart_value_axis_bounds(chart)?,
                editor.body_chart_value_axis_steps(chart)?,
                editor.body_chart_value_axis_scale(chart)?,
            )
        },
        Target::Keynote { .. } => return Err("Keynote has no host value-axis route".into()),
    })
}

fn render(settings: ValueAxisSettings) -> String {
    format!(
        "{:?} {:?} {:?}",
        settings.bounds(),
        settings.steps(),
        settings.scale()
    )
}

/// The facade read in the host read's rendering.
fn facade_read(format: Format, bytes: &[u8]) -> TestResult<String> {
    Ok(render(match format {
        Format::Numbers => {
            litchi_numbers::Package::from_bytes(bytes)?.sheet_chart_value_axis_settings(0, 1)?
        },
        Format::Pages => {
            litchi_pages::Package::from_bytes(bytes)?.body_chart_value_axis_settings(0)?
        },
        Format::Keynote => return Err("Keynote has no host value-axis route".into()),
    }))
}

/// The facade writes of `writes`, staged on one edit, with the inverse
/// replayed.
fn facade_write(format: Format, source: &[u8], writes: &[Write]) -> TestResult<LeafRun> {
    Ok(match format {
        Format::Numbers => {
            let package = litchi_numbers::Package::from_bytes(source)?;
            let mut edit = package.edit_sheet_chart_value_axis_settings(0, 1)?;
            for write in writes {
                match *write {
                    Write::Bounds(bounds) => edit.set_bounds(bounds)?,
                    Write::Steps(steps) => edit.set_steps(steps)?,
                    Write::Scale(scale) => edit.set_scale(scale)?,
                };
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_sheet_chart_value_axis_settings(&commit.patch().inverse())?;
            LeafRun::with_inverse(
                crate::leaf_bytes!(commit.package()),
                crate::leaf_bytes!(restored.package()),
            )
        },
        Format::Pages => {
            let package = litchi_pages::Package::from_bytes(source)?;
            let mut edit = package.edit_body_chart_value_axis_settings(0)?;
            for write in writes {
                edit = match *write {
                    Write::Bounds(bounds) => edit.set_bounds(bounds)?,
                    Write::Steps(steps) => edit.set_steps(steps)?,
                    Write::Scale(scale) => edit.set_scale(scale)?,
                };
            }
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_body_chart_value_axis_settings(&commit.patch().inverse())?;
            LeafRun::with_inverse(
                crate::leaf_bytes!(commit.package()),
                crate::leaf_bytes!(restored.package()),
            )
        },
        Format::Keynote => return Err("Keynote has no host value-axis route".into()),
    })
}

fn bound(value: f64) -> TestResult<Bound> {
    Ok(Bound::new(value)?)
}

fn fixed(low: f64, high: f64) -> TestResult<Bounds> {
    Ok(Bounds::fixed(bound(low)?, bound(high)?)?)
}

fn steps(major: u32, minor: u32) -> TestResult<Steps> {
    Ok(Steps::fixed(
        MajorStepCount::new(major)?,
        MinorStepCount::new(minor)?,
    ))
}

/// The sequences: (starting host writes, the written change).
fn sequences() -> TestResult<Vec<(Vec<Write>, Vec<Write>)>> {
    Ok(vec![
        // Into an axis without the generated fields.
        (vec![], vec![Write::Bounds(fixed(0.0, 50.0)?)]),
        // The host-added chart stores the native 5 major and 1 minor steps.
        (vec![], vec![Write::Steps(steps(7, 3)?)]),
        (vec![], vec![Write::Scale(Scale::Logarithmic)]),
        (
            vec![],
            vec![Write::Bounds(Bounds::new(None, Some(bound(80.0)?))?)],
        ),
        // In place.
        (
            vec![Write::Bounds(fixed(-5.0, 50.0)?)],
            vec![Write::Bounds(fixed(0.0, 30.0)?)],
        ),
        (
            vec![Write::Steps(steps(4, 2)?)],
            vec![Write::Steps(steps(6, 0)?)],
        ),
        // Both step fields absent, then appended.
        (
            vec![Write::Steps(Steps::automatic())],
            vec![Write::Steps(steps(7, 3)?)],
        ),
        (
            vec![Write::Steps(Steps::automatic())],
            vec![
                Write::Bounds(fixed(2.0, 20.0)?),
                Write::Steps(steps(2, 1)?),
                Write::Scale(Scale::Logarithmic),
            ],
        ),
        (
            vec![Write::Bounds(fixed(1.0, 1000.0)?)],
            vec![Write::Scale(Scale::Logarithmic)],
        ),
        // Back to automatic, the stored scale kept explicit.
        (
            vec![
                Write::Bounds(fixed(-5.0, 50.0)?),
                Write::Steps(steps(4, 2)?),
            ],
            vec![Write::Bounds(Bounds::automatic())],
        ),
        (
            vec![Write::Steps(steps(4, 2)?)],
            vec![Write::Steps(Steps::automatic())],
        ),
        (
            vec![Write::Scale(Scale::Logarithmic)],
            vec![Write::Scale(Scale::Linear)],
        ),
        // Several components in one edit (the host writes them in order).
        (
            vec![Write::Steps(steps(3, 1)?)],
            vec![
                Write::Bounds(fixed(10.0, 100.0)?),
                Write::Steps(steps(9, 0)?),
                Write::Scale(Scale::Logarithmic),
            ],
        ),
    ])
}

fn set_case(name: &'static str, format: Format) -> TestResult {
    for (before, after) in sequences()? {
        let mut input = input(format, Kind::Column2d)?;
        let (target, chart) = (input.target, input.chart);
        if !before.is_empty() {
            input.fixture.bytes = host_write(&input.fixture.bytes, target, chart, &before)?;
        }
        let writes = after.clone();
        Case::new(name, format, Tier::A, Evidence::Queued, &input.fixture)
            .host(move |source| host_write(source, target, chart, &writes))
            .leaf(move |source| facade_write(format, source, &after))
            .read("host value-axis read", move |bytes| {
                host_read(bytes, target, chart)
            })
            .read("leaf value-axis read", move |bytes| {
                facade_read(format, bytes)
            })
            .run()
            .map_err(|error| format!("{before:?}: {error}"))?;
    }
    Ok(())
}

#[test]
fn numbers_sheet_chart_value_axis_set_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    set_case("numbers_sheet_chart_value_axis_set_parity", Format::Numbers)
}

#[test]
fn pages_body_chart_value_axis_set_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    set_case("pages_body_chart_value_axis_set_parity", Format::Pages)
}

fn read_case(format: Format) -> TestResult {
    let mut states = vec![Vec::new()];
    for (before, after) in sequences()? {
        states.push(before.clone());
        states.push(before.into_iter().chain(after).collect());
    }
    for writes in states {
        let input = input(format, Kind::Column2d)?;
        let bytes = host_write(&input.fixture.bytes, input.target, input.chart, &writes)?;
        assert_eq!(
            host_read(&bytes, input.target, input.chart)?,
            facade_read(format, &bytes)?,
            "{writes:?}"
        );
    }
    Ok(())
}

#[test]
fn numbers_sheet_chart_value_axis_read_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    read_case(Format::Numbers)
}

#[test]
fn pages_body_chart_value_axis_read_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    read_case(Format::Pages)
}

/// Stricter admission (record 0422, macOS-queued): the host stores a
/// logarithmic scale over a non-positive manual endpoint, which the native
/// applications cannot display; the leaf refuses the write with
/// `Refusal::InvalidSettings` before touching the source, as Keynote's
/// value-axis family always has. A source that stores such settings is an
/// invalid source for the leaf read.
#[test]
fn chart_value_axis_logarithmic_non_positive_bound_refuses() -> TestResult {
    for format in [Format::Numbers, Format::Pages] {
        let input = input(format, Kind::Column2d)?;
        let (target, chart) = (input.target, input.chart);
        let source = host_write(
            &input.fixture.bytes,
            target,
            chart,
            &[Write::Bounds(fixed(-5.0, 50.0)?)],
        )?;
        // The host accepts the combination.
        let host = host_write(&source, target, chart, &[Write::Scale(Scale::Logarithmic)])?;
        match format {
            Format::Numbers => {
                let package = litchi_numbers::Package::from_bytes(&source)?;
                let mut edit = package.edit_sheet_chart_value_axis_settings(0, 1)?;
                let refused = edit.set_scale(Scale::Logarithmic).map(|_| ()).unwrap_err();
                assert_eq!(
                    refused,
                    litchi_numbers::sheet::chart::value_axis::Error::Refused(
                        litchi_numbers::sheet::chart::value_axis::Refusal::InvalidSettings
                    )
                );
                assert!(
                    litchi_numbers::Package::from_bytes(&host)?
                        .sheet_chart_value_axis_settings(0, 1)
                        .is_err()
                );
            },
            Format::Pages => {
                let package = litchi_pages::Package::from_bytes(&source)?;
                let refused = package
                    .edit_body_chart_value_axis_settings(0)?
                    .set_scale(Scale::Logarithmic)
                    .map(|_| ())
                    .unwrap_err();
                assert_eq!(
                    refused,
                    litchi_pages::chart::value_axis::Error::Refused(
                        litchi_pages::chart::value_axis::Refusal::InvalidSettings
                    )
                );
                assert!(
                    litchi_pages::Package::from_bytes(&host)?
                        .body_chart_value_axis_settings(0)
                        .is_err()
                );
            },
            Format::Keynote => {},
        }
    }
    Ok(())
}
