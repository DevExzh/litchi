//! The reference-line graph engine (wave W1-E, record 0375a): a chart's
//! primary value-axis reference lines through the host editors and through
//! `litchi_iwa_style::reference_line::set_lines` over a
//! `litchi_iwa_style::graph` workspace of the same input — in-place patch,
//! removal (object, UUID binding, watermark release), allocation (object,
//! item UUID, component UUID, watermark) and the chart's
//! `reference_lines` extension. The chart is a column chart the host added
//! to a native-saved fixture.

use litchi_iwa::{keynote, numbers, pages};
use litchi_iwa_archive::iwa::Limits as CoreLimits;
use litchi_iwa_common::chart::kind::Kind;
use litchi_iwa_common::chart::reference_line::{Line, Value};
use litchi_iwa_style::graph::Workspace;
use litchi_iwa_style::reference_line;

use crate::chart_style_cases::{Components, Target, input, publish};
use crate::harness::{Case, Evidence, Format, LeafRun, TestResult, Tier};
use crate::known_fields::Decoded;
use crate::series_graph_cases::publish_changes;

const CHART_MESSAGE_TYPE: u32 = 5_021;

/// The chart drawable with reference-line item UUIDs zeroed for the known-field
/// decoder. The structural parity policy separately preserves source-existing
/// UUID submessages and masks only known UUID halves on newly allocated items.
fn chart_without_item_uuids(data: &[u8]) -> Result<Decoded, prost::DecodeError> {
    use litchi_iwa_protos::chart_style_codec::{self as codec, Limits};
    fn fail<E>(_: E) -> prost::DecodeError {
        // A truncated varint: the generated decoder's own error.
        <litchi_iwa_protos::tsp::Reference as prost::Message>::decode(&[0x08, 0xff][..])
            .expect_err("a truncated varint fails")
    }
    let normalised = match codec::reference_lines(data, Limits::DEFAULT).map_err(fail)? {
        Some(extension) => {
            let zero = [0x08, 0x00, 0x10, 0x00];
            let extension = litchi_iwa_common::wire::transform_length_delimited_fields_at_path(
                extension,
                &[1, 2],
                |item| -> Result<Vec<u8>, litchi_iwa_common::Error> {
                    litchi_iwa_common::wire::patch_length_delimited_field(
                        item,
                        2,
                        true,
                        Some(&zero),
                    )
                },
            )
            .map_err(fail)?;
            codec::rewrite_reference_lines(data, Some(&extension), Limits::DEFAULT)
                .map_err(fail)?
                .into_bytes()
                .into_owned()
        },
        None => data.to_vec(),
    };
    Ok(Decoded {
        debug: format!("{normalised:02x?}"),
        canonical: normalised,
        unknown_fields: std::collections::BTreeSet::new(),
    })
}

fn host_write(source: &[u8], target: Target, chart: u64, lines: &[Line]) -> TestResult<Vec<u8>> {
    Ok(match target {
        Target::Keynote { slide } => {
            let mut editor = keynote::Draft::from_bytes(source)?;
            editor.set_slide_chart_reference_lines(slide, chart, lines)?;
            editor.to_bytes()?
        },
        Target::Numbers { sheet } => {
            let mut editor = numbers::Draft::from_bytes(source)?;
            editor.set_sheet_chart_reference_lines(sheet, chart, lines)?;
            editor.to_bytes()?
        },
        Target::Pages => {
            let mut editor = pages::Draft::from_bytes(source)?;
            editor.set_body_chart_reference_lines(chart, lines)?;
            editor.to_bytes()?
        },
    })
}

fn host_read(bytes: &[u8], target: Target, chart: u64) -> TestResult<String> {
    Ok(match target {
        Target::Keynote { slide } => format!(
            "{:?}",
            keynote::Draft::from_bytes(bytes)?.slide_chart_reference_lines(slide, chart)?
        ),
        Target::Numbers { sheet } => format!(
            "{:?}",
            numbers::Draft::from_bytes(bytes)?.sheet_chart_reference_lines(sheet, chart)?
        ),
        Target::Pages => format!(
            "{:?}",
            pages::Draft::from_bytes(bytes)?.body_chart_reference_lines(chart)?
        ),
    })
}

fn engine_write(source: &[u8], chart: u64, lines: &[Line]) -> TestResult<Vec<u8>> {
    let components = Components::read(source)?;
    let mut workspace = Workspace::new(&components.archives, CoreLimits::default())?;
    reference_line::set_lines(&mut workspace, chart, lines)?;
    publish_changes(source, &components, workspace.finish())
}

fn corrupt_chart_unknown_payload(source: &[u8], chart: u64) -> TestResult<Vec<u8>> {
    let components = Components::read(source)?;
    let (component, message) = components.locate(chart, CHART_MESSAGE_TYPE)?;
    let payload = components.archives[component]
        .1
        .object(chart)
        .ok_or("chart object vanished")?
        .messages[message]
        .data
        .clone();
    let mut corrupted = payload;
    litchi_iwa_common::wire::append_varint_field(&mut corrupted, 12_345, 1)?;
    publish(
        source,
        components,
        vec![(chart, CHART_MESSAGE_TYPE, corrupted)],
    )
}

fn first_reference_line_item_id(payload: &[u8]) -> TestResult<u64> {
    use litchi_iwa_common::wire::{repeated_length_delimited_payloads, repeated_varint_values};
    use litchi_iwa_protos::chart_style_codec::{self as codec, Limits};

    let extension =
        codec::reference_lines(payload, Limits::DEFAULT)?.ok_or("no reference lines")?;
    let axis = repeated_length_delimited_payloads(extension, 1)?
        .into_iter()
        .next()
        .ok_or("no reference-line axis")?;
    let item = repeated_length_delimited_payloads(axis, 2)?
        .into_iter()
        .next()
        .ok_or("no reference-line item")?;
    let reference = repeated_length_delimited_payloads(item, 1)?
        .into_iter()
        .next()
        .ok_or("no reference-line reference")?;
    repeated_varint_values(reference, 1)?
        .into_iter()
        .next()
        .ok_or_else(|| "reference-line identifier is missing".into())
}

fn corrupt_existing_chart_item_uuid(
    source: &[u8],
    output: &[u8],
    chart: u64,
) -> TestResult<Vec<u8>> {
    use litchi_iwa_common::wire::{
        patch_varint_field, repeated_length_delimited_payloads, repeated_varint_values,
        transform_length_delimited_fields_at_path,
    };
    use litchi_iwa_protos::chart_style_codec::{self as codec, Limits};

    let source_components = Components::read(source)?;
    let source_payload = {
        let (component, message) = source_components.locate(chart, CHART_MESSAGE_TYPE)?;
        source_components.payload(chart, (component, message))?
    };
    let existing_id = first_reference_line_item_id(source_payload)?;

    let components = Components::read(output)?;
    let (component, message) = components.locate(chart, CHART_MESSAGE_TYPE)?;
    let payload = components.payload(chart, (component, message))?;
    let extension =
        codec::reference_lines(payload, Limits::DEFAULT)?.ok_or("no reference lines")?;
    let extension = transform_length_delimited_fields_at_path(
        extension,
        &[1, 2],
        |item| -> Result<Vec<u8>, litchi_iwa_common::Error> {
            let reference = repeated_length_delimited_payloads(item, 1)
                .map_err(|error| litchi_iwa_common::Error::InvalidFormat(error.to_string()))?
                .into_iter()
                .next()
                .ok_or_else(|| {
                    litchi_iwa_common::Error::InvalidFormat(
                        "reference-line reference is missing".to_owned(),
                    )
                })?;
            let identifier = repeated_varint_values(reference, 1)
                .map_err(|error| litchi_iwa_common::Error::InvalidFormat(error.to_string()))?
                .into_iter()
                .next()
                .ok_or_else(|| {
                    litchi_iwa_common::Error::InvalidFormat(
                        "reference-line identifier is missing".to_owned(),
                    )
                })?;
            if identifier != existing_id {
                return Ok(item.to_vec());
            }
            transform_length_delimited_fields_at_path(item, &[2], |uuid| {
                patch_varint_field(uuid, 1, true, Some(999_999))
            })
        },
    )?;
    let payload = codec::rewrite_reference_lines(payload, Some(&extension), Limits::DEFAULT)?
        .into_bytes()
        .into_owned();
    publish(
        output,
        components,
        vec![(chart, CHART_MESSAGE_TYPE, payload)],
    )
}

/// Write `after`, first written with `before` through the host. A case
/// that allocates lines is tier B: the host draws the new items' UUIDs,
/// which live in the chart payload, at random; the engine derives them
/// deterministically, so the chart payload is compared byte for byte with
/// the item UUIDs zeroed.
fn line_case(
    name: &'static str,
    format: Format,
    tier: Tier,
    evidence: Evidence,
    before: &[Line],
    after: &[Line],
) -> TestResult {
    let mut input = input(format, Kind::Column2d)?;
    let (target, chart) = (input.target, input.chart);
    if !before.is_empty() {
        input.fixture.bytes = host_write(&input.fixture.bytes, target, chart, before)?;
    }
    let mut case = Case::new(name, format, tier, evidence, &input.fixture)
        .host(|source| host_write(source, target, chart, after))
        .leaf(|source| Ok(LeafRun::output(engine_write(source, chart, after)?)))
        .read("host reference-line read", move |bytes| {
            host_read(bytes, target, chart)
        });
    if tier == Tier::B {
        case = case
            .leaf_structural_metadata()
            .reference_line_allocation_metadata(chart)
            .known_fields(|known| {
                known.with(
                    CHART_MESSAGE_TYPE,
                    "TSCH.ChartDrawableArchive (item UUIDs zeroed)",
                    chart_without_item_uuids,
                )
            });
    }
    case.run()?;
    Ok(())
}

fn named_average() -> TestResult<Line> {
    Ok(Line::average()
        .try_with_name("Mean")?
        .with_value_visibility(true))
}

fn threshold() -> TestResult<Line> {
    Ok(Line::custom(Value::new(42.5)?).try_with_name("Target")?)
}

#[test]
fn keynote_chart_reference_lines_allocate_parity() -> TestResult {
    line_case(
        "keynote_chart_reference_lines_allocate_parity",
        Format::Keynote,
        Tier::B,
        Evidence::Queued,
        &[],
        &[Line::average()],
    )
}

#[test]
fn numbers_chart_reference_lines_allocate_parity() -> TestResult {
    line_case(
        "numbers_chart_reference_lines_allocate_parity",
        Format::Numbers,
        Tier::B,
        Evidence::Queued,
        &[],
        &[threshold()?],
    )
}

#[test]
fn pages_chart_reference_lines_allocate_parity() -> TestResult {
    line_case(
        "pages_chart_reference_lines_allocate_parity",
        Format::Pages,
        Tier::B,
        Evidence::Queued,
        &[Line::minimum()],
        &[Line::minimum(), Line::maximum()],
    )
}

#[test]
fn pages_chart_reference_lines_allocate_rejects_changed_unknown_payload() -> TestResult {
    let mut input = input(Format::Pages, Kind::Column2d)?;
    let (target, chart) = (input.target, input.chart);
    input.fixture.bytes = host_write(&input.fixture.bytes, target, chart, &[Line::minimum()])?;
    let case = Case::new(
        "pages_chart_reference_lines_allocate_rejects_changed_unknown_payload",
        Format::Pages,
        Tier::B,
        Evidence::Queued,
        &input.fixture,
    )
    .host(|source| host_write(source, target, chart, &[Line::minimum(), Line::maximum()]))
    .leaf(|source| {
        let output = engine_write(source, chart, &[Line::minimum(), Line::maximum()])?;
        Ok(LeafRun::output(corrupt_chart_unknown_payload(
            &output, chart,
        )?))
    })
    .read("host reference-line read", move |bytes| {
        host_read(bytes, target, chart)
    })
    .leaf_structural_metadata()
    .reference_line_allocation_metadata(chart)
    .known_fields(|known| {
        known.with(
            CHART_MESSAGE_TYPE,
            "TSCH.ChartDrawableArchive (item UUIDs zeroed)",
            chart_without_item_uuids,
        )
    });
    let failure = match case.run() {
        Ok(_) => return Err("changed unknown chart payload unexpectedly passed".into()),
        Err(error) => error.to_string(),
    };
    assert!(
        failure.contains("payload differs outside allocation UUIDs"),
        "{failure}"
    );
    Ok(())
}

#[test]
fn pages_chart_reference_lines_allocate_rejects_existing_uuid_corruption() -> TestResult {
    let mut input = input(Format::Pages, Kind::Column2d)?;
    let (target, chart) = (input.target, input.chart);
    input.fixture.bytes = host_write(&input.fixture.bytes, target, chart, &[Line::minimum()])?;
    let case = Case::new(
        "pages_chart_reference_lines_allocate_rejects_existing_uuid_corruption",
        Format::Pages,
        Tier::B,
        Evidence::Queued,
        &input.fixture,
    )
    .host(|source| host_write(source, target, chart, &[Line::minimum(), Line::maximum()]))
    .leaf(move |source| {
        let output = engine_write(source, chart, &[Line::minimum(), Line::maximum()])?;
        Ok(LeafRun::output(corrupt_existing_chart_item_uuid(
            source, &output, chart,
        )?))
    })
    .read("host reference-line read", move |bytes| {
        host_read(bytes, target, chart)
    })
    .leaf_structural_metadata()
    .reference_line_allocation_metadata(chart)
    .known_fields(|known| {
        known.with(
            CHART_MESSAGE_TYPE,
            "TSCH.ChartDrawableArchive (item UUIDs zeroed)",
            chart_without_item_uuids,
        )
    });
    let failure = match case.run() {
        Ok(_) => return Err("existing UUID corruption unexpectedly passed".into()),
        Err(error) => error.to_string(),
    };
    assert!(
        failure.contains("existing reference-line item UUID"),
        "{failure}"
    );
    Ok(())
}

#[test]
fn keynote_chart_reference_lines_patch_parity() -> TestResult {
    line_case(
        "keynote_chart_reference_lines_patch_parity",
        Format::Keynote,
        Tier::A,
        Evidence::Queued,
        &[Line::average()],
        &[named_average()?],
    )
}

#[test]
fn numbers_chart_reference_lines_patch_parity() -> TestResult {
    line_case(
        "numbers_chart_reference_lines_patch_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Queued,
        &[Line::median()],
        &[threshold()?],
    )
}

#[test]
fn pages_chart_reference_lines_remove_parity() -> TestResult {
    line_case(
        "pages_chart_reference_lines_remove_parity",
        Format::Pages,
        Tier::A,
        Evidence::Queued,
        &[Line::minimum(), Line::maximum()],
        &[Line::minimum()],
    )
}

#[test]
fn keynote_chart_reference_lines_remove_all_parity() -> TestResult {
    line_case(
        "keynote_chart_reference_lines_remove_all_parity",
        Format::Keynote,
        Tier::A,
        Evidence::Queued,
        &[Line::median()],
        &[],
    )
}

#[test]
fn numbers_chart_reference_lines_remove_all_parity() -> TestResult {
    line_case(
        "numbers_chart_reference_lines_remove_all_parity",
        Format::Numbers,
        Tier::A,
        Evidence::Queued,
        &[threshold()?, Line::maximum()],
        &[],
    )
}

#[inline(never)]
fn leaf_reference_line_route(
    format: Format,
    source: &[u8],
    chart: u64,
    lines: &[Line],
) -> TestResult<Vec<u8>> {
    crate::series_style_cases::leaf_reopen(format, source)?;
    let output = engine_write(source, chart, lines)?;
    format.reopen(&output)?;
    Ok(output)
}

#[inline(never)]
fn host_reference_line_route(
    source: &[u8],
    target: Target,
    chart: u64,
    lines: &[Line],
) -> TestResult<Vec<u8>> {
    host_write(source, target, chart, lines)
}

/// ABBA cost probe of the reference-line writes with source and candidate
/// admission (record 0375a §4): an allocation, a patch and a removal per
/// format; toggle `*reference_line_cases::host_reference_line_route*` or
/// `*reference_line_cases::leaf_reference_line_route*`.
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn reference_line_write_cost_abba() -> TestResult {
    for format in [Format::Keynote, Format::Numbers, Format::Pages] {
        let runs = [
            (Vec::new(), vec![Line::average()]),
            (vec![Line::average()], vec![named_average()?]),
            (
                vec![Line::minimum(), Line::maximum()],
                vec![Line::minimum()],
            ),
        ];
        for (before, after) in runs {
            let mut input = input(format, Kind::Column2d)?;
            let (target, chart) = (input.target, input.chart);
            if !before.is_empty() {
                input.fixture.bytes = host_write(&input.fixture.bytes, target, chart, &before)?;
            }
            let source = input.fixture.bytes.as_slice();
            std::hint::black_box(host_reference_line_route(source, target, chart, &after)?);
            std::hint::black_box(leaf_reference_line_route(format, source, chart, &after)?);
            std::hint::black_box(leaf_reference_line_route(format, source, chart, &after)?);
            std::hint::black_box(host_reference_line_route(source, target, chart, &after)?);
        }
    }
    Ok(())
}

// ---------------------------------------------------------------------------
// The format facades (record 0419): `edit_{sheet,slide,body}_chart_reference_lines`.
// ---------------------------------------------------------------------------

/// The leaf position of the host-added chart.
const fn leaf_chart(format: Format) -> usize {
    match format {
        Format::Pages => 0,
        Format::Keynote | Format::Numbers => 1,
    }
}

/// Replace the chart's lines with `after` through a format facade, with
/// the inverse replayed.
fn facade_write(format: Format, source: &[u8], after: &[Line]) -> TestResult<LeafRun> {
    let chart = leaf_chart(format);
    Ok(match format {
        Format::Numbers => {
            let package = litchi_numbers::Package::from_bytes(source)?;
            let mut edit = package.edit_sheet_chart_reference_lines(0, chart)?;
            edit.set(after.to_vec())?;
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_sheet_chart_reference_lines(&commit.patch().inverse())?;
            LeafRun::with_inverse(
                crate::leaf_bytes!(commit.package()),
                crate::leaf_bytes!(restored.package()),
            )
        },
        Format::Keynote => {
            let package = litchi_keynote::Package::from_bytes(source)?;
            let mut edit = package.edit_slide_chart_reference_lines(0, chart)?;
            edit.set(after.to_vec())?;
            let commit = edit.commit()?;
            let restored = commit
                .package()
                .apply_slide_chart_reference_lines(&commit.patch().inverse())?;
            LeafRun::with_inverse(
                crate::leaf_bytes!(commit.package()),
                crate::leaf_bytes!(restored.package()),
            )
        },
        Format::Pages => {
            let package = litchi_pages::Package::from_bytes(source)?;
            let commit = package
                .edit_body_chart_reference_lines(chart)?
                .set(after.to_vec())?
                .commit()?;
            let restored = commit
                .package()
                .apply_body_chart_reference_lines(&commit.patch().inverse())?;
            LeafRun::with_inverse(
                crate::leaf_bytes!(commit.package()),
                crate::leaf_bytes!(restored.package()),
            )
        },
    })
}

/// The facade read in the host read's rendering.
fn facade_read(format: Format, bytes: &[u8]) -> TestResult<String> {
    let chart = leaf_chart(format);
    Ok(match format {
        Format::Numbers => format!(
            "{:?}",
            litchi_numbers::Package::from_bytes(bytes)?.sheet_chart_reference_lines(0, chart)?
        ),
        Format::Keynote => format!(
            "{:?}",
            litchi_keynote::Package::from_bytes(bytes)?.slide_chart_reference_lines(0, chart)?
        ),
        Format::Pages => format!(
            "{:?}",
            litchi_pages::Package::from_bytes(bytes)?.body_chart_reference_lines(chart)?
        ),
    })
}

/// `after` written through the host and the facade, after a host write
/// of `before`; allocations are tier B (random host item UUIDs), Keynote
/// always (root previews).
fn facade_case(
    name: &'static str,
    format: Format,
    tier: Tier,
    before: &[Line],
    after: &[Line],
) -> TestResult {
    let mut input = input(format, Kind::Column2d)?;
    let (target, chart) = (input.target, input.chart);
    if !before.is_empty() {
        input.fixture.bytes = host_write(&input.fixture.bytes, target, chart, before)?;
    }
    let mut case = Case::new(name, format, tier, Evidence::Queued, &input.fixture)
        .host(|source| host_write(source, target, chart, after))
        .leaf(|source| facade_write(format, source, after))
        .read("host reference-line read", move |bytes| {
            host_read(bytes, target, chart)
        })
        .read("leaf reference-line read", move |bytes| {
            facade_read(format, bytes)
        });
    if format == Format::Keynote {
        case = case.leaf_invalidates_root_previews();
    }
    if tier == Tier::B {
        case = case
            .leaf_structural_metadata()
            .reference_line_allocation_metadata(chart)
            .known_fields(|known| {
                known.with(
                    CHART_MESSAGE_TYPE,
                    "TSCH.ChartDrawableArchive (item UUIDs zeroed)",
                    chart_without_item_uuids,
                )
            });
    }
    case.run()?;
    Ok(())
}

/// The facade sequences: allocate into an empty axis, patch in place,
/// append, remove one and remove all.
fn facade_sequences() -> TestResult<Vec<(bool, Vec<Line>, Vec<Line>)>> {
    Ok(vec![
        (true, vec![], vec![Line::average(), threshold()?]),
        (false, vec![Line::average()], vec![named_average()?]),
        (
            true,
            vec![Line::minimum()],
            vec![Line::minimum(), Line::maximum()],
        ),
        (
            false,
            vec![Line::minimum(), Line::median()],
            vec![Line::minimum()],
        ),
        (false, vec![Line::median(), threshold()?], vec![]),
    ])
}

fn facade_set(name: &'static str, format: Format, keynote_tier: bool) -> TestResult {
    for (allocates, before, after) in facade_sequences()? {
        let tier = if allocates || keynote_tier {
            Tier::B
        } else {
            Tier::A
        };
        facade_case(name, format, tier, &before, &after)
            .map_err(|error| format!("{before:?} -> {after:?}: {error}"))?;
    }
    Ok(())
}

#[test]
fn numbers_sheet_chart_reference_lines_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    facade_set(
        "numbers_sheet_chart_reference_lines_set_parity",
        Format::Numbers,
        false,
    )
}

#[test]
fn keynote_slide_chart_reference_lines_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    facade_set(
        "keynote_slide_chart_reference_lines_set_parity",
        Format::Keynote,
        true,
    )
}

#[test]
fn pages_body_chart_reference_lines_set_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    facade_set(
        "pages_body_chart_reference_lines_set_parity",
        Format::Pages,
        false,
    )
}

fn facade_reads(format: Format) -> TestResult {
    for (_, before, _) in facade_sequences()? {
        let input = input(format, Kind::Column2d)?;
        let bytes = if before.is_empty() {
            input.fixture.bytes.clone()
        } else {
            host_write(&input.fixture.bytes, input.target, input.chart, &before)?
        };
        assert_eq!(
            host_read(&bytes, input.target, input.chart)?,
            facade_read(format, &bytes)?,
            "{before:?}"
        );
    }
    Ok(())
}

#[test]
fn numbers_sheet_chart_reference_lines_read_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    facade_reads(Format::Numbers)
}

#[test]
fn keynote_slide_chart_reference_lines_read_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    facade_reads(Format::Keynote)
}

#[test]
fn pages_body_chart_reference_lines_read_parity() -> TestResult {
    let _ = (Tier::A, Evidence::Queued);
    facade_reads(Format::Pages)
}
