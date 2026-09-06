//! Strict, source-preserving integration coverage for Keynote chart Arrange.
//!
//! The fixture is deliberately small and source-built.  It contains two
//! charts on one slide, stable semantic titles, the native drawable envelope,
//! and metadata/foreign components supplied by the shared chart fixture.  The
//! public tests never pass a native identifier to the arrangement API; IDs are
//! used only while constructing hostile input and checking physical locality.

use std::io;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::wire::{WireView, append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{Archive, RawMessage};
use litchi_iwa_protos::{kn, tsch, tsd, tsp};
use litchi_keynote::{
    ChartArrangement, ChartArrangementCommit, ChartArrangementDiagnostics, ChartArrangementEdit,
    ChartArrangementError, ChartArrangementLimitKind, ChartArrangementPatch, ChartSelector,
    Package, Position, ReadOptions, SemanticLimits, SlideSelector,
};

use prost::Message as _;

#[path = "support/chart_axis_fixture.rs"]
mod chart_fixture;

use chart_fixture::{
    CHART_MESSAGE_TYPE, CHART_NON_STYLE_MESSAGE_TYPE, CHART_NON_STYLES, CHARTS, DOCUMENT_MEMBER,
    FOREIGN_MEMBER, PREVIEWS, TITLES,
};

const DRAWABLE_SUPER_FIELD: u32 = 1;
const DRAWABLE_LOCKED_FIELD: u32 = 5;
const DRAWABLE_ASPECT_RATIO_LOCKED_FIELD: u32 = 7;
const GENERATED_CHART_EXTENSION_FIELD: u32 = 10_000;
const UNKNOWN_DRAWABLE_FIELD: u32 = 4_090;
const UNKNOWN_DRAWABLE_BYTES_FIELD: u32 = 4_091;

type TestResult<T> = Result<T, Box<dyn std::error::Error>>;

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn base_source() -> TestResult<Vec<u8>> {
    chart_fixture::synthetic_metadata_package(false)
}

/// Add or omit the two optional native fields in one chart drawable.
///
/// The helper appends fields rather than decoding/re-encoding the generated
/// `DrawableArchive`; this keeps the fixture honest about chart-drawable
/// z-order and makes the exact inverse assertion sensitive to field presence.
fn with_drawable_arrangement(
    source: &[u8],
    chart: usize,
    state: (Option<bool>, Option<bool>),
    unknown: bool,
) -> TestResult<Vec<u8>> {
    let original = chart_fixture::message_payload(source, CHARTS[chart], CHART_MESSAGE_TYPE)?;
    let drawable = WireView::parse(&original)?
        .fields()
        .find(|field| field.number() == DRAWABLE_SUPER_FIELD)
        .ok_or_else(|| io::Error::other("missing synthetic chart drawable"))?;
    let mut drawable_payload = drawable.payload().to_vec();
    if let Some(locked) = state.0 {
        append_varint_field(
            &mut drawable_payload,
            DRAWABLE_LOCKED_FIELD,
            u64::from(locked),
        )?;
    }
    if let Some(constrain_proportions) = state.1 {
        append_varint_field(
            &mut drawable_payload,
            DRAWABLE_ASPECT_RATIO_LOCKED_FIELD,
            u64::from(constrain_proportions),
        )?;
    }
    if unknown {
        append_varint_field(&mut drawable_payload, UNKNOWN_DRAWABLE_FIELD, 0x5a5a)?;
        append_length_delimited_field(
            &mut drawable_payload,
            UNKNOWN_DRAWABLE_BYTES_FIELD,
            b"opaque drawable arrange metadata",
        )?;
    }
    let replacement = chart_fixture::length_delimited_field_with_key_width(
        DRAWABLE_SUPER_FIELD,
        &drawable_payload,
        drawable.key().len(),
    );
    let payload =
        chart_fixture::replace_first_field(&original, DRAWABLE_SUPER_FIELD, &replacement)?;
    chart_fixture::with_chart_payload(source, chart, payload)
}

fn synthetic_package_with_states(
    states: [(Option<bool>, Option<bool>); 2],
    unknown_drawable: bool,
    unknown_group: bool,
) -> TestResult<Vec<u8>> {
    let mut source = base_source()?;
    for (chart, state) in states.into_iter().enumerate() {
        source = with_drawable_arrangement(&source, chart, state, unknown_drawable && chart == 0)?;
    }
    if unknown_group {
        source = with_unknown_title_group(&source)?;
    }
    Ok(source)
}

/// Add a balanced unknown protobuf group to the generated chart-title view.
///
/// The chart-title codec admits and retains unknown groups.  Keeping the
/// group in the unselected chart non-style object means the Arrange rewrite
/// must preserve it while the selector still traverses the same graph.
fn with_unknown_title_group(source: &[u8]) -> TestResult<Vec<u8>> {
    let original =
        chart_fixture::message_payload(source, CHART_NON_STYLES[0], CHART_NON_STYLE_MESSAGE_TYPE)?;
    let extension = WireView::parse(&original)?
        .fields()
        .find(|field| field.number() == GENERATED_CHART_EXTENSION_FIELD)
        .ok_or_else(|| io::Error::other("missing synthetic chart title extension"))?;
    let mut generated = extension.payload().to_vec();
    // field 91, start-group; unknown scalar field 1; matching end-group.
    generated.extend([0xdb, 0x05, 0x08, 0x07, 0xdc, 0x05]);
    let replacement = chart_fixture::length_delimited_field_with_key_width(
        GENERATED_CHART_EXTENSION_FIELD,
        &generated,
        extension.key().len(),
    );
    let payload = chart_fixture::replace_first_field(
        &original,
        GENERATED_CHART_EXTENSION_FIELD,
        &replacement,
    )?;
    chart_fixture::with_document_message_payload(
        source,
        CHART_NON_STYLES[0],
        CHART_NON_STYLE_MESSAGE_TYPE,
        payload,
    )
}

fn arrangement<'slide, 'chart>(
    package: &Package,
    slide: impl Into<SlideSelector<'slide>>,
    chart: impl Into<ChartSelector<'chart>>,
) -> TestResult<ChartArrangement> {
    Ok(package.slide_chart_arrangement(slide, chart)?)
}

fn commit_arrangement<'slide, 'chart>(
    package: &Package,
    slide: impl Into<SlideSelector<'slide>>,
    chart: impl Into<ChartSelector<'chart>>,
    target: ChartArrangement,
) -> TestResult<ChartArrangementCommit> {
    Ok(package
        .edit_slide_chart_arrangement(slide, chart)?
        .set(target)
        .commit()?)
}

fn error_is_redacted<E: std::fmt::Debug + std::fmt::Display>(error: &E) {
    let text = format!("{error:?} {error}");
    for identifier in CHARTS.into_iter().chain(TITLES).chain(CHART_NON_STYLES) {
        assert!(
            !text.contains(&identifier.to_string()),
            "native identifier leaked from arrangement error: {text}"
        );
    }
}

fn assert_document_only_changed(before: &[u8], after: &[u8]) -> TestResult<()> {
    let source = Catalog::from_bytes(before)?;
    let target = Catalog::from_bytes(after)?;
    assert_eq!(source.len(), target.len());
    let mut changed = 0;
    for source_entry in source.iter() {
        let target_entry = target
            .iter()
            .find(|entry| entry.name() == source_entry.name())
            .ok_or_else(|| io::Error::other("arrangement removed a package member"))?;
        if source_entry.name() == DOCUMENT_MEMBER {
            assert_ne!(source_entry.data(), target_entry.data());
            changed += 1;
        } else {
            assert_eq!(source_entry.raw_name(), target_entry.raw_name());
            assert_eq!(source_entry.data(), target_entry.data());
            assert_eq!(source_entry.metadata(), target_entry.metadata());
            assert_eq!(source_entry.is_opaque(), target_entry.is_opaque());
        }
    }
    assert_eq!(changed, 1);
    Ok(())
}

fn assert_previews_preserved(before: &[u8], after: &[u8]) -> TestResult<()> {
    let source = Catalog::from_bytes(before)?;
    let target = Catalog::from_bytes(after)?;
    for preview in PREVIEWS {
        let source_data = source
            .iter()
            .filter(|entry| entry.name() == preview)
            .map(|entry| entry.data())
            .collect::<Vec<_>>();
        let target_data = target
            .iter()
            .filter(|entry| entry.name() == preview)
            .map(|entry| entry.data())
            .collect::<Vec<_>>();
        assert_eq!(
            target_data, source_data,
            "Arrange must not invalidate {preview}"
        );
    }
    Ok(())
}

fn drawable_payload(package: &[u8], chart: usize) -> TestResult<Vec<u8>> {
    let payload = chart_fixture::message_payload(package, CHARTS[chart], CHART_MESSAGE_TYPE)?;
    let field = WireView::parse(&payload)?
        .fields()
        .find(|field| field.number() == DRAWABLE_SUPER_FIELD)
        .ok_or_else(|| io::Error::other("missing chart drawable"))?;
    Ok(field.payload().to_vec())
}

fn raw_drawable_fields(package: &[u8], chart: usize, number: u32) -> TestResult<Vec<Vec<u8>>> {
    Ok(WireView::parse(&drawable_payload(package, chart)?)?
        .fields()
        .filter(|field| field.number() == number)
        .map(|field| field.raw().to_vec())
        .collect())
}

fn with_drawable_wire(
    source: &[u8],
    chart: usize,
    rewrite: impl FnOnce(&[u8]) -> TestResult<Vec<u8>>,
) -> TestResult<Vec<u8>> {
    let original = chart_fixture::message_payload(source, CHARTS[chart], CHART_MESSAGE_TYPE)?;
    let drawable = WireView::parse(&original)?
        .fields()
        .find(|field| field.number() == DRAWABLE_SUPER_FIELD)
        .ok_or_else(|| io::Error::other("missing chart drawable"))?;
    let rewritten = rewrite(drawable.payload())?;
    let replacement = chart_fixture::length_delimited_field_with_key_width(
        DRAWABLE_SUPER_FIELD,
        &rewritten,
        drawable.key().len(),
    );
    let payload =
        chart_fixture::replace_first_field(&original, DRAWABLE_SUPER_FIELD, &replacement)?;
    chart_fixture::with_chart_payload(source, chart, payload)
}

fn empty_slide_source(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut archive = Archive::parse(&chart_fixture::document_stream(source)?)?;
    let slide = archive
        .object_mut(chart_fixture::SLIDE)
        .ok_or_else(|| io::Error::other("missing synthetic slide"))?;
    let message = slide
        .messages
        .iter_mut()
        .find(|message| message.type_ == 5)
        .ok_or_else(|| io::Error::other("missing synthetic slide message"))?;
    let mut value = kn::SlideArchive::decode(message.data.as_slice())?;
    value.owned_drawables.clear();
    value.drawables_z_order.clear();
    message.data = value.encode_to_vec();
    chart_fixture::replace_document_stream(source, archive)
}

fn noncanonical_bool_field(number: u32) -> Vec<u8> {
    let mut encoded = chart_fixture::varint_field(number, 0);
    let key_width = encoded.len() - 1;
    encoded.truncate(key_width);
    encoded.extend([0x80, 0x00]);
    encoded
}

fn malformed_arrangement_source(kind: &str) -> TestResult<Vec<u8>> {
    let source = synthetic_package_with_states([(None, None), (None, None)], false, false)?;
    match kind {
        "duplicate-locked" => with_drawable_wire(&source, 0, |drawable| {
            let mut output = drawable.to_vec();
            append_varint_field(&mut output, DRAWABLE_LOCKED_FIELD, 1)?;
            append_varint_field(&mut output, DRAWABLE_LOCKED_FIELD, 0)?;
            Ok(output)
        }),
        "duplicate-aspect-ratio" => with_drawable_wire(&source, 0, |drawable| {
            let mut output = drawable.to_vec();
            append_varint_field(&mut output, DRAWABLE_ASPECT_RATIO_LOCKED_FIELD, 1)?;
            append_varint_field(&mut output, DRAWABLE_ASPECT_RATIO_LOCKED_FIELD, 0)?;
            Ok(output)
        }),
        "wrong-locked-wire" => with_drawable_wire(&source, 0, |drawable| {
            let mut output = drawable.to_vec();
            append_length_delimited_field(&mut output, DRAWABLE_LOCKED_FIELD, b"not a bool")?;
            Ok(output)
        }),
        "wrong-aspect-ratio-wire" => with_drawable_wire(&source, 0, |drawable| {
            let mut output = drawable.to_vec();
            append_length_delimited_field(
                &mut output,
                DRAWABLE_ASPECT_RATIO_LOCKED_FIELD,
                b"not a bool",
            )?;
            Ok(output)
        }),
        "noncanonical-locked" => with_drawable_wire(&source, 0, |drawable| {
            let mut output = drawable.to_vec();
            output.extend(noncanonical_bool_field(DRAWABLE_LOCKED_FIELD));
            Ok(output)
        }),
        "noncanonical-aspect-ratio" => with_drawable_wire(&source, 0, |drawable| {
            let mut output = drawable.to_vec();
            output.extend(noncanonical_bool_field(DRAWABLE_ASPECT_RATIO_LOCKED_FIELD));
            Ok(output)
        }),
        "invalid-locked" => with_drawable_wire(&source, 0, |drawable| {
            let mut output = drawable.to_vec();
            append_varint_field(&mut output, DRAWABLE_LOCKED_FIELD, 2)?;
            Ok(output)
        }),
        "invalid-aspect-ratio" => with_drawable_wire(&source, 0, |drawable| {
            let mut output = drawable.to_vec();
            append_varint_field(&mut output, DRAWABLE_ASPECT_RATIO_LOCKED_FIELD, 2)?;
            Ok(output)
        }),
        "wrong-outer-wire" => {
            let original = chart_fixture::message_payload(&source, CHARTS[0], CHART_MESSAGE_TYPE)?;
            let replacement = chart_fixture::varint_field(DRAWABLE_SUPER_FIELD, 1);
            let payload =
                chart_fixture::replace_first_field(&original, DRAWABLE_SUPER_FIELD, &replacement)?;
            chart_fixture::with_chart_payload(&source, 0, payload)
        },
        "duplicate-outer" => {
            let original = chart_fixture::message_payload(&source, CHARTS[0], CHART_MESSAGE_TYPE)?;
            let field = WireView::parse(&original)?
                .fields()
                .find(|field| field.number() == DRAWABLE_SUPER_FIELD)
                .ok_or_else(|| io::Error::other("missing chart drawable"))?;
            let mut payload = original.clone();
            payload.extend_from_slice(field.raw());
            chart_fixture::with_chart_payload(&source, 0, payload)
        },
        other => {
            Err(io::Error::other(format!("unknown malformed arrangement case: {other}")).into())
        },
    }
}

fn with_shared_non_style(source: &[u8]) -> TestResult<Vec<u8>> {
    let original = chart_fixture::message_payload(source, CHARTS[1], CHART_MESSAGE_TYPE)?;
    let extension = WireView::parse(&original)?
        .fields()
        .find(|field| field.number() == GENERATED_CHART_EXTENSION_FIELD)
        .ok_or_else(|| io::Error::other("missing chart extension"))?;
    let mut chart = tsch::ChartArchive::decode(extension.payload())?;
    chart.chart_non_style = Some(reference(CHART_NON_STYLES[0]));
    let replacement = chart_fixture::length_delimited_field_with_key_width(
        GENERATED_CHART_EXTENSION_FIELD,
        &chart.encode_to_vec(),
        extension.key().len(),
    );
    let payload = chart_fixture::replace_first_field(
        &original,
        GENERATED_CHART_EXTENSION_FIELD,
        &replacement,
    )?;
    chart_fixture::with_chart_payload(source, 1, payload)
}

fn with_shared_title(source: &[u8]) -> TestResult<Vec<u8>> {
    let original = chart_fixture::message_payload(source, CHARTS[1], CHART_MESSAGE_TYPE)?;
    let drawable = WireView::parse(&original)?
        .fields()
        .find(|field| field.number() == DRAWABLE_SUPER_FIELD)
        .ok_or_else(|| io::Error::other("missing chart drawable"))?;
    let mut value = tsd::DrawableArchive::decode(drawable.payload())?;
    value.title = Some(reference(TITLES[0]));
    let replacement = chart_fixture::length_delimited_field_with_key_width(
        DRAWABLE_SUPER_FIELD,
        &value.encode_to_vec(),
        drawable.key().len(),
    );
    let payload =
        chart_fixture::replace_first_field(&original, DRAWABLE_SUPER_FIELD, &replacement)?;
    chart_fixture::with_chart_payload(source, 1, payload)
}

fn with_parent_mismatch(source: &[u8]) -> TestResult<Vec<u8>> {
    let original = chart_fixture::message_payload(source, CHARTS[0], CHART_MESSAGE_TYPE)?;
    let drawable = WireView::parse(&original)?
        .fields()
        .find(|field| field.number() == DRAWABLE_SUPER_FIELD)
        .ok_or_else(|| io::Error::other("missing chart drawable"))?;
    let mut value = tsd::DrawableArchive::decode(drawable.payload())?;
    value.parent = Some(reference(999));
    let replacement = chart_fixture::length_delimited_field_with_key_width(
        DRAWABLE_SUPER_FIELD,
        &value.encode_to_vec(),
        drawable.key().len(),
    );
    let payload =
        chart_fixture::replace_first_field(&original, DRAWABLE_SUPER_FIELD, &replacement)?;
    chart_fixture::with_chart_payload(source, 0, payload)
}

fn with_duplicate_chart_message(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut archive = Archive::parse(&chart_fixture::document_stream(source)?)?;
    let chart = archive
        .object_mut(CHARTS[0])
        .ok_or_else(|| io::Error::other("missing synthetic chart"))?;
    let data = chart
        .messages
        .iter()
        .find(|message| message.type_ == CHART_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("missing synthetic chart message"))?
        .data
        .clone();
    chart.push_message(RawMessage {
        type_: CHART_MESSAGE_TYPE,
        data,
    })?;
    chart_fixture::replace_document_stream(source, archive)
}

fn with_foreign_chart_owner(source: &[u8]) -> TestResult<Vec<u8>> {
    let chart_payload = chart_fixture::message_payload(source, CHARTS[0], CHART_MESSAGE_TYPE)?;
    let mut archive = Archive::parse(&chart_fixture::component_stream(source, FOREIGN_MEMBER)?)?;
    archive.objects.push(litchi_iwa_core::ArchiveObject::new(
        901,
        vec![RawMessage {
            type_: CHART_MESSAGE_TYPE,
            data: chart_payload,
        }],
    )?);
    chart_fixture::replace_component_stream(source, FOREIGN_MEMBER, archive)
}

fn assert_unknowns_unchanged(before: &[u8], after: &[u8]) -> TestResult<()> {
    assert_eq!(
        raw_drawable_fields(before, 0, UNKNOWN_DRAWABLE_FIELD)?,
        raw_drawable_fields(after, 0, UNKNOWN_DRAWABLE_FIELD)?,
    );
    assert_eq!(
        raw_drawable_fields(before, 0, UNKNOWN_DRAWABLE_BYTES_FIELD)?,
        raw_drawable_fields(after, 0, UNKNOWN_DRAWABLE_BYTES_FIELD)?,
    );
    assert_eq!(
        chart_fixture::message_payload(before, CHART_NON_STYLES[0], CHART_NON_STYLE_MESSAGE_TYPE,)?,
        chart_fixture::message_payload(after, CHART_NON_STYLES[0], CHART_NON_STYLE_MESSAGE_TYPE,)?,
        "unknown title group/non-style bytes changed during Arrange",
    );
    Ok(())
}

fn assert_semantic_handles_are_send_sync<T: Send + Sync>() {}

#[test]
fn arrangement_batch_preserves_z_order_for_all_boolean_combinations() -> TestResult<()> {
    let combinations = [(false, false), (false, true), (true, false), (true, true)];
    for first in combinations {
        for second in combinations {
            let source = synthetic_package_with_states(
                [
                    (Some(first.0), Some(first.1)),
                    (Some(second.0), Some(second.1)),
                ],
                false,
                false,
            )?;
            let package = Package::from_bytes(&source)?;
            let batch = package.slide_chart_arrangements(0usize)?;
            let expected = [
                ChartArrangement::new(first.0, first.1),
                ChartArrangement::new(second.0, second.1),
            ];
            assert_eq!(batch.as_ref(), expected.as_slice());
            assert_eq!(batch.len(), CHARTS.len());
        }
    }
    Ok(())
}

#[test]
fn arrangement_batch_slide_name_and_index_are_equivalent_and_match_individual_reads()
-> TestResult<()> {
    let source =
        synthetic_package_with_states([(None, None), (Some(true), Some(false))], false, false)?;
    let package = Package::from_bytes(&source)?;
    let by_index = package.slide_chart_arrangements(0usize)?;
    let by_name = package.slide_chart_arrangements(SlideSelector::name("Charts"))?;
    let by_str = package.slide_chart_arrangements("Charts")?;
    assert_eq!(by_index.as_ref(), by_name.as_ref());
    assert_eq!(by_index.as_ref(), by_str.as_ref());

    for (chart, value) in by_index.iter().enumerate() {
        assert_eq!(
            *value,
            package.slide_chart_arrangement(0usize, chart)?,
            "batch value diverged from index-selected chart {chart}",
        );
        let name = if chart == 0 { "Revenue" } else { "Costs" };
        assert_eq!(
            *value,
            package.slide_chart_arrangement("Charts", name)?,
            "batch value diverged from name-selected chart {name}",
        );
    }
    Ok(())
}

#[test]
fn arrangement_batch_empty_slide_returns_empty_without_mutation() -> TestResult<()> {
    let source = empty_slide_source(&base_source()?)?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    let by_index = package.slide_chart_arrangements(0usize)?;
    let by_name = package.slide_chart_arrangements("Charts")?;
    assert!(by_index.is_empty());
    assert!(by_name.is_empty());
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn arrangement_batch_malformed_second_chart_fails_atomically_and_redacted() -> TestResult<()> {
    let source =
        synthetic_package_with_states([(Some(false), Some(false)), (None, None)], false, false)?;
    let malformed = with_drawable_wire(&source, 1, |drawable| {
        let mut output = drawable.to_vec();
        append_varint_field(&mut output, DRAWABLE_LOCKED_FIELD, 1)?;
        append_varint_field(&mut output, DRAWABLE_LOCKED_FIELD, 0)?;
        Ok(output)
    })?;
    let package = Package::from_bytes(&malformed)?;
    let before = exact_bytes(&package)?;
    let error = package
        .slide_chart_arrangements(0usize)
        .expect_err("a malformed later chart must invalidate the whole batch");
    assert!(matches!(error, ChartArrangementError::InvalidSource));
    error_is_redacted(&error);
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn arrangement_batch_semantic_reference_limit_refuses_without_mutation() -> TestResult<()> {
    let source = base_source()?;
    let semantic = SemanticLimits::new(
        SemanticLimits::MAX_OBJECTS,
        1,
        1,
        SemanticLimits::MAX_TEXT_STORAGES,
        SemanticLimits::MAX_TEXT_FRAGMENTS,
        SemanticLimits::MAX_TEXT_BYTES,
    )?;
    let package =
        Package::from_bytes_with_options(&source, ReadOptions::new(Limits::default(), semantic))?;
    let before = exact_bytes(&package)?;
    let error = package
        .slide_chart_arrangements(0usize)
        .expect_err("the batch traversal must honor the semantic reference ceiling");
    assert!(matches!(
        error,
        ChartArrangementError::LimitExceeded {
            kind: ChartArrangementLimitKind::References,
            ..
        }
    ));
    error_is_redacted(&error);
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn arrangement_defaults_and_selectors_cover_all_boolean_combinations() -> TestResult<()> {
    for (locked, constrain_proportions) in
        [(false, false), (false, true), (true, false), (true, true)]
    {
        let source = synthetic_package_with_states(
            [(Some(locked), Some(constrain_proportions)), (None, None)],
            false,
            false,
        )?;
        let package = Package::from_bytes(&source)?;
        let expected = ChartArrangement::new(locked, constrain_proportions);
        assert_eq!(arrangement(&package, 0usize, 0usize)?, expected);
        assert_eq!(
            arrangement(
                &package,
                SlideSelector::name("Charts"),
                ChartSelector::name("Revenue")
            )?,
            expected
        );
        assert_eq!(
            arrangement(&package, "Charts", "Costs")?,
            ChartArrangement::default()
        );
    }

    // An absent field and an explicit false field have the same semantic
    // value, but both are accepted as exact no-op sources.
    let absent = Package::from_bytes(&synthetic_package_with_states(
        [(None, None), (None, None)],
        false,
        false,
    )?)?;
    assert_eq!(
        absent.slide_chart_arrangement(0usize, 0usize)?,
        ChartArrangement::default()
    );
    let explicit_false = Package::from_bytes(&synthetic_package_with_states(
        [(Some(false), Some(false)), (None, None)],
        false,
        false,
    )?)?;
    assert_eq!(
        explicit_false.slide_chart_arrangement("Charts", "Revenue")?,
        ChartArrangement::default()
    );
    Ok(())
}

#[test]
fn arrangement_edit_exposes_only_semantic_positions_and_values() -> TestResult<()> {
    let package = Package::from_bytes(&base_source()?)?;
    let edit = package.edit_slide_chart_arrangement("Charts", "Revenue")?;
    assert_eq!(edit.slide_position(), Position::new(0));
    assert_eq!(edit.chart_position(), Position::new(0));
    assert_eq!(edit.before(), ChartArrangement::default());
    assert_eq!(edit.after(), ChartArrangement::default());

    let debug = format!("{edit:?}");
    for identifier in CHARTS.into_iter().chain(TITLES).chain(CHART_NON_STYLES) {
        assert!(
            !debug.contains(&identifier.to_string()),
            "native identifier leaked from semantic arrangement edit debug output: {debug}"
        );
    }
    Ok(())
}

#[test]
fn arrangement_noop_is_byte_exact_and_reapplicable() -> TestResult<()> {
    let source =
        synthetic_package_with_states([(None, None), (Some(false), Some(false))], true, true)?;
    let package = Package::from_bytes(&source)?;
    let noop = package
        .edit_slide_chart_arrangement("Charts", "Revenue")?
        .set(ChartArrangement::default())
        .commit()?;

    assert!(noop.patch().is_noop());
    assert_eq!(noop.patch().before(), ChartArrangement::default());
    assert_eq!(noop.patch().after(), ChartArrangement::default());
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert!(!noop.diagnostics().full_reparse_performed());
    assert_eq!(exact_bytes(noop.package())?, source);
    assert_previews_preserved(&source, &exact_bytes(noop.package())?)?;

    let reapplied = package.apply_slide_chart_arrangement(noop.patch())?;
    assert!(reapplied.patch().is_noop());
    assert!(!reapplied.diagnostics().changed());
    assert_eq!(exact_bytes(reapplied.package())?, source);
    Ok(())
}

#[test]
fn arrangement_changed_commit_reopens_and_inverse_restores_exact_source() -> TestResult<()> {
    let source = synthetic_package_with_states(
        [(Some(false), Some(false)), (Some(true), Some(false))],
        true,
        true,
    )?;
    let package = Package::from_bytes(&source)?;
    let target = ChartArrangement::new(true, true);
    let commit = commit_arrangement(&package, "Charts", "Revenue", target)?;

    assert_eq!(commit.patch().before(), ChartArrangement::default());
    assert_eq!(commit.patch().after(), target);
    assert!(!commit.patch().is_noop());
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert!(commit.diagnostics().full_reparse_performed());
    assert_eq!(
        commit.package().slide_chart_arrangement(0usize, 0usize)?,
        target
    );
    assert_eq!(exact_bytes(&package)?, source);

    let target_bytes = exact_bytes(commit.package())?;
    assert_document_only_changed(&source, &target_bytes)?;
    assert_previews_preserved(&source, &target_bytes)?;
    assert_unknowns_unchanged(&source, &target_bytes)?;
    assert_eq!(
        chart_fixture::message_payload(&source, CHARTS[1], CHART_MESSAGE_TYPE)?,
        chart_fixture::message_payload(&target_bytes, CHARTS[1], CHART_MESSAGE_TYPE)?,
        "an Arrange edit must remain local to the selected chart",
    );

    let reopened = Package::from_bytes(&target_bytes)?;
    assert_eq!(arrangement(&reopened, "Charts", "Revenue")?, target);
    let applied = package.apply_slide_chart_arrangement(commit.patch())?;
    assert_eq!(exact_bytes(applied.package())?, target_bytes);

    let inverse = commit.patch().inverse();
    assert_eq!(inverse.before(), target);
    assert_eq!(inverse.after(), ChartArrangement::default());
    assert_eq!(inverse.inverse(), commit.patch().clone());
    let restored = reopened.apply_slide_chart_arrangement(&inverse)?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_eq!(
        restored.package().slide_chart_arrangement(0usize, 0usize)?,
        ChartArrangement::default()
    );
    Ok(())
}

#[test]
fn arrangement_changed_bool_combinations_preserve_presence_and_inverse() -> TestResult<()> {
    for (locked, constrain_proportions) in [(false, true), (true, false), (true, true)] {
        let source = synthetic_package_with_states(
            [(Some(false), Some(false)), (None, None)],
            false,
            false,
        )?;
        let package = Package::from_bytes(&source)?;
        let target = ChartArrangement::new(locked, constrain_proportions);
        let commit = commit_arrangement(&package, 0usize, 0usize, target)?;
        assert_eq!(arrangement(commit.package(), 0usize, 0usize)?, target);
        let restored = commit
            .package()
            .apply_slide_chart_arrangement(&commit.patch().inverse())?;
        assert_eq!(exact_bytes(restored.package())?, source);
        assert_eq!(
            arrangement(restored.package(), 0usize, 0usize)?,
            ChartArrangement::default()
        );
    }

    // Setting semantic false on an absent field must not synthesize an
    // explicit wire scalar; the candidate remains a physical no-op.
    let source = synthetic_package_with_states([(None, None), (None, None)], false, false)?;
    let package = Package::from_bytes(&source)?;
    let commit = commit_arrangement(&package, 0usize, 0usize, ChartArrangement::default())?;
    assert!(commit.patch().is_noop());
    assert_eq!(exact_bytes(commit.package())?, source);
    assert!(raw_drawable_fields(&source, 0, DRAWABLE_LOCKED_FIELD)?.is_empty());
    assert!(raw_drawable_fields(&source, 0, DRAWABLE_ASPECT_RATIO_LOCKED_FIELD)?.is_empty());
    Ok(())
}

#[test]
fn arrangement_changed_commit_preserves_absent_flags_and_opaque_selected_payload_on_inverse()
-> TestResult<()> {
    // The semantic value of an omitted flag is false, but its physical
    // absence is part of the source contract.  Exercise every source shape
    // that still contains at least one omitted flag while the selected
    // drawable also carries opaque wire data outside the two edited fields.
    for state in [(None, None), (Some(false), None), (None, Some(false))] {
        let source = synthetic_package_with_states([state, (Some(true), Some(false))], true, true)?;
        assert_eq!(
            raw_drawable_fields(&source, 0, DRAWABLE_LOCKED_FIELD)?.is_empty(),
            state.0.is_none(),
        );
        assert_eq!(
            raw_drawable_fields(&source, 0, DRAWABLE_ASPECT_RATIO_LOCKED_FIELD)?.is_empty(),
            state.1.is_none(),
        );

        let package = Package::from_bytes(&source)?;
        let target = ChartArrangement::new(true, true);
        let commit = commit_arrangement(&package, "Charts", "Revenue", target)?;
        let candidate = exact_bytes(commit.package())?;

        assert_eq!(arrangement(commit.package(), "Charts", "Revenue")?, target);
        assert_unknowns_unchanged(&source, &candidate)?;

        let applied = package.apply_slide_chart_arrangement(commit.patch())?;
        assert_eq!(exact_bytes(applied.package())?, candidate);

        let reopened = Package::from_bytes(&candidate)?;
        let restored = reopened.apply_slide_chart_arrangement(&commit.patch().inverse())?;
        let restored_bytes = exact_bytes(restored.package())?;
        assert_eq!(restored_bytes, source);
        assert_eq!(
            raw_drawable_fields(&restored_bytes, 0, DRAWABLE_LOCKED_FIELD)?,
            raw_drawable_fields(&source, 0, DRAWABLE_LOCKED_FIELD)?,
        );
        assert_eq!(
            raw_drawable_fields(&restored_bytes, 0, DRAWABLE_ASPECT_RATIO_LOCKED_FIELD)?,
            raw_drawable_fields(&source, 0, DRAWABLE_ASPECT_RATIO_LOCKED_FIELD)?,
        );
    }
    Ok(())
}

#[test]
fn arrangement_unknown_fields_and_groups_remain_byte_exact_and_local() -> TestResult<()> {
    let source = synthetic_package_with_states(
        [(Some(false), Some(false)), (Some(true), Some(false))],
        true,
        true,
    )?;
    let package = Package::from_bytes(&source)?;
    let selected_before = chart_fixture::message_payload(&source, CHARTS[0], CHART_MESSAGE_TYPE)?;
    let other_before = chart_fixture::message_payload(&source, CHARTS[1], CHART_MESSAGE_TYPE)?;
    let title_before =
        chart_fixture::message_payload(&source, CHART_NON_STYLES[0], CHART_NON_STYLE_MESSAGE_TYPE)?;
    let commit = commit_arrangement(
        &package,
        SlideSelector::name("Charts"),
        ChartSelector::name("Revenue"),
        ChartArrangement::new(true, true),
    )?;
    let target = exact_bytes(commit.package())?;
    let selected_after = chart_fixture::message_payload(&target, CHARTS[0], CHART_MESSAGE_TYPE)?;
    let title_after =
        chart_fixture::message_payload(&target, CHART_NON_STYLES[0], CHART_NON_STYLE_MESSAGE_TYPE)?;

    assert_ne!(selected_before, selected_after);
    assert_eq!(
        other_before,
        chart_fixture::message_payload(&target, CHARTS[1], CHART_MESSAGE_TYPE)?
    );
    assert_eq!(
        title_before, title_after,
        "unknown generated group was not retained"
    );
    assert_unknowns_unchanged(&source, &target)?;
    assert_document_only_changed(&source, &target)?;
    assert_previews_preserved(&source, &target)?;
    Ok(())
}

#[test]
fn arrangement_stale_and_foreign_patches_conflict_without_mutation() -> TestResult<()> {
    let source = base_source()?;
    let package = Package::from_bytes(&source)?;
    let commit = commit_arrangement(&package, 0usize, 0usize, ChartArrangement::new(true, false))?;

    let target = Package::from_bytes(&exact_bytes(commit.package())?)?;
    let target_before = exact_bytes(&target)?;
    let stale_error = target
        .apply_slide_chart_arrangement(commit.patch())
        .expect_err("a patch must not apply to its already-changed target");
    assert!(matches!(stale_error, ChartArrangementError::PatchConflict));
    error_is_redacted(&stale_error);
    assert_eq!(exact_bytes(&target)?, target_before);

    let foreign_source = chart_fixture::synthetic_metadata_package(true)?;
    let foreign = Package::from_bytes(&foreign_source)?;
    let foreign_before = exact_bytes(&foreign)?;
    let foreign_error = foreign
        .apply_slide_chart_arrangement(commit.patch())
        .expect_err("a patch must not cross exact package artifacts");
    assert!(matches!(
        foreign_error,
        ChartArrangementError::PatchConflict
    ));
    error_is_redacted(&foreign_error);
    assert_eq!(exact_bytes(&foreign)?, foreign_before);
    Ok(())
}

#[test]
fn arrangement_selectors_fail_closed_for_missing_names_positions_and_ambiguity() -> TestResult<()> {
    let source = base_source()?;
    let package = Package::from_bytes(&source)?;
    assert!(matches!(
        package.slide_chart_arrangement("missing", 0usize),
        Err(ChartArrangementError::SlideNameNotFound)
    ));
    assert!(matches!(
        package.slide_chart_arrangement(SlideSelector::name(""), 0usize),
        Err(ChartArrangementError::EmptySlideName)
    ));
    assert!(matches!(
        package.slide_chart_arrangement(0usize, "missing"),
        Err(ChartArrangementError::ChartNameNotFound)
    ));
    assert!(matches!(
        package.slide_chart_arrangement(0usize, ChartSelector::name("")),
        Err(ChartArrangementError::EmptyChartName)
    ));
    assert!(matches!(
        package.slide_chart_arrangement(Position::new(9), 0usize),
        Err(ChartArrangementError::SlidePositionNotFound { position })
            if position == Position::new(9)
    ));
    assert!(matches!(
        package.slide_chart_arrangement(0usize, Position::new(9)),
        Err(ChartArrangementError::ChartPositionNotFound { position })
            if position == Position::new(9)
    ));

    let duplicate_source = chart_fixture::with_document_message_payload(
        &source,
        CHART_NON_STYLES[1],
        CHART_NON_STYLE_MESSAGE_TYPE,
        chart_fixture::chart_non_style_payload("Revenue")?,
    )?;
    let duplicate = Package::from_bytes(&duplicate_source)?;
    let before = exact_bytes(&duplicate)?;
    let error = duplicate
        .edit_slide_chart_arrangement(0usize, ChartSelector::name("Revenue"))
        .expect_err("duplicate chart names must be ambiguous");
    assert!(matches!(error, ChartArrangementError::AmbiguousSelector));
    error_is_redacted(&error);
    assert_eq!(exact_bytes(&duplicate)?, before);
    Ok(())
}

#[test]
fn arrangement_wire_validation_rejects_duplicate_wrong_noncanonical_and_invalid_bools()
-> TestResult<()> {
    for kind in [
        "duplicate-locked",
        "duplicate-aspect-ratio",
        "wrong-locked-wire",
        "wrong-aspect-ratio-wire",
        "noncanonical-locked",
        "noncanonical-aspect-ratio",
        "invalid-locked",
        "invalid-aspect-ratio",
        "wrong-outer-wire",
        "duplicate-outer",
    ] {
        let malformed = malformed_arrangement_source(kind)?;
        let package = Package::from_bytes(&malformed)?;
        let before = exact_bytes(&package)?;
        let read = package.slide_chart_arrangement(0usize, 0usize);
        let error = read.expect_err(kind);
        assert!(matches!(error, ChartArrangementError::InvalidSource));
        error_is_redacted(&error);
        assert!(
            package
                .edit_slide_chart_arrangement(0usize, 0usize)
                .is_err(),
            "malformed {kind} source unexpectedly opened for mutation"
        );
        assert_eq!(exact_bytes(&package)?, before);
    }
    Ok(())
}

#[test]
fn arrangement_graph_ambiguity_and_ownership_guards_are_atomic() -> TestResult<()> {
    let source = base_source()?;
    for hostile in [
        with_shared_non_style(&source)?,
        with_shared_title(&source)?,
        with_parent_mismatch(&source)?,
        with_duplicate_chart_message(&source)?,
    ] {
        let package = Package::from_bytes(&hostile)?;
        let before = exact_bytes(&package)?;
        let error = package
            .edit_slide_chart_arrangement(0usize, 0usize)
            .expect_err("ambiguous or malformed chart graph must fail closed");
        assert!(matches!(error, ChartArrangementError::InvalidSource));
        error_is_redacted(&error);
        assert_eq!(exact_bytes(&package)?, before);
    }

    let foreign = with_foreign_chart_owner(&chart_fixture::synthetic_metadata_package(true)?)?;
    let package = Package::from_bytes(&foreign)?;
    let before = exact_bytes(&package)?;
    let error = package
        .edit_slide_chart_arrangement(0usize, 0usize)
        .expect_err("foreign chart ownership must fail closed");
    assert!(matches!(error, ChartArrangementError::InvalidSource));
    error_is_redacted(&error);
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn arrangement_semantic_limits_fail_before_publication_and_output_limits_are_atomic()
-> TestResult<()> {
    let unpadded = base_source()?;
    let catalog = Catalog::from_bytes(&unpadded)?;
    // Give graph verification headroom so this case reaches the independent
    // output ceiling instead of an earlier aggregate retained-byte refusal.
    let padding = [0u8; 16 * 1024];
    let entries: Vec<_> = catalog
        .iter()
        .map(|entry| (entry.name(), entry.data()))
        .chain(std::iter::once((
            "Data/output-limit-padding.bin",
            padding.as_slice(),
        )))
        .collect();
    let source = litchi_iwa_archive::package::to_bytes(entries.iter().copied(), Limits::default())?;
    let semantic = SemanticLimits::new(1, 1, 1, 1, 1, 1)?;
    let limited =
        Package::from_bytes_with_options(&source, ReadOptions::new(Limits::default(), semantic));
    let ingress_error = limited.expect_err("semantic object limit must reject the chart graph");
    error_is_redacted(&ingress_error);
    assert!(format!("{ingress_error:?}").len() < 512);

    let package = Package::from_bytes(&source)?;
    let unrestricted =
        commit_arrangement(&package, 0usize, 0usize, ChartArrangement::new(true, true))?;
    let candidate = exact_bytes(unrestricted.package())?;
    assert!(candidate.len() > source.len());
    let defaults = Limits::default();
    let limits = Limits::new(
        u64::try_from(candidate.len() - 1)?,
        defaults.max_entries(),
        defaults.max_entry_bytes(),
        defaults.max_total_bytes(),
        defaults.max_iwa_stream_bytes(),
    )?;
    let bounded = Package::from_bytes_with_options(
        &source,
        ReadOptions::new(limits, SemanticLimits::default()),
    )?;
    let before = exact_bytes(&bounded)?;
    let error = bounded
        .edit_slide_chart_arrangement(0usize, 0usize)
        .and_then(|edit| edit.set(ChartArrangement::new(true, true)).commit())
        .expect_err("arrangement candidate should exceed the output ceiling");
    assert!(
        matches!(
            error,
            ChartArrangementError::LimitExceeded {
                kind: ChartArrangementLimitKind::OutputBytes,
                ..
            }
        ),
        "unexpected resource refusal: {error:?}"
    );
    error_is_redacted(&error);
    assert_eq!(exact_bytes(&bounded)?, before);
    Ok(())
}

#[test]
fn arrangement_public_types_are_send_sync_and_debug_is_redacted() -> TestResult<()> {
    assert_semantic_handles_are_send_sync::<ChartArrangement>();
    assert_semantic_handles_are_send_sync::<ChartArrangementEdit<'static>>();
    assert_semantic_handles_are_send_sync::<ChartArrangementPatch>();
    assert_semantic_handles_are_send_sync::<ChartArrangementDiagnostics>();
    assert_semantic_handles_are_send_sync::<ChartArrangementCommit>();
    assert_semantic_handles_are_send_sync::<ChartArrangementError>();
    assert_semantic_handles_are_send_sync::<ChartArrangementLimitKind>();

    let package = Package::from_bytes(&base_source()?)?;
    let edit = package.edit_slide_chart_arrangement(
        SlideSelector::name("Charts"),
        ChartSelector::name("Revenue"),
    )?;
    let edit_debug = format!("{edit:?}");
    for identifier in CHARTS.into_iter().chain(TITLES).chain(CHART_NON_STYLES) {
        assert!(!edit_debug.contains(&identifier.to_string()));
    }
    let commit = edit.set(ChartArrangement::new(true, false)).commit()?;
    let patch_debug = format!("{:?}", commit.patch());
    let diagnostics_debug = format!("{:?}", commit.diagnostics());
    for identifier in CHARTS.into_iter().chain(TITLES).chain(CHART_NON_STYLES) {
        assert!(!patch_debug.contains(&identifier.to_string()));
        assert!(!diagnostics_debug.contains(&identifier.to_string()));
    }

    let error = package
        .slide_chart_arrangement("missing", 0usize)
        .expect_err("missing selector should produce a redacted error");
    error_is_redacted(&error);
    Ok(())
}
