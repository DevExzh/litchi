//! Strict, source-preserving Keynote chart-legend visibility coverage.
//!
//! The fixture is intentionally small, but it contains the same semantic
//! ownership graph as a native Keynote chart: a slide owns two chart
//! drawables, each chart owns a chart non-style object, and the chart title
//! extension supplies stable human-readable selectors.  All assertions use
//! positions and names.  The native object identifiers below exist only to
//! construct and inspect this test fixture; they are never part of the
//! public API under test.

use std::io;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::wire::{append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsch, tsd, tsk, tsp};
use litchi_keynote::{
    ChartLegendVisibilityDiagnostics, ChartLegendVisibilityError, ChartLegendVisibilityLimitKind,
    ChartLegendVisibilityPatch, ChartSelector, Package, Position, SlideSelector,
};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const UNRELATED_MEMBER: &str = "Index/Unrelated.iwa";
const ROOT_PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];
const SLIDE_NODE: u64 = 3;
const SLIDE: u64 = 4;
const CHARTS: [u64; 2] = [100, 101];
const TITLES: [u64; 2] = [110, 111];
const NON_STYLES: [u64; 2] = [120, 121];
const CHART_MESSAGE_TYPE: u32 = 5_021;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
const CHART_NON_STYLE_MESSAGE_TYPE: u32 = 5_023;
const CHART_EXTENSION_FIELD: u32 = 10_000;
const LEGEND_VISIBLE_FIELD: u32 = 20;

type TestResult<T> = Result<T, Box<dyn std::error::Error>>;

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn object(identifier: u64, type_: u32, data: Vec<u8>) -> TestResult<ArchiveObject> {
    object_with_messages(identifier, vec![RawMessage { type_, data }])
}

fn object_with_messages(identifier: u64, messages: Vec<RawMessage>) -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(identifier, messages)?)
}

fn chart_payload(chart: usize) -> TestResult<Vec<u8>> {
    chart_payload_with_graph(TITLES[chart], NON_STYLES[chart])
}

fn chart_payload_with_non_style(chart: usize, non_style_identifier: u64) -> TestResult<Vec<u8>> {
    chart_payload_with_graph(TITLES[chart], non_style_identifier)
}

fn chart_payload_with_graph(
    title_identifier: u64,
    non_style_identifier: u64,
) -> TestResult<Vec<u8>> {
    let drawable = tsd::DrawableArchive {
        parent: Some(reference(SLIDE)),
        title: Some(reference(title_identifier)),
        ..tsd::DrawableArchive::default()
    };
    let chart_data = tsch::ChartArchive {
        chart_non_style: Some(reference(non_style_identifier)),
        ..tsch::ChartArchive::default()
    };
    let mut payload = Vec::new();
    append_length_delimited_field(&mut payload, 1, &drawable.encode_to_vec())?;
    append_length_delimited_field(
        &mut payload,
        CHART_EXTENSION_FIELD,
        &chart_data.encode_to_vec(),
    )?;
    Ok(payload)
}

/// Build the generated chart non-style extension used by the lazy legend
/// projection.  `None` means that the native field is absent; the semantic
/// API must still report the documented default (`false`).
fn non_style_payload(
    visible: Option<bool>,
    title: Option<&str>,
    with_unknown_fields: bool,
) -> TestResult<Vec<u8>> {
    let extension = non_style_extension(visible, title, with_unknown_fields)?;
    let mut payload = Vec::new();
    append_length_delimited_field(&mut payload, CHART_EXTENSION_FIELD, &extension)?;
    Ok(payload)
}

fn non_style_extension(
    visible: Option<bool>,
    title: Option<&str>,
    with_unknown_fields: bool,
) -> TestResult<Vec<u8>> {
    let mut extension = Vec::new();
    if with_unknown_fields {
        append_varint_field(&mut extension, 4_000, 7)?;
    }
    if let Some(visible) = visible {
        append_varint_field(&mut extension, LEGEND_VISIBLE_FIELD, u64::from(visible))?;
    }
    if title.is_some() {
        // ChartSelector::Name resolves through the chart-title projection;
        // keep the generated title-presence bit alongside its text.
        append_varint_field(&mut extension, 21, 1)?;
    }
    if with_unknown_fields {
        append_length_delimited_field(&mut extension, 4_001, b"opaque legend metadata")?;
    }
    if let Some(title) = title {
        append_length_delimited_field(&mut extension, 23, title.as_bytes())?;
    }
    if with_unknown_fields {
        append_varint_field(&mut extension, 4_002, 9)?;
    }
    Ok(extension)
}

fn component(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

fn synthetic_package() -> TestResult<Vec<u8>> {
    synthetic_package_with_states(
        [(None, Some("Revenue")), (Some(true), Some("Costs"))],
        false,
    )
}

fn synthetic_package_with_states(
    states: [(Option<bool>, Option<&str>); 2],
    with_unknown_fields: bool,
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
        object(SLIDE, 5, slide.encode_to_vec())?,
    ];
    for chart in 0..CHARTS.len() {
        objects.push(object(
            CHARTS[chart],
            CHART_MESSAGE_TYPE,
            chart_payload(chart)?,
        )?);
        objects.push(object(TITLES[chart], STANDIN_MESSAGE_TYPE, Vec::new())?);
        objects.push(object_with_messages(
            NON_STYLES[chart],
            vec![
                RawMessage {
                    type_: CHART_NON_STYLE_MESSAGE_TYPE,
                    data: non_style_payload(states[chart].0, states[chart].1, with_unknown_fields)?,
                },
                RawMessage {
                    type_: 7_777,
                    data: b"unselected message sentinel".to_vec(),
                },
            ],
        )?);
    }
    let document_component = component(objects)?;
    let unrelated_component = component(vec![object(
        900,
        7_778,
        b"unselected component sentinel".to_vec(),
    )?])?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Data/sentinel.bin", b"unrelated ZIP sentinel".as_slice()),
            ("preview.jpg", b"large root preview".as_slice()),
            ("preview-micro.jpg", b"micro root preview".as_slice()),
            ("preview-web.jpg", b"web root preview".as_slice()),
            (UNRELATED_MEMBER, unrelated_component.as_slice()),
            (DOCUMENT_MEMBER, document_component.as_slice()),
        ],
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

fn message_payload(package: &[u8], identifier: u64, type_: u32) -> TestResult<Vec<u8>> {
    let archive = Archive::parse(&document_stream(package)?)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing synthetic object"))?;
    object
        .messages
        .iter()
        .find(|message| message.type_ == type_)
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("missing synthetic message").into())
}

fn replace_message_payload(
    package: &[u8],
    identifier: u64,
    payload: Vec<u8>,
) -> TestResult<Vec<u8>> {
    replace_typed_message_payload(package, identifier, CHART_NON_STYLE_MESSAGE_TYPE, payload)
}

fn replace_typed_message_payload(
    package: &[u8],
    identifier: u64,
    type_: u32,
    payload: Vec<u8>,
) -> TestResult<Vec<u8>> {
    let mut stream = document_stream(package)?;
    let mut archive = Archive::parse(&stream)?;
    let object = archive
        .object_mut(identifier)
        .ok_or_else(|| io::Error::other("missing synthetic object"))?;
    let message = object
        .messages
        .iter_mut()
        .find(|message| message.type_ == type_)
        .ok_or_else(|| io::Error::other("missing synthetic typed message"))?;
    message.data = payload;
    stream = archive.to_bytes()?;
    let compressed = SnappyStream::compress(&stream)?;
    Ok(Catalog::from_bytes(package)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            DOCUMENT_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?)
}

fn same_central_record_except_offset(source: &[u8], candidate: &[u8]) -> bool {
    const LOCAL_HEADER_OFFSET: std::ops::Range<usize> = 42..46;
    source.len() == candidate.len()
        && source.len() >= LOCAL_HEADER_OFFSET.end
        && source[..LOCAL_HEADER_OFFSET.start] == candidate[..LOCAL_HEADER_OFFSET.start]
        && source[LOCAL_HEADER_OFFSET.end..] == candidate[LOCAL_HEADER_OFFSET.end..]
}

fn assert_only_document_payload_changed(before: &[u8], after: &[u8]) -> TestResult<()> {
    let before_catalog = Catalog::from_bytes(before)?;
    let after_catalog = Catalog::from_bytes(after)?;
    let mut changed = 0;
    for before_entry in before_catalog.iter() {
        if ROOT_PREVIEWS.contains(&before_entry.name()) {
            assert!(
                after_catalog
                    .iter()
                    .all(|entry| entry.name() != before_entry.name()),
                "root preview survived changed legend: {}",
                before_entry.name()
            );
            continue;
        }
        let after_entry = after_catalog
            .iter()
            .find(|entry| entry.name() == before_entry.name())
            .ok_or_else(|| io::Error::other("candidate removed an unselected member"))?;
        if before_entry.name() == DOCUMENT_MEMBER {
            assert_ne!(before_entry.data(), after_entry.data());
            changed += 1;
        } else {
            assert_eq!(before_entry.raw_name(), after_entry.raw_name());
            assert_eq!(before_entry.is_opaque(), after_entry.is_opaque());
            assert_eq!(before_entry.data(), after_entry.data());
            assert_eq!(before_entry.metadata(), after_entry.metadata());
            assert_eq!(
                before_entry.raw_record().local_record(),
                after_entry.raw_record().local_record()
            );
            assert_eq!(
                before_entry.raw_record().compressed_data(),
                after_entry.raw_record().compressed_data()
            );
            assert!(same_central_record_except_offset(
                before_entry.raw_record().central_directory_record(),
                after_entry.raw_record().central_directory_record()
            ));
        }
    }
    for after_entry in after_catalog.iter() {
        assert!(
            !ROOT_PREVIEWS.contains(&after_entry.name())
                && before_catalog
                    .iter()
                    .any(|entry| entry.name() == after_entry.name()),
            "candidate added an unexpected member: {}",
            after_entry.name()
        );
    }
    assert_eq!(changed, 1);
    Ok(())
}

fn assert_root_previews_absent(package: &[u8]) -> TestResult<()> {
    let catalog = Catalog::from_bytes(package)?;
    for preview in ROOT_PREVIEWS {
        assert!(catalog.iter().all(|entry| entry.name() != preview));
    }
    Ok(())
}

fn assert_semantic_handles_are_send_sync<T: Send + Sync>() {}

#[test]
fn legend_visibility_uses_position_and_name_selectors_and_absent_defaults_false() -> TestResult<()>
{
    let source = synthetic_package()?;
    let package = Package::from_bytes(&source)?;

    assert_eq!(package.slide_chart_catalog("Charts")?.len(), 2);
    assert!(!package.slide_chart_legend_visible(0usize, 0usize)?);
    assert!(package.slide_chart_legend_visible("Charts", "Costs")?);
    assert!(!package.slide_chart_legend_visible(
        SlideSelector::position(Position::new(0)),
        ChartSelector::name("Revenue"),
    )?);

    let edit = package.edit_slide_chart_legend("Charts", "Revenue")?;
    assert_eq!(edit.slide_position(), Position::new(0));
    assert_eq!(edit.chart_position(), Position::new(0));
    assert!(!edit.before());
    assert!(!edit.after());

    // Semantic handles must not accidentally expose native graph identity.
    let debug = format!("{edit:?}");
    for identifier in CHARTS.into_iter().chain(TITLES).chain(NON_STYLES) {
        assert!(
            !debug.contains(&identifier.to_string()),
            "native identifier leaked from semantic legend edit debug output: {debug}"
        );
    }
    Ok(())
}

#[test]
fn legend_noop_is_byte_exact_and_reapplicable() -> TestResult<()> {
    let source = synthetic_package()?;
    let package = Package::from_bytes(&source)?;
    let noop = package
        .edit_slide_chart_legend(0usize, 0usize)?
        .set(false)
        .commit()?;

    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(exact_bytes(noop.package())?, source);

    let reapplied = package.apply_slide_chart_legend(noop.patch())?;
    assert!(reapplied.patch().is_noop());
    assert!(!reapplied.diagnostics().changed());
    assert_eq!(exact_bytes(reapplied.package())?, source);
    Ok(())
}

#[test]
fn legend_transitions_reopen_and_inverse_to_the_exact_source() -> TestResult<()> {
    let source = synthetic_package()?;
    let package = Package::from_bytes(&source)?;

    let enabled = package
        .edit_slide_chart_legend("Charts", "Revenue")?
        .set(true)
        .commit()?;
    assert!(
        enabled
            .package()
            .slide_chart_legend_visible(0usize, 0usize)?
    );
    assert!(!enabled.patch().is_noop());
    assert!(!enabled.patch().before());
    assert!(enabled.patch().after());
    assert!(enabled.diagnostics().changed());
    assert!(enabled.diagnostics().full_reparse_performed());
    assert!(enabled.diagnostics().touched_components() >= 1);
    assert_eq!(
        enabled.diagnostics().deleted_previews(),
        ROOT_PREVIEWS.len()
    );
    let enabled_bytes = exact_bytes(enabled.package())?;
    assert_root_previews_absent(&enabled_bytes)?;
    assert_only_document_payload_changed(&source, &enabled_bytes)?;

    let reopened = Package::from_bytes(&exact_bytes(enabled.package())?)?;
    assert!(reopened.slide_chart_legend_visible("Charts", "Revenue")?);

    let disabled = enabled
        .package()
        .edit_slide_chart_legend(0usize, 0usize)?
        .set(false)
        .commit()?;
    assert!(
        !disabled
            .package()
            .slide_chart_legend_visible(0usize, 0usize)?
    );
    assert!(disabled.patch().before());
    assert!(!disabled.patch().after());
    assert_eq!(disabled.diagnostics().deleted_previews(), 0);

    let restored = enabled
        .package()
        .apply_slide_chart_legend(&enabled.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert!(
        !restored
            .package()
            .slide_chart_legend_visible(0usize, 0usize)?
    );
    Ok(())
}

#[test]
fn legend_unknown_fields_and_other_chart_payloads_remain_local() -> TestResult<()> {
    let source = synthetic_package_with_states(
        [(Some(false), Some("Revenue")), (Some(true), Some("Costs"))],
        true,
    )?;
    let selected_before = message_payload(&source, NON_STYLES[0], CHART_NON_STYLE_MESSAGE_TYPE)?;
    let other_before = message_payload(&source, NON_STYLES[1], CHART_NON_STYLE_MESSAGE_TYPE)?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_slide_chart_legend("Charts", "Revenue")?
        .set(true)
        .commit()?;
    let target = exact_bytes(commit.package())?;
    let selected_after = message_payload(&target, NON_STYLES[0], CHART_NON_STYLE_MESSAGE_TYPE)?;
    let other_after = message_payload(&target, NON_STYLES[1], CHART_NON_STYLE_MESSAGE_TYPE)?;

    assert_ne!(selected_before, selected_after);
    assert_eq!(
        selected_after,
        non_style_payload(Some(true), Some("Revenue"), true)?
    );
    assert_eq!(other_before, other_after);
    assert_eq!(commit.diagnostics().deleted_previews(), ROOT_PREVIEWS.len());
    assert_root_previews_absent(&target)?;
    assert_only_document_payload_changed(&source, &target)?;
    assert!(commit.diagnostics().touched_components() >= 1);

    let patch_debug = format!("{:?}", commit.patch());
    for identifier in CHARTS.into_iter().chain(TITLES).chain(NON_STYLES) {
        assert!(
            !patch_debug.contains(&identifier.to_string()),
            "native identifier leaked from semantic legend patch debug output: {patch_debug}"
        );
    }
    Ok(())
}

#[test]
fn legend_patches_are_source_checked_and_do_not_mutate_conflicts() -> TestResult<()> {
    let source = synthetic_package()?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_slide_chart_legend(0usize, 0usize)?
        .set(true)
        .commit()?;
    let stale_patch = commit.patch().clone();

    let independently_changed = package
        .edit_slide_chart_legend(0usize, 0usize)?
        .set(true)
        .commit()?;
    let before = exact_bytes(independently_changed.package())?;
    let error = independently_changed
        .package()
        .apply_slide_chart_legend(&stale_patch)
        .expect_err("a patch from another exact artifact must conflict");
    assert!(matches!(error, ChartLegendVisibilityError::PatchConflict));
    assert_eq!(exact_bytes(independently_changed.package())?, before);
    Ok(())
}

#[test]
fn legend_selectors_fail_closed_for_missing_positions_names_and_ambiguity() -> TestResult<()> {
    let source = synthetic_package()?;
    let package = Package::from_bytes(&source)?;
    assert!(package.slide_chart_legend_visible(0usize, 3usize).is_err());
    assert!(
        package
            .slide_chart_legend_visible(0usize, ChartSelector::name("Missing"))
            .is_err()
    );
    assert!(
        package
            .slide_chart_legend_visible("missing", 0usize)
            .is_err()
    );
    assert!(
        package
            .edit_slide_chart_legend(0usize, ChartSelector::name(""))
            .is_err()
    );

    let duplicate = Package::from_bytes(&synthetic_package_with_states(
        [
            (Some(true), Some("Revenue")),
            (Some(false), Some("Revenue")),
        ],
        false,
    )?)?;
    assert!(
        duplicate
            .slide_chart_legend_visible(0usize, ChartSelector::name("Revenue"))
            .is_err()
    );
    let duplicate_source = exact_bytes(&duplicate)?;
    assert!(
        duplicate
            .edit_slide_chart_legend(0usize, ChartSelector::name("Revenue"))
            .is_err()
    );
    assert_eq!(exact_bytes(&duplicate)?, duplicate_source);
    Ok(())
}

fn malformed_payload(kind: &str) -> TestResult<Vec<u8>> {
    let mut extension = Vec::new();
    append_varint_field(&mut extension, 4_000, 7)?;
    match kind {
        "duplicate" => {
            append_varint_field(&mut extension, LEGEND_VISIBLE_FIELD, 1)?;
            append_varint_field(&mut extension, LEGEND_VISIBLE_FIELD, 0)?;
        },
        "wrong-wire" => {
            append_length_delimited_field(&mut extension, LEGEND_VISIBLE_FIELD, b"wrong wire")?;
        },
        "noncanonical" => {
            // field 20, wire type 0, followed by a non-canonical zero varint.
            extension.extend_from_slice(&[0xa0, 0x01, 0x80, 0x00]);
        },
        other => return Err(io::Error::other(format!("unknown malformed case: {other}")).into()),
    }
    append_length_delimited_field(&mut extension, 23, b"Revenue")?;
    let mut payload = Vec::new();
    append_length_delimited_field(&mut payload, CHART_EXTENSION_FIELD, &extension)?;
    Ok(payload)
}

fn malformed_outer_non_style_payload(kind: &str) -> TestResult<Vec<u8>> {
    let extension = non_style_extension(Some(false), Some("Revenue"), false)?;
    let mut payload = Vec::new();
    match kind {
        "duplicate" => {
            append_length_delimited_field(&mut payload, CHART_EXTENSION_FIELD, &extension)?;
            append_length_delimited_field(&mut payload, CHART_EXTENSION_FIELD, &extension)?;
        },
        "wrong-wire" => append_varint_field(&mut payload, CHART_EXTENSION_FIELD, 1)?,
        other => {
            return Err(io::Error::other(format!("unknown malformed outer case: {other}")).into());
        },
    }
    Ok(payload)
}

#[test]
fn legend_wire_validation_rejects_duplicate_wrong_wire_and_noncanonical_fields_atomically()
-> TestResult<()> {
    let source = synthetic_package()?;
    for kind in ["duplicate", "wrong-wire", "noncanonical"] {
        let malformed = replace_message_payload(&source, NON_STYLES[0], malformed_payload(kind)?)?;
        let package = Package::from_bytes(&malformed)?;
        let before = exact_bytes(&package)?;
        assert!(
            package.slide_chart_legend_visible(0usize, 0usize).is_err(),
            "malformed {kind} legend field unexpectedly decoded"
        );
        assert!(
            package.edit_slide_chart_legend(0usize, 0usize).is_err(),
            "malformed {kind} legend field unexpectedly opened for mutation"
        );
        assert_eq!(exact_bytes(&package)?, before);
    }
    Ok(())
}

#[test]
fn legend_outer_extension_validation_rejects_duplicate_and_wrong_wire_atomically() -> TestResult<()>
{
    let source = synthetic_package()?;
    for kind in ["duplicate", "wrong-wire"] {
        let malformed = replace_message_payload(
            &source,
            NON_STYLES[0],
            malformed_outer_non_style_payload(kind)?,
        )?;
        let package = Package::from_bytes(&malformed)?;
        let before = exact_bytes(&package)?;
        assert!(matches!(
            package.slide_chart_legend_visible(0usize, 0usize),
            Err(ChartLegendVisibilityError::InvalidSource)
        ));
        assert!(matches!(
            package.edit_slide_chart_legend(0usize, 0usize),
            Err(ChartLegendVisibilityError::InvalidSource)
        ));
        assert_eq!(exact_bytes(&package)?, before);
    }
    Ok(())
}

#[test]
fn legend_rejects_shared_non_style_ownership_atomically() -> TestResult<()> {
    let source = synthetic_package()?;
    let malformed = replace_typed_message_payload(
        &source,
        CHARTS[1],
        CHART_MESSAGE_TYPE,
        chart_payload_with_non_style(1, NON_STYLES[0])?,
    )?;
    let package = Package::from_bytes(&malformed)?;
    let before = exact_bytes(&package)?;

    for chart_position in 0..CHARTS.len() {
        assert!(matches!(
            package.slide_chart_legend_visible(0usize, chart_position),
            Err(ChartLegendVisibilityError::InvalidSource)
        ));
        assert!(matches!(
            package.edit_slide_chart_legend(0usize, chart_position),
            Err(ChartLegendVisibilityError::InvalidSource)
        ));
    }
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn legend_semantic_public_types_are_send_sync() {
    assert_semantic_handles_are_send_sync::<ChartLegendVisibilityPatch>();
    assert_semantic_handles_are_send_sync::<ChartLegendVisibilityDiagnostics>();
    assert_semantic_handles_are_send_sync::<ChartLegendVisibilityLimitKind>();
    assert_eq!(
        ChartLegendVisibilityLimitKind::Allocations.to_string(),
        "allocations"
    );
}
