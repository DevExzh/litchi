use std::io;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::wire::{append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsch, tsd, tsk, tsp};
use litchi_keynote::{ChartSelector, ChartTitleError, Package, Position, SlideSelector};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const SLIDE_NODE: u64 = 3;
const SLIDE: u64 = 4;
const CHARTS: [u64; 2] = [100, 101];
const TITLES: [u64; 2] = [110, 111];
const NON_STYLES: [u64; 2] = [120, 121];
const CHART_MESSAGE_TYPE: u32 = 5_021;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
const CHART_NON_STYLE_MESSAGE_TYPE: u32 = 5_023;

type TestResult<T> = Result<T, Box<dyn std::error::Error>>;

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

fn chart_payload(chart: usize) -> TestResult<Vec<u8>> {
    chart_payload_with_graph(TITLES[chart], NON_STYLES[chart])
}

fn chart_payload_with_title(chart: usize, title_identifier: u64) -> TestResult<Vec<u8>> {
    chart_payload_with_graph(title_identifier, NON_STYLES[chart])
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
    append_length_delimited_field(&mut payload, 10_000, &chart_data.encode_to_vec())?;
    Ok(payload)
}

fn non_style_payload_state(
    title_visible: Option<bool>,
    title: Option<&str>,
) -> TestResult<Vec<u8>> {
    let mut extension = Vec::new();
    if let Some(title_visible) = title_visible {
        append_varint_field(&mut extension, 21, u64::from(title_visible))?;
    }
    if let Some(title) = title {
        append_length_delimited_field(&mut extension, 23, title.as_bytes())?;
    }
    let mut payload = Vec::new();
    append_length_delimited_field(&mut payload, 10_000, &extension)?;
    Ok(payload)
}

fn non_style_payload_state_with_unknown_title_fields(
    title_visible: Option<bool>,
    title: Option<&str>,
) -> TestResult<Vec<u8>> {
    let mut extension = Vec::new();
    append_varint_field(&mut extension, 4_000, 7)?;
    if let Some(title_visible) = title_visible {
        append_varint_field(&mut extension, 21, u64::from(title_visible))?;
    }
    append_length_delimited_field(&mut extension, 4_001, b"opaque title metadata")?;
    if let Some(title) = title {
        append_length_delimited_field(&mut extension, 23, title.as_bytes())?;
    }
    append_varint_field(&mut extension, 4_002, 9)?;

    let mut payload = Vec::new();
    append_length_delimited_field(&mut payload, 10_000, &extension)?;
    Ok(payload)
}

fn component(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

fn synthetic_package() -> TestResult<Vec<u8>> {
    synthetic_package_with_states([(Some(true), Some("Revenue")), (Some(true), Some("Costs"))])
}

fn synthetic_package_with_states(states: [(Option<bool>, Option<&str>); 2]) -> TestResult<Vec<u8>> {
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
        objects.push(object(
            NON_STYLES[chart],
            CHART_NON_STYLE_MESSAGE_TYPE,
            non_style_payload_state(states[chart].0, states[chart].1)?,
        )?);
    }
    let document_component = component(objects)?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Data/sentinel.bin", b"unrelated ZIP sentinel".as_slice()),
            (DOCUMENT_MEMBER, document_component.as_slice()),
        ],
        Limits::default(),
    )?)
}

fn synthetic_package_with_unknown_title_fields() -> TestResult<Vec<u8>> {
    let source = synthetic_package()?;
    let mut stream = document_stream(&source)?;
    let mut archive = Archive::parse(&stream)?;
    let object = archive
        .object_mut(NON_STYLES[0])
        .ok_or_else(|| io::Error::other("missing synthetic chart non-style object"))?;
    let message = object
        .messages
        .iter_mut()
        .find(|message| message.type_ == CHART_NON_STYLE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("missing synthetic chart non-style message"))?;
    message.data = non_style_payload_state_with_unknown_title_fields(Some(true), Some("Revenue"))?;
    stream = archive.to_bytes()?;
    let compressed = SnappyStream::compress(&stream)?;
    Ok(Catalog::from_bytes(&source)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            DOCUMENT_MEMBER,
            &compressed,
        )],
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

fn assert_only_document_payload_changed(before: &[u8], after: &[u8]) -> TestResult<()> {
    let before_catalog = Catalog::from_bytes(before)?;
    let after_catalog = Catalog::from_bytes(after)?;
    let mut changed = 0;
    for (before_entry, after_entry) in before_catalog.iter().zip(after_catalog.iter()) {
        assert_eq!(before_entry.name(), after_entry.name());
        if before_entry.name() == DOCUMENT_MEMBER {
            assert_ne!(before_entry.data(), after_entry.data());
            changed += 1;
        } else {
            assert_eq!(before_entry.data(), after_entry.data());
            assert_eq!(before_entry.metadata(), after_entry.metadata());
        }
    }
    assert_eq!(changed, 1);
    Ok(())
}

#[test]
fn selector_first_chart_title_transaction_is_exact_and_reversible() -> TestResult<()> {
    let bytes = synthetic_package()?;
    let package = Package::from_bytes(&bytes)?;
    assert_eq!(package.slide_chart_catalog("Charts")?.len(), 2);
    assert_eq!(
        package.slide_chart_title(SlideSelector::index(0), ChartSelector::name("Revenue"))?,
        Some("Revenue".to_owned())
    );
    assert_eq!(
        package.slide_chart_title("Charts", ChartSelector::index(1))?,
        Some("Costs".to_owned())
    );

    let source = exact_bytes(&package)?;
    let edit = package.edit_slide_chart_title("Charts", "Revenue")?;
    assert_eq!(edit.slide_position(), Position::new(0));
    assert_eq!(edit.chart_position(), Position::new(0));
    assert_eq!(edit.before(), Some("Revenue"));
    let commit = edit.set("Revenue by region")?.commit()?;
    assert_eq!(
        commit.package().slide_chart_title(0usize, 0usize)?,
        Some("Revenue by region".to_owned())
    );
    assert_eq!(commit.patch().before(), Some("Revenue"));
    assert_eq!(commit.patch().after(), Some("Revenue by region"));
    assert!(!commit.patch().is_noop());
    assert!(commit.diagnostics().changed());
    assert!(commit.diagnostics().full_reparse_performed());
    assert_only_document_payload_changed(&source, &exact_bytes(commit.package())?)?;
    assert_eq!(
        message_payload(&source, NON_STYLES[1], CHART_NON_STYLE_MESSAGE_TYPE)?,
        message_payload(
            &exact_bytes(commit.package())?,
            NON_STYLES[1],
            CHART_NON_STYLE_MESSAGE_TYPE
        )?
    );

    let applied = package.apply_slide_chart_title(commit.patch())?;
    assert_eq!(
        exact_bytes(applied.package())?,
        exact_bytes(commit.package())?
    );
    let inverse = commit.patch().inverse();
    assert_eq!(inverse.before(), commit.patch().after());
    assert_eq!(inverse.after(), commit.patch().before());
    assert_eq!(inverse.inverse(), commit.patch().clone());
    let restored = commit.package().apply_slide_chart_title(&inverse)?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_eq!(
        restored.package().slide_chart_title(0usize, 0usize)?,
        Some("Revenue".to_owned())
    );
    Ok(())
}

#[test]
fn chart_title_noop_and_clear_preserve_exact_source() -> TestResult<()> {
    let package = Package::from_bytes(&synthetic_package()?)?;
    let source = exact_bytes(&package)?;
    let noop = package
        .edit_slide_chart_title(0usize, 0usize)?
        .set("Revenue")?
        .commit()?;
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(exact_bytes(noop.package())?, source);

    let cleared = package
        .edit_slide_chart_title("Charts", "Revenue")?
        .clear()?
        .commit()?;
    assert_eq!(cleared.package().slide_chart_title(0usize, 0usize)?, None);
    let restored = cleared
        .package()
        .apply_slide_chart_title(&cleared.patch().inverse())?;
    assert_eq!(
        restored.package().slide_chart_title(0usize, 0usize)?,
        Some("Revenue".to_owned())
    );
    Ok(())
}

#[test]
fn chart_title_set_accepts_owned_text_without_changing_transaction_wire_behavior() -> TestResult<()>
{
    let package = Package::from_bytes(&synthetic_package()?)?;
    let title = String::from("Revenue by region");
    let committed = package
        .edit_slide_chart_title(0usize, 0usize)?
        .set(title)?
        .commit()?;

    assert_eq!(
        committed.package().slide_chart_title(0usize, 0usize)?,
        Some("Revenue by region".to_owned())
    );
    assert!(!committed.patch().is_noop());
    Ok(())
}

#[test]
fn chart_title_empty_visible_clear_inverse_restores_presence_and_bytes() -> TestResult<()> {
    let bytes = synthetic_package_with_states([(Some(true), None), (Some(true), Some("Costs"))])?;
    let package = Package::from_bytes(&bytes)?;
    let source = exact_bytes(&package)?;
    assert_eq!(
        package.slide_chart_title(0usize, 0usize)?,
        Some(String::new())
    );

    let cleared = package
        .edit_slide_chart_title(0usize, 0usize)?
        .clear()?
        .commit()?;
    assert_eq!(cleared.package().slide_chart_title(0usize, 0usize)?, None);
    assert_eq!(
        message_payload(
            &exact_bytes(cleared.package())?,
            NON_STYLES[0],
            CHART_NON_STYLE_MESSAGE_TYPE
        )?,
        non_style_payload_state(Some(false), None)?
    );

    let restored = cleared
        .package()
        .apply_slide_chart_title(&cleared.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_eq!(
        restored.package().slide_chart_title(0usize, 0usize)?,
        Some(String::new())
    );
    assert_eq!(
        message_payload(
            &exact_bytes(restored.package())?,
            NON_STYLES[0],
            CHART_NON_STYLE_MESSAGE_TYPE
        )?,
        non_style_payload_state(Some(true), None)?
    );
    Ok(())
}

#[test]
fn chart_title_explicit_empty_text_clear_inverse_restores_exact_presence() -> TestResult<()> {
    let bytes =
        synthetic_package_with_states([(Some(true), Some("")), (Some(true), Some("Costs"))])?;
    let package = Package::from_bytes(&bytes)?;
    let source = exact_bytes(&package)?;
    assert_eq!(
        package.slide_chart_title(0usize, 0usize)?,
        Some(String::new())
    );

    let cleared = package
        .edit_slide_chart_title(0usize, 0usize)?
        .clear()?
        .commit()?;
    assert_eq!(cleared.package().slide_chart_title(0usize, 0usize)?, None);
    assert_eq!(
        message_payload(
            &exact_bytes(cleared.package())?,
            NON_STYLES[0],
            CHART_NON_STYLE_MESSAGE_TYPE,
        )?,
        non_style_payload_state(Some(false), None)?,
    );

    let restored = cleared
        .package()
        .apply_slide_chart_title(&cleared.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_eq!(
        restored.package().slide_chart_title(0usize, 0usize)?,
        Some(String::new())
    );
    assert_eq!(
        message_payload(
            &exact_bytes(restored.package())?,
            NON_STYLES[0],
            CHART_NON_STYLE_MESSAGE_TYPE,
        )?,
        non_style_payload_state(Some(true), Some(""))?,
    );
    Ok(())
}

#[test]
fn chart_title_clear_hidden_or_absent_stale_text_is_exact_noop() -> TestResult<()> {
    for state in [(Some(false), Some("stale")), (None, Some("stale"))] {
        let bytes = synthetic_package_with_states([state, (Some(true), Some("Costs"))])?;
        let package = Package::from_bytes(&bytes)?;
        let source = exact_bytes(&package)?;
        assert_eq!(package.slide_chart_title(0usize, 0usize)?, None);

        let cleared = package
            .edit_slide_chart_title(0usize, 0usize)?
            .clear()?
            .commit()?;
        assert!(cleared.patch().is_noop());
        assert!(!cleared.diagnostics().changed());
        assert_eq!(exact_bytes(cleared.package())?, source);
        assert_eq!(
            message_payload(
                &exact_bytes(cleared.package())?,
                NON_STYLES[0],
                CHART_NON_STYLE_MESSAGE_TYPE
            )?,
            non_style_payload_state(state.0, state.1)?
        );
    }
    Ok(())
}

#[test]
fn selector_edit_preserves_unknown_generated_title_spans_and_other_owners() -> TestResult<()> {
    let package = Package::from_bytes(&synthetic_package_with_unknown_title_fields()?)?;
    let source_bytes = exact_bytes(&package)?;
    let source_other_owner =
        message_payload(&source_bytes, NON_STYLES[1], CHART_NON_STYLE_MESSAGE_TYPE)?;

    let committed = package
        .edit_slide_chart_title("Charts", "Revenue")?
        .set("Revenue by region")?
        .commit()?;
    assert_eq!(
        message_payload(
            &exact_bytes(committed.package())?,
            NON_STYLES[0],
            CHART_NON_STYLE_MESSAGE_TYPE,
        )?,
        non_style_payload_state_with_unknown_title_fields(Some(true), Some("Revenue by region"))?,
    );
    assert_eq!(
        message_payload(
            &exact_bytes(committed.package())?,
            NON_STYLES[1],
            CHART_NON_STYLE_MESSAGE_TYPE,
        )?,
        source_other_owner,
    );
    assert_eq!(
        committed
            .package()
            .slide_chart_title(0usize, ChartSelector::name("Revenue by region"))?,
        Some("Revenue by region".to_owned()),
    );
    Ok(())
}

#[test]
fn chart_title_selectors_and_graph_guards_fail_closed() -> TestResult<()> {
    let bytes = synthetic_package()?;
    let package = Package::from_bytes(&bytes)?;
    assert_eq!(
        package.slide_chart_title(0usize, ChartSelector::name("Revenue"))?,
        Some("Revenue".to_owned())
    );
    assert!(matches!(
        package.slide_chart_title(0usize, ChartSelector::name("revenue")),
        Err(ChartTitleError::ChartNameNotFound)
    ));
    assert!(matches!(
        package.slide_chart_title(0usize, ChartSelector::name("Missing")),
        Err(ChartTitleError::ChartNameNotFound)
    ));
    assert!(matches!(
        package.slide_chart_title("charts", 0usize),
        Err(ChartTitleError::SlideNameNotFound)
    ));
    assert!(matches!(
        package.edit_slide_chart_title("Missing", 0usize),
        Err(ChartTitleError::SlideNameNotFound)
    ));
    assert!(matches!(
        package.edit_slide_chart_title(0usize, 4usize),
        Err(ChartTitleError::ChartPositionNotFound { position }) if position == Position::new(4)
    ));
    assert!(matches!(
        package.edit_slide_chart_title(0usize, ChartSelector::name("")),
        Err(ChartTitleError::EmptyChartName)
    ));
    assert_eq!(exact_bytes(&package)?, bytes);

    let duplicate = Package::from_bytes(&synthetic_package_with_states([
        (Some(true), Some("Revenue")),
        (Some(true), Some("Revenue")),
    ])?)?;
    assert_eq!(
        duplicate.slide_chart_title(0usize, ChartSelector::index(0))?,
        Some("Revenue".to_owned())
    );
    assert_eq!(
        duplicate.slide_chart_title(0usize, ChartSelector::index(1))?,
        Some("Revenue".to_owned())
    );
    assert!(matches!(
        duplicate.slide_chart_title(0usize, ChartSelector::name("Revenue")),
        Err(ChartTitleError::AmbiguousSelector)
    ));
    let duplicate_source = exact_bytes(&duplicate)?;
    assert!(matches!(
        duplicate.edit_slide_chart_title(0usize, ChartSelector::name("Revenue")),
        Err(ChartTitleError::AmbiguousSelector)
    ));
    assert_eq!(exact_bytes(&duplicate)?, duplicate_source);

    let malformed = {
        let mut stream = document_stream(&bytes)?;
        let mut archive = Archive::parse(&stream)?;
        let object = archive
            .object_mut(CHARTS[0])
            .ok_or_else(|| io::Error::other("missing synthetic chart"))?;
        let message = object
            .messages
            .iter_mut()
            .find(|message| message.type_ == CHART_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing synthetic chart message"))?;
        let drawable = tsd::DrawableArchive {
            parent: Some(reference(999)),
            title: Some(reference(TITLES[0])),
            ..tsd::DrawableArchive::default()
        }
        .encode_to_vec();
        let mut malformed_chart = Vec::new();
        append_length_delimited_field(&mut malformed_chart, 1, &drawable)?;
        message.data = malformed_chart;
        stream = archive.to_bytes()?;
        let compressed = SnappyStream::compress(&stream)?;
        Catalog::from_bytes(&bytes)?.reassemble_to_bytes(
            &[litchi_iwa_archive::package::EntryEdit::new(
                DOCUMENT_MEMBER,
                &compressed,
            )],
            Limits::default(),
        )?
    };
    let malformed_package = Package::from_bytes(&malformed)?;
    assert!(matches!(
        malformed_package.slide_chart_title(0usize, 0usize),
        Err(ChartTitleError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&malformed_package)?, malformed);
    Ok(())
}

#[test]
fn chart_title_rejects_shared_title_standin_ownership() -> TestResult<()> {
    let source = synthetic_package()?;
    let mut stream = document_stream(&source)?;
    let mut archive = Archive::parse(&stream)?;
    let chart = archive
        .object_mut(CHARTS[1])
        .ok_or_else(|| io::Error::other("missing second synthetic chart"))?;
    let message = chart
        .messages
        .iter_mut()
        .find(|message| message.type_ == CHART_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("missing second synthetic chart message"))?;
    message.data = chart_payload_with_title(1, TITLES[0])?;
    stream = archive.to_bytes()?;
    let compressed = SnappyStream::compress(&stream)?;
    let malformed = Catalog::from_bytes(&source)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            DOCUMENT_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?;

    let package = Package::from_bytes(&malformed)?;
    let before = exact_bytes(&package)?;
    assert!(matches!(
        package.slide_chart_title(0usize, 0usize),
        Err(ChartTitleError::InvalidSource)
    ));
    assert!(matches!(
        package.edit_slide_chart_title(0usize, 0usize),
        Err(ChartTitleError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn chart_title_rejects_shared_non_style_ownership() -> TestResult<()> {
    let source = synthetic_package()?;
    let mut stream = document_stream(&source)?;
    let mut archive = Archive::parse(&stream)?;
    let chart = archive
        .object_mut(CHARTS[1])
        .ok_or_else(|| io::Error::other("missing second synthetic chart"))?;
    let message = chart
        .messages
        .iter_mut()
        .find(|message| message.type_ == CHART_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("missing second synthetic chart message"))?;
    message.data = chart_payload_with_graph(TITLES[1], NON_STYLES[0])?;
    stream = archive.to_bytes()?;
    let compressed = SnappyStream::compress(&stream)?;
    let malformed = Catalog::from_bytes(&source)?.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            DOCUMENT_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?;

    let package = Package::from_bytes(&malformed)?;
    let before = exact_bytes(&package)?;
    assert!(matches!(
        package.slide_chart_title(0usize, 0usize),
        Err(ChartTitleError::InvalidSource)
    ));
    assert!(matches!(
        package.edit_slide_chart_title(0usize, 0usize),
        Err(ChartTitleError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}
