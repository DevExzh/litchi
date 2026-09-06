use std::path::PathBuf;

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_core::{Archive, ArchiveObject, SnappyStream};

use super::*;

fn native_chart_caption_package() -> (Vec<u8>, Package, CaptionSelection) {
    let bytes = std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/keynote/chart-caption-native.key"),
    )
    .expect("native chart-caption fixture");
    let package = Package::from_bytes(&bytes).expect("native chart-caption package");
    let selection = select_caption(
        &package,
        SlideSelector::index(0),
        ChartSelector::index(0),
        false,
    )
    .expect("native chart-caption selection");
    (bytes, package, selection)
}

fn unlimited_budget(source: &Package) -> CaptionBudget {
    let mut budget = CaptionBudget::for_package(source).expect("caption budget");
    budget.maximum_input = usize::MAX;
    budget.maximum_output = usize::MAX;
    budget.maximum_fields = usize::MAX;
    budget.maximum_work = usize::MAX;
    budget.maximum_depth = u32::MAX;
    budget.maximum_components = usize::MAX;
    budget.maximum_references = usize::MAX;
    budget.maximum_allocations = usize::MAX;
    budget
}

fn immutable_style_cost(object: &ArchiveObject) -> (usize, usize) {
    let work = usize::try_from(object.header_length).expect("fixture header length")
        + object
            .messages
            .iter()
            .map(|message| message.data.len())
            .sum::<usize>();
    let references = object
        .archive_info
        .message_infos
        .iter()
        .map(|info| {
            info.object_references.len()
                + info.data_references.len()
                + info
                    .field_infos
                    .iter()
                    .map(|field| field.object_references.len() + field.data_references.len())
                    .sum::<usize>()
        })
        .sum();
    (work, references)
}

/// Replace one decoded IWA component, preserving every other ZIP member
/// exactly. `destination_component` moves the selected object after the
/// optional mutation, which gives the verifier a realistic component fence
/// regression without constructing a synthetic package graph.
fn reassemble_style_object<F>(
    source_bytes: &[u8],
    style_identifier: u64,
    destination_component: Option<&str>,
    mutate: F,
) -> Vec<u8>
where
    F: FnOnce(&mut ArchiveObject),
{
    let catalog = Catalog::from_bytes(source_bytes).expect("fixture catalog");
    let mut source_name = None;
    let mut source_archive = None;
    let mut destination_archive = None;

    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data()).expect("fixture IWA stream");
        let archive = Archive::parse(stream.as_bytes()).expect("fixture IWA archive");
        if archive.object(style_identifier).is_some() {
            source_name = Some(entry.name().to_owned());
            source_archive = Some(archive);
        } else if destination_component == Some(entry.name()) {
            destination_archive = Some(archive);
        }
    }

    let source_name = source_name.expect("external style component");
    let mut source_archive = source_archive.expect("external style archive");
    mutate(
        source_archive
            .object_mut(style_identifier)
            .expect("external style object"),
    );

    let mut destination_compressed = None;
    if let Some(destination_name) = destination_component {
        assert_ne!(destination_name, source_name.as_str());
        let mut destination_archive = destination_archive.expect("destination slide archive");
        let style = source_archive
            .remove_object(style_identifier)
            .expect("external style object to move");
        destination_archive.objects.push(style);
        let destination_bytes = destination_archive
            .to_bytes()
            .expect("destination IWA bytes");
        destination_compressed =
            Some(SnappyStream::compress(&destination_bytes).expect("destination Snappy bytes"));
    }

    let source_bytes = source_archive.to_bytes().expect("source IWA bytes");
    let source_compressed = SnappyStream::compress(&source_bytes).expect("source Snappy bytes");
    let mut edits = Vec::with_capacity(2);
    if let (Some(destination_name), Some(destination_compressed)) =
        (destination_component, destination_compressed.as_deref())
    {
        edits.push(EntryEdit::new(destination_name, destination_compressed));
    }
    edits.push(EntryEdit::new(&source_name, source_compressed.as_slice()));
    catalog
        .reassemble_to_bytes(&edits, Limits::default())
        .expect("reassembled fixture")
}

#[test]
fn immutable_external_style_accepts_exact_bytes_and_enforces_comparison_budget() {
    let (_bytes, source, selection) = native_chart_caption_package();
    let style_identifier = selection.style_identifier.expect("caption style");
    let (style_component, style_object) = source
        .object_with_component(style_identifier)
        .expect("caption style object");
    assert_ne!(style_component, selection.slide_component_name);
    let (comparison_work, comparison_references) = immutable_style_cost(style_object);
    assert!(comparison_work > 0);

    let mut exact = unlimited_budget(&source);
    exact.maximum_work = comparison_work * 2;
    exact.maximum_references = comparison_references * 2;
    verify_external_style_transition(
        &source,
        &source,
        &selection.slide_component_name,
        Some(style_identifier),
        Some(style_identifier),
        &mut exact,
    )
    .expect("unchanged external style");
    assert_eq!(exact.work, comparison_work * 2);
    assert_eq!(exact.references, comparison_references * 2);

    let mut one_under = unlimited_budget(&source);
    one_under.maximum_work = comparison_work * 2 - 1;
    one_under.maximum_references = usize::MAX;
    assert!(matches!(
        verify_external_style_transition(
            &source,
            &source,
            &selection.slide_component_name,
            Some(style_identifier),
            Some(style_identifier),
            &mut one_under,
        ),
        Err(ChartCaptionError::LimitExceeded {
            kind: ChartCaptionLimitKind::WireWork,
            ..
        })
    ));
}

#[test]
fn immutable_external_style_rejects_payload_or_header_mutation() {
    let (source_bytes, source, selection) = native_chart_caption_package();
    let style_identifier = selection.style_identifier.expect("caption style");
    let mutations: [fn(&mut ArchiveObject); 2] = [
        |style: &mut ArchiveObject| style.messages[0].data.push(0),
        |style: &mut ArchiveObject| {
            style.archive_info.message_infos[0]
                .object_references
                .push(0)
        },
    ];
    for mutate in mutations {
        let candidate_bytes =
            reassemble_style_object(&source_bytes, style_identifier, None, mutate);
        let candidate = Package::from_bytes(&candidate_bytes).expect("mutated candidate package");
        let mut budget = unlimited_budget(&source);
        assert!(matches!(
            verify_external_style_transition(
                &source,
                &candidate,
                &selection.slide_component_name,
                Some(style_identifier),
                Some(style_identifier),
                &mut budget,
            ),
            Err(ChartCaptionError::Verification)
        ));
    }
}

#[test]
fn immutable_external_style_rejects_component_migration() {
    let (source_bytes, source, selection) = native_chart_caption_package();
    let style_identifier = selection.style_identifier.expect("caption style");
    let candidate_bytes = reassemble_style_object(
        &source_bytes,
        style_identifier,
        Some(selection.slide_component_name.as_str()),
        |_| {},
    );
    let candidate = Package::from_bytes(&candidate_bytes).expect("migrated candidate package");
    let mut budget = unlimited_budget(&source);
    assert!(matches!(
        verify_external_style_transition(
            &source,
            &candidate,
            &selection.slide_component_name,
            Some(style_identifier),
            Some(style_identifier),
            &mut budget,
        ),
        Err(ChartCaptionError::Verification)
    ));
}
