use std::{path::PathBuf, sync::Arc};

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryEdit, EntryInsertion, ExactArtifacts},
};
use litchi_iwa_common::wire::{WireView, append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{
    Archive, ArchiveObject, FieldInfo, FieldPath, FieldType, RawMessage, SnappyStream,
};
use litchi_iwa_protos::tsp;
use prost::Message as _;

use super::*;

const CHART_MESSAGE_TYPE: u32 = 5_021;
const CHART_NON_STYLE_MESSAGE_TYPE: u32 = 5_023;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
const CHART_EXTENSION_FIELD: u32 = 10_000;
const CHART_MEDIATOR_FIELD: u32 = 8;
const STYLESHEET_STYLES_FIELD: u32 = 1;
const ROOT_PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

fn native_chart_title_package() -> (Vec<u8>, Package, ChartSelection) {
    let bytes = std::fs::read(
        PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/keynote/chart-titles-native.key"),
    )
    .expect("native chart-title fixture");
    let package = Package::from_bytes(&bytes).expect("native chart-title package");
    let selection = select_chart(
        &package,
        SlideSelector::index(0),
        ChartSelector::index(0),
        false,
    )
    .expect("native chart-title selection");
    (bytes, package, selection)
}

fn native_chart_title_package_with_stylesheet_registration() -> (Vec<u8>, Package, ChartSelection) {
    let (source, package, selection) = native_chart_title_package();
    let stylesheet_identifier = fixture_stylesheet_identifier(&package);
    let registered_source = reassemble_object(&source, stylesheet_identifier, |object| {
        let message_index = message_index(object, STYLESHEET_MESSAGE_TYPE);
        replace_message_data(object, message_index, |payload| {
            let reference = tsp::Reference {
                identifier: selection.non_style_identifier,
                ..tsp::Reference::default()
            }
            .encode_to_vec();
            let has_reference = WireView::parse(payload)
                .expect("stylesheet payload")
                .fields()
                .filter(|field| field.number() == STYLESHEET_STYLES_FIELD)
                .map(|field| field.payload())
                .filter_map(|payload| tsp::Reference::decode(payload).ok())
                .any(|candidate| candidate.identifier == selection.non_style_identifier);
            if !has_reference {
                append_length_delimited_field(payload, STYLESHEET_STYLES_FIELD, &reference)
                    .expect("stylesheet wire registration");
            }
        });
        let info = &mut object.archive_info.message_infos[message_index];
        if !info
            .object_references
            .contains(&selection.non_style_identifier)
        {
            info.object_references.push(selection.non_style_identifier);
        }
    });
    let package = Package::from_bytes(&registered_source).expect("registered chart package");
    let selection = select_chart(
        &package,
        SlideSelector::index(0),
        ChartSelector::index(0),
        false,
    )
    .expect("registered chart selection");
    (registered_source, package, selection)
}

// Add a source-built registry edge to an existing native stylesheet. Native
// non-style objects need not themselves carry a stylesheet parent reference.
fn fixture_stylesheet_identifier(package: &Package) -> u64 {
    package
        .state
        .source
        .components()
        .iter()
        .flat_map(|component| &component.archive().objects)
        .find(|object| {
            object.messages.len() == 1 && object.messages[0].type_ == STYLESHEET_MESSAGE_TYPE
        })
        .and_then(|object| object.archive_info.identifier)
        .expect("native stylesheet object for synthetic registration")
}

fn reassemble_object<F>(source: &[u8], identifier: u64, mutate: F) -> Vec<u8>
where
    F: FnOnce(&mut ArchiveObject),
{
    let catalog = Catalog::from_bytes(source).expect("source catalog");
    let mut component_name = None;
    let mut selected_archive = None;
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let Ok(stream) = SnappyStream::decompress(entry.data()) else {
            continue;
        };
        let Ok(archive) = Archive::parse(stream.as_bytes()) else {
            continue;
        };
        if archive.object(identifier).is_some() {
            assert!(
                selected_archive.is_none(),
                "object identifier is duplicated"
            );
            component_name = Some(entry.name().to_owned());
            selected_archive = Some(archive);
        }
    }
    let component_name = component_name.expect("selected object component");
    let mut archive = selected_archive.expect("selected archive");
    mutate(
        archive
            .object_mut(identifier)
            .expect("selected object in archive"),
    );
    let encoded = archive.to_bytes().expect("mutated archive");
    let compressed = SnappyStream::compress(&encoded).expect("mutated component");
    catalog
        .reassemble_to_bytes(
            &[EntryEdit::new(&component_name, compressed.as_slice())],
            Limits::default(),
        )
        .expect("reassembled object mutation")
}

fn replace_message_data<F>(object: &mut ArchiveObject, message_index: usize, mutate: F)
where
    F: FnOnce(&mut Vec<u8>),
{
    let message_type = object
        .messages
        .get(message_index)
        .map(|message| message.type_)
        .expect("message index");
    let mut data = object.messages[message_index].data.clone();
    mutate(&mut data);
    object
        .replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )
        .expect("replace message payload");
}

fn message_index(object: &ArchiveObject, message_type: u32) -> usize {
    object
        .messages
        .iter()
        .position(|message| message.type_ == message_type)
        .expect("message type")
}

fn replace_first_wire_field(payload: &[u8], number: u32, replacement: &[u8]) -> Vec<u8> {
    let view = WireView::parse(payload).expect("wire payload");
    let mut output = Vec::with_capacity(payload.len() + replacement.len());
    let mut replaced = false;
    for field in view.fields() {
        if !replaced && field.number() == number {
            output.extend_from_slice(replacement);
            replaced = true;
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    assert!(replaced, "wire field to replace");
    output
}

fn stylesheet_registration(package: &Package, target_identifier: u64) -> (u64, usize) {
    let mut selected = None;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            if object.messages.len() != 1 {
                continue;
            }
            let Some(object_identifier) = object.archive_info.identifier else {
                continue;
            };
            for (message_index, message) in object.messages.iter().enumerate() {
                if message.type_ != STYLESHEET_MESSAGE_TYPE {
                    continue;
                }
                let info = object
                    .archive_info
                    .message_infos
                    .get(message_index)
                    .expect("stylesheet message metadata");
                if info
                    .object_references
                    .iter()
                    .any(|identifier| *identifier == target_identifier)
                {
                    assert!(
                        selected
                            .replace((object_identifier, message_index))
                            .is_none(),
                        "native chart has one stylesheet registration"
                    );
                }
            }
        }
    }
    selected.expect("native chart stylesheet registration")
}

fn mutate_stylesheet_wire<F>(
    source: &[u8],
    package: &Package,
    target_identifier: u64,
    mutate: F,
) -> Vec<u8>
where
    F: FnOnce(&[u8]) -> Vec<u8>,
{
    let (identifier, message_index) = stylesheet_registration(package, target_identifier);
    reassemble_object(source, identifier, |object| {
        replace_message_data(object, message_index, |payload| {
            *payload = mutate(payload);
        });
    })
}

fn mutate_stylesheet_metadata<F>(
    source: &[u8],
    package: &Package,
    target_identifier: u64,
    mutate: F,
) -> Vec<u8>
where
    F: FnOnce(&mut litchi_iwa_core::MessageInfo),
{
    let (identifier, message_index) = stylesheet_registration(package, target_identifier);
    reassemble_object(source, identifier, |object| {
        let info = object
            .archive_info
            .message_infos
            .get_mut(message_index)
            .expect("stylesheet message metadata");
        mutate(info);
    })
}

fn assert_budgeted_rejects(source: &[u8]) {
    let package = Package::from_bytes(source).expect("mutated fixture package");
    let mut budget = ChartGraphScanBudget::new(&package).expect("graph budget");
    assert!(matches!(
        select_chart_with_budget(
            &package,
            SlideSelector::index(0),
            ChartSelector::index(0),
            true,
            &mut budget,
        ),
        Err(AxisSupportError::InvalidSource)
    ));
}

fn append_nonzero_mediator(payload: &mut Vec<u8>) {
    let outer = WireView::parse(payload).expect("chart payload");
    let chart = outer
        .fields()
        .find(|field| field.number() == CHART_EXTENSION_FIELD)
        .expect("chart extension");
    let reference = tsp::Reference {
        identifier: 9_999_999,
        ..tsp::Reference::default()
    }
    .encode_to_vec();
    let mut chart_payload = chart.payload().to_vec();
    let mut mediator = Vec::new();
    append_length_delimited_field(&mut mediator, CHART_MEDIATOR_FIELD, &reference)
        .expect("mediator field");
    if WireView::parse(&chart_payload)
        .expect("chart extension payload")
        .fields()
        .any(|field| field.number() == CHART_MEDIATOR_FIELD)
    {
        chart_payload = replace_first_wire_field(&chart_payload, CHART_MEDIATOR_FIELD, &mediator);
    } else {
        chart_payload.extend_from_slice(&mediator);
    }
    let mut replacement = Vec::new();
    append_length_delimited_field(&mut replacement, CHART_EXTENSION_FIELD, &chart_payload)
        .expect("chart extension replacement");
    *payload = replace_first_wire_field(payload, CHART_EXTENSION_FIELD, &replacement);
}

fn first_unrelated_object(source: &[u8], selection: &ChartSelection) -> (u64, usize) {
    let catalog = Catalog::from_bytes(source).expect("source catalog");
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let Ok(stream) = SnappyStream::decompress(entry.data()) else {
            continue;
        };
        let Ok(archive) = Archive::parse(stream.as_bytes()) else {
            continue;
        };
        for object in &archive.objects {
            let Some(identifier) = object.archive_info.identifier else {
                continue;
            };
            if [
                selection.slide_identifier,
                selection.chart_identifier,
                selection.non_style_identifier,
            ]
            .contains(&identifier)
            {
                continue;
            }
            if let Some((index, _)) = object.messages.iter().enumerate().find(|(_, message)| {
                message.type_ != CHART_MESSAGE_TYPE
                    && message.type_ != CHART_NON_STYLE_MESSAGE_TYPE
                    && message.type_ != STANDIN_MESSAGE_TYPE
            }) {
                return (identifier, index);
            }
        }
    }
    panic!("fixture has no unrelated object");
}

fn assert_public_rejects(source: &[u8]) {
    let package = Package::from_bytes(source).expect("mutated fixture package");
    assert!(matches!(
        package.slide_chart_title(SlideSelector::index(0), ChartSelector::index(0)),
        Err(ChartTitleError::InvalidSource)
    ));
    assert!(matches!(
        package.edit_slide_chart_title(SlideSelector::index(0), ChartSelector::index(0)),
        Err(ChartTitleError::InvalidSource)
    ));
}

#[test]
fn chart_title_rejects_mediator_and_merge_diff_metadata() {
    let (source, _package, selection) = native_chart_title_package();
    let mediator = reassemble_object(&source, selection.chart_identifier, |object| {
        let index = message_index(object, CHART_MESSAGE_TYPE);
        replace_message_data(object, index, append_nonzero_mediator);
    });
    assert_public_rejects(&mediator);

    for mutation in 0..6 {
        let adversarial =
            reassemble_object(
                &source,
                selection.non_style_identifier,
                |object| match mutation {
                    0 => object.archive_info.should_merge = Some(true),
                    1 => {
                        object.archive_info.message_infos[selection.non_style_message_index]
                            .base_message_index = Some(0)
                    },
                    2 => object.archive_info.message_infos[selection.non_style_message_index]
                        .diff_merge_version
                        .push(1),
                    3 => {
                        object.archive_info.message_infos[selection.non_style_message_index]
                            .diff_field_path = Some(FieldPath::new(vec![1]))
                    },
                    4 => object.archive_info.message_infos[selection.non_style_message_index]
                        .fields_to_remove
                        .push(FieldPath::new(vec![1])),
                    5 => object.archive_info.message_infos[selection.non_style_message_index]
                        .diff_read_version
                        .push(1),
                    _ => unreachable!(),
                },
            );
        assert_public_rejects(&adversarial);
    }
}

#[test]
fn chart_title_rejects_foreign_inbound_metadata_reference() {
    let (source, _package, selection) = native_chart_title_package();
    let (foreign_identifier, message_index) = first_unrelated_object(&source, &selection);
    let adversarial = reassemble_object(&source, foreign_identifier, |object| {
        object.archive_info.message_infos[message_index]
            .object_references
            .push(selection.non_style_identifier);
    });
    assert_public_rejects(&adversarial);

    let adversarial_package = Package::from_bytes(&adversarial).expect("foreign-ref package");
    let mut budget = ChartGraphScanBudget::new(&adversarial_package).expect("graph budget");
    assert!(matches!(
        select_chart_with_budget(
            &adversarial_package,
            SlideSelector::index(0),
            ChartSelector::index(0),
            true,
            &mut budget,
        ),
        Err(AxisSupportError::InvalidSource)
    ));
}

#[test]
fn chart_title_accepts_native_stylesheet_registration_on_both_paths() {
    let (_source, package, selection) = native_chart_title_package_with_stylesheet_registration();
    let (stylesheet_identifier, stylesheet_message_index) =
        stylesheet_registration(&package, selection.non_style_identifier);
    let stylesheet = package
        .object(stylesheet_identifier)
        .expect("stylesheet object");
    assert_eq!(stylesheet.messages.len(), 1);
    assert_eq!(
        stylesheet.messages[stylesheet_message_index].type_,
        STYLESHEET_MESSAGE_TYPE
    );
    let stylesheet_info = stylesheet
        .archive_info
        .message_infos
        .get(stylesheet_message_index)
        .expect("stylesheet metadata");
    assert_eq!(
        stylesheet_info
            .object_references
            .iter()
            .filter(|identifier| **identifier == selection.non_style_identifier)
            .count(),
        1
    );

    let selected = select_chart(
        &package,
        SlideSelector::index(0),
        ChartSelector::index(0),
        true,
    )
    .expect("ordinary stylesheet registration path");
    assert_eq!(
        selected.non_style_identifier,
        selection.non_style_identifier
    );

    let mut budget = ChartGraphScanBudget::new(&package).expect("graph budget");
    let selected_with_budget = select_chart_with_budget(
        &package,
        SlideSelector::index(0),
        ChartSelector::index(0),
        true,
        &mut budget,
    )
    .expect("budgeted stylesheet registration path");
    assert_eq!(
        selected_with_budget.non_style_identifier,
        selection.non_style_identifier
    );
}

#[test]
fn chart_title_rejects_malformed_stylesheet_registration_wire_membership() {
    let (source, package, selection) = native_chart_title_package_with_stylesheet_registration();
    for mutation in 0..3 {
        let candidate = mutate_stylesheet_wire(
            &source,
            &package,
            selection.non_style_identifier,
            |payload| {
                let view = WireView::parse(payload).expect("stylesheet wire");
                let mut output = Vec::new();
                let mut matched = false;
                for field in view.fields() {
                    let selected = field.number() == STYLESHEET_STYLES_FIELD
                        && tsp::Reference::decode(field.payload()).is_ok_and(|reference| {
                            reference.identifier == selection.non_style_identifier
                        });
                    if !selected {
                        output.extend_from_slice(field.raw());
                        continue;
                    }
                    assert!(!matched, "one selected registration");
                    matched = true;
                    match mutation {
                        0 => {},
                        1 => {
                            output.extend_from_slice(field.raw());
                            output.extend_from_slice(field.raw());
                        },
                        2 => append_varint_field(&mut output, STYLESHEET_STYLES_FIELD, 1)
                            .expect("malformed stylesheet field"),
                        _ => unreachable!(),
                    }
                }
                assert!(matched, "selected registration to mutate");
                output
            },
        );
        assert_public_rejects(&candidate);
        assert_budgeted_rejects(&candidate);
    }
}

#[test]
fn chart_title_rejects_duplicate_foreign_or_field_stylesheet_metadata() {
    let (source, package, selection) = native_chart_title_package_with_stylesheet_registration();
    for mutation in 0..3 {
        let candidate =
            mutate_stylesheet_metadata(&source, &package, selection.non_style_identifier, |info| {
                match mutation {
                    0 => info.object_references.push(selection.non_style_identifier),
                    1 => info.data_references.push(selection.non_style_identifier),
                    2 => {
                        let mut field = FieldInfo::new(vec![1]);
                        field.r#type = Some(FieldType::ObjectReference);
                        field.object_references.push(selection.non_style_identifier);
                        info.field_infos.push(field);
                    },
                    _ => unreachable!(),
                }
            });
        assert_public_rejects(&candidate);
        assert_budgeted_rejects(&candidate);
    }
}

#[test]
fn chart_title_accepts_one_expected_chart_field_metadata_edge() {
    let (source, package, selection) = native_chart_title_package();
    let chart = package
        .object(selection.chart_identifier)
        .expect("chart object");
    let chart_message_index = message_index(chart, CHART_MESSAGE_TYPE);
    let existing_field_edges = chart.archive_info.message_infos[chart_message_index]
        .field_infos
        .iter()
        .flat_map(|field| field.object_references.iter())
        .filter(|identifier| **identifier == selection.non_style_identifier)
        .count();
    assert!(existing_field_edges <= 1);

    let candidate = if existing_field_edges == 0 {
        reassemble_object(&source, selection.chart_identifier, |object| {
            let info = &mut object.archive_info.message_infos[chart_message_index];
            let mut field = FieldInfo::new(vec![1]);
            field.r#type = Some(FieldType::ObjectReference);
            field.object_references.push(selection.non_style_identifier);
            info.field_infos.push(field);
        })
    } else {
        source
    };
    let candidate_package = Package::from_bytes(&candidate).expect("chart-field package");
    select_chart(
        &candidate_package,
        SlideSelector::index(0),
        ChartSelector::index(0),
        true,
    )
    .expect("ordinary expected chart field edge");
    let mut budget = ChartGraphScanBudget::new(&candidate_package).expect("graph budget");
    select_chart_with_budget(
        &candidate_package,
        SlideSelector::index(0),
        ChartSelector::index(0),
        true,
        &mut budget,
    )
    .expect("budgeted expected chart field edge");
}

fn exact_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes).expect("package bytes");
    bytes
}

fn forge_patch(
    source_bytes: &[u8],
    patch: &ChartTitlePatch,
    target_bytes: Vec<u8>,
) -> ChartTitlePatch {
    let mut forged = patch.clone();
    forged.artifacts = ExactArtifacts::new(
        Arc::<[u8]>::from(source_bytes.to_vec()),
        Arc::<[u8]>::from(target_bytes),
    );
    forged
}

fn mutate_unrelated_member(source: &[u8]) -> Vec<u8> {
    let catalog = Catalog::from_bytes(source).expect("candidate catalog");
    let entry = catalog
        .iter()
        .find(|entry| {
            !entry.name().ends_with(".iwa")
                && !ROOT_PREVIEWS.contains(&entry.name())
                && !entry.is_opaque()
        })
        .expect("unrelated decoded member");
    let mut data = entry.data().to_vec();
    data.push(0x5a);
    catalog
        .reassemble_to_bytes(
            &[EntryEdit::new(entry.name(), data.as_slice())],
            Limits::default(),
        )
        .expect("unrelated member mutation")
}

fn mutate_unrelated_object(source: &[u8], selection: &ChartSelection) -> Vec<u8> {
    let (identifier, message_index) = first_unrelated_object(source, selection);
    reassemble_object(source, identifier, |object| {
        replace_message_data(object, message_index, |data| {
            append_varint_field(data, 4_777, 1).expect("unknown object field");
        });
    })
}

fn mutate_selected_payload(source: &[u8], selection: &ChartSelection) -> Vec<u8> {
    reassemble_object(source, selection.non_style_identifier, |object| {
        replace_message_data(object, selection.non_style_message_index, |data| {
            append_varint_field(data, 4_778, 1).expect("unknown selected field");
        });
    })
}

fn restore_one_preview(source: &[u8], preview_source: &[u8]) -> Vec<u8> {
    let source_catalog = Catalog::from_bytes(preview_source).expect("preview source catalog");
    let preview = source_catalog
        .iter()
        .find(|entry| entry.name() == "preview.jpg")
        .expect("native preview");
    Catalog::from_bytes(source)
        .expect("candidate catalog")
        .reassemble_with_insertions_to_bytes(
            &[EntryInsertion::new("preview.jpg", preview.data())],
            Limits::default(),
        )
        .expect("preview insertion")
}

#[test]
fn chart_title_candidate_verification_rejects_unrelated_drift_and_stale_preview() {
    let (source_bytes, source, selection) = native_chart_title_package();
    let changed = source
        .edit_slide_chart_title(SlideSelector::index(0), ChartSelector::index(0))
        .expect("chart title edit")
        .set("candidate verification title")
        .expect("chart title value")
        .commit()
        .expect("source candidate");
    let patch = changed.patch().clone();
    let target_bytes = exact_bytes(changed.package());
    let candidates = [
        mutate_unrelated_member(&target_bytes),
        mutate_unrelated_object(&target_bytes, &selection),
        mutate_selected_payload(&target_bytes, &selection),
        restore_one_preview(&target_bytes, &source_bytes),
    ];
    for candidate_bytes in candidates {
        let forged = forge_patch(&source_bytes, &patch, candidate_bytes);
        assert!(matches!(
            source.apply_slide_chart_title(&forged),
            Err(ChartTitleError::Verification)
        ));
    }
}

#[test]
fn chart_graph_owner_scan_admits_exact_work_and_rejects_one_under() {
    let (source, package, selection) = native_chart_title_package();
    let mut measured = ChartGraphScanBudget::new(&package).expect("graph budget");
    scan_chart_graph_owners_with_budget(&package, &mut measured).expect("graph owner scan");
    assert!(measured.work > 0);

    let mut exact = ChartGraphScanBudget::new(&package).expect("exact graph budget");
    exact.maximum_work = measured.work;
    scan_chart_graph_owners_with_budget(&package, &mut exact).expect("inclusive work ceiling");
    assert_eq!(exact.work, measured.work);

    let mut one_under = ChartGraphScanBudget::new(&package).expect("one-under graph budget");
    one_under.maximum_work = measured.work - 1;
    assert!(matches!(
        scan_chart_graph_owners_with_budget(&package, &mut one_under),
        Err(AxisSupportError::LimitExceeded {
            kind: chart_axis_support::AxisSupportLimitKind::WireWork,
            ..
        })
    ));

    let (unrelated_identifier, unrelated_message_index) =
        first_unrelated_object(&source, &selection);
    let with_empty_field_info = reassemble_object(&source, unrelated_identifier, |object| {
        object.archive_info.message_infos[unrelated_message_index]
            .field_infos
            .push(FieldInfo::new(vec![99]));
    });
    let augmented_package =
        Package::from_bytes(&with_empty_field_info).expect("empty field-info package");
    let mut augmented = ChartGraphScanBudget::new(&augmented_package).expect("graph budget");
    scan_chart_graph_owners_with_budget(&augmented_package, &mut augmented)
        .expect("empty field-info graph owner scan");
    assert!(augmented.work >= measured.work + 2);
}

#[test]
fn unknown_reference_metadata_is_preserved_for_reads_and_noops_but_blocks_title_publication() {
    let (source, package, selection) = native_chart_title_package();
    let (identifier, _) = first_unrelated_object(&source, &selection);
    let (component, _) = package
        .object_with_component(identifier)
        .expect("unrelated component");
    let catalog = Catalog::from_bytes(&source).expect("source catalog");
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == component)
        .expect("component entry");
    let stream = SnappyStream::decompress(entry.data()).expect("component stream");
    let archive = Archive::parse(stream.as_bytes()).expect("component archive");
    let object = archive.object(identifier).expect("unrelated object");
    let offset = usize::try_from(object.header_offset).expect("header offset");
    let (length, prefix) =
        decode_varint_from_bytes(&stream.as_bytes()[offset..]).expect("header prefix");
    let end = offset + prefix + usize::try_from(length).expect("header length");
    let mut header = stream.as_bytes()[offset + prefix..end].to_vec();
    append_varint_field(&mut header, 9_999, selection.non_style_identifier)
        .expect("opaque owner metadata");
    let mut rewritten = stream.as_bytes()[..offset].to_vec();
    encode_varint_into(
        &mut rewritten,
        u64::try_from(header.len()).expect("header length"),
    );
    rewritten.extend_from_slice(&header);
    rewritten.extend_from_slice(&stream.as_bytes()[end..]);
    let compressed = SnappyStream::compress(&rewritten).expect("changed component");
    let opaque = catalog
        .reassemble_to_bytes(&[EntryEdit::new(component, &compressed)], Limits::default())
        .expect("opaque package");
    let opaque_package = Package::from_bytes(&opaque).expect("opaque header is readable");
    assert_eq!(
        opaque_package
            .slide_chart_title(0usize, 0usize)
            .expect("title read"),
        selection.title
    );
    let noop = opaque_package
        .edit_slide_chart_title(0usize, 0usize)
        .expect("noop edit")
        .commit()
        .expect("exact noop");
    assert_eq!(exact_bytes(noop.package()), opaque);
    assert!(matches!(
        opaque_package
            .edit_slide_chart_title(0usize, 0usize)
            .expect("edit handle")
            .set("changed")
            .expect("title")
            .commit(),
        Err(ChartTitleError::InvalidSource)
    ));
    let changed = package
        .edit_slide_chart_title(0usize, 0usize)
        .expect("canonical edit")
        .set("changed")
        .expect("title")
        .commit()
        .expect("canonical commit");
    let forged = forge_patch(&opaque, changed.patch(), exact_bytes(changed.package()));
    assert!(matches!(
        opaque_package.apply_slide_chart_title(&forged),
        Err(ChartTitleError::InvalidSource)
    ));
    assert_eq!(exact_bytes(&opaque_package), opaque);
}
